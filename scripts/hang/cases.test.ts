import { describe, expect, test } from 'bun:test';
import { readZip } from '../../packages/ooxml-core/src/zip.ts';
import { readXlsx } from '../../packages/core/src/store/index.ts';
import {
  externalReferences,
  HANG_CASES,
  HOSTILE_SHEET,
  patchPart,
  randomSharedStrings,
  seeded,
  styleChain,
  zipRatioBomb,
} from './cases.ts';
import { NO_FIRST_CELL } from './first-cells.ts';

const dec = new TextDecoder();

/** The text of one part of the workbook of the case `id`. */
function partOf(id: string, part: string): string {
  const bytes = HANG_CASES.find((c) => c.id === id)?.bytes?.();
  if (!bytes) throw new Error(`no workbook for ${id}`);
  const read = readZip(bytes);
  if (!read.ok) throw new Error(read.reason);
  return dec.decode(read.entries.get(part));
}

const SHEET = '/xl/worksheets/sheet1.xml';

describe('the hostile workbooks of the hang test', () => {
  test('use the options of an opened file only on a case that opens a file', () => {
    const misplaced = HANG_CASES.filter(
      (c) => !c.bytes && (c.headingsOff || c.race || c.network === 'offline-after-start')
    );
    expect(misplaced.map((c) => c.id)).toEqual([]);
  });

  test('check the text of A1 on each case that opens a workbook, or say why not', () => {
    const opening = HANG_CASES.filter((c) => c.bytes && c.expect === 'open' && !c.race);
    const unchecked = opening.filter((c) => !c.firstCell && !NO_FIRST_CELL.has(c.id));
    expect(unchecked.map((c) => c.id)).toEqual([]);
    const both = opening.filter((c) => c.firstCell && NO_FIRST_CELL.has(c.id));
    expect(both.map((c) => c.id)).toEqual([]);
  });

  test('have unique ids', () => {
    const ids = HANG_CASES.map((c) => c.id);
    expect(new Set(ids).size).toBe(ids.length);
  });

  test('cover each hostile input, large table, and network fault of the test', () => {
    const ids = new Set(HANG_CASES.map((c) => c.id));
    for (const id of [
      'zip-leading-slash',
      'relationship-target-traversal',
      'sheet-names-object-keys',
      'active-content-inert',
      'dimension-whole-sheet',
      'sst-count-huge',
      'row-number-hostile',
      'cell-reference-hostile',
      'row-spans-huge',
      'col-range-huge',
      'merge-whole-sheet',
      'style-chain-cycle',
      'style-xml-deep',
      'sst-50mb',
      'fonts-stalled',
      'fonts-blocked',
      'network-blocked',
    ])
      expect(ids.has(id)).toBe(true);
  });

  test.each([
    ['dimension-whole-sheet', SHEET, 'ref="A1:XFD1048576"'],
    ['sst-count-huge', '/xl/sharedStrings.xml', 'uniqueCount="4294967295"'],
    ['row-number-hostile', SHEET, 'r="4294967295"'],
    ['cell-reference-hostile', SHEET, 'r="ZZZZZZ1"'],
    ['row-spans-huge', SHEET, 'spans="1:2147483647 5:4294967295 0:-1"'],
    ['col-range-huge', SHEET, 'max="2147483647"'],
    ['merge-whole-sheet', SHEET, 'ref="A1:XFD1048576"'],
  ])('%s holds its hostile value', (id, part, value) => {
    expect(partOf(id, part)).toContain(value);
  });

  test('name the first sheet so that the test can wait for it', () => {
    expect(partOf('dimension-whole-sheet', '/xl/workbook.xml')).toContain(
      `name="${HOSTILE_SHEET}"`
    );
  });

  test('build the same random text each time', () => {
    expect(seeded(1)()).toBe(seeded(1)());
    expect(randomSharedStrings(2000)).toEqual(randomSharedStrings(2000));
  });

  test('parse each modeled part and name the hostile sheet first in a workbook that opens', () => {
    // These parts stay as bytes on purpose: one passes the depth limit, and one declares a DTD.
    const staysBinary = new Set(['style-xml-deep', 'xml-entity-expansion']);
    const modeled = /\/(sharedStrings|styles|workbook|sheet\d+|theme\d*|drawing\d+)\.xml$/i;
    for (const c of HANG_CASES) {
      // A race case opens the workbook of `sheet-dense` again. Each workbook is released here.
      if (!c.bytes || c.expect !== 'open' || c.race) continue;
      const bytes = c.bytes();
      const zip = readZip(bytes);
      if (!zip.ok) throw new Error(`${c.id}: ${zip.reason}`);
      const workbook = dec.decode(zip.entries.get('/xl/workbook.xml'));
      expect(`${c.id}: ${/<sheet [^>]*name="([^"]*)"/.exec(workbook)?.[1]}`).toBe(
        `${c.id}: ${HOSTILE_SHEET}`
      );
      if (staysBinary.has(c.id)) continue;
      const read = readXlsx(bytes);
      if (!read.ok) throw new Error(`${c.id}: ${read.error.reason}`);
      const unread = [...read.pkg.parts.values()]
        .filter((p) => p.kind === 'binary' && modeled.test(p.partName))
        .map((p) => p.partName);
      expect(`${c.id}: ${unread.join()}`).toBe(`${c.id}: `);
    }
  }, 120_000);

  test('build a style chain that never ends and that a cell enters', () => {
    const chain = styleChain(100).map((xf) => Number(/xfId="(\d+)"/.exec(xf)![1]));
    // Index 0 is the default format, so the item at list position `i` has the index `i + 1`.
    const seen = new Set<number>();
    let at = 1;
    while (!seen.has(at)) {
      seen.add(at);
      at = chain[at - 1]!;
    }
    expect(seen.size).toBe(100);
    expect(at).toBe(1);
    expect(partOf('style-chain-cycle', '/xl/styles.xml')).toContain('<cellXfs');
  });

  test('name remote content that resolves through relationships', () => {
    const read = readZip(externalReferences());
    if (!read.ok) throw new Error(read.reason);
    const text = (part: string) => dec.decode(read.entries.get(part));
    expect(text('/xl/drawings/drawing1.xml')).toContain('r:link="rIdRemote"');
    expect(text('/xl/drawings/_rels/drawing1.xml.rels')).toContain('Id="rIdRemote"');
    expect(text('/xl/externalLinks/_rels/externalLink1.xml.rels')).toContain(
      'TargetMode="External"'
    );
    expect(text('/xl/workbook.xml')).not.toContain('<externalReferences><externalReferences>');
  });

  test('build a ratio bomb that the zip reader refuses for its ratio', () => {
    const read = readZip(zipRatioBomb());
    expect(read.ok).toBe(false);
    if (!read.ok) expect(read.limit).toBe('zip.maxRatio');
  });

  test('refuse a patch of a part that the package does not have', () => {
    const bytes = HANG_CASES.find((c) => c.id === 'dimension-whole-sheet')!.bytes!();
    expect(() => patchPart(bytes, '/xl/none.xml', (x) => x)).toThrow('no part');
  });
});
