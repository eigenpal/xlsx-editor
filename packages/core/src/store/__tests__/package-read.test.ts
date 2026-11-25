import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { runInNewContext } from 'node:vm';
import { strToU8, unzipSync, zipSync } from 'fflate';
import {
  buildXlsx,
  c,
  row,
  type FixtureWorkbook,
} from '../../../../../scripts/fixtures/builder.ts';
import {
  binaryPartOf,
  canonicalPackageFingerprint,
  MAX_STORED_SHEET_NAME,
  readXlsx,
  relationshipsOf,
  xmlPartOf,
  type XlsxPackage,
} from '../package/package.ts';
import { writeXlsx } from '../package-write/write.ts';
import { sheetByName, workbookIndex } from '../workbook/workbook-index.ts';
import { cellKey, sheetIndex } from '../workbook/sheet-index.ts';
import { attr, child } from '../package/query.ts';
import { parseXmlPart, type Element } from '../package/tree.ts';
import { XMLParser } from 'fast-xml-parser';
import { workbookOracles } from '../package-write/save.ts';
import { NS } from '../package/namespaces.ts';
import { parseCellRef } from '../cell-ref.ts';
import { FIXTURES, fixtureFile } from '../../../../../scripts/fixtures/catalog.ts';

function open(wb: FixtureWorkbook): XlsxPackage {
  const read = readXlsx(buildXlsx(wb));
  if (!read.ok) throw new Error(JSON.stringify(read.error));
  return read.pkg;
}

const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(`${import.meta.dir}/../../../../../fixtures/${fixtureFile(name)}`));

function sheetPart(pkg: XlsxPackage, name: string) {
  const wb = workbookIndex(pkg)!;
  const part = xmlPartOf(pkg, sheetByName(wb, name)!.partName);
  if (!part) throw new Error(`no sheet ${name}`);
  return part;
}

describe('bounded package read', () => {
  test('a zip bomb is refused with a typed reason', () => {
    const big = new Uint8Array(64 * 1024 * 1024);
    const bomb = zipSync({ '[Content_Types].xml': big }, { level: 9 });
    const read = readXlsx(bomb);
    expect(read.ok).toBe(false);
    if (!read.ok) expect(read.error.reason).toBe('zip');
  });

  test('a relationship target that escapes the package is not resolved', () => {
    const pkg = open({
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 1)]),
          rels: [['rIdBad', 'http://example.com/rel', '../../../../etc/passwd']],
        },
      ],
    });
    const rels = relationshipsOf(pkg, '/xl/worksheets/sheet1.xml');
    expect(rels.find((r) => r.id === 'rIdBad')).toBeUndefined();
  });

  test('a package with no workbook part is refused', () => {
    const bytes = zipSync({
      '[Content_Types].xml': strToU8(
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'
      ),
    });
    const read = readXlsx(bytes);
    expect(read.ok ? null : read.error.reason).toBe('not-a-workbook');
  });

  // `x:sheet/@name` (`CT_Sheet`, ECMA-376 Part 1, 18.2.19) has no length limit in the standard.
  // The read takes a stored name of at most 32,767 UTF-16 code units, the most text of one cell.
  test('a stored sheet name of more than 32,767 code units refuses the workbook', () => {
    const named = (name: string) =>
      readXlsx(
        buildXlsx({
          sheets: [
            { name: 'First', rows: row(1, [c('A1', 1)]) },
            { name, rows: row(1, [c('A1', 2)]) },
          ],
        })
      );
    expect(MAX_STORED_SHEET_NAME).toBe(32_767);
    const longest = named('N'.repeat(32_767));
    expect(longest.ok).toBe(true);
    if (longest.ok) expect(workbookIndex(longest.pkg)?.sheets[1]?.name.length).toBe(32_767);
    const past = named('N'.repeat(32_768));
    expect(past.ok ? null : past.error).toEqual({
      reason: 'workbook',
      detail: 'sheet-name-too-long',
      limit: 'workbook.maxSheetNameLength',
    });
  });
});

describe('the bytes that the read accepts', () => {
  const bytes = buildXlsx({ sheets: [{ name: 'Data', rows: row(1, [c('A1', 'text')]) }] });
  const names = (read: ReturnType<typeof readXlsx>): string[] =>
    read.ok ? (workbookIndex(read.pkg)?.sheets.map((sheet) => sheet.name) ?? []) : [];

  test('an ArrayBuffer reads as the same bytes in a Uint8Array', () => {
    const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
    expect(names(readXlsx(buffer as ArrayBuffer))).toEqual(['Data']);
  });

  test('an ArrayBuffer and a Uint8Array of another realm read', () => {
    const buffer = runInNewContext(`new ArrayBuffer(${bytes.byteLength})`) as ArrayBuffer;
    new Uint8Array(buffer).set(bytes);
    expect(buffer instanceof ArrayBuffer).toBe(false);
    expect(names(readXlsx(buffer))).toEqual(['Data']);

    // The view starts after the first byte, so the read must keep its offset and length.
    const view = runInNewContext(
      `new Uint8Array(new ArrayBuffer(${bytes.byteLength + 2}), 1, ${bytes.byteLength})`
    ) as Uint8Array;
    view.set(bytes);
    expect(view instanceof Uint8Array).toBe(false);
    expect(names(readXlsx(view))).toEqual(['Data']);
  });

  test('a detached ArrayBuffer and a value of another type fail as a package', () => {
    const detached = new ArrayBuffer(8);
    structuredClone(detached, { transfer: [detached] });
    for (const source of [detached, 'book.xlsx', null, { [Symbol.toStringTag]: 'Uint8Array' }]) {
      const read = readXlsx(source as never);
      expect(read.ok).toBe(false);
      if (!read.ok) expect(read.error.reason).toBe('zip');
    }
  });
});

describe('canonical ordered tree per part', () => {
  test('known and unknown siblings load together in source order', () => {
    const pkg = readXlsx(fixture('preserved-parts'));
    if (!pkg.ok) throw new Error('read');
    const root = sheetPart(pkg.pkg, 'Kept').root;
    const names = root.children.flatMap((n) =>
      n.type === 'element' ? [`${n.local}:${n.kind}`] : []
    );
    expect(names).toContain('conditionalFormatting:generic');
    expect(names).toContain('extLst:generic');
    expect(names.indexOf('sheetData:sheetData')).toBeLessThan(
      names.indexOf('conditionalFormatting:generic')
    );
  });

  test('a misplaced known element demotes to generic', () => {
    const pkg = open({
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 1)]),
          ext: '<ext uri="{X}"><c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="Z9"/></ext>',
        },
      ],
    });
    const root = sheetPart(pkg, 'S').root;
    const ext = child(root, 'x:extLst')!;
    const misplaced = (ext.children[0] as typeof root).children[0] as typeof root;
    expect(misplaced.local).toBe('c');
    expect(misplaced.kind).toBe('generic');
    expect(sheetIndex(sheetPart(pkg, 'S')).cells.size).toBe(1);
  });
});

describe('sparse cell index', () => {
  test('a false dimension does not allocate', () => {
    const pkg = open({
      sheets: [
        {
          name: 'S',
          rows: [row(1, [c('A1', 1)]), row(1048576, [c('A1048576', 2), c('XFD1048576', 3)])].join(
            ''
          ),
        },
      ],
    });
    const index = sheetIndex(sheetPart(pkg, 'S'));
    expect(index.cells.size).toBe(3);
    expect(index.cells.get(cellKey(1048575, 16383))?.el.kind).toBe('cell');
  });

  test('other file-stated sizes do not allocate or bound a loop', () => {
    const start = performance.now();
    const pkg = open({
      strings: ['x'],
      sheets: [
        {
          name: 'S',
          cols: '<cols><col min="1" max="16384" width="9"/></cols>',
          rows: row(1, [c('A1', 1, { f: 'B1', t: 'n' })], 'spans="1:16384"').replace(
            '<f>B1</f>',
            '<f t="shared" ref="A1:XFD1048576" si="0">B1</f>'
          ),
          merges: ['A1:XFD1048576'],
        },
      ],
    });
    const index = sheetIndex(sheetPart(pkg, 'S'));
    expect(index.cells.size).toBe(1);
    expect(performance.now() - start).toBeLessThan(1000);
  });

  test('duplicate and out-of-order cells are recorded, never guessed', () => {
    const pkg = open({
      sheets: [{ name: 'S', rows: row(1, [c('B1', 1), c('A1', 2), c('B1', 3)]) }],
    });
    const index = sheetIndex(sheetPart(pkg, 'S'));
    expect(index.problems.map((p) => p.code)).toEqual(['cell-out-of-order', 'duplicate-cell']);
    expect(index.cells.size).toBe(2);
  });

  test('too many elements refuses the part with a typed limit reason', () => {
    const cells = Array.from({ length: 20 }, (_, i) => c(`A${i + 1}`, i));
    const sheetXml = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${row(1, cells)}</sheetData></worksheet>`;
    const result = parseXmlPart('/xl/worksheets/sheet1.xml', sheetXml, {
      maxBytes: 1 << 20,
      maxElements: 10,
    });
    expect(result.ok ? null : result.reason).toBe('too-many-elements');
  });
});

describe('untrusted values', () => {
  test('a sheet named __proto__ is an ordinary key', () => {
    const pkg = open({
      sheets: [
        { name: '__proto__', rows: '' },
        { name: 'constructor', rows: '' },
      ],
    });
    const index = workbookIndex(pkg)!;
    expect(sheetByName(index, '__proto__')?.index).toBe(0);
    expect(sheetByName(index, 'CONSTRUCTOR')?.index).toBe(1);
    expect(sheetByName(index, 'toString')).toBeUndefined();
    expect(({} as Record<string, unknown>).index).toBeUndefined();
  });

  test('external targets are recorded, never fetched, and active content survives byte for byte', () => {
    const original = globalThis.fetch;
    let fetched = 0;
    globalThis.fetch = (() => {
      fetched += 1;
      throw new Error('no fetch');
    }) as unknown as typeof fetch;
    try {
      const read = readXlsx(fixture('preserved-parts'));
      if (!read.ok) throw new Error('read');
      const rels = relationshipsOf(read.pkg, '/xl/worksheets/sheet1.xml');
      expect(rels.every((r) => r.external)).toBe(true);
      const saved = readXlsx(writeXlsx(read.pkg));
      if (!saved.ok) throw new Error('reread');
      expect(binaryPartOf(saved.pkg, '/xl/fixtureData.bin')).toEqual(
        binaryPartOf(read.pkg, '/xl/fixtureData.bin')!
      );
      expect(fetched).toBe(0);
    } finally {
      globalThis.fetch = original;
    }
  });
});

describe('save without an edit', () => {
  for (const name of ['basic-values', 'preserved-parts'])
    test(`${name} keeps its fingerprint through save and reopen`, () => {
      const read = readXlsx(fixture(name));
      if (!read.ok) throw new Error('read');
      expect(workbookOracles(read.pkg).digest).toBe(canonicalPackageFingerprint(read.pkg));
    });

  test('the workbook index reads sheets, date system, and calculation state', () => {
    const pkg = open({
      date1904: true,
      calcPr: 'calcId="1" fullCalcOnLoad="1"',
      sheets: [
        { name: 'A', rows: '' },
        { name: 'B', state: 'hidden', rows: '' },
      ],
    });
    const index = workbookIndex(pkg)!;
    expect(index.sheets.map((s) => [s.name, s.state, s.kind])).toEqual([
      ['A', 'visible', 'worksheet'],
      ['B', 'hidden', 'worksheet'],
    ]);
    expect(index.date1904).toBe(true);
    expect(index.fullCalcOnLoad).toBe(true);
    expect(attr(child(sheetPart(pkg, 'A').root, 'x:sheetFormatPr'), 'defaultRowHeight')).toBe('15');
  });
});

describe('hostile element names and wrappers', () => {
  test('a root named after an Object member is generic', () => {
    for (const name of ['constructor', 'toString', '__proto__', 'hasOwnProperty']) {
      const r = parseXmlPart('/p.xml', `<x:${name} xmlns:x="${NS.x}"/>`);
      expect(r.ok && r.part.root.kind).toBe('generic');
    }
  });

  test('a cell inside a markup-compatibility wrapper is not typed', () => {
    const r = parseXmlPart(
      '/s.xml',
      `<worksheet xmlns="${NS.x}" xmlns:mc="${NS.mc}"><sheetData><row r="1"><mc:AlternateContent><mc:Choice xmlns:x="${NS.x}" Requires="x"><c r="A1"/></mc:Choice></mc:AlternateContent></row></sheetData></worksheet>`
    );
    if (!r.ok) throw new Error('parse');
    const rowEl = child(r.part.root, 'x:sheetData')!.children[0] as typeof r.part.root;
    const wrapper = rowEl.children[0] as typeof r.part.root;
    const cell = (wrapper.children[0] as typeof r.part.root).children[0] as typeof r.part.root;
    expect(cell.local).toBe('c');
    expect(cell.kind).toBe('generic');
  });

  test('a repeated row is left out of the index and recorded', () => {
    const pkg = open({
      sheets: [{ name: 'S', rows: [row(3, [c('A3', 1)]), row(3, [c('B3', 2)])].join('') }],
    });
    const index = sheetIndex(sheetPart(pkg, 'S'));
    expect(index.problems.map((p) => p.code)).toEqual(['row-out-of-order']);
    expect(index.cells.size).toBe(1);
  });
});

describe('fixture catalog', () => {
  test('every fixture edit names a sheet and a cell, a range, or a band of its workbook', () => {
    for (const fixture of FIXTURES) {
      const pkg = open(fixture.workbook());
      const index = workbookIndex(pkg)!;
      for (const edit of fixture.edits) {
        // A sheet add names the new sheet, which the workbook does not have yet.
        const added = 'sheetOp' in edit && edit.sheetOp === 'add';
        expect(sheetByName(index, edit.sheet)?.kind).toBe(added ? undefined : 'worksheet');
        if ('sheetOp' in edit) continue;
        if ('structure' in edit) {
          expect(Number.isInteger(edit.start) && edit.start >= 1).toBe(true);
          expect(Number.isInteger(edit.count) && edit.count >= 1).toBe(true);
          continue;
        }
        if ('resize' in edit) {
          expect(Number.isInteger(edit.first) && edit.first >= 1).toBe(true);
          expect(edit.last >= edit.first).toBe(true);
          continue;
        }
        // A format edit or a copy can name a range; a value edit or a paste names one cell.
        const corners =
          'copy' in edit
            ? [...edit.copy.split(':'), edit.paste]
            : 'pasteText' in edit
              ? [edit.at]
              : 'decimalStep' in edit
                ? [...edit.range.split(':'), edit.active]
                : 'format' in edit
                  ? edit.ref.split(':')
                  : [edit.ref];
        for (const corner of corners) expect(parseCellRef(corner)).not.toBeNull();
      }
    }
  });
});

describe('part names compare without ASCII case', () => {
  function renamed(from: string, to: string): Uint8Array {
    const entries = unzipSync(fixture('basic-values'));
    entries[to] = entries[from]!;
    delete entries[from];
    return zipSync(entries);
  }

  test('package relationships under _RELS are found', () => {
    const read = readXlsx(renamed('_rels/.rels', '_RELS/.rels'));
    expect(read.ok).toBe(true);
  });

  test('a worksheet entry spelled with other case is the relationship target', () => {
    const read = readXlsx(renamed('xl/worksheets/sheet1.xml', 'xl/worksheets/Sheet1.xml'));
    if (!read.ok) throw new Error('read');
    const sheet = workbookIndex(read.pkg)!.sheets[0]!;
    expect([sheet.partName, sheet.kind]).toEqual(['/xl/worksheets/Sheet1.xml', 'worksheet']);
  });
});

describe('read oracle: the tree keeps the source content', () => {
  // The save oracles compare a tree with its reopened tree, so they cannot see content that the
  // reader itself drops. This check reads each modeled part a second way, with an independent
  // XML parser, and compares the elements and attributes that each side holds.
  type Counts = Map<string, number>;
  const bump = (m: Counts, k: string) => m.set(k, (m.get(k) ?? 0) + 1);
  const localOf = (name: string) => name.slice(name.indexOf(':') + 1);

  function sourceCounts(xml: string): Counts {
    const parser = new XMLParser({
      preserveOrder: true,
      ignoreAttributes: false,
      attributeNamePrefix: '',
      processEntities: false,
      ignoreDeclaration: true,
    });
    const counts: Counts = new Map();
    const walk = (nodes: Record<string, unknown>[]) => {
      for (const node of nodes) {
        for (const [key, value] of Object.entries(node)) {
          if (key === ':@' || key === '#text') continue;
          bump(counts, `<${localOf(key)}>`);
          const attrs = (node[':@'] ?? {}) as Record<string, string>;
          for (const [a, v] of Object.entries(attrs))
            if (a !== 'xmlns' && !a.startsWith('xmlns:')) bump(counts, `@${localOf(a)}=${v}`);
          walk(value as Record<string, unknown>[]);
        }
      }
    };
    walk(parser.parse(xml) as Record<string, unknown>[]);
    return counts;
  }

  function treeCounts(el: Element, counts: Counts = new Map()): Counts {
    bump(counts, `<${el.local}>`);
    for (const a of el.attrs) bump(counts, `@${a.local}=${a.value}`);
    for (const c of el.children) if (c.type === 'element') treeCounts(c, counts);
    return counts;
  }

  const decodeEntities = (s: string) =>
    s.replace(
      /&(lt|gt|amp|quot|apos);/g,
      (_, e: string) => ({ lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" })[e]!
    );

  test.each(FIXTURES.map((f) => f.name))(
    '%s: every element and attribute of the source is in the tree',
    (name) => {
      const bytes = fixture(name);
      const read = readXlsx(bytes);
      if (!read.ok) throw new Error('read');
      const entries = unzipSync(bytes);
      let checked = 0;
      for (const part of read.pkg.parts.values()) {
        if (part.kind !== 'xml') continue;
        const source = new TextDecoder().decode(entries[part.partName.slice(1)]!);
        const expected = new Map(
          [...sourceCounts(source)].map(([k, v]) => [decodeEntities(k), v] as const)
        );
        expect([part.partName, treeCounts(part.xml.root)]).toEqual([part.partName, expected]);
        checked += 1;
      }
      expect(checked).toBeGreaterThan(5);
    }
  );

  test('the check fails when the tree loses an attribute', () => {
    const xml = '<a xmlns="urn:example:a" x="1"><b y="2"/></a>';
    const parsed = parseXmlPart('/a.xml', xml);
    if (!parsed.ok) throw new Error('parse');
    const root = parsed.part.root;
    const b = root.children[0] as Element;
    const lossy = { ...root, children: [{ ...b, attrs: [] }] };
    expect(treeCounts(root)).toEqual(sourceCounts(xml));
    expect(treeCounts(lossy)).not.toEqual(sourceCounts(xml));
  });
});
