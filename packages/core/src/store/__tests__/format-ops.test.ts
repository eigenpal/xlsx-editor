// The format ops of task 7.4: `applyFormat`, `clearFormats`, and their undo. Each test reads the
// styles part and the cells back from the store, and from a saved and reopened copy where the
// save rules matter.

import { describe, expect, test } from 'bun:test';
import {
  buildXlsx,
  c,
  row,
  type FixtureWorkbook,
} from '../../../../../scripts/fixtures/builder.ts';
import {
  canonicalPackageFingerprint,
  partOf,
  readXlsx,
  xmlPartOf,
  type XlsxPackage,
} from '../package/package.ts';
import { attr, child, children, elements } from '../package/query.ts';
import { sheetByName, styleLists, workbookIndex } from '../workbook/workbook-index.ts';
import { cellKey, sheetIndex } from '../workbook/sheet-index.ts';
import { saveXlsx } from '../package-write/save.ts';
import { serializeXmlPart } from '../package-write/serialize.ts';
import { WorkbookDocumentStore } from '../store/document-store.ts';
import type { WorkbookOp } from '../store/ops.ts';
import { applyHistory } from './history-helpers.ts';
import type { FormatPatch } from '../store/format-patch.ts';
import { MAX_FORMAT_NEW_CELLS } from '../store/format-range.ts';
import { relationshipsOf } from '../package/package.ts';
import { IdSource, plainAttr, smlElement } from '../store/tree-edit.ts';

function open(wb: FixtureWorkbook): XlsxPackage {
  const read = readXlsx(buildXlsx(wb));
  if (!read.ok) throw new Error(JSON.stringify(read.error));
  return read.pkg;
}

const fmt = (ref: string, patch: FormatPatch, sheet = 'S'): WorkbookOp => ({
  op: 'applyFormat',
  sheet,
  ref,
  patch,
});

function cellAt(pkg: XlsxPackage, ref: string, sheet = 'S') {
  const index = workbookIndex(pkg)!;
  const part = xmlPartOf(pkg, sheetByName(index, sheet)!.partName)!;
  const m = /^([A-Z]+)(\d+)$/.exec(ref)!;
  let col = 0;
  for (const ch of m[1]!) col = col * 26 + ch.charCodeAt(0) - 64;
  return sheetIndex(part).cells.get(cellKey(Number(m[2]) - 1, col - 1))?.el;
}

const styles = (pkg: XlsxPackage) => styleLists(pkg, workbookIndex(pkg)!);
const stylesRoot = (pkg: XlsxPackage) => xmlPartOf(pkg, workbookIndex(pkg)!.stylesPart)!.root;

/** The cell format of a cell, and the records it names. */
function formatOf(pkg: XlsxPackage, ref: string) {
  const lists = styles(pkg);
  const xf = lists.cellXfs[Number(attr(cellAt(pkg, ref), 's') ?? '0')]!;
  return {
    xf,
    font: lists.fonts[Number(attr(xf, 'fontId') ?? '0')]!,
    fill: lists.fills[Number(attr(xf, 'fillId') ?? '0')]!,
    border: lists.borders[Number(attr(xf, 'borderId') ?? '0')]!,
  };
}

const locals = (el: Parameters<typeof elements>[0]) => elements(el).map((e) => e.local);

const base = (): FixtureWorkbook => ({
  sheets: [
    {
      name: 'S',
      rows: [
        row(1, [c('A1', 1), c('B1', 2, { s: 1 }), c('C1', 3, { s: 1 })]),
        row(2, [c('A2', 4)]),
      ].join(''),
    },
  ],
  styles: {
    fonts: ['<font><color rgb="FFFF0000"/><sz val="11"/><name val="Calibri"/></font>'],
    cellXfs: [
      '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"><alignment horizontal="center"/></xf>',
    ],
  },
});

describe('applyFormat', () => {
  test('a run counts the new cells of each op as the ops one by one count them', () => {
    const wb: FixtureWorkbook = { sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] };
    const ops = [fmt('A1:A16384', { bold: true }), fmt('A1:B16384', { bold: true })];
    const together = new WorkbookDocumentStore(open(wb));
    const apart = new WorkbookDocumentStore(open(wb));
    const one = together.transact(ops).ok;
    const two = apart.transact([ops[0]!]).ok && apart.transact([ops[1]!]).ok;
    expect(one).toBe(two);
    if (one)
      expect(canonicalPackageFingerprint(together.package)).toBe(
        canonicalPackageFingerprint(apart.package)
      );
  });

  test('a run of ops on one sheet applies in order, in one pass and one undo step', () => {
    const store = new WorkbookDocumentStore(open(base()));
    const before = canonicalPackageFingerprint(store.package);
    // Overlapping ops: the later op reads the cells that the earlier op wrote, also a new one.
    const r = store.transact([
      fmt('A1:B1', { bold: true }),
      fmt('B1', { italic: true }),
      fmt('D5', { numberFormat: { code: '0.000' } }),
      fmt('D5', { bold: true }),
    ]);
    expect(r.ok).toBe(true);
    const pkg = store.package;
    expect(locals(formatOf(pkg, 'A1').font)).toContain('b');
    expect(locals(formatOf(pkg, 'B1').font)).toEqual(expect.arrayContaining(['b', 'i']));
    expect(locals(formatOf(pkg, 'D5').font)).toContain('b');
    expect(attr(formatOf(pkg, 'D5').xf, 'numFmtId')).not.toBe('0');
    expect(store.undo().ok).toBe(true);
    expect(canonicalPackageFingerprint(store.package)).toBe(before);
  });

  test('bold keeps the other properties of each cell and never changes a shared record', () => {
    const store = new WorkbookDocumentStore(open(base()));
    const before = styles(store.package);
    const font1 = before.fonts[1];
    const xf1 = before.cellXfs[1];
    expect(store.transact([fmt('A1:C1', { bold: true })]).ok).toBe(true);
    const pkg = store.package;
    const b1 = formatOf(pkg, 'B1');
    expect(locals(b1.font)).toEqual(['b', 'sz', 'color', 'name']);
    expect(attr(child(b1.font, 'x:color'), 'rgb')).toBe('FFFF0000');
    expect(attr(child(b1.xf, 'x:alignment'), 'horizontal')).toBe('center');
    // B1 and C1 share a format before the edit, and one new record after it.
    expect(attr(cellAt(pkg, 'B1'), 's')).toBe(attr(cellAt(pkg, 'C1'), 's'));
    expect(locals(formatOf(pkg, 'A1').font)).toContain('b');
    // The records that other cells use are the same objects.
    expect(styles(pkg).fonts[1]).toBe(font1);
    expect(styles(pkg).cellXfs[1]).toBe(xf1);
    expect(attr(cellAt(pkg, 'A2'), 's')).toBeUndefined();
  });

  test('an identical record is reused, not appended again', () => {
    const store = new WorkbookDocumentStore(open(base()));
    store.transact([fmt('A1', { bold: true })]);
    const counts = (pkg: XlsxPackage) => [styles(pkg).fonts.length, styles(pkg).cellXfs.length];
    const after = counts(store.package);
    store.transact([fmt('A2', { bold: true })]);
    expect(counts(store.package)).toEqual(after);
    expect(attr(cellAt(store.package, 'A2'), 's')).toBe(attr(cellAt(store.package, 'A1'), 's'));
    // Setting the value a cell already has writes nothing.
    const revision = store.revision;
    expect(store.transact([fmt('A1', { bold: true })]).ok).toBe(true);
    expect(store.revision).toBe(revision);
    // Removing bold again finds the original default format.
    store.transact([fmt('A1', { bold: false })]);
    expect(attr(cellAt(store.package, 'A1'), 's')).toBeUndefined();
  });

  test('records are written in schema order, and the counts follow the content', () => {
    const store = new WorkbookDocumentStore(open(base()));
    const r = store.transact([
      fmt('A1:B2', {
        italic: true,
        underline: true,
        fontName: 'Courier New',
        fontSize: 14,
        fontColor: '1F4E79',
        fill: { type: 'solid', color: 'FFFF00' },
        borders: { edges: 'outline', style: 'thin', color: '00FF00' },
        horizontal: 'right',
        vertical: 'top',
        wrapText: true,
        numberFormat: { code: '0.000' },
      }),
    ]);
    expect(r.ok).toBe(true);
    const pkg = store.package;
    const a1 = formatOf(pkg, 'A1');
    expect(locals(a1.font)).toEqual(['i', 'u', 'sz', 'color', 'name', 'family']);
    expect(locals(child(a1.fill, 'x:patternFill'))).toEqual(['fgColor', 'bgColor']);
    expect(locals(a1.border)).toEqual(['left', 'right', 'top', 'bottom', 'diagonal']);
    expect(attr(child(a1.border, 'x:left'), 'style')).toBe('thin');
    expect(attr(child(a1.border, 'x:right'), 'style')).toBeUndefined();
    const b2 = formatOf(pkg, 'B2');
    expect(attr(child(b2.border, 'x:right'), 'style')).toBe('thin');
    expect(attr(child(b2.border, 'x:bottom'), 'style')).toBe('thin');
    expect(attr(child(b2.border, 'x:top'), 'style')).toBeUndefined();
    expect(locals(a1.xf)).toEqual(['alignment']);
    // The styles part keeps the `CT_Stylesheet` sequence, with `x:numFmts` created first.
    const root = stylesRoot(pkg);
    expect(locals(root).slice(0, 6)).toEqual([
      'numFmts',
      'fonts',
      'fills',
      'borders',
      'cellStyleXfs',
      'cellXfs',
    ]);
    for (const name of ['numFmts', 'fonts', 'fills', 'borders', 'cellXfs'])
      expect(Number(attr(child(root, `x:${name}`), 'count'))).toBe(
        elements(child(root, `x:${name}`)).length
      );
    // The saved part is the part the store holds: the save rules change no count.
    const saved = readXlsx(saveXlsx(pkg));
    if (!saved.ok) throw new Error('reopen');
    expect(serializeXmlPart(xmlPartOf(saved.pkg, '/xl/styles.xml')!)).toBe(
      serializeXmlPart(xmlPartOf(pkg, '/xl/styles.xml')!)
    );
    // B2 did not exist: the edit created it with the format and no value.
    const created = cellAt(pkg, 'B2')!;
    expect(created.children).toHaveLength(0);
  });

  test('a custom number format gets the lowest unused id from 164, and a code is reused', () => {
    const wb = base();
    const store = new WorkbookDocumentStore(
      open({
        ...wb,
        styles: {
          ...wb.styles,
          numFmts: [
            '<numFmt numFmtId="164" formatCode="0.0"/>',
            '<numFmt numFmtId="166" formatCode="#,##0.0"/>',
          ],
        },
      })
    );
    store.transact([fmt('A1', { numberFormat: { code: '0.000' } })]);
    store.transact([fmt('A2', { numberFormat: { code: '0.0000' } })]);
    store.transact([fmt('B1', { numberFormat: { code: '0.0' } })]);
    store.transact([fmt('C1', { numberFormat: { id: 10 } })]);
    const lists = styles(store.package);
    expect([...lists.numFmts]).toEqual([
      [164, '0.0'],
      [166, '#,##0.0'],
      [165, '0.000'],
      [167, '0.0000'],
    ]);
    const id = (ref: string) => attr(formatOf(store.package, ref).xf, 'numFmtId');
    expect([id('A1'), id('A2'), id('B1'), id('C1')]).toEqual(['165', '167', '164', '10']);
  });

  test('a font name drops the theme scheme, and markup in a name or code is escaped on save', () => {
    const store = new WorkbookDocumentStore(
      open({
        sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }],
        defaultFont: '<sz val="11"/><name val="Aptos"/><scheme val="minor"/>',
      })
    );
    store.transact([fmt('A1', { fontName: 'A"<&>B', numberFormat: { code: '"<&>"0' } })]);
    const font = formatOf(store.package, 'A1').font;
    expect(child(font, 'x:scheme')).toBeUndefined();
    const xml = serializeXmlPart(xmlPartOf(store.package, '/xl/styles.xml')!);
    expect(xml).toContain('val="A&quot;&lt;&amp;&gt;B"');
    expect(xml).toContain('formatCode="&quot;&lt;&amp;&gt;&quot;0"');
  });

  test('an automatic font color writes `x:color` with `auto` in place of `rgb`', () => {
    const store = new WorkbookDocumentStore(open(base()));
    expect(store.transact([fmt('A1', { fontColor: 'FF0000' })]).ok).toBe(true);
    expect(store.transact([fmt('A1', { fontColor: 'auto' })]).ok).toBe(true);
    const color = child(formatOf(store.package, 'A1').font, 'x:color');
    // `CT_Color/@auto` (ECMA-376 Part 1, 18.8.19), and no `rgb` from the earlier color.
    expect(attr(color, 'auto')).toBe('1');
    expect(attr(color, 'rgb')).toBeUndefined();
    expect(store.transact([fmt('A1', { fontColor: 'AUTO' })]).ok).toBe(false);
  });

  test('values outside their ranges are refused with bad-format', () => {
    const store = new WorkbookDocumentStore(open(base()));
    const bad: FormatPatch[] = [
      {},
      { fontColor: 'red' },
      { fontSize: 0 },
      { fontSize: Number.NaN },
      { fontName: '' },
      { fontName: 'x'.repeat(32) },
      { fontName: 'a\u0001b' },
      { fill: { type: 'solid', color: '12345' } },
      { borders: { edges: 'inner' as never } },
      { horizontal: 'middle' as never },
      { numberFormat: { id: 164 } },
      { numberFormat: { code: '' } },
      { numberFormat: { code: 'x'.repeat(256) } },
      { bold: 'yes' as never },
      { __proto__: { bold: true } } as never,
      JSON.parse('{"constructor": 1}') as FormatPatch,
    ];
    for (const patch of bad) {
      const r = store.transact([fmt('A1', patch)]);
      expect(r.ok ? 'ok' : r.rejection.code).toBe('bad-format');
    }
    expect(store.revision).toBe(0);
  });

  test('one command creates at most one whole row of cells, and never walks a column', () => {
    const store = new WorkbookDocumentStore(open(base()));
    const r = store.transact([fmt('A1:A1048576', { bold: true })]);
    expect(r.ok ? 'ok' : r.rejection.code).toBe('range-too-large');
    expect(store.transact([fmt('A1:XFD1048576', { bold: true })]).ok).toBe(false);
    expect(store.revision).toBe(0);
    const started = performance.now();
    expect(store.transact([fmt('A3:XFD3', { italic: true })]).ok).toBe(true);
    expect(performance.now() - started).toBeLessThan(2000);
    const rowEl = child(
      xmlPartOf(store.package, '/xl/worksheets/sheet1.xml')!.root,
      'x:sheetData'
    )!.children.at(-1)!;
    expect(children(rowEl as never, 'x:c')).toHaveLength(MAX_FORMAT_NEW_CELLS);
    // The cells that exist do not count against the cap.
    expect(store.transact([fmt('A1:XFD3', { italic: true })]).ok).toBe(false);
  });

  test('a new cell starts from its row format or its column format', () => {
    const store = new WorkbookDocumentStore(
      open({
        sheets: [
          {
            name: 'S',
            cols: '<cols><col min="3" max="3" width="9" style="1" customWidth="1"/></cols>',
            rows: '<row r="2" s="1" customFormat="1"><c r="A2" s="1"><v>1</v></c></row>',
          },
        ],
        styles: {
          fills: [
            '<fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill></fill>',
          ],
          cellXfs: ['<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>'],
        },
      })
    );
    store.transact([fmt('B1:C2', { bold: true })]);
    const pkg = store.package;
    expect(attr(formatOf(pkg, 'B2').xf, 'fillId')).toBe('2');
    expect(attr(formatOf(pkg, 'C1').xf, 'fillId')).toBe('2');
    expect(attr(formatOf(pkg, 'B1').xf, 'fillId')).toBe('0');
    expect(locals(formatOf(pkg, 'B1').font)).toContain('b');
  });

  test('undo restores the package exactly in one step, and redo applies it again', () => {
    const pkg = open(base());
    const store = new WorkbookDocumentStore(pkg);
    const original = canonicalPackageFingerprint(pkg);
    store.transact([
      fmt('A1:D4', {
        bold: true,
        fill: { type: 'solid', color: 'ABCDEF' },
        borders: { edges: 'all' },
        numberFormat: { code: '0.00;[Red]-0.00' },
      }),
    ]);
    const edited = canonicalPackageFingerprint(store.package);
    expect(edited).not.toBe(original);
    expect(store.undo().ok).toBe(true);
    expect(canonicalPackageFingerprint(store.package)).toBe(original);
    expect(store.canUndo).toBe(false);
    expect(store.redo().ok).toBe(true);
    expect(canonicalPackageFingerprint(store.package)).toBe(edited);
  });

  test('a workbook with no styles part gets one, and undo removes it with its links', () => {
    const pkg = open({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] });
    const parts = new Map(pkg.parts);
    parts.delete('/xl/styles.xml');
    const bare = readXlsx(saveXlsx(stripStyles({ parts })));
    if (!bare.ok) throw new Error('reopen');
    const store = new WorkbookDocumentStore(bare.pkg);
    const original = canonicalPackageFingerprint(bare.pkg);
    expect(store.transact([fmt('A1', { bold: true })]).ok).toBe(true);
    const index = workbookIndex(store.package)!;
    expect(index.stylesPart).toBe('/xl/styles.xml');
    const lists = styles(store.package);
    // The default records come first, so the unformatted cells keep the default format.
    expect(lists.cellXfs).toHaveLength(2);
    expect(lists.fills.map((f) => attr(child(f, 'x:patternFill'), 'patternType'))).toEqual([
      'none',
      'gray125',
    ]);
    expect(attr(cellAt(store.package, 'A1'), 's')).toBe('1');
    // ECMA-376 18.8.9 and 18.8.8: a new part carries the Normal cell style, and each x:xf links to it.
    const root = stylesRoot(store.package);
    expect(children(child(root, 'x:cellStyleXfs')!, 'x:xf')).toHaveLength(1);
    const normal = child(child(root, 'x:cellStyles')!, 'x:cellStyle')!;
    expect([attr(normal, 'name'), attr(normal, 'xfId'), attr(normal, 'builtinId')]).toEqual([
      'Normal',
      '0',
      '0',
    ]);
    expect(lists.cellXfs.map((xf) => attr(xf, 'xfId'))).toEqual(['0', '0']);
    const reopened = readXlsx(saveXlsx(store.package));
    expect(reopened.ok && workbookIndex(reopened.pkg)!.stylesPart).toBe('/xl/styles.xml');
    store.undo();
    expect(canonicalPackageFingerprint(store.package)).toBe(original);
    expect(partOf(store.package, '/xl/styles.xml')).toBeUndefined();
  });
});

/** The package without its style relationship and content type override. */
function stripStyles(pkg: XlsxPackage): XlsxPackage {
  const parts = new Map(pkg.parts);
  const rels = xmlPartOf(pkg, '/xl/_rels/workbook.xml.rels')!;
  const id = relationshipsOf(pkg, '/xl/workbook.xml').find((r) => r.type.endsWith('/styles'))!.id;
  parts.set(rels.partName, {
    kind: 'xml',
    partName: rels.partName,
    xml: {
      ...rels,
      root: {
        ...rels.root,
        children: rels.root.children.filter((n) => n.type !== 'element' || attr(n, 'Id') !== id),
      },
    },
  });
  const ct = xmlPartOf(pkg, '/[Content_Types].xml')!;
  parts.set(ct.partName, {
    kind: 'xml',
    partName: ct.partName,
    xml: {
      ...ct,
      root: {
        ...ct.root,
        children: ct.root.children.filter(
          (n) => n.type !== 'element' || attr(n, 'PartName') !== '/xl/styles.xml'
        ),
      },
    },
  });
  return { parts };
}

describe('clearFormats', () => {
  test('removes the cell formats, keeps the values, and drops cells that hold nothing', () => {
    const store = new WorkbookDocumentStore(open(base()));
    store.transact([fmt('D1', { bold: true })]);
    expect(cellAt(store.package, 'D1')).toBeDefined();
    const r = store.transact([{ op: 'clearFormats', sheet: 'S', ref: 'A1:XFD1048576' }]);
    expect(r.ok).toBe(true);
    const pkg = store.package;
    expect(attr(cellAt(pkg, 'B1'), 's')).toBeUndefined();
    expect(child(cellAt(pkg, 'B1'), 'x:v')).toBeDefined();
    expect(cellAt(pkg, 'D1')).toBeUndefined();
    store.undo();
    expect(attr(cellAt(store.package, 'B1'), 's')).toBe('1');
  });

  test('clearing values and formats is one transaction and one undo step', () => {
    const store = new WorkbookDocumentStore(open(base()));
    const original = canonicalPackageFingerprint(store.package);
    store.transact([
      { op: 'clearValues', sheet: 'S', ref: 'A1:C1' },
      { op: 'clearFormats', sheet: 'S', ref: 'A1:C1' },
    ]);
    expect(cellAt(store.package, 'B1')).toBeUndefined();
    expect(store.revision).toBe(1);
    store.undo();
    expect(canonicalPackageFingerprint(store.package)).toBe(original);
  });
});

describe('format ops and the store guards', () => {
  const dimension = (pkg: XlsxPackage) =>
    attr(child(xmlPartOf(pkg, '/xl/worksheets/sheet1.xml')!.root, 'x:dimension'), 'ref');

  test('cells that a format creates outside x:dimension widen it in the same undo step', () => {
    const store = new WorkbookDocumentStore(
      open({ sheets: [{ name: 'S', dimension: 'B2:C2', rows: row(2, [c('B2', 1), c('C2', 2)]) }] })
    );
    expect(store.transact([fmt('A1:D4', { bold: true })]).ok).toBe(true);
    expect(dimension(store.package)).toBe('A1:D4');
    expect(store.undo().ok).toBe(true);
    expect(dimension(store.package)).toBe('B2:C2');
    expect(store.redo().ok).toBe(true);
    expect(dimension(store.package)).toBe('A1:D4');
    // A format of existing cells inside the range keeps it.
    expect(store.transact([fmt('B2', { italic: true })]).ok).toBe(true);
    expect(dimension(store.package)).toBe('A1:D4');
  });

  test('a format that changes nothing commits nothing', () => {
    const store = new WorkbookDocumentStore(
      open({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] })
    );
    expect(store.transact([fmt('A1', { bold: true })]).ok).toBe(true);
    const [revision, pkg] = [store.revision, store.package];
    expect(store.transact([fmt('A1', { bold: true })]).ok).toBe(true);
    expect(store.transact([{ op: 'clearFormats', sheet: 'S', ref: 'B5:C9' }]).ok).toBe(true);
    expect([store.revision, store.package]).toEqual([revision, pkg]);
  });

  test('a style record with a character that XML cannot hold never commits', () => {
    const store = new WorkbookDocumentStore(open({ sheets: [{ name: 'S', rows: '' }] }));
    for (const patch of [{ fontName: 'A\u0001' }, { numberFormat: { code: '0\ud800' } }])
      expect(store.transact([fmt('A1', patch)]).ok).toBe(false);
    expect(store.revision).toBe(0);
    expect(() => saveXlsx(store.package)).not.toThrow();
  });

  test('the part check refuses a new style record that XML cannot hold, whatever op brings it', () => {
    const pkg = open({ sheets: [{ name: 'S', rows: '' }] });
    const part = xmlPartOf(pkg, workbookIndex(pkg)!.stylesPart)!;
    const fonts = child(part.root, 'x:fonts')!;
    const ids = new IdSource(part);
    const name = smlElement(ids, 'name', 'generic', [plainAttr('val', 'A\u0001')]);
    const font = smlElement(ids, 'font', 'generic', [], [name]);
    const root = {
      ...part.root,
      children: part.root.children.map((n) =>
        n === fonts ? { ...fonts, children: [...fonts.children, font] } : n
      ),
    };
    const r = applyHistory(pkg, [
      { op: 'restoreStyles', part: { ...part, root, nextId: ids.next } },
    ]);
    expect(r.ok ? null : r.rejection.code).toBe('invalid-part');
  });

  test('a code equal to a built-in code keeps its code when the workbook redefines that id', () => {
    const redefined = open({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }],
      styles: { numFmts: ['<numFmt numFmtId="14" formatCode="yyyy-mm-dd"/>'] },
    });
    const store = new WorkbookDocumentStore(redefined);
    store.transact([fmt('A1', { numberFormat: { code: 'mm-dd-yy' } })]);
    expect(attr(formatOf(store.package, 'A1').xf, 'numFmtId')).toBe('164');
    const plain = new WorkbookDocumentStore(
      open({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] })
    );
    plain.transact([fmt('A1', { numberFormat: { code: 'mm-dd-yy' } })]);
    expect(attr(formatOf(plain.package, 'A1').xf, 'numFmtId')).toBe('14');
  });

  test('a code equal to a redefined built-in id gets a custom id', () => {
    const store = new WorkbookDocumentStore(
      open({
        sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }],
        styles: { numFmts: ['<numFmt numFmtId="14" formatCode="yyyy-mm-dd"/>'] },
      })
    );
    store.transact([fmt('A1', { numberFormat: { code: 'yyyy-mm-dd' } })]);
    expect(attr(formatOf(store.package, 'A1').xf, 'numFmtId')).toBe('164');
  });

  test('a new code skips a custom id that a cell format names with no record', () => {
    const store = new WorkbookDocumentStore(
      open({
        sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 2, { s: 1 })]) }],
        styles: {
          cellXfs: [
            '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>',
          ],
          dxfs: ['<dxf><numFmt numFmtId="165" formatCode="0.0"/></dxf>'],
        },
      })
    );
    store.transact([fmt('A1', { numberFormat: { code: '0.000%' } })]);
    expect(attr(formatOf(store.package, 'A1').xf, 'numFmtId')).toBe('166');
    expect(attr(formatOf(store.package, 'B1').xf, 'numFmtId')).toBe('164');
  });

  test('a record id with white space is in use, and only the first record of an id is reused', () => {
    const store = new WorkbookDocumentStore(
      open({
        sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 2)]) }],
        styles: {
          numFmts: [
            '<numFmt numFmtId=" 164" formatCode="0.0"/>',
            '<numFmt numFmtId="0000000168" formatCode="0.0%"/>',
            '<numFmt numFmtId="165" formatCode="0.00%"/>',
            '<numFmt numFmtId="165" formatCode="0.000"/>',
          ],
        },
      })
    );
    store.transact([fmt('A1', { numberFormat: { code: '0.0000' } })]);
    expect(attr(formatOf(store.package, 'A1').xf, 'numFmtId')).toBe('166');
    store.transact([fmt('B1', { numberFormat: { code: '0.000' } })]);
    expect(attr(formatOf(store.package, 'B1').xf, 'numFmtId')).toBe('167');
    store.transact([fmt('A1', { numberFormat: { code: '0.00000' } })]);
    expect(attr(formatOf(store.package, 'A1').xf, 'numFmtId')).toBe('169');
  });
});
