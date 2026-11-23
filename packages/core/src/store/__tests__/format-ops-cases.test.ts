// More format op cases: a format edit and its undo never touch the stale flag
// or the calculation chain, the store reads every custom number format code, a font change lets
// the row fit again, and a clear keeps an empty cell from showing its row or column format.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
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
import { attr, child } from '../package/query.ts';
import { sheetByName, styleLists, workbookIndex } from '../workbook/workbook-index.ts';
import { cellKey, sheetIndex } from '../workbook/sheet-index.ts';
import { WorkbookDocumentStore } from '../store/document-store.ts';
import type { WorkbookOp } from '../store/ops.ts';
import type { FormatPatch } from '../store/format-patch.ts';

const FIXTURES = join(import.meta.dir, '../../../../../fixtures');

function open(wb: FixtureWorkbook): XlsxPackage {
  const read = readXlsx(buildXlsx(wb));
  if (!read.ok) throw new Error(JSON.stringify(read.error));
  return read.pkg;
}

function openFixture(name: string): XlsxPackage {
  const read = readXlsx(readFileSync(join(FIXTURES, `${name}.xlsx`)));
  if (!read.ok) throw new Error(JSON.stringify(read.error));
  return read.pkg;
}

const fmt = (ref: string, patch: FormatPatch, sheet = 'S'): WorkbookOp => ({
  op: 'applyFormat',
  sheet,
  ref,
  patch,
});

const sheetPart = (pkg: XlsxPackage, sheet = 'S') =>
  xmlPartOf(pkg, sheetByName(workbookIndex(pkg)!, sheet)!.partName)!;

function cellAt(pkg: XlsxPackage, row: number, col: number, sheet = 'S') {
  return sheetIndex(sheetPart(pkg, sheet)).cells.get(cellKey(row, col))?.el;
}

const rowAt = (pkg: XlsxPackage, r: number, sheet = 'S') =>
  sheetIndex(sheetPart(pkg, sheet)).rows.get(r);

const calcPr = (pkg: XlsxPackage) => child(xmlPartOf(pkg, '/xl/workbook.xml')!.root, 'x:calcPr');

describe('format edits and the calculation state', () => {
  const formulas = 'Formulas';

  test('a format and its undo keep calcPr and the calculation chain', () => {
    const pkg = openFixture('formulas');
    const original = canonicalPackageFingerprint(pkg);
    const store = new WorkbookDocumentStore(pkg);
    const r = store.transact([fmt('A1:B7', { bold: true, numberFormat: { id: 2 } }, formulas)]);
    expect(r.ok).toBe(true);
    expect(partOf(store.package, '/xl/calcChain.xml')).toBeDefined();
    expect(attr(calcPr(store.package), 'fullCalcOnLoad')).toBeUndefined();
    expect(store.undo().ok).toBe(true);
    expect(canonicalPackageFingerprint(store.package)).toBe(original);
    expect(store.redo().ok).toBe(true);
    expect(partOf(store.package, '/xl/calcChain.xml')).toBeDefined();
    expect(attr(calcPr(store.package), 'fullCalcOnLoad')).toBeUndefined();
  });

  test('a clear of formats and its undo keep them too', () => {
    const pkg = openFixture('formulas');
    const store = new WorkbookDocumentStore(pkg);
    store.transact([fmt('A1:B7', { italic: true }, formulas)]);
    const formatted = canonicalPackageFingerprint(store.package);
    store.transact([{ op: 'clearFormats', sheet: formulas, ref: 'A1:B7' }]);
    expect(store.undo().ok).toBe(true);
    expect(canonicalPackageFingerprint(store.package)).toBe(formatted);
    expect(partOf(store.package, '/xl/calcChain.xml')).toBeDefined();
    expect(attr(calcPr(store.package), 'fullCalcOnLoad')).toBeUndefined();
  });

  test('the undo of a value edit still reports a value change', () => {
    const store = new WorkbookDocumentStore(openFixture('formulas'));
    store.transact([
      { op: 'setCellValue', sheet: formulas, ref: 'A1', value: { type: 'number', value: 9 } },
    ]);
    expect(attr(calcPr(store.package), 'fullCalcOnLoad')).toBe('1');
    expect(store.undo().ok).toBe(true);
    // The stale flag stays (design D3), and the cell value is back.
    expect(attr(calcPr(store.package), 'fullCalcOnLoad')).toBe('1');
    expect(child(cellAt(store.package, 0, 0, formulas), 'x:v')).toBeDefined();
  });
});

describe('number format codes in the store', () => {
  const wb = (): FixtureWorkbook => ({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] });

  test('an op with a code that the parser cannot read is refused', () => {
    const store = new WorkbookDocumentStore(open(wb()));
    for (const code of ['[[[0;"', '"open', '[Red', '0;0;0;0;0']) {
      const r = store.transact([fmt('A1', { numberFormat: { code } })]);
      expect(r.ok ? null : r.rejection.code).toBe('unreadable-number-format');
    }
    expect(store.revision).toBe(0);
  });

  test('a code that equals a built-in code is written as its id', () => {
    const store = new WorkbookDocumentStore(open(wb()));
    const id = () => {
      const lists = styleLists(store.package, workbookIndex(store.package)!);
      const xf = lists.cellXfs[Number(attr(cellAt(store.package, 0, 0), 's') ?? '0')]!;
      return [attr(xf, 'numFmtId'), lists.numFmts.size];
    };
    store.transact([fmt('A1', { numberFormat: { code: '0.00' } })]);
    expect(id()).toEqual(['2', 0]);
    store.transact([fmt('A1', { numberFormat: { code: '0.0%' } })]);
    expect(id()).toEqual(['164', 1]);
    store.transact([fmt('A1', { numberFormat: { code: 'General' } })]);
    expect(id()).toEqual(['0', 1]);
  });
});

describe('row heights after a font change', () => {
  const wb = (): FixtureWorkbook => ({
    sheets: [
      {
        name: 'S',
        rows:
          '<row r="1" ht="15"><c r="A1"><v>1</v></c></row>' +
          '<row r="2" ht="15" customHeight="1"><c r="A2"><v>2</v></c></row>',
      },
    ],
  });

  test('a font size removes a fitted height, keeps a custom one, and undo puts it back', () => {
    const store = new WorkbookDocumentStore(open(wb()));
    store.transact([fmt('A1:A2', { fontSize: 24 })]);
    expect(attr(rowAt(store.package, 0), 'ht')).toBeUndefined();
    expect(attr(rowAt(store.package, 1), 'ht')).toBe('15');
    store.undo();
    expect(attr(rowAt(store.package, 0), 'ht')).toBe('15');
  });

  test('bold keeps the height; a font name, wrapping, or a clear of a larger font removes it', () => {
    const store = new WorkbookDocumentStore(open(wb()));
    store.transact([fmt('A1', { bold: true })]);
    expect(attr(rowAt(store.package, 0), 'ht')).toBe('15');
    for (const patch of [{ wrapText: true }, { fontName: 'Courier New' }]) {
      const fitted = new WorkbookDocumentStore(open(wb()));
      fitted.transact([fmt('A1', patch)]);
      expect(attr(rowAt(fitted.package, 0), 'ht')).toBeUndefined();
    }
    const withHeight = new WorkbookDocumentStore(
      open({
        sheets: [{ name: 'S', rows: '<row r="1" ht="30"><c r="A1" s="1"><v>1</v></c></row>' }],
        styles: {
          fonts: ['<font><sz val="24"/><name val="Calibri"/></font>'],
          cellXfs: ['<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'],
        },
      })
    );
    withHeight.transact([{ op: 'clearFormats', sheet: 'S', ref: 'A1' }]);
    expect(attr(rowAt(withHeight.package, 0), 'ht')).toBeUndefined();
  });
});

describe('clearFormats in a formatted row or column', () => {
  test('an empty cell keeps s="0" so it does not show its row or column format', () => {
    const store = new WorkbookDocumentStore(
      open({
        sheets: [
          {
            name: 'S',
            cols: '<cols><col min="3" max="3" width="9" style="1" customWidth="1"/></cols>',
            rows:
              '<row r="1"><c r="A1" s="1"/><c r="C1" s="1"/></row>' +
              '<row r="2" s="1" customFormat="1"><c r="A2" s="1"/></row>',
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
    store.transact([{ op: 'clearFormats', sheet: 'S', ref: 'A1:C2' }]);
    // A1: no row or column format, so the empty cell goes away.
    expect(cellAt(store.package, 0, 0)).toBeUndefined();
    expect(attr(cellAt(store.package, 0, 2), 's')).toBe('0');
    expect(attr(cellAt(store.package, 1, 0), 's')).toBe('0');
  });

  const formatted = () =>
    open({
      sheets: [
        {
          name: 'S',
          cols: '<cols><col min="3" max="4" width="9" style="1" customWidth="1"/></cols>',
          rows: '<row r="2" s="1" customFormat="1"><c r="A2"><v>1</v></c></row>',
        },
      ],
      styles: {
        cellXfs: ['<xf numFmtId="0" fontId="0" fillId="1" borderId="0" xfId="0" applyFill="1"/>'],
      },
    });

  test('an empty position of a formatted row or column gets s="0", and undo removes it', () => {
    const store = new WorkbookDocumentStore(formatted());
    const original = canonicalPackageFingerprint(store.package);
    // C1 has no cell and shows its column format; B2 has no cell and shows its row format.
    for (const ref of ['C1', 'B2']) {
      const r = store.transact([{ op: 'clearFormats', sheet: 'S', ref }]);
      expect(r.ok ? store.revision : r.rejection).toBeGreaterThan(0);
    }
    expect(attr(cellAt(store.package, 0, 2), 's')).toBe('0');
    expect(attr(cellAt(store.package, 1, 1), 's')).toBe('0');
    // A position whose row and column have no format gets no cell.
    expect(store.transact([{ op: 'clearFormats', sheet: 'S', ref: 'A1:B1' }]).ok).toBe(true);
    expect(cellAt(store.package, 0, 0)).toBeUndefined();
    expect(store.revision).toBe(2);
    store.undo();
    store.undo();
    expect(canonicalPackageFingerprint(store.package)).toBe(original);
  });

  test('clearing values and formats of the range writes s="0" where a format shows', () => {
    const store = new WorkbookDocumentStore(formatted());
    store.transact([
      { op: 'clearValues', sheet: 'S', ref: 'A2:D3' },
      { op: 'clearFormats', sheet: 'S', ref: 'A2:D3' },
    ]);
    // Row 2 has a format: each position keeps s="0". Row 3 has none: only C3 and D3.
    for (const col of [0, 1, 2, 3]) expect(attr(cellAt(store.package, 1, col), 's')).toBe('0');
    expect(cellAt(store.package, 2, 0)).toBeUndefined();
    expect(cellAt(store.package, 2, 1)).toBeUndefined();
    expect(attr(cellAt(store.package, 2, 2), 's')).toBe('0');
    expect(attr(cellAt(store.package, 2, 3), 's')).toBe('0');
  });

  test('a clear that would create more cells than the cap is refused with its count', () => {
    const store = new WorkbookDocumentStore(formatted());
    // Two formatted columns over 9 999 rows, and three positions of the formatted row 2.
    const r = store.transact([{ op: 'clearFormats', sheet: 'S', ref: 'A1:D10000' }]);
    expect(r.ok ? null : r.rejection.code).toBe('clear-too-large');
    expect(r.ok ? '' : r.rejection.detail).toContain('20001');
    // One whole row with a format fits, and a whole sheet without formats clears.
    expect(store.transact([{ op: 'clearFormats', sheet: 'S', ref: 'A2:XFD2' }]).ok).toBe(true);
    const plain = new WorkbookDocumentStore(
      open({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1, { s: 1 })]) }] })
    );
    expect(plain.transact([{ op: 'clearFormats', sheet: 'S', ref: 'A1:XFD1048576' }]).ok).toBe(
      true
    );
  });
});

describe('undo of a format edit', () => {
  test('never leaves a cell format index past the cell formats', () => {
    const store = new WorkbookDocumentStore(
      open({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 2)]) }] })
    );
    const count = () => styleLists(store.package, workbookIndex(store.package)!).cellXfs.length;
    const xfs = count();
    expect(store.transact([fmt('A1:C2', { bold: true })]).ok).toBe(true);
    expect(store.transact([fmt('B1', { italic: true })]).ok).toBe(true);
    expect(count()).toBeGreaterThan(xfs);
    for (let i = 0; i < 2; i += 1) {
      expect(store.undo().ok).toBe(true);
      const limit = count();
      for (let col = 0; col < 3; col += 1)
        for (let r = 0; r < 2; r += 1) {
          const s = attr(cellAt(store.package, r, col), 's');
          expect(s === undefined || Number(s) < limit).toBe(true);
        }
    }
    expect(count()).toBe(xfs);
  });
});

describe('the cell cap of one format command', () => {
  test('a range far over the cap is refused before the cells are counted', () => {
    const rows: string[] = [];
    for (let r = 1; r <= 2000; r += 1) rows.push(row(r, [c(`A${r}`, r), c(`B${r}`, r)]));
    const store = new WorkbookDocumentStore(open({ sheets: [{ name: 'S', rows: rows.join('') }] }));
    const r = store.transact([fmt('A1:XFD1048576', { bold: true })]);
    expect(r.ok ? null : r.rejection.detail).toContain('at least');
    // A range near the cap counts the cells that exist in it: none of the 4000 are in C:D.
    const near = store.transact([fmt('C1:D10000', { bold: true })]);
    expect(near.ok ? null : near.rejection.code).toBe('range-too-large');
    expect(near.ok ? '' : near.rejection.detail).not.toContain('at least');
  });
});
