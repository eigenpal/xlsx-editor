// The `setCellFormatIndex` op (design D2 of the change `workbook-structure-editing`): it sets
// the cell format index of a cell, keeps the value, creates or removes a cell only where the
// row or column format asks for it, refuses an index that the workbook does not have, and one
// undo puts the package back as it was.

import { describe, expect, test } from 'bun:test';
import {
  buildXlsx,
  c,
  row,
  type FixtureWorkbook,
} from '../../../../../scripts/fixtures/builder.ts';
import {
  canonicalPackageFingerprint,
  readXlsx,
  xmlPartOf,
  type XlsxPackage,
} from '../package/package.ts';
import { attr, child } from '../package/query.ts';
import { sheetByName, workbookIndex } from '../workbook/workbook-index.ts';
import { cellKey, sheetIndex } from '../workbook/sheet-index.ts';
import { WorkbookDocumentStore } from '../store/document-store.ts';
import type { WorkbookOp } from '../store/ops.ts';

function open(wb: FixtureWorkbook): XlsxPackage {
  const read = readXlsx(buildXlsx(wb));
  if (!read.ok) throw new Error(JSON.stringify(read.error));
  return read.pkg;
}

const set = (ref: string, index: number, sheet = 'S'): WorkbookOp => ({
  op: 'setCellFormatIndex',
  sheet,
  ref,
  index,
});

function sheetOf(pkg: XlsxPackage, sheet = 'S') {
  const index = workbookIndex(pkg)!;
  return xmlPartOf(pkg, sheetByName(index, sheet)!.partName)!;
}

function cellAt(pkg: XlsxPackage, row: number, col: number) {
  return sheetIndex(sheetOf(pkg)).cells.get(cellKey(row, col))?.el;
}

const base = (): FixtureWorkbook => ({
  sheets: [
    {
      name: 'S',
      dimension: 'A1:B2',
      rows: [row(1, [c('A1', 1), c('B1', 2, { s: 1 })]), row(2, [c('A2', 'x')])].join(''),
      cols: '<cols><col min="4" max="4" width="9" style="2" customWidth="1"/></cols>',
    },
  ],
  styles: {
    fonts: ['<font><b/><sz val="11"/><name val="Calibri"/></font>'],
    cellXfs: [
      '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
      '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
    ],
  },
});

describe('setCellFormatIndex', () => {
  test('sets and removes @s, keeps the value, and one undo restores the package', () => {
    const store = new WorkbookDocumentStore(open(base()));
    const before = canonicalPackageFingerprint(store.package);
    expect(store.transact([set('A1', 2), set('B1', 0)]).ok).toBe(true);
    expect(attr(cellAt(store.package, 0, 0), 's')).toBe('2');
    expect(child(cellAt(store.package, 0, 0)!, 'x:v')).toBeDefined();
    expect(attr(cellAt(store.package, 0, 1), 's')).toBeUndefined();
    expect(child(cellAt(store.package, 0, 1)!, 'x:v')).toBeDefined();
    expect(store.undo().ok).toBe(true);
    expect(canonicalPackageFingerprint(store.package)).toBe(before);
  });

  test('a cell is created only when its row or column shows another format', () => {
    const store = new WorkbookDocumentStore(open(base()));
    const revision = store.revision;
    // C3 shows format 0 already; D3 shows the column format 2.
    expect(store.transact([set('C3', 0), set('D3', 2)]).ok).toBe(true);
    expect(store.revision).toBe(revision);
    expect(store.transact([set('C3', 1), set('D3', 0)]).ok).toBe(true);
    expect(attr(cellAt(store.package, 2, 2), 's')).toBe('1');
    // Format 0 under a formatted column is explicit, or the cell would show the column format.
    expect(attr(cellAt(store.package, 2, 3), 's')).toBe('0');
    // The new cells widen the dimension in the same undo step.
    const dim = (pkg: XlsxPackage) => attr(child(sheetOf(pkg).root, 'x:dimension'), 'ref');
    expect(dim(store.package)).toBe('A1:D3');
    expect(store.undo().ok).toBe(true);
    expect(dim(store.package)).toBe('A1:B2');
    expect(cellAt(store.package, 2, 2)).toBeUndefined();
  });

  test('an emptied cell that shows the inherited format goes away', () => {
    const store = new WorkbookDocumentStore(open(base()));
    expect(store.transact([set('D1', 1)]).ok).toBe(true);
    expect(attr(cellAt(store.package, 0, 3), 's')).toBe('1');
    expect(store.transact([set('D1', 2)]).ok).toBe(true);
    expect(cellAt(store.package, 0, 3)).toBeUndefined();
  });

  test('the last op for a cell wins, and the value of the cell does not change', () => {
    const store = new WorkbookDocumentStore(open(base()));
    expect(store.transact([set('A2', 1), set('A2', 2)]).ok).toBe(true);
    expect(attr(cellAt(store.package, 1, 0), 's')).toBe('2');
    expect(attr(cellAt(store.package, 1, 0), 't')).toBe('inlineStr');
  });

  test('an index outside x:cellXfs and a bad reference are refused, and nothing changes', () => {
    const store = new WorkbookDocumentStore(open(base()));
    const before = canonicalPackageFingerprint(store.package);
    for (const index of [3, -1, 1.5, Number.NaN]) {
      const r = store.transact([set('A1', 1), set('B2', index)]);
      expect(r.ok ? null : r.rejection.code).toBe('bad-format');
    }
    const bad = store.transact([set('A0', 1)]);
    expect(bad.ok ? null : bad.rejection.code).toBe('bad-reference');
    const sheet = store.transact([set('A1', 1, 'Nope')]);
    expect(sheet.ok ? null : sheet.rejection.code).toBe('unknown-sheet');
    expect(canonicalPackageFingerprint(store.package)).toBe(before);
    expect(store.canUndo).toBe(false);
  });

  test('a format edit marks no formula stale', () => {
    const pkg = open({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 2, { f: '1+1' })]) }],
      calcPr: 'calcId="191029"',
      styles: { cellXfs: ['<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0"/>'] },
    });
    const store = new WorkbookDocumentStore(pkg);
    expect(store.transact([set('A1', 1)]).ok).toBe(true);
    const calcPr = child(xmlPartOf(store.package, '/xl/workbook.xml')!.root, 'x:calcPr');
    expect(attr(calcPr, 'fullCalcOnLoad')).toBeUndefined();
  });

  test('a changed cell format marks the callers of CELL and GET.* stale, also in a clear and a format', () => {
    const book = (f: string, defined = ''): XlsxPackage =>
      open({
        sheets: [
          {
            name: 'S',
            rows: row(1, [c('A1', 0, { f }), c('B1', 1)]) + row(2, [c('B2', 1, { s: 1 })]),
          },
        ],
        definedNames: defined,
        calcPr: 'calcId="191029"',
        styles: { cellXfs: ['<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0"/>'] },
      });
    const stale = (store: WorkbookDocumentStore) =>
      attr(child(xmlPartOf(store.package, '/xl/workbook.xml')!.root, 'x:calcPr'), 'fullCalcOnLoad');
    const clear: WorkbookOp = { op: 'clearFormats', sheet: 'S', ref: 'B2' };
    const bold: WorkbookOp = { op: 'applyFormat', sheet: 'S', ref: 'B1', patch: { bold: true } };
    for (const op of [set('B1', 1), clear, bold]) {
      for (const pkg of [
        book('CELL("format",B1)'),
        book('SUM(1)', '<definedName name="f">GET.CELL(7,S!B1)</definedName>'),
      ]) {
        const store = new WorkbookDocumentStore(pkg);
        expect(store.transact([op]).ok).toBe(true);
        expect(stale(store)).toBe('1');
      }
      // Without such a call the values stay current.
      const plain = new WorkbookDocumentStore(book('SUM(1)'));
      expect(plain.transact([op]).ok).toBe(true);
      expect(stale(plain)).toBeUndefined();
    }
  });
});
