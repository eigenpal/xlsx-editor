// The sheet ops and the calculation chain of design D5: an edit drops the chain outside the undo
// history, so an undo of a sheet op can put back a workbook relationship part or a content type
// part that still names the chain. No op sequence may leave a workbook relationship or a
// content type override that names a missing part, and a saved copy reopens with neither.

import { describe, expect, test } from 'bun:test';
import { c, row, type FixtureWorkbook } from '../../../../../scripts/fixtures/builder.ts';
import { attr } from '../package/query.ts';
import {
  CONTENT_TYPES_PART,
  equivalentPartName,
  readXlsx,
  relationshipsOf,
  xmlPartOf,
  type XlsxPackage,
} from '../package/package.ts';
import { NS } from '../package/namespaces.ts';
import { saveXlsx } from '../package-write/save.ts';
import { workbookIndex } from '../workbook/workbook-index.ts';
import { touchedPartViolation } from '../store/part-checks.ts';
import type { WorkbookDocumentStore } from '../store/document-store.ts';
import { storeOf } from './structure-helpers.ts';
import { sheetStale } from '../../cascade/index.ts';

const CHAIN_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml';
const CHAIN_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain';
const CHAIN_XML =
  '<?xml version="1.0" encoding="UTF-8"?><calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="B1" i="1"/></calcChain>';

const chain = (): FixtureWorkbook['extraParts'] => [
  ['xl/calcChain.xml', CHAIN_TYPE, new TextEncoder().encode(CHAIN_XML), CHAIN_REL],
];

const twoSheets = (): FixtureWorkbook => ({
  sheets: [
    { name: 'S', rows: row(1, [c('A1', 1), c('B1', 2, { f: 'A1*2' })]) },
    { name: 'T', rows: row(1, [c('A1', 1), c('B1', 2, { f: 'S!A1*2' })]) },
  ],
  extraParts: chain(),
});

/** The workbook relationships and content type overrides that name a missing part. */
function dangling(pkg: XlsxPackage): string[] {
  const out: string[] = [];
  const main = workbookIndex(pkg)!.mainPart;
  for (const r of relationshipsOf(pkg, main))
    if (!r.external && r.target && !equivalentPartName(pkg, r.target)) out.push(`rel ${r.target}`);
  for (const n of xmlPartOf(pkg, CONTENT_TYPES_PART)!.root.children)
    if (n.type === 'element' && n.ns === NS.ct && n.local === 'Override') {
      const name = attr(n, 'PartName') ?? '';
      if (!equivalentPartName(pkg, name)) out.push(`override ${name}`);
    }
  return out;
}

/** The current package and its saved and reopened copy have no dangling entry. */
function expectClean(s: WorkbookDocumentStore): void {
  expect(dangling(s.package)).toEqual([]);
  const reopened = readXlsx(saveXlsx(s.package));
  if (!reopened.ok) throw new Error(JSON.stringify(reopened.error));
  expect(dangling(reopened.pkg)).toEqual([]);
}

describe('sheet ops and the calculation chain', () => {
  test('delete a sheet, undo, and redo leave no relationship or override to the chain', () => {
    const s = storeOf(twoSheets());
    expect(s.transact([{ op: 'deleteSheet', sheet: 'T' }]).ok).toBe(true);
    expectClean(s);
    expect(s.undo().ok).toBe(true);
    expectClean(s);
    expect(workbookIndex(s.package)!.sheets.map((x) => x.name)).toEqual(['S', 'T']);
    expect(s.redo().ok).toBe(true);
    expectClean(s);
  });

  test('add a sheet, edit a formula, undo twice, and redo twice leave no dangling entry', () => {
    const s = storeOf({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 2, { f: 'A1*2' })]) }],
      extraParts: chain(),
    });
    expect(s.transact([{ op: 'addSheet', name: 'N', index: 1 }]).ok).toBe(true);
    const edit = s.transact([
      { op: 'setCellValue', sheet: 'S', ref: 'B1', value: { type: 'formula', text: 'A1*3' } },
    ]);
    expect(edit.ok).toBe(true);
    expectClean(s);
    expect(s.undo().ok).toBe(true);
    expectClean(s);
    expect(s.undo().ok).toBe(true);
    expectClean(s);
    expect(s.redo().ok).toBe(true);
    expectClean(s);
    expect(s.redo().ok).toBe(true);
    expectClean(s);
  });

  test('a workbook with a chain keeps it through an op that touches no formula', () => {
    const s = storeOf(twoSheets());
    expect(s.transact([{ op: 'moveSheet', sheet: 'T', index: 0 }]).ok).toBe(true);
    expect(relationshipsOf(s.package, '/xl/workbook.xml').some((r) => r.type === CHAIN_REL)).toBe(
      true
    );
    expect(equivalentPartName(s.package, '/xl/calcChain.xml')).toBeDefined();
  });

  test('the part check refuses a new workbook relationship or override to a missing part', () => {
    const s = storeOf(twoSheets());
    const before = s.package;
    // The package without the chain part, with its relationship and override kept.
    const parts = new Map(before.parts);
    parts.delete('/xl/calcChain.xml');
    const after: XlsxPackage = { parts };
    // Before is clean; after names a missing part through both touched parts.
    expect(
      touchedPartViolation(before, after, ['/xl/_rels/workbook.xml.rels', CONTENT_TYPES_PART])
    ).toContain('/xl/calcChain.xml');
    // An entry that already named a missing part does not refuse.
    expect(
      touchedPartViolation(after, after, ['/xl/_rels/workbook.xml.rels', CONTENT_TYPES_PART])
    ).toBeUndefined();
    // A transaction that touches neither part is not checked for it.
    expect(touchedPartViolation(before, after, [])).toBeUndefined();
  });
});

describe('a sheet add or rename to a name that a formula already names', () => {
  const fullCalc = (s: WorkbookDocumentStore): string | undefined => {
    const index = workbookIndex(s.package)!;
    const root = xmlPartOf(s.package, index.mainPart)!.root;
    const calc = root.children.find((n) => n.type === 'element' && n.local === 'calcPr');
    return calc?.type === 'element' ? attr(calc, 'fullCalcOnLoad') : undefined;
  };
  const book = (): FixtureWorkbook => ({
    sheets: [{ name: 'S', rows: row(1, [c('A1', 0, { f: 'IFERROR(Later!A1,0)' })]) }],
  });

  test('an add of that sheet marks the cached results stale', () => {
    const s = storeOf(book());
    expect(fullCalc(s)).toBeUndefined();
    expect(s.transact([{ op: 'addSheet', name: 'later', index: 1 }]).ok).toBe(true);
    expect(fullCalc(s)).toBe('1');
  });

  test('a rename to that name marks them stale; another name does not', () => {
    const s = storeOf({ ...book(), sheets: [...book().sheets, { name: 'T', rows: '' }] });
    expect(s.transact([{ op: 'renameSheet', sheet: 'T', name: 'Other' }]).ok).toBe(true);
    expect(fullCalc(s)).toBeUndefined();
    expect(s.transact([{ op: 'renameSheet', sheet: 'Other', name: 'LATER' }]).ok).toBe(true);
    expect(fullCalc(s)).toBe('1');
  });

  test('a defined name that names the new sheet marks them stale too', () => {
    const s = storeOf({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 0, { f: 'SUM(1)' })]) }],
      definedNames: '<definedName name="x">Later!$A$1</definedName>',
    });
    expect(s.transact([{ op: 'addSheet', name: 'Later', index: 1 }]).ok).toBe(true);
    expect(fullCalc(s)).toBe('1');
  });
});

describe('a sheet op and a function that reads sheet names or positions', () => {
  const stale = (s: WorkbookDocumentStore): [boolean, boolean] => {
    const index = workbookIndex(s.package)!;
    const part = xmlPartOf(s.package, index.sheets[0]!.partName)!;
    return [index.fullCalcOnLoad, sheetStale(index, part.root)];
  };
  const book = (f: string): FixtureWorkbook => ({
    sheets: [
      { name: 'S', rows: row(1, [c('A1', 5), c('B1', 1, { f })]) },
      { name: 'T', rows: row(1, [c('A1', 2)]) },
    ],
  });
  const ops = [
    { op: 'addSheet', name: 'N', index: 2 },
    { op: 'renameSheet', sheet: 'T', name: 'U' },
    { op: 'moveSheet', sheet: 'T', index: 0 },
  ] as const;

  for (const op of ops)
    test(`${op.op} marks the cached results stale for each of the functions`, () => {
      for (const f of [
        'INDIRECT("T!A1")',
        'CELL("address",A1)',
        'INFO("numfile")',
        '_xlfn.SHEET()',
        '_xlfn.SHEETS()',
        '_xlfn.FORMULATEXT(C1)',
      ]) {
        const s = storeOf(book(f));
        expect(stale(s)).toEqual([false, false]);
        expect(s.transact([op]).ok).toBe(true);
        expect([f, ...stale(s)]).toEqual([f, true, true]);
      }
    });

  test('a defined name that calls one marks them stale too; other formulas do not', () => {
    const named = storeOf({
      ...book('SUM(1)'),
      definedNames: '<definedName name="here">_xlfn.SHEET(S!A1)</definedName>',
    });
    expect(named.transact([ops[2]]).ok).toBe(true);
    expect(stale(named)).toEqual([true, true]);
    // A name or a string that only holds the letters of a function is not a call.
    for (const op of ops) {
      const s = storeOf(book('LEN("INDIRECT")+SUM(S!A1)'));
      expect(s.transact([op]).ok).toBe(true);
      expect(stale(s)).toEqual([false, false]);
    }
  });

  test('a defined name that calls the macro function GET.WORKBOOK marks them stale', () => {
    const named = storeOf({
      ...book('SUM(1)'),
      definedNames: '<definedName name="sheets">GET.WORKBOOK(1)</definedName>',
    });
    expect(named.transact([ops[1]]).ok).toBe(true);
    expect(stale(named)).toEqual([true, true]);
  });

  test('a column width marks them stale when a formula or a defined name calls CELL', () => {
    const width = { op: 'setColumnWidth', sheet: 'S', first: 1, last: 1, width: 20 } as const;
    const s = storeOf(book('CELL("width",A1)'));
    expect(s.transact([width]).ok).toBe(true);
    expect(stale(s)).toEqual([true, true]);
    const band = storeOf(book('CELL("width",A1)'));
    const widths = {
      op: 'setColumnWidths',
      sheet: 'S',
      columns: [{ column: 1, width: 20 }],
    } as const;
    expect(band.transact([widths]).ok).toBe(true);
    expect(stale(band)).toEqual([true, true]);
    const named = storeOf({
      ...book('SUM(1)'),
      definedNames: '<definedName name="w">CELL("width",S!A1)</definedName>',
    });
    expect(named.transact([width]).ok).toBe(true);
    expect(stale(named)).toEqual([true, true]);
    // Another function, and a row height, keep the values current.
    const other = storeOf(book('INDIRECT("T!A1")'));
    expect(other.transact([width]).ok).toBe(true);
    const height = { op: 'setRowHeight', sheet: 'S', first: 1, last: 1, height: 30 } as const;
    const tall = storeOf(book('CELL("width",A1)'));
    expect(tall.transact([height]).ok).toBe(true);
    expect([stale(other), stale(tall)]).toEqual([
      [false, false],
      [false, false],
    ]);
  });

  test('every macro function GET.* marks them stale for a sheet op', () => {
    for (const call of ['GET.DOCUMENT(1)', 'GET.CELL(17,S!A1)', 'get.workbook(1)']) {
      const named = storeOf({
        ...book('SUM(1)'),
        definedNames: `<definedName name="m">${call}</definedName>`,
      });
      expect(named.transact([ops[1]]).ok).toBe(true);
      expect(stale(named)).toEqual([true, true]);
    }
  });

  test('a GET.* macro marks them stale for a column width, a row height, and a fitted row height', () => {
    const width = { op: 'setColumnWidth', sheet: 'S', first: 1, last: 1, width: 20 } as const;
    const height = { op: 'setRowHeight', sheet: 'S', first: 1, last: 1, height: 30 } as const;
    const fitted = {
      op: 'setFittedRowHeights',
      sheet: 'S',
      rows: [{ row: 1, height: 30 }],
    } as const;
    for (const op of [width, height, fitted]) {
      const named = storeOf({
        ...book('SUM(1)'),
        definedNames: '<definedName name="h">GET.CELL(17,S!A1)</definedName>',
      });
      expect(named.transact([op]).ok).toBe(true);
      expect(stale(named)).toEqual([true, true]);
    }
    // A name that only starts with the letters of the prefix is not a macro call.
    const plain = storeOf(book('GETPIVOTDATA("x",S!A1)'));
    expect(plain.transact([height]).ok).toBe(true);
    expect(stale(plain)).toEqual([false, false]);
  });
});
