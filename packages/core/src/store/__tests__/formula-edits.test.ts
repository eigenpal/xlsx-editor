// Formula edits through the store (tasks 7.2 and 7.3): formula text as `x:f` with no cached
// value, refusal of text the tokenizer cannot read, the shared formula rule of design D4, the
// dropped calculation chain of design D5, and the stale flag of design D3 in the same
// transaction and outside the undo history.

import { describe, expect, test } from 'bun:test';
import {
  buildXlsx,
  c,
  row,
  type FixtureWorkbook,
} from '../../../../../scripts/fixtures/builder.ts';
import {
  contentTypeOf,
  readXlsx,
  relationshipsOf,
  xmlPartOf,
  type XlsxPackage,
} from '../package/package.ts';
import { attr, child, textOf } from '../package/query.ts';
import { serializeXmlPart } from '../package-write/serialize.ts';
import { saveXlsx } from '../package-write/save.ts';
import { sheetByName, workbookIndex } from '../workbook/workbook-index.ts';
import { cellKey, sheetIndex } from '../workbook/sheet-index.ts';
import { parseCellRef } from '../cell-ref.ts';
import { WorkbookDocumentStore, type StoreChange } from '../store/document-store.ts';
import type { CellInput, WorkbookOp } from '../store/ops.ts';

const SHEET = '/xl/worksheets/sheet1.xml';
const CALC_CHAIN_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml';
const CALC_CHAIN_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain';

function open(wb: FixtureWorkbook): XlsxPackage {
  const read = readXlsx(buildXlsx(wb));
  if (!read.ok) throw new Error(JSON.stringify(read.error));
  return read.pkg;
}

function cell(pkg: XlsxPackage, ref: string) {
  const a = parseCellRef(ref)!;
  return sheetIndex(xmlPartOf(pkg, SHEET)!).cells.get(cellKey(a.row, a.col))?.el;
}

const set = (ref: string, value: CellInput): WorkbookOp => ({
  op: 'setCellValue',
  sheet: 'S',
  ref,
  value,
});
const formula = (text: string): CellInput => ({ type: 'formula', text });
const fullCalc = (pkg: XlsxPackage) => workbookIndex(pkg)!.fullCalcOnLoad;
const sheetXml = (pkg: XlsxPackage) => serializeXmlPart(xmlPartOf(pkg, SHEET)!);

describe('formula input', () => {
  test('a range end with the shape of a column is a defined name when the workbook holds one', () => {
    const store = new WorkbookDocumentStore(
      open({
        sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }],
        definedNames: '<definedName name="tax">S!$A$1</definedName>',
      })
    );
    expect(store.transact([set('D1', formula('SUM(tax:B5)'))]).ok).toBe(true);
    const r = store.transact([set('D2', formula('SUM(tay:B5)'))]);
    expect(r.ok ? null : r.rejection.code).toBe('unreadable-formula');
    const r2 = store.transact([set('D3', formula('A1:XFE1'))]);
    expect(r2.ok ? null : r2.rejection.code).toBe('unreadable-formula');
  });

  test('a typed call outside the predefined functions is refused; a file formula keeps its text', () => {
    const store = new WorkbookDocumentStore(
      open({
        sheets: [
          {
            name: 'S',
            rows: row(1, [c('A1', 1), c('B1', 2), '<c r="C1"><f>_xlfn.CONCAT(A1,B1)</f></c>']),
          },
        ],
        definedNames: '<definedName name="Rate">S!$A$1</definedName>',
      })
    );
    for (const text of [
      'LET(x,1,x)',
      'XLOOKUP(1,A1:A2,B1:B2)',
      'IFS(A1>0,1)',
      'CONCAT(A1,B1)',
      'TEXTJOIN(",",TRUE,A1:B1)',
      'FILTER(A1:B1,A1:B1>0)',
      'UNIQUE(A1:B1)',
      '_xlfn.XLOOKUP(1,A1:A2,B1:B2)',
      'SUM(_xlws.SORT(A1:B1))',
    ]) {
      const r = store.transact([set('D1', formula(text))]);
      expect(r.ok ? null : r.rejection.code).toBe('unsupported-function');
    }
    for (const text of [
      'SUM(A1:B1)',
      'IF(A1>0,1,2)',
      'VLOOKUP(1,A1:B1,2,FALSE)',
      'INDEX(A1:B1,1,2)',
      'IFERROR(A1,0)',
      'CONCATENATE(A1,B1)',
      'sum(a1:b1)',
      'ERROR.TYPE(A1)',
      'Rate*2',
    ])
      expect(store.transact([set('D1', formula(text))]).ok).toBe(true);
    const reopened = readXlsx(saveXlsx(store.package));
    if (!reopened.ok) throw new Error('reopen');
    expect(textOf(child(cell(reopened.pkg, 'C1')!, 'x:f'))).toBe('_xlfn.CONCAT(A1,B1)');
  });

  test('writes x:f with no cached value, keeps the format, and undo restores the cell', () => {
    const store = new WorkbookDocumentStore(
      open({ sheets: [{ name: 'S', rows: row(1, ['<c r="A1" s="0" t="str"><v>text</v></c>']) }] })
    );
    const before = sheetXml(store.package);
    expect(store.transact([set('A1', formula('sum(b1:b3)+"a1"'))]).ok).toBe(true);
    const a1 = cell(store.package, 'A1')!;
    expect(textOf(child(a1, 'x:f'))).toBe('SUM(B1:B3)+"a1"');
    expect(child(a1, 'x:v')).toBeUndefined();
    expect(attr(a1, 't')).toBeUndefined();
    expect(attr(a1, 's')).toBe('0');
    expect(store.undo().ok).toBe(true);
    expect(sheetXml(store.package)).toBe(before);
  });

  test('text that the tokenizer cannot read is refused with a typed reason', () => {
    const store = new WorkbookDocumentStore(open({ sheets: [{ name: 'S', rows: '' }] }));
    // A range operator between operands that are not references commits nothing either.
    for (const text of [
      'SUM(A1',
      '"open',
      '',
      '1;2',
      'A1 @ B1',
      '"a":"b"',
      '1.5:2',
      'A1:1',
      'SUM(0:3)',
      'SUM(1:1048577)',
      'TRUE:FALSE',
      '{1,2}:A1',
    ]) {
      const r = store.transact([set('A1', formula(text))]);
      expect(r.ok ? null : r.rejection.code).toBe('unreadable-formula');
    }
    expect(store.revision).toBe(0);
    expect(store.canUndo).toBe(false);
  });
});

/** `B1` is the master of `=A1*2` shared over `B1:B5`; the children hold cached values. */
const shared = (masterText = 'A1*2'): FixtureWorkbook => ({
  sheets: [
    {
      name: 'S',
      rows: [1, 2, 3, 4, 5]
        .map((r) =>
          row(r, [
            c(`A${r}`, r),
            r === 1
              ? `<c r="B1"><f t="shared" ref="B1:B5" si="0">${masterText}</f><v>2</v></c>`
              : `<c r="B${r}"><f t="shared" si="0"/><v>${r * 2}</v></c>`,
          ])
        )
        .join(''),
    },
  ],
});

describe('the shared formula rule of design D4', () => {
  test('overwriting the master writes each child as its own formula', () => {
    const store = new WorkbookDocumentStore(open(shared()));
    const before = sheetXml(store.package);
    expect(store.transact([set('B1', { type: 'number', value: 7 })]).ok).toBe(true);
    const pkg = store.package;
    expect(child(cell(pkg, 'B1')!, 'x:f')).toBeUndefined();
    expect(textOf(child(cell(pkg, 'B1')!, 'x:v'))).toBe('7');
    for (const r of [2, 3, 4, 5]) {
      const f = child(cell(pkg, `B${r}`)!, 'x:f')!;
      expect(textOf(f)).toBe(`A${r}*2`);
      expect(attr(f, 't')).toBeUndefined();
      expect(attr(f, 'si')).toBeUndefined();
      expect(textOf(child(cell(pkg, `B${r}`)!, 'x:v'))).toBe(String(r * 2));
    }
    // The saved part has no shared formula child without a master.
    expect(sheetXml(pkg)).not.toContain('shared');
    const reopened = readXlsx(saveXlsx(pkg));
    expect(reopened.ok).toBe(true);
    // One undo step restores the master and every child.
    expect(store.undo().ok).toBe(true);
    expect(sheetXml(store.package)).toBe(before);
    expect(store.redo().ok).toBe(true);
    expect(textOf(child(cell(store.package, 'B4')!, 'x:f'))).toBe('A4*2');
  });

  test('clearing the master unshares too, and absolute parts stay', () => {
    const store = new WorkbookDocumentStore(open(shared('$A$1+A1')));
    expect(store.transact([set('B1', { type: 'empty' })]).ok).toBe(true);
    expect(textOf(child(cell(store.package, 'B3')!, 'x:f'))).toBe('$A$1+A3');
  });

  test('a master whose text cannot be read refuses the edit', () => {
    const store = new WorkbookDocumentStore(open(shared('A1+"open')));
    const r = store.transact([set('B1', { type: 'number', value: 7 })]);
    expect(r.ok ? null : r.rejection.code).toBe('shared-formula');
    // The subject names the shared range, not only the master cell; the cause is the master.
    expect(r.ok ? null : r.rejection.subject).toMatch(/!B1:B5$/);
    expect(r.ok ? null : r.rejection.causeCell).toBe('S!B1');
    expect(r.ok ? null : r.rejection.causeUnreadable).toBe(true);
    expect(r.ok ? null : r.rejection.masterCell).toBe('S!B1');
    expect(store.revision).toBe(0);
  });

  test('a child whose formula would move out of the grid refuses the edit and is its cause', () => {
    // B2 moves the reference to the last row; B3 would move it past the grid.
    const store = new WorkbookDocumentStore(open(shared('A1048575')));
    const r = store.transact([set('B1', { type: 'number', value: 7 })]);
    expect(r.ok ? null : [r.rejection.code, r.rejection.subject, r.rejection.causeCell]).toEqual([
      'shared-formula',
      'S!B1:B5',
      'S!B3',
    ]);
    expect(r.ok ? null : [r.rejection.causeUnreadable, r.rejection.masterCell]).toEqual([
      undefined,
      'S!B1',
    ]);
    expect(store.revision).toBe(0);
  });

  test('editing a child leaves the master and the other children shared', () => {
    const store = new WorkbookDocumentStore(open(shared()));
    expect(store.transact([set('B3', { type: 'number', value: 1 })]).ok).toBe(true);
    expect(attr(child(cell(store.package, 'B1')!, 'x:f'), 't')).toBe('shared');
    expect(attr(child(cell(store.package, 'B4')!, 'x:f'), 't')).toBe('shared');
  });
});

/** A workbook with a formula and a calculation chain. */
const chained = (): FixtureWorkbook => ({
  sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 2, { f: 'A1*2' })]) }],
  extraParts: [
    [
      'xl/calcChain.xml',
      CALC_CHAIN_TYPE,
      new TextEncoder().encode(
        '<?xml version="1.0" encoding="UTF-8"?><calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="B1" i="1"/></calcChain>'
      ),
      CALC_CHAIN_REL,
    ],
  ],
});

const hasChain = (pkg: XlsxPackage) => ({
  part: pkg.parts.has('/xl/calcChain.xml'),
  rel: relationshipsOf(pkg, '/xl/workbook.xml').some((r) => r.type === CALC_CHAIN_REL),
  type: contentTypeOf(pkg, '/xl/calcChain.xml') === CALC_CHAIN_TYPE,
});

describe('the calculation chain of design D5', () => {
  test('a value edit that touches no formula keeps the chain', () => {
    const store = new WorkbookDocumentStore(open(chained()));
    store.transact([set('C1', { type: 'number', value: 3 })]);
    expect(hasChain(store.package)).toEqual({ part: true, rel: true, type: true });
  });

  for (const [what, op] of [
    ['adding a formula', set('C1', formula('A1+1'))],
    ['removing a formula', set('B1', { type: 'number', value: 5 })],
    ['changing a formula', set('B1', formula('A1*3'))],
  ] as const)
    test(`${what} drops the part, its relationship, and its content type`, () => {
      const store = new WorkbookDocumentStore(open(chained()));
      const seen: StoreChange[] = [];
      store.subscribe((change) => seen.push(change));
      expect(store.transact([op]).ok).toBe(true);
      expect(hasChain(store.package)).toEqual({ part: false, rel: false, type: false });
      const expected: string[] = ['setCellValue', 'dropCalcChain'];
      if (fullCalc(store.package)) expected.push('setFullCalcOnLoad');
      expect(seen.map((s) => s.ops.map((o) => o.op as string))).toEqual([expected]);
      // Undo restores the cell, not the stale chain.
      expect(store.undo().ok).toBe(true);
      expect(hasChain(store.package).part).toBe(false);
      const saved = readXlsx(saveXlsx(store.package));
      expect(saved.ok && hasChain(saved.pkg)).toEqual({ part: false, rel: false, type: false });
    });
});

describe('the stale flag of design D3', () => {
  test('the first value edit in a workbook with a formula sets it in the same revision', () => {
    const store = new WorkbookDocumentStore(open(chained()));
    expect(fullCalc(store.package)).toBe(false);
    expect(store.transact([set('A1', { type: 'number', value: 4 })]).ok).toBe(true);
    expect(store.revision).toBe(1);
    expect(fullCalc(store.package)).toBe(true);
    // The undo history holds the value edit only, and undo keeps the flag.
    expect(store.undo().ok).toBe(true);
    expect(store.canUndo).toBe(false);
    expect(textOf(child(cell(store.package, 'A1')!, 'x:v'))).toBe('1');
    expect(fullCalc(store.package)).toBe(true);
    // A save writes it, and the reopened workbook is stale.
    const saved = readXlsx(saveXlsx(store.package));
    expect(saved.ok && workbookIndex(saved.pkg)!.fullCalcOnLoad).toBe(true);
  });

  test('a workbook with no formula after the edit is not marked', () => {
    const store = new WorkbookDocumentStore(
      open({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] })
    );
    store.transact([set('A1', { type: 'number', value: 2 })]);
    expect(fullCalc(store.package)).toBe(false);
    store.transact([set('B1', formula('A1'))]);
    expect(fullCalc(store.package)).toBe(true);
  });

  test('a new x:calcPr goes in its schema position', () => {
    const pkg = open({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 1, { f: '1' })]) }],
      definedNames: '<definedName name="Rate">S!$A$1</definedName>',
    });
    const wb = xmlPartOf(pkg, '/xl/workbook.xml')!;
    const root = {
      ...wb.root,
      children: wb.root.children.filter((n) => !(n.type === 'element' && n.local === 'calcPr')),
    };
    const parts = new Map(pkg.parts);
    parts.set(wb.partName, { kind: 'xml', partName: wb.partName, xml: { ...wb, root } });
    const store = new WorkbookDocumentStore({ parts });
    store.transact([set('B1', { type: 'number', value: 1 })]);
    const names = xmlPartOf(store.package, '/xl/workbook.xml')!.root.children.flatMap((n) =>
      n.type === 'element' ? [n.local] : []
    );
    expect(names).toEqual(['workbookPr', 'bookViews', 'sheets', 'definedNames', 'calcPr']);
    expect(sheetByName(workbookIndex(store.package)!, 'S')).toBeDefined();
  });
});

describe('clearValues', () => {
  const clear = (ref: string): WorkbookOp => ({ op: 'clearValues', sheet: 'S', ref });

  test('clears values and formulas in the range, keeps formats, and undoes in one step', () => {
    const store = new WorkbookDocumentStore(
      open({
        sheets: [
          {
            name: 'S',
            rows:
              row(1, ['<c r="A1" s="0"><v>1</v></c>', c('B1', 2, { f: 'A1*2' }), c('C1', 3)]) +
              row(2, [c('A2', 'x', { t: 'str' })]),
          },
        ],
      })
    );
    const before = sheetXml(store.package);
    expect(store.transact([clear('B2:A1')]).ok).toBe(true);
    const a1 = cell(store.package, 'A1')!;
    expect([attr(a1, 's'), child(a1, 'x:v')]).toEqual(['0', undefined]);
    expect(cell(store.package, 'B1')).toBeUndefined();
    expect(cell(store.package, 'A2')).toBeUndefined();
    expect(textOf(child(cell(store.package, 'C1')!, 'x:v'))).toBe('3');
    expect(store.undo().ok).toBe(true);
    expect(sheetXml(store.package)).toBe(before);
  });

  test('a whole-sheet range costs the cells that exist', () => {
    const store = new WorkbookDocumentStore(
      open({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) + row(900000, [c('Z900000', 2)]) }] })
    );
    const start = performance.now();
    expect(store.transact([clear('A1:XFD1048576')]).ok).toBe(true);
    expect(performance.now() - start).toBeLessThan(200);
    expect(sheetIndex(xmlPartOf(store.package, SHEET)!).cells.size).toBe(0);
  });

  test('a range that meets an array formula range is refused', () => {
    const store = new WorkbookDocumentStore(
      open({
        sheets: [
          {
            name: 'S',
            rows: row(1, [c('A1', 1, { f: 'B1:B2' })]).replace('<f>', '<f t="array" ref="A1:A2">'),
          },
        ],
      })
    );
    const r = store.transact([clear('A2:C9')]);
    expect(r.ok ? null : r.rejection.code).toBe('shared-formula');
  });

  test('a shared master in the range is unshared, so children outside keep their formulas', () => {
    const store = new WorkbookDocumentStore(open(shared()));
    expect(store.transact([clear('B1:B2')]).ok).toBe(true);
    expect(cell(store.package, 'B2')).toBeUndefined();
    expect(textOf(child(cell(store.package, 'B3')!, 'x:f'))).toBe('A3*2');
    expect(sheetXml(store.package)).not.toContain('shared');
  });

  test('a range with nothing to clear changes nothing', () => {
    const store = new WorkbookDocumentStore(open(shared()));
    const before = store.package;
    expect(store.transact([clear('H1:H9')]).ok).toBe(true);
    expect(store.package).toBe(before);
    expect([store.revision, store.canUndo]).toEqual([0, false]);
  });
});
