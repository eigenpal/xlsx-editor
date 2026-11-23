// Edits that remove a cell or a row keep the positions of the cells and rows after it that have
// no `@r`: a cell without `@r` follows the one before it, and a row without `@r` is the row after
// the one before it. Each path pins the next ones with an explicit `@r`, and the saved, reopened,
// undone, and redone packages hold every value at its cell.

import { describe, expect, test } from 'bun:test';
import { buildXlsx } from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx, xmlPartOf, type XlsxPackage } from '../package/package.ts';
import { child, textOf } from '../package/query.ts';
import { formatCellRef } from '../cell-ref.ts';
import { sheetIndex } from '../workbook/sheet-index.ts';
import { saveXlsx } from '../package-write/save.ts';
import { WorkbookDocumentStore } from '../store/document-store.ts';
import type { WorkbookOp } from '../store/ops.ts';

function storeOf(rows: string): WorkbookDocumentStore {
  const read = readXlsx(buildXlsx({ sheets: [{ name: 'S', rows }] }));
  if (!read.ok) throw new Error(JSON.stringify(read.error));
  return new WorkbookDocumentStore(read.pkg);
}

/** Each cell that holds a value, as `A1=1`, in sheet order. */
function values(pkg: XlsxPackage): string[] {
  const part = xmlPartOf(pkg, '/xl/worksheets/sheet1.xml')!;
  const out: string[] = [];
  for (const e of sheetIndex(part).cells.values()) {
    const v = child(e.el, 'x:v');
    if (v) out.push(`${formatCellRef({ row: e.row, col: e.col })}=${textOf(v)}`);
  }
  return out.sort();
}

function check(rows: string, ops: WorkbookOp[], after: string[]) {
  const store = storeOf(rows);
  const before = values(store.package);
  const r = store.transact(ops);
  expect(r.ok ? 'ok' : r.rejection).toBe('ok');
  expect(values(store.package)).toEqual(after);
  const reopened = readXlsx(saveXlsx(store.package));
  if (!reopened.ok) throw new Error('reopen failed');
  expect(values(reopened.pkg)).toEqual(after);
  expect(store.undo().ok).toBe(true);
  expect(values(store.package)).toEqual(before);
  expect(store.redo().ok).toBe(true);
  expect(values(store.package)).toEqual(after);
  // Every later edit still applies.
  const more = store.transact([
    { op: 'setCellValue', sheet: 'S', ref: 'D9', value: { type: 'number', value: 9 } },
  ]);
  expect(more.ok).toBe(true);
}

const set = (ref: string, value: number | null): WorkbookOp => ({
  op: 'setCellValue',
  sheet: 'S',
  ref,
  value: value === null ? { type: 'empty' } : { type: 'number', value },
});
const CELLS = '<row r="1"><c r="A1"><v>1</v></c><c><v>2</v></c><c><v>3</v></c></row>';
const ROWS =
  '<row r="1"><c r="A1"><v>1</v></c></row><row><c r="A2"><v>2</v></c></row><row><c><v>3</v></c></row>';

describe('removing a cell keeps the columns of the cells after it without @r', () => {
  test('one cleared value', () => check(CELLS, [set('A1', null)], ['B1=2', 'C1=3']));
  test('a batch of values', () =>
    check(CELLS, [set('A1', null), set('E1', 5)], ['B1=2', 'C1=3', 'E1=5']));
  test('clearValues', () =>
    check(CELLS, [{ op: 'clearValues', sheet: 'S', ref: 'A1:B1' }], ['C1=3']));
  test('clearFormats of a cell that holds only a format', () =>
    check(
      '<row r="1"><c r="A1" s="0"/><c><v>2</v></c></row>',
      [{ op: 'clearFormats', sheet: 'S', ref: 'A1' }],
      ['B1=2']
    ));
});

describe('removing a row keeps the numbers of the rows after it without @r', () => {
  test('one cleared value', () => check(ROWS, [set('A1', null)], ['A2=2', 'A3=3']));
  test('a batch of values', () =>
    check(ROWS, [set('A1', null), set('C5', 5)], ['A2=2', 'A3=3', 'C5=5']));
  test('clearValues', () =>
    check(ROWS, [{ op: 'clearValues', sheet: 'S', ref: 'A1:A2' }], ['A3=3']));
});

describe('a row number with a leading zero', () => {
  // `x:row/@r` is an xsd:unsignedInt, so `05` is row 5. Undo compares numbers, not strings.
  test('edits, undo, and redo of row `05` hold every value at its cell', () => {
    const store = storeOf('<row r="05"><c r="A5"><v>1</v></c></row>');
    const steps: WorkbookOp[][] = [
      [set('B5', 2)],
      [set('C5', 3), set('D5', 4)],
      [{ op: 'clearValues', sheet: 'S', ref: 'A5:Z5' }],
    ];
    const states = [values(store.package)];
    for (const ops of steps) {
      const r = store.transact(ops);
      expect(r.ok ? 'ok' : r.rejection).toBe('ok');
      states.push(values(store.package));
    }
    expect(states.map((s) => s.join())).toEqual(['A5=1', 'A5=1,B5=2', 'A5=1,B5=2,C5=3,D5=4', '']);
    for (let i = steps.length - 1; i >= 0; i -= 1) {
      expect(store.undo().ok).toBe(true);
      expect(values(store.package)).toEqual(states[i]!);
    }
    for (let i = 1; i <= steps.length; i += 1) {
      expect(store.redo().ok).toBe(true);
      expect(values(store.package)).toEqual(states[i]!);
    }
  });
});
