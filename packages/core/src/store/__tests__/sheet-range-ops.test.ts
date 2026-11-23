// Sheet ranges (`Jan:Mar!A1`) and the sheet ops of design D1 of the change
// `workbook-structure-editing`: a move is refused only when it changes the sheets between the
// ends of a range, and a delete only when it removes an end. A sheet between the ends needs no
// text change. Content that the op does not read refuses any move.

import { describe, expect, test } from 'bun:test';
import { c, row } from '../../../../../scripts/fixtures/builder.ts';
import { workbookIndex } from '../workbook/workbook-index.ts';
import type { XlsxPackage } from '../package/package.ts';
import type { WorkbookDocumentStore } from '../store/document-store.ts';
import type { WorkbookOp } from '../store/ops.ts';
import { definedNames, formulaOf, storeOf } from './structure-helpers.ts';

const names = (pkg: XlsxPackage) => workbookIndex(pkg)!.sheets.map((s) => s.name);

const rejection = (s: WorkbookDocumentStore, ops: readonly WorkbookOp[]) => {
  const before = s.package;
  const r = s.transact(ops);
  expect(s.package).toBe(before);
  return r.ok ? null : { code: r.rejection.code, subject: r.rejection.subject };
};

const book = (extra = {}) =>
  storeOf({
    sheets: [
      { name: 'Jan', rows: row(1, [c('A1', 1)]) },
      { name: 'Feb', rows: row(1, [c('A1', 2)]) },
      { name: 'Mar', rows: row(1, [c('A1', 3)]) },
      { name: 'Sum', rows: row(1, [c('A1', 6, { f: 'SUM(Jan:Mar!A1)' })]) },
      { name: 'Notes', rows: '' },
    ],
    definedNames: '<definedName name="q1">Jan:Mar!$A$1</definedName>',
    ...extra,
  });

const move = (sheet: string, index: number): WorkbookOp[] => [{ op: 'moveSheet', sheet, index }];
const blocked = { code: 'sheet-range-blocked', subject: 'Sum!A1' } as const;

describe('a move and a sheet range', () => {
  test('a move outside the range, or inside it, is applied and keeps the text', () => {
    const s = book();
    expect(s.transact(move('Notes', 0)).ok).toBe(true);
    expect(s.transact(move('Sum', 0)).ok).toBe(true);
    expect(names(s.package)).toEqual(['Sum', 'Notes', 'Jan', 'Feb', 'Mar']);
    expect(formulaOf(s.package, 'Sum', 'A1')).toBe('SUM(Jan:Mar!A1)');
    expect(definedNames(s.package).q1).toBe('Jan:Mar!$A$1');
  });

  test('a swap of the two ends is refused, though it spans the same sheets', () => {
    const s = storeOf({
      sheets: [
        { name: 'Jan', rows: '' },
        { name: 'Feb', rows: '' },
        { name: 'Sum', rows: row(1, [c('A1', 0, { f: 'SUM(Jan:Feb!A1)' })]) },
      ],
    });
    expect(rejection(s, move('Jan', 1))).toEqual({
      code: 'sheet-range-blocked',
      subject: 'Sum!A1',
    });
    expect(rejection(s, move('Feb', 0))).toEqual({
      code: 'sheet-range-blocked',
      subject: 'Sum!A1',
    });
    expect(names(s.package)).toEqual(['Jan', 'Feb', 'Sum']);
  });

  test('a move into the range, out of it, or of an end past another sheet is refused', () => {
    const s = book();
    expect(rejection(s, move('Notes', 1))).toEqual(blocked);
    expect(rejection(s, move('Feb', 0))).toEqual(blocked);
    expect(rejection(s, move('Feb', 4))).toEqual(blocked);
    expect(rejection(s, move('Mar', 1))).toEqual(blocked);
    expect(rejection(s, move('Jan', 2))).toEqual(blocked);
  });

  test('a defined name with a range refuses the move with its subject', () => {
    const s = storeOf({
      sheets: [
        { name: 'Jan', rows: '' },
        { name: 'Feb', rows: '' },
        { name: 'Notes', rows: '' },
      ],
      definedNames: '<definedName name="q1">Jan:Feb!$A$1</definedName>',
    });
    expect(rejection(s, move('Notes', 1))).toEqual({
      code: 'sheet-range-blocked',
      subject: 'x:definedName q1',
    });
  });

  test('a range in content that the op does not read refuses any move', () => {
    const dv =
      '<dataValidations count="1"><dataValidation type="list" sqref="B1">' +
      '<formula1>Jan:Mar!$A$1</formula1></dataValidation></dataValidations>';
    const s = storeOf({
      sheets: [
        { name: 'Jan', rows: '' },
        { name: 'Mar', rows: '' },
        { name: 'Notes', rows: row(1, [c('A1', 1)]), afterMerges: dv },
      ],
    });
    expect(rejection(s, move('Notes', 0))?.code).toBe('sheet-range-blocked');
  });
});

describe('a delete and a sheet range', () => {
  test('a delete of a sheet between the ends keeps the text', () => {
    const s = book();
    expect(s.transact([{ op: 'deleteSheet', sheet: 'Feb' }]).ok).toBe(true);
    expect(formulaOf(s.package, 'Sum', 'A1')).toBe('SUM(Jan:Mar!A1)');
    expect(definedNames(s.package).q1).toBe('Jan:Mar!$A$1');
  });

  test('a delete of a sheet outside the range keeps the text', () => {
    const s = book();
    expect(s.transact([{ op: 'deleteSheet', sheet: 'Notes' }]).ok).toBe(true);
    expect(formulaOf(s.package, 'Sum', 'A1')).toBe('SUM(Jan:Mar!A1)');
  });

  test('a delete of an end is refused', () => {
    const s = book();
    expect(rejection(s, [{ op: 'deleteSheet', sheet: 'Jan' }])).toEqual(blocked);
    expect(rejection(s, [{ op: 'deleteSheet', sheet: 'Mar' }])).toEqual(blocked);
  });
});

describe('a sheet range of this workbook with the prefix [0]', () => {
  const zero = (f: string) =>
    storeOf({
      sheets: [
        { name: 'Jan', rows: row(1, [c('A1', 1)]) },
        { name: 'Feb', rows: row(1, [c('A1', 2)]) },
        { name: 'Mar', rows: row(1, [c('A1', 3)]) },
        { name: 'Sum', rows: row(1, [c('A1', 6, { f })]) },
      ],
    });

  test('a move into the range or a swap of its ends is refused, as for a local range', () => {
    for (const f of ['SUM([0]Jan:Mar!A1)', "SUM('[0]Jan:Mar'!A1)"]) {
      const s = zero(f);
      expect(rejection(s, move('Sum', 1))).toEqual(blocked);
      expect(rejection(s, move('Mar', 0))).toEqual(blocked);
      expect(formulaOf(s.package, 'Sum', 'A1')).toBe(f);
    }
  });

  test('a delete of an end is refused with sheet-range-blocked, as for a local range', () => {
    const s = zero('SUM([0]Jan:Mar!A1)');
    expect(rejection(s, [{ op: 'deleteSheet', sheet: 'Jan' }])).toEqual(blocked);
    expect(rejection(s, [{ op: 'deleteSheet', sheet: 'Mar' }])).toEqual(blocked);
  });

  test('a move outside the range is applied and keeps the text', () => {
    const s = zero('SUM([0]Jan:Mar!A1)');
    expect(s.transact(move('Sum', 0)).ok).toBe(true);
    expect(names(s.package)).toEqual(['Sum', 'Jan', 'Feb', 'Mar']);
    expect(formulaOf(s.package, 'Sum', 'A1')).toBe('SUM([0]Jan:Mar!A1)');
  });
});

describe('an external sheet range', () => {
  // `[1]Jan:Mar!A1` names sheets of another workbook: the local `Jan` and `Mar` do not count.
  const ext = () =>
    storeOf({
      sheets: [
        { name: 'Jan', rows: row(1, [c('A1', 1)]) },
        { name: 'Feb', rows: '' },
        { name: 'Mar', rows: '' },
        {
          name: 'Sum',
          rows: row(1, [
            c('A1', 0, { f: 'SUM([1]Jan:Mar!A1)' }),
            c('B1', 0, { f: '[1]Jan!A1+[1]Jan!B2' }),
          ]),
        },
        { name: 'Notes', rows: '' },
      ],
    });
  const kept = (s: WorkbookDocumentStore) => {
    expect(formulaOf(s.package, 'Sum', 'A1')).toBe('SUM([1]Jan:Mar!A1)');
    expect(formulaOf(s.package, 'Sum', 'B1')).toBe('[1]Jan!A1+[1]Jan!B2');
  };

  test('a rename of a local sheet with the same name keeps the text', () => {
    const s = ext();
    expect(s.transact([{ op: 'renameSheet', sheet: 'Jan', name: 'Q1' }]).ok).toBe(true);
    kept(s);
  });

  test('a delete of a local end keeps the text', () => {
    const s = ext();
    expect(s.transact([{ op: 'deleteSheet', sheet: 'Jan' }]).ok).toBe(true);
    kept(s);
  });

  test('a move into or past the local ends is applied', () => {
    const s = ext();
    expect(s.transact(move('Notes', 1)).ok).toBe(true);
    expect(s.transact(move('Mar', 0)).ok).toBe(true);
    kept(s);
  });

  test('a row insert does not move the external references', () => {
    const s = ext();
    expect(s.transact([{ op: 'insertRows', sheet: 'Jan', start: 1, count: 1 }]).ok).toBe(true);
    kept(s);
  });
});
