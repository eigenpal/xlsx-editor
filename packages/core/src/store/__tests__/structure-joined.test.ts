// A range that the range operator builds from two cells (`Data!A1:Data!A4`) moves as one span
// in a structural edit; an end of a range built from other operands refuses the edit when the
// edit deletes it. A move that makes a formula or a defined name longer than the tokenizer
// reads refuses the edit, and names the cell or the defined name.

import { describe, expect, test } from 'bun:test';
import { c, row } from '../../../../../scripts/fixtures/builder.ts';
import { MAX_FORMULA_LENGTH } from '../formula/tokenizer.ts';
import { moveReferences, type ReferenceContext } from '../structure/references.ts';
import type { AxisEdit } from '../structure/axis.ts';
import { definedNames, formulaOf, op, storeOf } from './structure-helpers.ts';

const del = (at: number, count = 1): AxisEdit => ({ axis: 'rows', kind: 'delete', at, count });
const ctx = (edit: AxisEdit, host: string | undefined = 'DATA'): ReferenceContext => ({
  edit,
  sheet: 'DATA',
  host,
  order: new Map([
    ['DATA', 0],
    ['OTHER', 1],
  ]),
});
const moved = (text: string, edit: AxisEdit, host?: string) => {
  const m = moveReferences(text, ctx(edit, host));
  return m.ok ? m.text : m.reason;
};

describe('two cells joined by the range operator', () => {
  test('a qualified range shrinks when the edit deletes one end', () => {
    expect(moved('SUM(Data!A1:Data!A4)', del(3), 'OTHER')).toBe('SUM(Data!A1:Data!A3)');
    expect(moved('SUM(Data!A1:Data!A4)', del(0), 'OTHER')).toBe('SUM(Data!A1:Data!A3)');
    expect(moved('SUM(A4:A1)', del(3))).toBe('SUM(A3:A1)');
    expect(moved('SUM(Data!A2 : Data!A4)', del(0), 'OTHER')).toBe('SUM(Data!A1 : Data!A3)');
  });

  test('the delete of the whole span makes it #REF!', () => {
    expect(moved('SUM(Data!A2:Data!A3)', del(1, 2), 'OTHER')).toBe('SUM(Data!#REF!)');
  });

  test('a range on another sheet does not move', () => {
    expect(moved('SUM(Other!A1:Other!A4)', del(0))).toBe('SUM(Other!A1:Other!A4)');
  });

  test('an end built from other operands refuses only a delete of that end', () => {
    expect(moved('SUM(INDEX(A:A,1):A10)', del(9))).toBe(
      'the reference A10 cannot be moved with certainty'
    );
    expect(moved('SUM(INDEX(A:A,1):A10)', del(2))).toBe('SUM(INDEX(A:A,1):A9)');
    expect(moved('SUM(A10:INDEX(A:A,20))', del(9))).toBe(
      'the reference A10 cannot be moved with certainty'
    );
  });
});

describe('a move past the length of a formula', () => {
  // `A2` (2 characters) becomes `#REF!` (5 characters) when row 2 is deleted.
  const long = `"${'x'.repeat(MAX_FORMULA_LENGTH - 6)}"&A2`;

  test('a formula that would grow too long refuses the edit, and names the cell', () => {
    expect(long.length).toBeLessThanOrEqual(MAX_FORMULA_LENGTH);
    const store = storeOf({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) + row(5, [c('B5', 0, { f: long })]) }],
    });
    const r = store.transact([op('deleteRows', 'S', 2)]);
    expect(r.ok ? 'ok' : r.rejection.code).toBe('structure-blocked');
    if (!r.ok) {
      expect(r.rejection.detail).toContain('B5');
      expect(r.rejection.detail).toContain(`longer than ${MAX_FORMULA_LENGTH}`);
    }
    expect(formulaOf(store.package, 'S', 'B5')).toBe(long);
  });

  test('a defined name that would grow too long refuses the edit, and names it', () => {
    const text = `"${'x'.repeat(MAX_FORMULA_LENGTH - 8)}"&S!A2`;
    expect(text.length).toBeLessThanOrEqual(MAX_FORMULA_LENGTH);
    const store = storeOf({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }],
      definedNames: `<definedName name="Big">${text.replaceAll('&', '&amp;')}</definedName>`,
    });
    const r = store.transact([op('deleteRows', 'S', 2)]);
    expect(r.ok ? 'ok' : r.rejection.code).toBe('structure-blocked');
    if (!r.ok) expect(r.rejection.detail).toContain('Big');
    expect(definedNames(store.package).Big).toBe(text);
  });
});
