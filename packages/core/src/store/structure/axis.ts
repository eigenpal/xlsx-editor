// The position rules of a structural edit (design D4): where a row or column index, and a span
// of them, goes when rows or columns are inserted or deleted. Every structural element moves by
// these rules, so a cell, a merged range, a formula reference, and a drawing anchor agree.
//
// - Insert `count` at `at`: an index at or after `at` moves by `count`. A span that starts
//   before `at` and ends at or after it grows. A span end pushed past the grid is cut at the
//   last index; a span or an index that leaves the grid as a whole is gone.
// - Delete `count` at `at`: an index inside the deleted band is gone; one after it moves back by
//   `count`. A span that loses some of its indexes shrinks; one that loses all of them is gone.

import { MAX_COLS, MAX_ROWS } from '../cell-ref.ts';

/** One structural edit along one axis of one sheet. Indexes are 0-based. */
export interface AxisEdit {
  readonly axis: 'rows' | 'cols';
  readonly kind: 'insert' | 'delete';
  readonly at: number;
  readonly count: number;
}

/** The number of indexes of an axis in the grid. */
export const axisLimit = (axis: AxisEdit['axis']): number =>
  axis === 'rows' ? MAX_ROWS : MAX_COLS;

/** The new index of `p`, or null when the edit deletes it or pushes it out of the grid. */
export function movePosition(e: AxisEdit, p: number): number | null {
  if (p < e.at) return p;
  if (e.kind === 'insert') {
    const q = p + e.count;
    return q < axisLimit(e.axis) ? q : null;
  }
  return p < e.at + e.count ? null : p - e.count;
}

/** The new span of `[a, b]` (`a <= b`), or null when nothing of it is left in the grid. */
export function moveSpan(e: AxisEdit, a: number, b: number): readonly [number, number] | null {
  if (e.kind === 'insert') {
    const last = axisLimit(e.axis) - 1;
    const a2 = a >= e.at ? a + e.count : a;
    const b2 = b >= e.at ? b + e.count : b;
    if (a2 > last) return null;
    return [a2, Math.min(b2, last)];
  }
  const end = e.at + e.count;
  const a2 = a < e.at ? a : a >= end ? a - e.count : e.at;
  const b2 = b < e.at ? b : b >= end ? b - e.count : e.at - 1;
  return a2 <= b2 ? [a2, b2] : null;
}

/** Whether an insertion pushes index `b` (the end of some content) past the grid. */
export function pushedOut(e: AxisEdit, b: number): boolean {
  return e.kind === 'insert' && b >= e.at && b + e.count >= axisLimit(e.axis);
}

/** Whether an insertion lands strictly inside `[a, b]`, so it would split the span. */
export function splits(e: AxisEdit, a: number, b: number): boolean {
  return e.kind === 'insert' && e.at > a && e.at <= b;
}

/** Whether a deletion removes part, but not all, of `[a, b]`. */
export function cutsPart(e: AxisEdit, a: number, b: number): boolean {
  if (e.kind !== 'delete') return false;
  const end = e.at + e.count - 1;
  const meets = e.at <= b && end >= a;
  const covers = e.at <= a && end >= b;
  return meets && !covers;
}
