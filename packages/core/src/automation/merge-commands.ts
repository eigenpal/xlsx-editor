// The merge commands, DOM-free: merge the range of a selection into one cell, or unmerge the
// merged ranges that the selection meets, lowered to store ops. The editor calls this lowering.
//
// A merge keeps the value of the top-left cell only (`x:mergeCell`, `CT_MergeCell`, ECMA-376
// Part 1, 18.3.1.54), so the store refuses a merge over other values (`merge-drops-values`). The
// caller asks the user first, and then merges with `dropValues`: the lowering then clears the
// values and formulas of the other cells (`clearValues`) before the merge, in the same
// transaction, so one undo puts the values back. A merge and center also centers the merged
// cell (`horizontal`).

import type { GridSelection } from '../contracts/neutral.ts';
import { formatCellRef, type WorkbookOp } from '../store/index.ts';

/**
 * How a merge treats the selection.
 *
 * @public
 */
export interface MergeOptions {
  /** Also center the text of the merged cell, as `horizontal="center"`. */
  readonly center?: boolean;
  /**
   * Clear the values and formulas of the cells other than the top-left cell, so the merge keeps
   * only the top-left value. Without it, a merge over other values is refused with
   * `merge-drops-values`, so the host can ask the user first. The clear can refuse with the
   * reasons of a clear of values, such as `shared-formula` for a part of an array formula range.
   */
  readonly dropValues?: boolean;
}

const ref = (top: number, left: number, bottom: number, right: number): string =>
  `${formatCellRef({ row: top, col: left })}:${formatCellRef({ row: bottom, col: right })}`;

/** The ops that merge the range of `selection` on `sheet`. */
export function mergeOps(
  sheet: string,
  selection: GridSelection,
  options: MergeOptions = {}
): WorkbookOp[] {
  const { start, end } = selection.range;
  const range = ref(start.row, start.col, end.row, end.col);
  const ops: WorkbookOp[] = [];
  // The center goes first: a clear removes a cell that has no format, so a format after it would
  // count more new cells than the state of the command counts (`formatRefusal`).
  if (options.center)
    ops.push({ op: 'applyFormat', sheet, ref: range, patch: { horizontal: 'center' } });
  if (options.dropValues) {
    // The cells other than the top-left cell: the rest of the top row, and the rows below it.
    if (end.col > start.col)
      ops.push({
        op: 'clearValues',
        sheet,
        ref: ref(start.row, start.col + 1, start.row, end.col),
      });
    if (end.row > start.row)
      ops.push({ op: 'clearValues', sheet, ref: ref(start.row + 1, start.col, end.row, end.col) });
  }
  ops.push({ op: 'mergeCells', sheet, ref: range });
  return ops;
}

/** The op that unmerges the merged ranges that the range of `selection` meets on `sheet`. */
export function unmergeOp(sheet: string, selection: GridSelection): WorkbookOp {
  const { start, end } = selection.range;
  return { op: 'unmergeCells', sheet, ref: ref(start.row, start.col, end.row, end.col) };
}
