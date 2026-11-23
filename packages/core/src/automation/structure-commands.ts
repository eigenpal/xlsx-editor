// The structural commands, DOM-free (task 7.5): insert or delete the rows or columns of a
// selection, lowered to one store op (design D4). The editor and the fixture edits of the save
// check call the same lowering.
//
// A row command acts on the rows of the selection: the selected rows of a whole-row selection,
// or the rows of a range. A selection of whole columns has every row, so a row command acts on
// the active cell's row instead. Column commands mirror this. A whole-sheet selection is
// refused (`isWholeSheet`): it would insert or delete every row or every column.

import type { GridSelection } from '../contracts/neutral.ts';
import { MAX_COLS, MAX_ROWS, type StructuralOpName, type WorkbookOp } from '../store/index.ts';

/** The kind of band that a selection is: whole rows, whole columns, or neither. */
export type SelectionBand = 'rows' | 'columns' | null;

/** Whether a selection covers whole rows or whole columns. A whole sheet is neither. */
export function selectionBand(selection: GridSelection): SelectionBand {
  const { start, end } = selection.range;
  const allCols = start.col === 0 && end.col === MAX_COLS - 1;
  const allRows = start.row === 0 && end.row === MAX_ROWS - 1;
  if (allCols && !allRows) return 'rows';
  if (allRows && !allCols) return 'columns';
  return null;
}

/** Whether a selection covers the whole sheet, which a structural command refuses. */
export function isWholeSheet(selection: GridSelection): boolean {
  const { start, end } = selection.range;
  return start.col === 0 && end.col === MAX_COLS - 1 && start.row === 0 && end.row === MAX_ROWS - 1;
}

/** The structural op of a command on a selection of a sheet that is not the whole sheet. */
export function structureOp(
  sheet: string,
  selection: GridSelection,
  name: StructuralOpName
): WorkbookOp {
  const rows = name === 'insertRows' || name === 'deleteRows';
  const band = selectionBand(selection);
  const { start, end } = selection.range;
  const active = selection.active;
  // A band across the other axis has every index of this axis: act on the active cell's.
  const across = rows ? band === 'columns' : band === 'rows';
  const first = across ? (rows ? active.row : active.col) : rows ? start.row : start.col;
  const last = across ? first : rows ? end.row : end.col;
  return { op: name, sheet, start: first + 1, count: last - first + 1 };
}
