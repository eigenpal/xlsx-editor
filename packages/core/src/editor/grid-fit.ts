// The fit of columns or rows to their content, for a double-click on a heading edge: the sizes come
// from the layout (`fitColumnWidths`, `fitRowHeights`), in zoom-1 units, so that the text shows
// at each of `FIT_ZOOMS` and the current zoom. They are written as one transaction: one
// `setColumnWidths` op for all the columns, or one `setFittedRowHeights` op for all the rows,
// and each op sets its columns or rows in one pass. So one undo restores every size
// of the band. The file keeps `@width` with `@customWidth` (ECMA-376 Part 1, 18.3.1.13), or `@ht`
// without `@customHeight`, because a fitted height is not a height set by hand (18.3.1.73): a later
// font change or value edit fits the row again. A fit writes no `@bestFit`. The text of 18.3.1.13
// gives a best fit column a width that was never set by hand and that grows as numbers are typed,
// and a fitted width carries `@customWidth` and does not grow.
//
// The open cell edit commits first, so the fit measures the text that the edit wrote. A band
// with nothing to fit (empty columns, or rows that the file does not hold) writes nothing.

import type { NumberLocale } from '../cascade/index.ts';
import { fitColumnWidths, fitRowHeights, zoomProfile, type TextMeasurer } from '../layout/index.ts';
import type { CellEditing } from './cell-editing.ts';
import { ZOOM_STEPS } from './editor-model.ts';
import type { EditResult } from './cell-input.ts';
import type { ResizeAxis } from './grid-resize.ts';

/**
 * The zoom steps at which a fitted column or row shows its text: 50% and more. Below 50%,
 * whole-pixel glyph widths of small fonts would make a column far wider than its text needs at
 * 100%. A fit also counts the current zoom, below 50% too, so the text shows in the view of the
 * fit, at the cost of a wider column.
 */
export const FIT_ZOOMS: readonly number[] = ZOOM_STEPS.filter((zoom) => zoom >= 0.5);

/**
 * Fit the columns or rows `first` to `last` (0-based) of the shown sheet to their content, as
 * one transaction and one undo step. A column or row fits at each of `FIT_ZOOMS` and at the
 * current `zoom`.
 */
export function runFit(
  edits: CellEditing,
  axis: ResizeAxis,
  [first, last]: readonly [number, number],
  measurer: TextMeasurer | null,
  locale: NumberLocale,
  zoom = 1
): EditResult {
  const open = edits.commitOpenEdit();
  if (!open.ok) return open;
  const target = edits.target;
  if (!target || !measurer) return edits.refuse({ code: 'not-editable' });
  const sheet = target.entry.name;
  const profile = zoomProfile(1);
  const columns = axis === 'columns';
  const zooms = [...new Set([...FIT_ZOOMS, zoom])];
  const sizes = columns
    ? fitColumnWidths(target.pkg, sheet, first, last, { measurer, profile, locale, zooms })
    : fitRowHeights(target.pkg, sheet, first, last, { measurer, profile, zooms });
  if (!sizes || sizes.size === 0) return { ok: true };
  if (!columns) {
    const rows = [...sizes]
      .sort(([a], [b]) => a - b)
      .map(([at, height]) => ({ row: at + 1, height }));
    return edits.run([{ op: 'setFittedRowHeights', sheet, rows }]);
  }
  const columnWidths = [...sizes]
    .sort(([a], [b]) => a - b)
    .map(([at, width]) => ({ column: at + 1, width }));
  return edits.run([{ op: 'setColumnWidths', sheet, columns: columnWidths }]);
}
