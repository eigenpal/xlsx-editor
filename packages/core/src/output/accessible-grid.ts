// The accessible text of a painted sheet window (design D7). The canvas paint holds no text that
// assistive technology reads, so the viewer puts these rows of the same layout inside its
// `grid` element: one row for each visible row that shows text, and one grid cell for each cell
// text, in reading order, with its row and column index. The rows are visually hidden; the
// canvas shows the cells.
//
// Every text goes through `textContent` and every attribute through `setAttribute`
// (AGENTS.md "Security").

import { formatCellRef } from '../store/index.ts';
import type { CellTextBox, SheetLayout } from '../layout/index.ts';

/** The shown text of a cell: its lines, joined by a space. */
export function cellText(box: CellTextBox): string {
  if (box.shownText !== undefined) return box.shownText;
  return box.text.lines.map((l) => l.spans.map((s) => s.text).join('')).join(' ');
}

/** The words that the accessible rows add to a cell text. */
export interface AccessibleGridLabels {
  /**
   * The text of a formula cell whose shown result can be out of date, from the
   * cell's shown text: the editor passes the catalog's `grid.staleResult`.
   */
  staleCell(text: string): string;
  /**
   * The words of a stale formula cell that shows no text, because it has no cached result: the
   * editor passes the catalog's `grid.staleNoResult`, with no placeholder.
   */
  readonly staleNoResult: string;
}

/**
 * A visually hidden row group of the cell texts of `layout`, in `doc`. A stale formula cell
 * reads its shown text with the stale words of `labels`, so assistive technology hears the
 * stale state of each cell; `data-stale` marks it for tests and styles.
 */
export function accessibleGrid(
  doc: Document,
  layout: SheetLayout,
  labels: AccessibleGridLabels
): HTMLElement {
  const table = doc.createElement('div');
  table.setAttribute('role', 'rowgroup');
  table.className = 'xlsx-a11y-rows';
  const boxes = layout.panes
    .flatMap((p) => p.texts)
    .slice()
    .sort((a, b) => a.row - b.row || a.col - b.col);
  let row: HTMLElement | null = null;
  let rowIndex = -1;
  for (const box of boxes) {
    if (box.row !== rowIndex) {
      row = doc.createElement('div');
      row.setAttribute('role', 'row');
      row.setAttribute('aria-rowindex', String(box.row + 1));
      table.appendChild(row);
      rowIndex = box.row;
    }
    const cell = doc.createElement('div');
    cell.setAttribute('role', 'gridcell');
    cell.setAttribute('aria-colindex', String(box.col + 1));
    cell.setAttribute('data-cell', formatCellRef({ row: box.row, col: box.col }));
    const text = cellText(box);
    if (box.stale) cell.setAttribute('data-stale', '1');
    // A stale cell with no shown text reads as having no result, never as a bare stale word.
    cell.textContent = !box.stale ? text : text ? labels.staleCell(text) : labels.staleNoResult;
    row!.appendChild(cell);
  }
  return table;
}
