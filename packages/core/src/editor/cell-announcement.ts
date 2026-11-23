// What the live region says about the selection: the active cell's reference, its shown text,
// and its formula, or the reference of a range. Split from `editor-model.ts`, so that each file
// stays within its line limit.

import type { GridPoint, GridRange, GridSelection } from '../contracts/index.ts';
import {
  attr,
  cellKey,
  formatCellRef,
  sheetIndex,
  xmlPartOf,
  type WorkbookIndex,
  type XlsxPackage,
} from '../store/index.ts';
import {
  cellFormat,
  cellValue,
  conditionalFormat,
  sheetStale,
  shownText,
  type NumberLocale,
} from '../cascade/index.ts';
import type { EditorTranslate } from './cell-input.ts';
import type { ShownGeneralText } from './decimal-selection.ts';
import { selectCell, type Merges } from './selection.ts';
import { contentFieldText } from './cell-editor.ts';
import { cellContent, selectionRef } from './editor-model.ts';
import { isWholeSheet, selectionBand } from '../automation/index.ts';

/**
 * What the live region says about one active cell: its reference, the text that the grid
 * shows (number format applied, a formula result, a date formatted, a `General` number fitted
 * to its column by `shown`), then a formula's text, and `grid.staleResult` when the result can
 * be out of date (design D3). The text comes from `shownText`, which layout paints with.
 *
 * Words longer than `MAX_CONTENT_FIELD_CHARS` give only their start, as the formula bar does,
 * and say that the text continues. The stale mark comes after the cut, so it always stays. The
 * live region is a page element, and each layout reads its whole text.
 *
 * @param rules - False after a budget cut of the sheet (`ConditionalCut.holds`): the text then
 * uses the cell format, as the hint does. True applies the true conditional formatting rules.
 */
export function cellAnnouncement(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  partName: string | undefined,
  at: GridPoint,
  locale: NumberLocale,
  t: EditorTranslate,
  rules = true,
  shown?: ShownGeneralText
): string {
  const ref = formatCellRef(at);
  const part = xmlPartOf(pkg, partName);
  const entry = part ? sheetIndex(part).cells.get(cellKey(at.row, at.col)) : undefined;
  if (!part || !entry) return t('grid.activeCellEmpty', { ref, content: '' });
  const resolved = cellValue(pkg, index, entry.el, sheetStale(index, part.root));
  // The text that the grid shows: the differential format under a true rule, the cell format
  // after a budget cut (`rules` false), and a `General` number as fitted (`shown`), unless `#`.
  const base = cellFormat(pkg, index, attr(entry.el, 's')).format;
  const { value } = resolved;
  const format =
    value.kind === 'empty' || !rules
      ? base
      : conditionalFormat(pkg, index, part.root, at.row, at.col, value, base);
  const fit = value.kind === 'number' ? shown?.(at.row, at.col) : undefined;
  const content =
    fit && /\d/.test(fit) ? fit : shownText(value, format, { locale, date1904: index.date1904 });
  // A shared formula whose text cannot be read gives the value back: it announces no formula.
  const source = resolved.formula ? cellContent(pkg, index, partName, at, locale) : '';
  const formula = source.startsWith('=') ? source : '';
  const said = formula
    ? t(content ? 'grid.activeFormula' : 'grid.activeFormulaNoResult', {
        ref,
        content,
        formula,
      })
    : t(content ? 'grid.activeCell' : 'grid.activeCellEmpty', { ref, content });
  // Cut the words first, so the stale mark always stays.
  const heard = heardStart(said, t);
  return resolved.stale ? t('grid.staleResult', { text: heard }) : heard;
}

/** The start of `text` that the live region says, and that the text continues when it does. */
function heardStart(text: string, t: EditorTranslate): string {
  const start = contentFieldText(text);
  return start.length < text.length ? t('grid.textContinues', { text: start }) : text;
}

/**
 * What the live region says about a selection: the range of a selection of more than one
 * cell, else the active cell as `cellAnnouncement` reads it. A merged cell is one cell, so its
 * range announces as the cell. `rules` is the flag of `cellAnnouncement`.
 */
export function selectionAnnouncement(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  partName: string | undefined,
  sel: GridSelection,
  merges: Merges,
  locale: NumberLocale,
  t: EditorTranslate,
  rules = true,
  shown?: ShownGeneralText
): string {
  const own = selectCell(sel.sheetIndex, sel.active, merges).range;
  if (!sameRange(own, sel.range)) return rangeAnnouncement(sel, t);
  return cellAnnouncement(pkg, index, partName, sel.active, locale, t, rules, shown);
}

/**
 * A selection of more than one cell in words: the whole sheet, whole rows (`Rows 2 to 4`),
 * whole columns (`Column B`), or the A1 reference of a range.
 */
function rangeAnnouncement(sel: GridSelection, t: EditorTranslate): string {
  if (isWholeSheet(sel)) return t('grid.selectedSheet');
  const band = selectionBand(sel);
  const { start, end } = sel.range;
  if (band === 'rows')
    return start.row === end.row
      ? t('grid.selectedRow', { first: start.row + 1 })
      : t('grid.selectedRows', { first: start.row + 1, last: end.row + 1 });
  if (band === 'columns') {
    const col = (c: number) => formatCellRef({ row: 0, col: c }).replace(/\d+$/, '');
    return start.col === end.col
      ? t('grid.selectedColumn', { first: col(start.col) })
      : t('grid.selectedColumns', { first: col(start.col), last: col(end.col) });
  }
  return t('grid.selectedRange', { range: selectionRef(sel) });
}

function sameRange(a: GridRange, b: GridRange): boolean {
  return (
    a.start.row === b.start.row &&
    a.start.col === b.start.col &&
    a.end.row === b.end.row &&
    a.end.col === b.end.col
  );
}
