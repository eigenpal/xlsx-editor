// Column widths and row heights of one worksheet (design D6), in points, as two axis indexes.
//
// Column widths follow ECMA-376 Part 1, 18.3.1.13 (`x:col`, `CT_Col`, sml.xsd): `@width` in
// characters of the maximum digit width (MDW) of the default font becomes whole pixels. The
// sheet defaults come from `x:sheetFormatPr` (`CT_SheetFormatPr`): `@baseColWidth` (default 8),
// `@defaultColWidth`, `@defaultRowHeight`, and `@customHeight`. Row heights come from
// `x:row/@ht` (`CT_Row`). The standard does not state the pixel density, the padding of a
// default column for each MDW, or the row snap. Measured on the reference PDFs
// (`basic-values`, `grid-structure`) and, where marked, in the screen profile
// (reference/quick-ref/spreadsheetml.md), which no committed reference shows:
// - It lays out at 72 pixels per inch, so 1 pixel is 1 point, and Carlito 11 has an MDW of
//   6 pixels. The formula gives every measured width: 1, 2, 5, 8.43, 10, 16, 30, and 50
//   characters are 6, 12, 30, 51, 60, 96, 180, and 300 points (screen; the PDFs print 3, 16, 18,
//   and 24 at 18, 96, 108, and 144).
// - A column with no `x:col` and no `defaultColWidth` is `baseColWidth` characters plus padding:
//   53 points for 8 characters (PDFs) and 65 points for 10 (screen), at MDW 6. Padding is
//   5 pixels at MDW 6 and 7 pixels at MDW 10 (Carlito 20). The text of the standard gives a fixed
//   padding of 5 pixels. The engine follows the measured values instead: `2 * ceil(MDW / 4) + 1`
//   is a formula fitted to these two points, and other widths have no measurement.
// - A row with `@ht` keeps whole pixels, rounded down: 33.75 is 33 points on screen. The
//   print profile scales every row height first (`rowScale`, metrics.ts).
// - A row without `@ht` takes the height of the default font, even when `defaultRowHeight` says
//   otherwise, unless `x:sheetFormatPr/@customHeight` is set.
//
// Nothing here loops over a range that the file states: an `x:col` range is one segment.

import type { Element } from '../store/package/tree.ts';
import { attr, child } from '../store/package/query.ts';
import { MAX_COLS, MAX_ROWS } from '../store/cell-ref.ts';
import type { SheetIndex } from '../store/workbook/sheet-index.ts';
import { AxisIndex, type AxisSegment } from './axis.ts';

/** The facts of the default font and the pixel grid that geometry needs. */
export interface GridMetrics {
  /** The maximum digit width of the default font, in whole pixels. */
  readonly maxDigitWidth: number;
  /** Pixels per inch of the layout grid: 72 in both profiles (metrics.ts). */
  readonly pixelsPerInch: number;
  /** The row height that the default font needs, in points, before `rowScale`. */
  readonly defaultRowHeight: number;
  /** The factor of every row height before it rounds down to whole pixels (`GridProfile`). */
  readonly rowScale: number;
}

/** The geometry of one sheet. */
export interface SheetGeometry {
  readonly columns: AxisIndex;
  readonly rows: AxisIndex;
}

const truncate = Math.trunc;

/**
 * The padding of a default column, in pixels, for an MDW. The standard's formula (Part 1,
 * 18.3.1.13) has a fixed padding of 5 pixels. This formula is fitted to the measured paddings,
 * 5 pixels at MDW 6 and 7 at MDW 10; it is not from the standard.
 */
export function columnPadding(mdw: number): number {
  return 2 * Math.ceil(mdw / 4) + 1;
}

/** Whole pixels of a column `width` in characters, `x:col/@width` (`CT_Col`, Part 1, 18.3.1.13). */
export function columnPixels(width: number, mdw: number): number {
  if (!(width > 0) || !(mdw > 0)) return 0;
  return truncate(((256 * width + truncate(128 / mdw)) / 256) * mdw);
}

/**
 * The inverse of `columnPixels`: the smallest width in characters, in steps of 1/256 of a
 * character as `columnPixels` reads them, whose whole pixels are `pixels` (Part 1, 18.3.1.13).
 * `columnPixels(columnWidthForPixels(p, mdw), mdw)` is `p` for every whole `p` up to the widest
 * column of 255 characters, the editor's limit of a column width. Clamped to 0 to 255.
 */
export function columnWidthForPixels(pixels: number, mdw: number): number {
  if (!(pixels > 0) || !(mdw > 0)) return 0;
  // columnPixels(w) = trunc((256w + trunc(128 / mdw)) * mdw / 256), so the smallest 256w is
  // ceil(256 * pixels / mdw) - trunc(128 / mdw), and the product stays below pixels + 1.
  const steps = Math.ceil((256 * Math.round(pixels)) / mdw - 1e-9) - truncate(128 / mdw);
  return Math.min(255, Math.max(0, steps) / 256);
}

/**
 * The inverse of the row snap of `sheetGeometry`: a row height in points, to 1/100 of a point,
 * whose height on the layout grid is `pixels` whole pixels. Clamped to 0 to 409, the editor's
 * limit of a row height (`MAX_ROW_HEIGHT` of the store).
 */
export function rowHeightForPixels(
  pixels: number,
  metrics: Pick<GridMetrics, 'pixelsPerInch' | 'rowScale'>
): number {
  const px = Math.round(pixels);
  if (!(px > 0) || !(metrics.pixelsPerInch > 0) || !(metrics.rowScale > 0)) return 0;
  const ptPerPx = 72 / metrics.pixelsPerInch;
  let pt = Math.round(((px * ptPerPx) / metrics.rowScale) * 100) / 100;
  if (Math.floor((pt / ptPerPx) * metrics.rowScale + 1e-9) < px) pt += 0.01;
  return Math.min(409, Math.round(pt * 100) / 100);
}

/**
 * The width in characters of a default column of `base` characters,
 * `x:sheetFormatPr/@baseColWidth` (`CT_SheetFormatPr`, Part 1, 18.3.1.81). The width formula is
 * the one of Part 1, 18.3.1.13, with the measured padding of `columnPadding`.
 */
export function baseColumnWidth(base: number, mdw: number): number {
  return truncate(((base * mdw + columnPadding(mdw)) / mdw) * 256) / 256;
}

/** A finite number attribute in a range, or undefined. */
export function numberAttr(
  el: Element | undefined,
  name: string,
  min: number,
  max: number
): number | undefined {
  const raw = attr(el, name);
  if (raw === undefined || raw.length === 0 || raw.length > 32) return undefined;
  const n = Number(raw);
  return Number.isFinite(n) && n >= min && n <= max ? n : undefined;
}

/** Whether a boolean attribute (`xsd:boolean`) is true (`1` or `true`). */
export const flag = (el: Element | undefined, name: string): boolean =>
  attr(el, name) === '1' || attr(el, name) === 'true';

/**
 * The height in points of a row without `@ht`, from the fonts and wrapped text of its cells, or
 * undefined for the default row height. `rowIndex` is 0-based, and `columns` is the column axis
 * of the sheet, for the width that wrapped text breaks at. Layout passes one that reads the
 * cell formats (`row-height.ts`).
 */
export type AutoRowHeight = (
  row: Element,
  rowIndex: number,
  columns: AxisIndex
) => number | undefined;

/**
 * The column and row geometry of a worksheet. A row without `@ht` takes `autoHeight` when the
 * sheet's default row height is not custom; the cost is the cells of those rows, which is in
 * proportion to the content of the file.
 */
export function sheetGeometry(
  sheet: SheetIndex,
  metrics: GridMetrics,
  autoHeight?: AutoRowHeight
): SheetGeometry {
  const ptPerPx = 72 / metrics.pixelsPerInch;
  const mdw = metrics.maxDigitWidth;
  const formatPr = child(sheet.root, 'x:sheetFormatPr');

  const defaultWidth =
    numberAttr(formatPr, 'defaultColWidth', 0, 255) ??
    baseColumnWidth(numberAttr(formatPr, 'baseColWidth', 0, 255) ?? 8, mdw);
  const defaultColumn = columnPixels(defaultWidth, mdw) * ptPerPx;
  const columns: AxisSegment[] = [];
  for (const col of child(sheet.root, 'x:cols')?.children ?? []) {
    if (col.type !== 'element' || col.local !== 'col') continue;
    const min = numberAttr(col, 'min', 1, MAX_COLS);
    const max = numberAttr(col, 'max', 1, MAX_COLS);
    if (min === undefined || max === undefined || !Number.isInteger(min) || !Number.isInteger(max))
      continue;
    const width = numberAttr(col, 'width', 0, 255);
    const size = flag(col, 'hidden')
      ? 0
      : width === undefined
        ? defaultColumn
        : columnPixels(width, mdw) * ptPerPx;
    columns.push({ first: min - 1, last: max - 1, size });
  }

  const columnAxis = new AxisIndex(MAX_COLS, defaultColumn, columns);
  const snap = (pt: number) => Math.floor((pt / ptPerPx) * metrics.rowScale + 1e-9) * ptPerPx;
  const customDefault = flag(formatPr, 'customHeight')
    ? numberAttr(formatPr, 'defaultRowHeight', 0, 409)
    : undefined;
  const defaultRow = snap(customDefault ?? metrics.defaultRowHeight);
  const rows: AxisSegment[] = [];
  const auto = customDefault === undefined ? autoHeight : undefined;
  for (const [r, el] of sheet.rows) {
    const ht = numberAttr(el, 'ht', 0, 409);
    if (flag(el, 'hidden')) rows.push({ first: r, last: r, size: 0 });
    else if (ht !== undefined) rows.push({ first: r, last: r, size: snap(ht) });
    else {
      const size = auto?.(el, r, columnAxis);
      if (size !== undefined && snap(size) !== defaultRow)
        rows.push({ first: r, last: r, size: snap(size) });
    }
  }
  return {
    columns: columnAxis,
    rows: new AxisIndex(MAX_ROWS, defaultRow, rows),
  };
}
