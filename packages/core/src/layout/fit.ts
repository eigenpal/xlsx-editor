// The size that fits a column or a row to its content, for a double-click on a heading edge.
//
// A column fits the widest shown text of its cells: the text after the number format, in the
// resolved font of the cell, measured with the measurer that paints it (`textWidth`). The cell
// box around the text is the one that `textArea` takes away, read backward: the gridline pixel
// that each cell owns, a margin on each side for the font, and the indent. The pixels become a
// width in characters of the maximum digit width of the workbook's default font (ECMA-376
// Part 1, 18.3.1.13), the smallest one that gives those pixels (`columnWidthForPixels`).
//
// A row fits the height that layout gives a row without `@ht` (`rowAutoHeight`): the tallest
// font of its cells, and the lines of its wrapped cells at the width of their columns. Line
// heights round to whole pixels, so a row with a wrapped cell takes the tallest of these heights
// in points at each zoom of `zooms`, less one pixel of each zoom other than 1: that pixel is the
// noise of the rounding, and cuts no line. A row with no wrapped cell (`wrapText`, `justify`, or
// `distributed`) takes the height at zoom 1. A row with no cell fits the default row height of
// the sheet. On a sheet whose `x:sheetFormatPr/@customHeight` is set, a row also fits its
// content, and a fitted height different from the custom default height is written as `@ht`. A
// row with rotated or stacked text keeps its height, because the automatic height does not count
// that text.
//
// Only the cells that exist count, so the cost is in proportion to the content of the file,
// never to the size of a range. These cells do not fit a column: a cell that shows no text, a
// cell whose text wraps, a rotated cell, justified or distributed text, text centered across a
// selection, a cell of a merged range that spans more than one column, and a cell of a hidden
// row. A column with no other cell keeps its width, so a fit never makes a column of empty
// cells narrower. A hidden column or row keeps its size, so a fit does not show it.
//
// The width is in characters at zoom 1: the caller passes the profile at zoom 1, and each
// counted cell must show its text at each zoom of `zooms`. Cells with the same text, font, and
// alignment are measured once.

import { attr, child, children } from '../store/package/query.ts';
import type { Element } from '../store/package/tree.ts';
import { xmlPartOf, type XlsxPackage } from '../store/package/package.ts';
import { sheetIndex } from '../store/workbook/sheet-index.ts';
import { sharedStringItems, sheetByName, workbookIndex } from '../store/workbook/workbook-index.ts';
import {
  cellFormat,
  cellValue,
  conditionalFormat,
  displayValue,
  EN_US,
  formatCodeOf,
  LAYOUT_CHECK_BUDGET,
  type NumberLocale,
  type ResolvedAlignment,
  type ResolvedFont,
  richTextRuns,
  sheetStale,
} from '../cascade/index.ts';
import {
  indentWidth,
  textArea,
  textWidth,
  valueKind,
  type TextContext,
  type TextPad,
  type ValueKind,
} from './cell-text.ts';
import {
  columnPixels,
  columnWidthForPixels,
  flag,
  numberAttr,
  sheetGeometry,
} from './grid-geometry.ts';
import { gridMetrics, type GridProfile } from './metrics.ts';
import { AxisIndex, type AxisSegment } from './axis.ts';
import { fitsWidth } from './number-fit.ts';
import { mergeCoverage, mergedRanges } from './panes.ts';
import { rowAutoHeight, wraps } from './row-height.ts';
import { MAX_COLUMN_WIDTH, MAX_ROW_HEIGHT } from '../store/index.ts';
import { MAX_COLS } from '../store/cell-ref.ts';
import type { TextMeasurer } from './text/measurer.ts';

/** What a fit measures with. */
export interface FitOptions {
  readonly measurer: TextMeasurer;
  /** The profile at zoom 1. A width and a height are in zoom-1 units. */
  readonly profile: GridProfile;
  /** The zooms at which every fitted cell must show its text. Zoom 1 counts always. */
  readonly zooms?: readonly number[];
}

/** What a column fit measures with. */
export interface FitColumnOptions extends FitOptions {
  /** The locale of the number formats. */
  readonly locale?: NumberLocale;
}

/** The worksheet and its indexes, or undefined for an unknown sheet. */
function sheetOf(pkg: XlsxPackage, sheetName: string) {
  const index = workbookIndex(pkg);
  const entry = index ? sheetByName(index, sheetName) : undefined;
  const part = xmlPartOf(pkg, entry?.partName);
  if (!index || !entry || entry.kind !== 'worksheet' || !part) return undefined;
  return { index, sheet: sheetIndex(part) };
}

/** Whether a cell element holds a value (`x:v` or `x:is`). */
const hasValue = (cell: Element): boolean =>
  children(cell, 'x:v').length + children(cell, 'x:is').length > 0;

/**
 * The hidden columns of `x:cols`, read once and as `sheetGeometry` reads them: an index whose
 * size is 1 for a hidden column and 0 for any other.
 */
function hiddenColumns(root: Element): AxisIndex {
  const segments: AxisSegment[] = [];
  for (const c of child(root, 'x:cols')?.children ?? []) {
    if (c.type !== 'element' || c.local !== 'col') continue;
    const min = numberAttr(c, 'min', 1, MAX_COLS);
    const max = numberAttr(c, 'max', 1, MAX_COLS);
    if (min === undefined || max === undefined || !Number.isInteger(min) || !Number.isInteger(max))
      continue;
    segments.push({ first: min - 1, last: max - 1, size: flag(c, 'hidden') ? 1 : 0 });
  }
  return new AxisIndex(MAX_COLS, 0, segments);
}

/** One cell that fits its column: what layout draws, before a zoom. */
interface FitCell {
  readonly col: number;
  readonly runs: readonly { readonly text: string; readonly font: ResolvedFont }[];
  /** The font of the first run, empty or not, that gives the margins of the cell box. */
  readonly marginFont: ResolvedFont;
  readonly pads: readonly TextPad[];
  readonly kind: ValueKind;
  readonly alignment: ResolvedAlignment;
  /** A number in the General format, which layout shortens unless it has room to spare. */
  readonly general: boolean;
}

/**
 * In a cell that does not wrap, as in `measureRuns`: a run of line breaks between two other
 * characters takes the width of a space, and a line break at the start or the end takes none.
 */
const unbroken = (text: string) => text.replace(/^[\n\r]+|[\n\r]+$/g, '').replace(/[\n\r]+/g, ' ');

/** The width in points of the text of a cell, as layout measures it in `ctx`. */
function textOf(cell: FitCell, ctx: TextContext): number {
  if (cell.runs.length > 1)
    return cell.runs.reduce((w, r) => w + textWidth(ctx, unbroken(r.text), r.font), 0);
  const { text, font } = cell.runs[0]!;
  return textWidth(ctx, cell.pads.length ? text : unbroken(text), font, cell.pads);
}

/**
 * Whether a cell shows all of its text, unchanged, in a box `width` points wide: the text area
 * of the font of its first run less the indent (`textArea`, `indentWidth`). A number in the
 * General format needs room to spare; a formatted number and text may fill the area
 * (`fitsWidth`, which `fitNumber` also uses).
 */
function fitsBox(cell: FitCell, text: number, width: number, ctx: TextContext): boolean {
  const available =
    textArea({ x: 0, y: 0, width, height: 0 }, ctx, cell.marginFont) -
    indentWidth(cell.alignment, cell.kind, ctx);
  return fitsWidth(text, available, cell.general);
}

/**
 * The cells of the columns `first` to `last` that fit them, read from the cells that exist, and
 * the columns that a fit must keep: a cell of such a column was read after the conditional
 * formatting check budget ran out, so its shown format is not known.
 */
function fitCells(
  pkg: XlsxPackage,
  found: NonNullable<ReturnType<typeof sheetOf>>,
  first: number,
  last: number,
  locale: NumberLocale
): { readonly cells: FitCell[]; readonly unknown: ReadonlySet<number> } {
  const { index, sheet } = found;
  const stale = sheetStale(index, sheet.root);
  const checks = { left: LAYOUT_CHECK_BUDGET };
  const cells: FitCell[] = [];
  const unknown = new Set<number>();
  const hiddenRow = (row: number) => {
    const el = sheet.rows.get(row);
    return el !== undefined && flag(el, 'hidden');
  };
  // A merged range over several columns shows across them, and the cells of a merged range
  // below its first row show nothing. Each test sweeps the rows in ascending order
  // (`mergeCoverage`), so the cost depends on the merges that reach each row, not on a scan of
  // every merge for each cell.
  const ranges = mergedRanges(sheet.root).ranges.filter(
    (m) => m.lastCol >= first && m.firstCol <= last
  );
  const across = mergeCoverage(ranges.filter((m) => m.lastCol > m.firstCol));
  const below = mergeCoverage(
    ranges
      .filter((m) => m.lastCol === m.firstCol && m.lastRow > m.firstRow)
      .map((m) => ({ ...m, firstRow: m.firstRow + 1 }))
  );
  const entries = [...sheet.cells.values()]
    .filter((e) => e.col >= first && e.col <= last && !hiddenRow(e.row))
    .sort((a, b) => a.row - b.row || a.col - b.col);
  for (const entry of entries) {
    if (across(entry.row, entry.col) || below(entry.row, entry.col)) continue;
    const resolved = cellValue(pkg, index, entry.el, stale);
    const v = resolved.value;
    if (v.kind === 'empty') continue;
    // The shown format, with the conditional formats that paint the cell (`resolveCell`). A
    // formula without a cached value has an empty value, which is skipped before this.
    const base = cellFormat(pkg, index, attr(entry.el, 's')).format;
    const format = conditionalFormat(pkg, index, sheet.root, entry.row, entry.col, v, base, checks);
    if (checks.left <= 0) unknown.add(entry.col);
    const a = format.alignment;
    if (a.textRotation !== 0) continue;
    // Only text wraps: a number shows on one line in any alignment (`layoutCellText`).
    const number = v.kind === 'number' || v.kind === 'date';
    if (!number && (a.wrapText || a.horizontal === 'justify' || a.horizontal === 'distributed'))
      continue;
    // Text centered across a selection shows over the empty cells beside it, as a merge does.
    if (a.horizontal === 'centerContinuous') continue;
    if (v.kind === 'string' && v.runs && v.runs.length > 0) {
      const all = richTextRuns(v.runs, format.font, pkg, index);
      const runs = all.filter((r) => r.text !== '');
      if (runs.length > 0)
        cells.push({
          col: entry.col,
          runs,
          marginFont: all[0]!.font,
          pads: [],
          kind: 'text',
          alignment: a,
          general: false,
        });
      continue;
    }
    const shown = displayValue(v, format, { locale, date1904: index.date1904 });
    // An empty string, such as the cached result "" of a formula, shows nothing to fit.
    if (shown.text === '') continue;
    const kind = shown.alignAsText ? 'text' : valueKind(v);
    cells.push({
      col: entry.col,
      runs: [{ text: shown.text, font: format.font }],
      marginFont: format.font,
      pads: shown.pads ?? [],
      kind,
      alignment: a,
      general:
        kind === 'number' &&
        v.kind === 'number' &&
        formatCodeOf(format, locale).toLowerCase() === 'general',
    });
  }
  return { cells, unknown };
}

/**
 * The widths in characters that fit the columns `first` to `last` (0-based) of a worksheet to
 * their content, by column. Each counted cell shows all of its text, unchanged, at each zoom of
 * `options.zooms`. A column with no cell that fits it, a hidden column, and a column with a cell
 * read after the conditional formatting check budget ran out have no entry.
 * Undefined for an unknown sheet.
 */
export function fitColumnWidths(
  pkg: XlsxPackage,
  sheetName: string,
  first: number,
  last: number,
  options: FitColumnOptions
): Map<number, number> | undefined {
  const found = sheetOf(pkg, sheetName);
  if (!found) return undefined;
  const locale = options.locale ?? EN_US;
  const byColumn = new Map<number, FitCell[]>();
  // Cells that layout draws alike are measured once: the same text, fonts, and alignment.
  const seen = new Set<string>();
  const ids = new Map<object, number>();
  const id = (o: object) => ids.get(o) ?? (ids.set(o, ids.size), ids.size - 1);
  const { cells: fitted, unknown } = fitCells(pkg, found, first, last, locale);
  for (const cell of fitted) {
    const key =
      cell.runs.length === 1 && cell.pads.length === 0
        ? `${cell.col}:${id(cell.runs[0]!.font)}:${id(cell.marginFont)}:${id(cell.alignment)}:` +
          `${cell.kind}:${cell.general}:${cell.runs[0]!.text}`
        : undefined;
    if (key !== undefined) {
      if (seen.has(key)) continue;
      seen.add(key);
    }
    const list = byColumn.get(cell.col);
    if (list) list.push(cell);
    else byColumn.set(cell.col, [cell]);
  }
  const hidden = hiddenColumns(found.sheet.root);
  for (const col of byColumn.keys())
    if (hidden.sizeOf(col) > 0 || unknown.has(col)) byColumn.delete(col);
  const views = [...new Set([1, ...(options.zooms ?? [])])].map((zoom) => {
    const profile = { ...options.profile, pixelsPerInch: options.profile.pixelsPerInch * zoom };
    const metrics = gridMetrics(pkg, found.index, options.measurer, profile);
    const ctx: TextContext = {
      measurer: options.measurer,
      profile,
      maxDigitWidth: metrics.maxDigitWidth,
    };
    return { ctx, mdw: metrics.maxDigitWidth, pointsPerPixel: 72 / profile.pixelsPerInch };
  });
  const base = gridMetrics(pkg, found.index, options.measurer, options.profile).maxDigitWidth;
  const most = columnPixels(MAX_COLUMN_WIDTH, base);
  const widths = new Map<number, number>();
  for (const [col, cells] of byColumn) {
    // Whether a cell fits grows with the box and shrinks with its text, so each group of cells
    // with one margin font, alignment, kind, and number rule needs only its widest text.
    // Each cell is measured at every zoom in turn, so the bounded advance cache serves the
    // repeats of its text.
    const groups = views.map(() => new Map<string, { cell: FitCell; text: number }>());
    for (const cell of cells) {
      const key = `${id(cell.marginFont)}:${id(cell.alignment)}:${cell.kind}:${cell.general}`;
      const texts = views.map((view) => textOf(cell, view.ctx));
      // Text that measures nothing, such as a line break alone, shows nothing to fit.
      if (texts.every((text) => text === 0)) continue;
      texts.forEach((text, v) => {
        const hit = groups[v]!.get(key);
        if (!hit || text > hit.text) groups[v]!.set(key, { cell, text });
      });
    }
    // A column with no text to fit keeps its width.
    if (groups.every((g) => g.size === 0)) continue;
    const widest = groups.map((g) => [...g.values()]);
    const fits = (pixels: number) => {
      const width = columnWidthForPixels(pixels, base);
      return views.every((view, v) =>
        widest[v]!.every(({ cell, text }) =>
          fitsBox(cell, text, columnPixels(width, view.mdw) * view.pointsPerPixel, view.ctx)
        )
      );
    };
    // The fewest pixels of zoom 1 at which every cell fits at every zoom. Text too wide for the
    // widest column gets the widest column.
    let lo = 0;
    let hi = most;
    if (!fits(hi)) lo = hi;
    while (lo < hi) {
      const mid = Math.floor((lo + hi) / 2);
      if (fits(mid)) hi = mid;
      else lo = mid + 1;
    }
    widths.set(col, columnWidthForPixels(lo, base));
  }
  return widths;
}

/**
 * The heights in points that fit the rows `first` to `last` (0-based) of a worksheet to their
 * content, by row: the automatic height at zoom 1, raised for a row with a wrapped cell where
 * a zoom of `options.zooms` needs more than one pixel of that zoom beyond it. Only rows that the
 * file holds and whose `@ht` or `@customHeight` changes have an entry. A hidden row and a row
 * with rotated or stacked text have none. Undefined for an unknown sheet.
 */
export function fitRowHeights(
  pkg: XlsxPackage,
  sheetName: string,
  first: number,
  last: number,
  options: FitOptions
): Map<number, number> | undefined {
  const found = sheetOf(pkg, sheetName);
  if (!found) return undefined;
  const { index, sheet } = found;
  const { measurer } = options;
  const metrics = gridMetrics(pkg, index, measurer, options.profile);
  const items = sharedStringItems(pkg, index);
  const merges = mergedRanges(sheet.root);
  // Line heights round to whole pixels at each zoom, so the lines of a wrapped cell can need
  // more points at one zoom than at zoom 1. A row with a wrapped cell takes the automatic height
  // at zoom 1, or more where another zoom needs more than one pixel of that zoom beyond it: one
  // pixel is the noise of the rounding, and cuts no line. A row with no wrapped cell keeps the
  // height at zoom 1, as the rows around it do.
  const views = [...new Set([1, ...(options.zooms ?? [])])].map((zoom) => {
    const profile = { ...options.profile, pixelsPerInch: options.profile.pixelsPerInch * zoom };
    const zoomed = gridMetrics(pkg, index, measurer, profile);
    const ctx: TextContext = { measurer, profile, maxDigitWidth: zoomed.maxDigitWidth };
    return {
      auto: rowAutoHeight(pkg, index, items, ctx, zoomed.defaultRowHeight, merges),
      columns: sheetGeometry(sheet, zoomed).columns,
      pixel: zoom === 1 ? 0 : 72 / profile.pixelsPerInch,
    };
  });
  const auto = (el: Element, r: number, wrapped: boolean): number | undefined => {
    let most: number | undefined;
    for (const view of wrapped ? views : views.slice(0, 1)) {
      const h = view.auto(el, r, view.columns);
      // Zoom 1 comes first, so it sets the height that the other zooms raise.
      const need = h === undefined ? undefined : h - view.pixel;
      if (need !== undefined && (most === undefined || need > most)) most = need;
    }
    return most;
  };
  // A sheet with a custom default height gives that height to each row without `@ht`.
  const formatPr = child(sheet.root, 'x:sheetFormatPr');
  const customDefault = flag(formatPr, 'customHeight')
    ? numberAttr(formatPr, 'defaultRowHeight', 0, MAX_ROW_HEIGHT)
    : undefined;
  const heights = new Map<number, number>();
  for (const [r, el] of sheet.rows) {
    if (r < first || r > last || flag(el, 'hidden')) continue;
    // The automatic height does not count rotated or stacked text, so such a row keeps its
    // height rather than cut that text.
    const formats = children(el, 'x:c')
      .filter((cell) => hasValue(cell))
      .map((cell) => cellFormat(pkg, index, attr(cell, 's')).format);
    if (formats.some((format) => format.alignment.textRotation !== 0)) continue;
    // The content decides the height, also on a sheet with a custom default height, so a fit
    // never cuts the content. On such a sheet, a row with no value fits the custom default.
    const valued = customDefault === undefined || formats.length > 0;
    const fitted =
      (valued ? auto(el, r, formats.some(wraps)) : undefined) ??
      customDefault ??
      metrics.defaultRowHeight;
    // Up to the 8 decimals that the file keeps (`numberText`), so the stored height loses no
    // pixel, and a second fit reads back the same height. The 1e-4 keeps a height that is at 8
    // decimals already from going up by 1e-8 through floating-point noise.
    const height = Math.min(MAX_ROW_HEIGHT, Math.max(0, Math.ceil(fitted * 1e8 - 1e-4) / 1e8));
    // A row that shows its fitted height already changes nothing: the same `@ht` without
    // `@customHeight`, or no `@ht` and no `@customHeight` where layout gives the row that
    // height (the automatic height, or the custom default height that is the same).
    const ht = numberAttr(el, 'ht', 0, MAX_ROW_HEIGHT);
    const shown = ht ?? customDefault ?? height;
    if (!flag(el, 'customHeight') && shown === height) continue;
    heights.set(r, height);
  }
  return heights;
}
