// The layout of the visible window of one worksheet (design D6): target-neutral records that a
// painter draws (design D7). Cost is in proportion to the visible cells, plus content-bounded
// work that is cached per part revision: the sorted cell keys, the geometry, the merged ranges,
// and the drawing anchors.
//
// Each pane of the window (frozen rows and columns, then the scrolling body) holds its columns and
// rows, the cell boxes that have a fill or a border, the text of each cell, and the drawings that
// meet it, all in viewport points. A merged range is one box with the text and fill of its top-left
// cell. A cell with no `x:c` element paints the fill and the borders of its row format (`x:row/@s`
// with `@customFormat`, `CT_Row`) or its column format (`x:col/@style`, `CT_Col`), as
// ECMA-376 Part 1, 18.3.1.73 and 18.3.1.13 give it; the cost is the visible cells of those rows
// and columns.
// Unwrapped text shows across empty neighbor cells up to the first cell that holds a value, or a
// merged range (quick reference, "Text in cells"); the search goes at most `OVERFLOW_REACH` columns
// (the approximation "Overflow reach").

import { cellKey, type SheetIndex, sheetIndex } from '../store/workbook/sheet-index.ts';
import { columnFormats, rowFormat } from '../store/workbook/inherited-format.ts';
import {
  sharedStringItems,
  sheetByName,
  styleLists,
  workbookIndex,
  type WorkbookIndex,
} from '../store/workbook/workbook-index.ts';
import { xmlPartOf, type XlsxPackage } from '../store/package/package.ts';
import { attr, child } from '../store/package/query.ts';
import type { Element, XmlPart } from '../store/package/tree.ts';
import {
  type CascadeProblem,
  type CellValue,
  cellFormat,
  type CheckBudget,
  conditionalFormat,
  rowHasShownRules,
  LAYOUT_CHECK_BUDGET,
  sheetRules,
  cellValue,
  defaultFont,
  displayValue,
  EN_US,
  formatCodeOf,
  type NumberLocale,
  type ResolvedBorder,
  type ResolvedCellFormat,
  type ResolvedFill,
  type ResolvedValue,
  richTextRuns,
  sheetStale,
  workbookTheme,
} from '../cascade/index.ts';
import { sheetGeometry, type GridMetrics, type SheetGeometry } from './grid-geometry.ts';
import { fontPixels, gridMetrics, type GridProfile } from './metrics.ts';
import { rowAutoHeight } from './row-height.ts';
import {
  type CellRange,
  frozenPanes,
  type FrozenPanes,
  mergeCoverage,
  mergeAnchoredAt,
  mergedRanges,
  mergesIn,
  windowPanes,
} from './panes.ts';
import {
  type CellTextLayout,
  indentWidth,
  layoutCellText,
  type Rect,
  textArea,
  textWidth,
  valueKind,
  type TextContext,
  type TextPad,
  type TextRunInput,
  type ValueKind,
} from './cell-text.ts';
import { MANY_RUNS, SharedCellTexts } from './shared-cell-text.ts';
import { fitNumber } from './number-fit.ts';
import { drawingBounds, drawingPartOf, type DrawingLayout, sheetDrawings } from './drawings.ts';
import type { TextMeasurer } from './text/measurer.ts';

/** The visible area, in points, and the scroll position of the scrolling pane. */
export interface Viewport {
  readonly width: number;
  readonly height: number;
  readonly scrollX: number;
  readonly scrollY: number;
}

export interface LayoutOptions {
  readonly measurer: TextMeasurer;
  /**
   * How the grid rounds: `PRINT_PROFILE` for the printed page of the reference PDFs (the
   * fidelity harness), `SCREEN_PROFILE` for a viewer. Required, so that no caller lays out
   * printed row heights on a screen by accident.
   */
  readonly profile: GridProfile;
  /** `EN_US` when absent. */
  readonly locale?: NumberLocale;
}

/** A cell (or merged range) that has a fill or a border to paint. */
export interface CellBox {
  readonly row: number;
  readonly col: number;
  /** In viewport points. */
  readonly rect: Rect;
  readonly fill: ResolvedFill;
  readonly border: ResolvedBorder;
  readonly merge?: CellRange;
}

/** The text of one cell, clipped to its cell or to the empty cells it shows across. */
export interface CellTextBox {
  readonly row: number;
  readonly col: number;
  readonly clip: Rect;
  readonly text: CellTextLayout;
  /** A cached formula result that may be out of date. */
  readonly stale: boolean;
  /** The complete text, when `text` holds only the spans that can reach `clip`. */
  readonly shownText?: string;
}

export interface AxisEntry {
  readonly index: number;
  /** Viewport position and size, in points. */
  readonly start: number;
  readonly size: number;
}

export interface PaneLayout {
  readonly range: CellRange;
  /** The part of the viewport that the pane paints into. */
  readonly clip: Rect;
  /**
   * Where the top-left corner of the sheet lies in this pane, in viewport points: it moves with
   * the scroll position, so a pattern fill that starts its tile grid here moves with its cells.
   */
  readonly sheetOrigin: { readonly x: number; readonly y: number };
  readonly columns: readonly AxisEntry[];
  readonly rows: readonly AxisEntry[];
  readonly cells: readonly CellBox[];
  /**
   * The text boxes of the pane. The box of a stale formula with no text to show has no lines,
   * and `PaintStats.texts` counts it with the others.
   */
  readonly texts: readonly CellTextBox[];
  /** Drawings that meet the pane, with `rect` in viewport points. */
  readonly drawings: readonly DrawingLayout[];
}

export interface LayoutProblem {
  readonly code:
    | CascadeProblem['code']
    | 'window-truncated'
    | 'merges-dropped'
    | 'merges-overlapping'
    | 'image-missing'
    /** Conditional formatting rules past `LAYOUT_CHECK_BUDGET` were not evaluated. */
    | 'conditional-budget';
  /** The cell reference, as row and column, where there is one. */
  readonly row?: number;
  readonly col?: number;
  readonly detail?: string;
}

export interface SheetLayout {
  readonly sheet: string;
  readonly metrics: GridMetrics;
  readonly geometry: SheetGeometry;
  readonly frozen: FrozenPanes;
  /** `x:sheetView/@showGridLines` of the first sheet view (`CT_SheetView`, default true). */
  readonly showGridLines: boolean;
  /** `x:printOptions/@gridLines` (`CT_PrintOptions`, default false). */
  readonly printGridLines: boolean;
  /** Viewport points of the frozen pane edges, when the sheet freezes rows or columns. */
  readonly frozenEdges: { readonly x?: number; readonly y?: number };
  readonly panes: readonly PaneLayout[];
  readonly problems: readonly LayoutProblem[];
  /**
   * The layout used up its conditional formatting check budget (`LAYOUT_CHECK_BUDGET`), so some
   * cells show their cell formats without their rules. Unlike the `conditional-budget` problem,
   * the problem cap never drops it.
   */
  readonly conditionalCut: boolean;
}

/** A pane shows at most this many rows and this many columns. */
export const MAX_WINDOW = 2000;
/** Unwrapped text shows across at most this many empty neighbor cells. */
export const OVERFLOW_REACH = 64;
const MAX_PROBLEMS = 200;

interface SheetContext {
  readonly pkg: XlsxPackage;
  readonly index: WorkbookIndex;
  readonly sheet: SheetIndex;
  readonly keys: Float64Array;
  readonly stale: boolean;
  readonly locale: NumberLocale;
  readonly text: TextContext;
  readonly geometry: SheetGeometry;
  readonly problems: LayoutProblem[];
  /** Rule-range checks left for this layout (`LAYOUT_CHECK_BUDGET`). */
  readonly checks: CheckBudget;
}

const keyCache = new WeakMap<SheetIndex, Float64Array>();

/** The cell keys of a sheet in ascending order. Cached per sheet index. */
function sortedKeys(sheet: SheetIndex): Float64Array {
  let keys = keyCache.get(sheet);
  if (!keys) {
    keys = Float64Array.from(sheet.order).sort();
    keyCache.set(sheet, keys);
  }
  return keys;
}

/** The first position in `keys` whose key is at least `key`. */
function lowerBound(keys: Float64Array, key: number): number {
  let lo = 0;
  let hi = keys.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (keys[mid]! < key) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

interface Geometry {
  readonly metrics: GridMetrics;
  readonly geometry: SheetGeometry;
}

interface CacheNode {
  readonly next: WeakMap<object, CacheNode>;
  value?: Map<string, Geometry>;
}

const geometryCache = new WeakMap<object, CacheNode>();

/**
 * The geometry cache of one combination of objects. Each object is a cached revision of what
 * the geometry reads: the sheet index, the style lists, the shared string items, the theme, and
 * the measurer. A change to any of them misses, and the old entries go with their objects.
 */
function geometrySlot(keys: readonly object[]): Map<string, Geometry> {
  let map = geometryCache;
  let node: CacheNode | undefined;
  for (const key of keys) {
    node = map.get(key);
    if (!node) {
      node = { next: new WeakMap() };
      map.set(key, node);
    }
    map = node.next;
  }
  return (node!.value ??= new Map());
}

function record(ctx: { problems: LayoutProblem[] }, problem: LayoutProblem): void {
  if (ctx.problems.length < MAX_PROBLEMS) ctx.problems.push(problem);
}

/** The geometry of a sheet in a profile, with row heights from the fonts of their cells. */
function geometryOf(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  sheet: SheetIndex,
  measurer: TextMeasurer,
  profile: GridProfile
): Geometry {
  const byProfile = geometrySlot([
    sheet,
    styleLists(pkg, index),
    sharedStringItems(pkg, index),
    workbookTheme(pkg, index),
    measurer,
  ]);
  const metrics = gridMetrics(pkg, index, measurer, profile);
  const key = `${profile.pixelsPerInch}|${profile.rowPadding}|${profile.rowScale}|${metrics.maxDigitWidth}|${metrics.defaultRowHeight}`;
  const hit = byProfile.get(key);
  if (hit) return hit;
  const autoHeight = rowAutoHeight(
    pkg,
    index,
    sharedStringItems(pkg, index),
    { measurer, profile, maxDigitWidth: metrics.maxDigitWidth },
    metrics.defaultRowHeight,
    mergedRanges(sheet.root)
  );
  const result = { metrics, geometry: sheetGeometry(sheet, metrics, autoHeight) };
  if (byProfile.size > 8) byProfile.clear();
  byProfile.set(key, result);
  return result;
}

const flagOf = (el: Element | undefined, name: string, fallback: boolean): boolean => {
  const v = attr(el, name);
  return v === undefined ? fallback : v === '1' || v === 'true';
};

/** The value of a cell without an `x:c` element. */
const EMPTY: CellValue = Object.freeze({ kind: 'empty' });

interface Resolved {
  readonly el: Element;
  readonly value: ResolvedValue;
  readonly format: ResolvedCellFormat;
}

/** `conditionalFormat` within the check budget of the layout; records the problem once. */
function conditionalWithin(
  ctx: SheetContext,
  row: number,
  col: number,
  value: CellValue,
  format: ResolvedCellFormat
): ResolvedCellFormat {
  const before = ctx.checks.left;
  const shown = conditionalFormat(
    ctx.pkg,
    ctx.index,
    ctx.sheet.root,
    row,
    col,
    value,
    format,
    ctx.checks
  );
  if (before > 0 && ctx.checks.left <= 0) record(ctx, { code: 'conditional-budget' });
  return shown;
}

/** {@link rowHasShownRules} under the budget of the layout. */
function rowHasShownRulesWithin(ctx: SheetContext, row: number): boolean {
  const before = ctx.checks.left;
  const meets = rowHasShownRules(ctx.sheet.root, row, ctx.checks);
  if (before > 0 && ctx.checks.left <= 0) record(ctx, { code: 'conditional-budget' });
  return meets;
}

function resolveCell(ctx: SheetContext, row: number, col: number): Resolved | undefined {
  const entry = ctx.sheet.cells.get(cellKey(row, col));
  if (!entry) return undefined;
  const value = cellValue(ctx.pkg, ctx.index, entry.el, ctx.stale);
  const format = cellFormat(ctx.pkg, ctx.index, attr(entry.el, 's'));
  if (value.problem) record(ctx, { ...value.problem, row, col });
  for (const p of format.problems) record(ctx, { ...p, row, col });
  // A formula without a cached value has no value to test: its rules do not paint.
  const unknown = value.formula && value.value.kind === 'empty';
  const shown = unknown
    ? format.format
    : conditionalWithin(ctx, row, col, value.value, format.format);
  return { el: entry.el, value, format: shown };
}

/**
 * Whether a cell holds a value or a formula: such a cell stops the text of a neighbor from
 * showing across it, also when the formula has no cached value and shows nothing.
 */
const hasValue = (r: Resolved | undefined): boolean =>
  !!r && (r.value.value.kind !== 'empty' || r.value.formula);

/** The runs of a cell whose unrotated text shares its layout in the pane, or undefined. */
function manyRunsOf(cell: Resolved): object | undefined {
  const v = cell.value.value;
  const { textRotation, horizontal } = cell.format.alignment;
  if (v.kind !== 'string' || !v.runs || v.runs.length < MANY_RUNS) return undefined;
  return textRotation === 0 && horizontal !== 'centerContinuous' ? v.runs : undefined;
}

/** The text runs, value kind, and color of a cell's value, fitted to `box` for numbers. */
function textOf(
  ctx: SheetContext,
  cell: Resolved,
  box: Rect
):
  | { runs: readonly TextRunInput[]; kind: ValueKind; alignKind?: ValueKind; color?: string }
  | undefined {
  const v = cell.value.value;
  const font = cell.format.font;
  if (v.kind === 'empty') return undefined;
  if (v.kind === 'string' && v.runs && v.runs.length > 0)
    return {
      kind: 'text',
      runs: richTextRuns(v.runs, font, ctx.pkg, ctx.index),
    };
  const shown = displayValue(v, cell.format, { locale: ctx.locale, date1904: ctx.index.date1904 });
  if (shown.problem) record(ctx, { ...shown.problem });
  const kind: ValueKind = valueKind(v);
  // `general` alignment: a number through `@` aligns as text (number-formats.pdf), and a cell
  // full of `#` centers it as it centers an error (number-formats-more.pdf: an out of range
  // date fills 139 points with 27 `#` of 5 points, 2 points from the left).
  let alignKind: ValueKind | undefined = shown.alignAsText
    ? 'text'
    : shown.fill && shown.text === '' && shown.fill.char === '#'
      ? 'error'
      : undefined;
  const available =
    textArea(box, ctx.text, font) - indentWidth(cell.format.alignment, alignKind ?? kind, ctx.text);
  const width = (s: string, at: readonly TextPad[] = []) => textWidth(ctx.text, s, font, at);
  let text = shown.text;
  let pads: readonly TextPad[] = shown.pads ?? [];
  if (shown.fill) {
    const room = available - width(text, pads);
    const unit = width(shown.fill.char);
    const count = unit > 0 && room > 0 ? Math.min(1000, Math.floor(room / unit)) : 0;
    const at = shown.fill.index;
    text = text.slice(0, at) + shown.fill.char.repeat(count) + text.slice(at);
    pads = pads.map((p) => (p.index >= at ? { index: p.index + count, char: p.char } : p));
  }
  if (
    kind === 'number' &&
    !cell.format.alignment.wrapText &&
    cell.format.alignment.textRotation === 0
  ) {
    const general = formatCodeOf(cell.format, ctx.locale).toLowerCase() === 'general';
    const value = v.kind === 'number' ? v.value : Number.NaN;
    const fitted = fitNumber(
      value,
      text,
      general && v.kind === 'number',
      ctx.locale.decimal,
      available,
      (s) => width(s, s === text ? pads : [])
    );
    if (fitted !== text) {
      pads = [];
      if (/^#+$/.test(fitted)) alignKind = 'error';
    }
    text = fitted;
  }
  return {
    kind,
    runs: [{ text, font, ...(pads.length ? { pads } : {}) }],
    ...(alignKind ? { alignKind } : {}),
    ...(shown.color ? { color: shown.color } : {}),
  };
}

interface PaneMap {
  readonly range: CellRange;
  /** Viewport x of the sheet's x = 0, and y likewise. */
  readonly dx: number;
  readonly dy: number;
  readonly clip: Rect;
}

function sheetRect(geometry: SheetGeometry, r: CellRange): Rect {
  const x = geometry.columns.offsetOf(r.firstCol);
  const y = geometry.rows.offsetOf(r.firstRow);
  return {
    x,
    y,
    width: geometry.columns.offsetOf(r.lastCol + 1) - x,
    height: geometry.rows.offsetOf(r.lastRow + 1) - y,
  };
}

const cellRect = (g: SheetGeometry, row: number, col: number): Rect =>
  sheetRect(g, { firstRow: row, lastRow: row, firstCol: col, lastCol: col });

const shift = (r: Rect, dx: number, dy: number): Rect => ({ ...r, x: r.x + dx, y: r.y + dy });

const intersect = (a: Rect, b: Rect): Rect => {
  const x = Math.max(a.x, b.x);
  const y = Math.max(a.y, b.y);
  return {
    x,
    y,
    width: Math.max(0, Math.min(a.x + a.width, b.x + b.width) - x),
    height: Math.max(0, Math.min(a.y + a.height, b.y + b.height) - y),
  };
};

const paints = (format: ResolvedCellFormat): boolean => {
  const b = format.border;
  return (
    format.fill.kind !== 'none' ||
    b.left.style !== 'none' ||
    b.right.style !== 'none' ||
    b.top.style !== 'none' ||
    b.bottom.style !== 'none' ||
    (b.diagonal.style !== 'none' && (b.diagonalUp || b.diagonalDown))
  );
};

/** The resolved cell format of an index from a row or a column, or undefined for format 0. */
function formatOfIndex(ctx: SheetContext, s: number | undefined): ResolvedCellFormat | undefined {
  if (s === undefined || s === 0) return undefined;
  return cellFormat(ctx.pkg, ctx.index, String(s)).format;
}

/** The format of one column (`x:col/@style`), or undefined. */
function columnFormatOf(ctx: SheetContext, col: number): number | undefined {
  return columnFormats(ctx.sheet.root, col, col).get(col);
}

/** Lay out one pane. */
function layoutPane(
  ctx: SheetContext,
  pane: PaneMap,
  merges: ReturnType<typeof mergedRanges>,
  drawings: readonly DrawingLayout[]
): PaneLayout {
  const g = ctx.geometry;
  const { range } = pane;
  const lastRow = Math.min(range.lastRow, range.firstRow + MAX_WINDOW - 1);
  const lastCol = Math.min(range.lastCol, range.firstCol + MAX_WINDOW - 1);
  if (lastRow < range.lastRow || lastCol < range.lastCol) record(ctx, { code: 'window-truncated' });
  const columns: AxisEntry[] = [];
  for (let c = range.firstCol; c <= lastCol; c += 1)
    if (g.columns.sizeOf(c) > 0)
      columns.push({ index: c, start: g.columns.offsetOf(c) + pane.dx, size: g.columns.sizeOf(c) });
  const rows: AxisEntry[] = [];
  for (let r = range.firstRow; r <= lastRow; r += 1)
    if (g.rows.sizeOf(r) > 0)
      rows.push({ index: r, start: g.rows.offsetOf(r) + pane.dy, size: g.rows.sizeOf(r) });

  const span: CellRange = { ...range, lastRow, lastCol };
  // The band of columns that text can show across into the window from either side.
  const band: CellRange = {
    ...span,
    firstCol: Math.max(0, span.firstCol - OVERFLOW_REACH),
    lastCol: Math.min(g.columns.count - 1, lastCol + OVERFLOW_REACH),
  };
  const inBand = mergesIn(merges, band);
  const inPane = inBand.filter((m) => m.firstCol <= lastCol && m.lastCol >= span.firstCol);
  // Rows go in ascending order, so one coverage sweep serves the pane.
  const isCovered = mergeCoverage(inBand);
  const blocked = (r: number, c: number): boolean =>
    isCovered(r, c) || hasValue(resolveCell(ctx, r, c));

  const cells: CellBox[] = [];
  const texts: CellTextBox[] = [];
  const sharedTexts = new SharedCellTexts();
  const addText = (row: number, col: number, cell: Resolved, sheetBox: Rect, overflow: boolean) => {
    const manyRuns = manyRunsOf(cell);
    let shared = manyRuns && sharedTexts.find(manyRuns, cell.format, sheetBox);
    const content = shared ? undefined : textOf(ctx, cell, sheetBox);
    if (!shared && !content) {
      // A stale formula with no text to show still gets a box with no lines, so the accessible
      // grid lists the cell with its stale mark (`data-stale`). That is a formula with no
      // cached value, which is always stale, or a formula whose cached value cannot be read,
      // on a sheet whose formula results are stale. A `t="str"` formula with a cached empty
      // string has text, so its box has one empty line.
      if (!cell.value.formula || !cell.value.stale) return;
      const clip = intersect(shift(sheetBox, pane.dx, pane.dy), pane.clip);
      if (clip.width <= 0 || clip.height <= 0) return;
      const none = { lines: [], left: 0, right: 0, areaLeft: 0, areaRight: 0, mayOverflow: false };
      texts.push({ row, col, clip, text: none, stale: true });
      return;
    }
    let own: CellTextLayout | undefined;
    let box = sheetBox;
    if (cell.format.alignment.horizontal === 'centerContinuous' && overflow) {
      let end = col;
      for (let c = col + 1; c <= col + OVERFLOW_REACH && c < g.columns.count; c += 1) {
        const next = resolveCell(ctx, row, c);
        if (hasValue(next) || next?.format.alignment.horizontal !== 'centerContinuous') break;
        end = c;
      }
      if (end > col)
        box = sheetRect(g, { firstRow: row, lastRow: row, firstCol: col, lastCol: end });
    }
    if (!shared && content) {
      const laid = layoutCellText(
        {
          runs: content.runs,
          kind: content.kind,
          ...(content.alignKind ? { alignKind: content.alignKind } : {}),
          alignment: cell.format.alignment,
          box,
          ...(content.color ? { color: content.color } : {}),
        },
        ctx.text
      );
      if (manyRuns) shared = sharedTexts.add(manyRuns, cell.format, box, laid);
      else own = laid;
    }
    const text = shared?.text ?? own!;
    let clipLeft = box.x;
    let clipRight = box.x + box.width;
    if (overflow && text.mayOverflow) {
      if (text.right > text.areaRight)
        for (let c = col + 1; c <= col + OVERFLOW_REACH && c < g.columns.count; c += 1) {
          if (blocked(row, c)) break;
          clipRight = g.columns.offsetOf(c + 1);
          if (clipRight >= text.right) break;
        }
      if (text.left < text.areaLeft)
        for (let c = col - 1; c >= Math.max(0, col - OVERFLOW_REACH); c -= 1) {
          if (blocked(row, c)) break;
          clipLeft = g.columns.offsetOf(c);
          if (clipLeft <= text.left) break;
        }
    }
    const clip = intersect(
      {
        x: clipLeft + pane.dx,
        y: box.y + pane.dy,
        width: clipRight - clipLeft,
        height: box.height,
      },
      pane.clip
    );
    if (clip.width <= 0 || clip.height <= 0) return;
    const stale = cell.value.stale;
    if (shared) {
      const inView = shared.inView(clip, pane.dx, pane.dy);
      texts.push({ row, col, clip, text: inView, stale, shownText: shared.shownText });
    } else texts.push({ row, col, clip, text: moveText(text, pane.dx, pane.dy), stale });
  };

  // The `x:col/@style` of each column of the window, read once for the pane: a blank cell reads
  // it from here, so its cost does not grow with the count of `x:col` elements.
  const windowColumns = columnFormats(ctx.sheet.root, span.firstCol, lastCol);
  const columnStyle = (c: number): number | undefined =>
    c >= span.firstCol && c <= lastCol ? windowColumns.get(c) : columnFormatOf(ctx, c);
  // The format of an empty cell, from its row or its column, when it is not the default.
  const emptyFormat = (r: number, c: number): ResolvedCellFormat | undefined =>
    formatOfIndex(ctx, rowFormat(ctx.sheet.rows.get(r)) ?? columnStyle(c));
  const formatAt = (r: number, c: number): ResolvedCellFormat | undefined =>
    resolveCell(ctx, r, c)?.format ?? emptyFormat(r, c);

  const hasRules = sheetRules(ctx.sheet.root).rules.some((r) => r.test);
  const defaultFormat = cellFormat(ctx.pkg, ctx.index, undefined).format;
  // The format of a cell without an `x:c` element under its conditional formatting rules: an
  // empty value is 0 or the empty string to a rule (`conditional-rules.pdf`, B17). Undefined
  // when nothing paints but the default.
  const blankFormat = (r: number, c: number): ResolvedCellFormat | undefined => {
    const own = emptyFormat(r, c);
    if (!hasRules) return own;
    const base = own ?? defaultFormat;
    const shown = conditionalWithin(ctx, r, c, EMPTY, base);
    return shown === base ? own : shown;
  };

  // Merged ranges: one box and the text of the top-left cell.
  for (const m of inPane) {
    const top = resolveCell(ctx, m.firstRow, m.firstCol);
    const topFormat = top?.format ?? blankFormat(m.firstRow, m.firstCol);
    const rect = sheetRect(g, m);
    if (topFormat) {
      const right = m.lastCol === m.firstCol ? topFormat : formatAt(m.firstRow, m.lastCol);
      const bottom = m.lastRow === m.firstRow ? topFormat : formatAt(m.lastRow, m.firstCol);
      const border: ResolvedBorder = {
        ...topFormat.border,
        right: right?.border.right ?? topFormat.border.right,
        bottom: bottom?.border.bottom ?? topFormat.border.bottom,
      };
      if (paints({ ...topFormat, border }))
        cells.push({
          row: m.firstRow,
          col: m.firstCol,
          rect: shift(rect, pane.dx, pane.dy),
          fill: topFormat.fill,
          border,
          merge: m,
        });
      if (top) addText(m.firstRow, m.firstCol, top, rect, false);
    }
  }

  // Columns of the window whose format paints an empty cell (`x:col/@style`).
  const paintedColumns: [number, ResolvedCellFormat][] = [];
  for (const [c, s] of windowColumns) {
    const format = formatOfIndex(ctx, s);
    if (format && paints(format) && g.columns.sizeOf(c) > 0) paintedColumns.push([c, format]);
  }

  const keys = ctx.keys;
  for (const { index: r } of rows) {
    // The nearest cell with a value on each side of the window may show across into it.
    const begin = lowerBound(keys, cellKey(r, span.firstCol));
    const end = lowerBound(keys, cellKey(r, lastCol + 1));
    for (let i = begin - 1; i >= 0; i -= 1) {
      const c = keys[i]! - cellKey(r, 0);
      if (c < band.firstCol) break;
      const covered = isCovered(r, c);
      const cell = resolveCell(ctx, r, c);
      if (!hasValue(cell) && !covered) continue;
      if (!covered) addText(r, c, cell!, cellRect(g, r, c), true);
      break;
    }
    // Empty cells that show the fill or the borders of their row or column format.
    const rowStyle = rowFormat(ctx.sheet.rows.get(r));
    const rowPaint = rowStyle === undefined ? undefined : formatOfIndex(ctx, rowStyle);
    const empties: Iterable<readonly [number, ResolvedCellFormat]> =
      rowStyle === undefined
        ? paintedColumns
        : rowPaint && paints(rowPaint)
          ? columns.map((col) => [col.index, rowPaint] as const)
          : [];
    // A row that no shown rule meets keeps the empty cells of its row and column formats.
    const rowRules = hasRules && rowHasShownRulesWithin(ctx, r);
    const blanks: Iterable<readonly [number, ResolvedCellFormat | undefined]> = rowRules
      ? columns.map((col) => [col.index, undefined] as const)
      : empties;
    for (const [c, rowOrColumn] of blanks) {
      if (ctx.sheet.cells.has(cellKey(r, c)) || isCovered(r, c) || g.columns.sizeOf(c) === 0)
        continue;
      // With rules, every visible column is a candidate, and one record holds both formats.
      const format = rowRules ? blankFormat(r, c) : rowOrColumn;
      if (!format || !paints(format)) continue;
      cells.push({
        row: r,
        col: c,
        rect: shift(cellRect(g, r, c), pane.dx, pane.dy),
        fill: format.fill,
        border: format.border,
      });
    }
    for (let i = begin; i < end; i += 1) {
      const c = keys[i]! - cellKey(r, 0);
      if (g.columns.sizeOf(c) === 0 || isCovered(r, c)) continue;
      const cell = resolveCell(ctx, r, c);
      if (!cell) continue;
      const rect = sheetRect(g, { firstRow: r, lastRow: r, firstCol: c, lastCol: c });
      if (paints(cell.format))
        cells.push({
          row: r,
          col: c,
          rect: shift(rect, pane.dx, pane.dy),
          fill: cell.format.fill,
          border: cell.format.border,
        });
      addText(r, c, cell, rect, true);
    }
    for (let i = end; i < keys.length; i += 1) {
      const c = keys[i]! - cellKey(r, 0);
      if (c > band.lastCol) break;
      const covered = isCovered(r, c);
      const cell = resolveCell(ctx, r, c);
      if (!hasValue(cell) && !covered) continue;
      if (!covered) addText(r, c, cell!, cellRect(g, r, c), true);
      break;
    }
  }

  const sheetClip = shift(pane.clip, -pane.dx, -pane.dy);
  // Closed intervals: a straight connector has no height or width, and still shows.
  const visibleDrawings = drawings
    .filter((d) => {
      const b = drawingBounds(d);
      return (
        b.x <= sheetClip.x + sheetClip.width &&
        b.x + b.width >= sheetClip.x &&
        b.y <= sheetClip.y + sheetClip.height &&
        b.y + b.height >= sheetClip.y
      );
    })
    .map((d) => ({ ...d, rect: shift(d.rect, pane.dx, pane.dy) }));

  return {
    range: span,
    clip: pane.clip,
    sheetOrigin: { x: pane.dx, y: pane.dy },
    columns,
    rows,
    cells,
    texts,
    drawings: visibleDrawings,
  };
}

/** A text layout moved by `dx`, `dy` points. */
function moveText(t: CellTextLayout, dx: number, dy: number): CellTextLayout {
  if (t.rotation)
    return {
      ...t,
      rotation: {
        ...t.rotation,
        originX: t.rotation.originX + dx,
        originY: t.rotation.originY + dy,
      },
      left: t.left + dx,
      right: t.right + dx,
      areaLeft: t.areaLeft + dx,
      areaRight: t.areaRight + dx,
    };
  return {
    ...t,
    lines: t.lines.map((l) => ({
      ...l,
      x: l.x + dx,
      baseline: l.baseline + dy,
      spans: l.spans.map((s) => ({
        ...s,
        x: s.x + dx,
        positions: s.positions.map((p) => p + dx),
      })),
    })),
    left: t.left + dx,
    right: t.right + dx,
    areaLeft: t.areaLeft + dx,
    areaRight: t.areaRight + dx,
  };
}

/** The context of one layout of a worksheet, or undefined when it is not a worksheet. */
function sheetContext(
  pkg: XlsxPackage,
  sheetName: string,
  options: LayoutOptions
): { ctx: SheetContext; part: XmlPart; metrics: GridMetrics; name: string } | undefined {
  const index = workbookIndex(pkg);
  const entry = index ? sheetByName(index, sheetName) : undefined;
  const part = xmlPartOf(pkg, entry?.partName);
  if (!index || !entry || entry.kind !== 'worksheet' || !part) return undefined;
  const { profile, measurer } = options;
  const sheet = sheetIndex(part);
  const { metrics, geometry } = geometryOf(pkg, index, sheet, measurer, profile);
  const ctx: SheetContext = {
    pkg,
    index,
    sheet,
    keys: sortedKeys(sheet),
    stale: sheetStale(index, sheet.root),
    locale: options.locale ?? EN_US,
    text: {
      measurer,
      profile,
      maxDigitWidth: metrics.maxDigitWidth,
      minDescent: fontPixels(measurer, defaultFont(pkg, index), profile).descent,
    },
    geometry,
    problems: [],
    checks: { left: LAYOUT_CHECK_BUDGET },
  };
  return { ctx, part, metrics, name: entry.name };
}

/**
 * A reader of the text that the grid shows for a `General` number in one cell of a worksheet,
 * fitted to the cell (or to the merged range that the cell anchors) as the paint fits it: a
 * narrow column shows `3.141593` for 3.14159265358. The reader gives undefined for a cell that
 * shows no number through `General`, such as a cell whose true conditional formatting rule
 * shows another number format. It builds its context on the first call, and is valid for one
 * revision of `pkg`. Each lookup of a merged range costs O(1) (`mergeAnchoredAt`).
 *
 * @internal
 */
export function shownGeneralReader(
  pkg: XlsxPackage,
  sheetName: string,
  options: LayoutOptions
): (row: number, col: number) => string | undefined {
  let opened: ReturnType<typeof sheetContext> | null = null;
  return (row, col) => {
    opened ??= sheetContext(pkg, sheetName, options);
    if (!opened) return undefined;
    const { ctx } = opened;
    const cell = resolveCell(ctx, row, col);
    if (!cell || cell.value.value.kind !== 'number') return undefined;
    if (formatCodeOf(cell.format, ctx.locale).toLowerCase() !== 'general') return undefined;
    const merge = mergeAnchoredAt(mergedRanges(ctx.sheet.root), row, col);
    const box = sheetRect(
      ctx.geometry,
      merge ?? { firstRow: row, lastRow: row, firstCol: col, lastCol: col }
    );
    return textOf(ctx, cell, box)?.runs[0]?.text;
  };
}

/**
 * The layout of the visible window of a worksheet, or undefined when the workbook has no
 * worksheet of that name. Viewport and records are in points; the grid rounds to the pixels of
 * `options.profile`.
 */
export function layoutSheet(
  pkg: XlsxPackage,
  sheetName: string,
  viewport: Viewport,
  options: LayoutOptions
): SheetLayout | undefined {
  const opened = sheetContext(pkg, sheetName, options);
  if (!opened) return undefined;
  const { ctx, part, metrics, name } = opened;
  const { index, sheet, geometry } = ctx;
  const frozen = frozenPanes(sheet.root);
  const merges = mergedRanges(sheet.root);
  if (merges.dropped > 0) record(ctx, { code: 'merges-dropped', detail: String(merges.dropped) });
  if (merges.overlapping > 0)
    record(ctx, { code: 'merges-overlapping', detail: String(merges.overlapping) });
  const drawings = sheetDrawings(
    pkg,
    drawingPartOf(pkg, part.partName, sheet.root),
    geometry,
    workbookTheme(pkg, index)
  );
  for (const d of drawings)
    if (d.kind === 'unsupported' && d.element === 'missingImage')
      record(ctx, { code: 'image-missing', detail: d.name });
  const frozenWidth = geometry.columns.offsetOf(frozen.cols);
  const frozenHeight = geometry.rows.offsetOf(frozen.rows);
  const panes = windowPanes(geometry.rows, geometry.columns, frozen, viewport).map((p) => {
    const left = p.range.firstCol < frozen.cols;
    const top = p.range.firstRow < frozen.rows;
    const clip: Rect = {
      x: left ? 0 : frozenWidth,
      y: top ? 0 : frozenHeight,
      width: Math.max(
        0,
        left ? Math.min(frozenWidth, viewport.width) : viewport.width - frozenWidth
      ),
      height: Math.max(
        0,
        top ? Math.min(frozenHeight, viewport.height) : viewport.height - frozenHeight
      ),
    };
    const map: PaneMap = {
      range: p.range,
      dx: p.x - geometry.columns.offsetOf(p.range.firstCol),
      dy: p.y - geometry.rows.offsetOf(p.range.firstRow),
      clip,
    };
    return layoutPane(ctx, map, merges, drawings);
  });
  const view = child(child(sheet.root, 'x:sheetViews'), 'x:sheetView');
  return {
    sheet: name,
    metrics,
    geometry,
    frozen,
    showGridLines: flagOf(view, 'showGridLines', true),
    printGridLines: flagOf(child(sheet.root, 'x:printOptions'), 'gridLines', false),
    frozenEdges: {
      ...(frozen.cols > 0 ? { x: Math.min(frozenWidth, viewport.width) } : {}),
      ...(frozen.rows > 0 ? { y: Math.min(frozenHeight, viewport.height) } : {}),
    },
    panes,
    problems: ctx.problems,
    conditionalCut: ctx.checks.left <= 0,
  };
}
