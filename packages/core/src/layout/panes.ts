// The visible window of a sheet (design D6): which rows and columns a viewport shows, the frozen
// panes that stay in place, and the merged ranges that become one cell rectangle.
//
// A frozen pane comes from `x:sheetView/x:pane` (`CT_Pane`, sml.xsd, ECMA-376 Part 1, 18.3.1.66)
// with `@state` (`ST_PaneState`, 18.18.53, default `split`) `frozen` or `frozenSplit`: `@xSplit`
// columns and `@ySplit` rows (each default 0) stay in place while the rest scrolls. A `split` pane
// does not freeze. The frozen band starts at `x:sheetView/@topLeftCell` (`CT_SheetView`,
// 18.3.1.87), so `@topLeftCell="A10"` with `@ySplit="2"` freezes rows 10 and 11. The text of
// 18.3.1.87 names the top-left visible cell, and then repeats the text of `x:pane/@topLeftCell`,
// so the standard does not state where the band starts: approximation "Frozen band origin". The
// rows and columns before the band do not show while the pane is frozen. The body starts at the
// first row and column after the band, at scroll 0, and scrolls from there. Merged ranges come from
// `x:mergeCells/x:mergeCell/@ref` (`CT_MergeCell`, `ST_Ref`). Both are read once per worksheet part
// revision. A merged range is never expanded into its cells: a lookup tests the ranges, so a merge
// of the whole grid costs one entry.

import type { Element } from '../store/package/tree.ts';
import { attr, child } from '../store/package/query.ts';
import { MAX_COLS, MAX_ROWS, parseCellRef } from '../store/cell-ref.ts';
import { cellKey } from '../store/workbook/sheet-index.ts';
import type { AxisIndex } from './axis.ts';

/** A cell range, 0-based and inclusive. */
export interface CellRange {
  readonly firstRow: number;
  readonly firstCol: number;
  readonly lastRow: number;
  readonly lastCol: number;
}

/**
 * Frozen rows and columns: the count of each that stays in place, and the first of each. The
 * frozen rows are `firstRow` to `firstRow + rows - 1`, and the frozen columns are `firstCol` to
 * `firstCol + cols - 1` (0-based). An axis with no frozen line has a first line of 0.
 */
export interface FrozenPanes {
  /** The count of frozen rows. */
  readonly rows: number;
  /** The count of frozen columns. */
  readonly cols: number;
  /** The first frozen row, 0-based: the row of `x:sheetView/@topLeftCell`, or 0. */
  readonly firstRow: number;
  /** The first frozen column, 0-based: the column of `x:sheetView/@topLeftCell`, or 0. */
  readonly firstCol: number;
}

/** No frozen row or column. */
export const NO_FROZEN_PANES: FrozenPanes = Object.freeze({
  rows: 0,
  cols: 0,
  firstRow: 0,
  firstCol: 0,
});

/** The frozen lines of one axis: `first` to `end - 1`, 0-based. `first === end` freezes none. */
export interface FrozenSpan {
  readonly first: number;
  readonly end: number;
}

/**
 * The frozen span of one axis. The lines before `first` do not show, and the scrolling pane
 * starts at `end`. An axis with no frozen line has the span from 0 to 0, so it hides nothing.
 */
export function frozenSpan(frozen: FrozenPanes, axis: 'rows' | 'cols'): FrozenSpan {
  const count = axis === 'rows' ? frozen.rows : frozen.cols;
  if (count <= 0) return { first: 0, end: 0 };
  const first = axis === 'rows' ? frozen.firstRow : frozen.firstCol;
  return { first, end: first + count };
}

/** At most this many merged ranges are read from one sheet; the rest are a problem. */
export const MAX_MERGES = 100_000;

/**
 * The frozen panes of the first `x:sheetView` of a worksheet (`CT_SheetView`), or zero rows and
 * columns. The band starts at `x:sheetView/@topLeftCell` (ECMA-376 Part 1, 18.3.1.87; approximation
 * "Frozen band origin"); a cell reference that does not parse, or no reference, starts it at row 1
 * and column A. Each count is cut so that at least one row and one column of the body stay in the
 * grid after the band, as a structural edit cuts a band that grows (`store/structure`). An axis
 * with no room for the body freezes nothing.
 */
export function frozenPanes(worksheet: Element): FrozenPanes {
  const view = child(child(worksheet, 'x:sheetViews'), 'x:sheetView');
  const pane = child(view, 'x:pane');
  const state = attr(pane, 'state');
  if (!pane || (state !== 'frozen' && state !== 'frozenSplit')) return NO_FROZEN_PANES;
  const top = parseCellRef(attr(view, 'topLeftCell') ?? '');
  const split = (name: string, first: number, max: number) => {
    const n = Number(attr(pane, name) ?? '0');
    return Number.isInteger(n) && n > 0 ? Math.max(0, Math.min(n, max - 1 - first)) : 0;
  };
  const rows = split('ySplit', top?.row ?? 0, MAX_ROWS);
  const cols = split('xSplit', top?.col ?? 0, MAX_COLS);
  return {
    rows,
    cols,
    firstRow: rows > 0 ? (top?.row ?? 0) : 0,
    firstCol: cols > 0 ? (top?.col ?? 0) : 0,
  };
}

/** A range reference such as `B2:D4`, or a single cell, as a range; undefined when invalid. */
export function parseRange(ref: string): CellRange | undefined {
  if (ref.length > 21) return undefined;
  const colon = ref.indexOf(':');
  const a = parseCellRef(colon < 0 ? ref : ref.slice(0, colon));
  const b = colon < 0 ? a : parseCellRef(ref.slice(colon + 1));
  if (!a || !b) return undefined;
  return {
    firstRow: Math.min(a.row, b.row),
    firstCol: Math.min(a.col, b.col),
    lastRow: Math.max(a.row, b.row),
    lastCol: Math.max(a.col, b.col),
  };
}

/** The merged ranges of a worksheet. */
export interface MergeIndex {
  /** Sorted by first row, then first column. */
  readonly ranges: readonly CellRange[];
  /** Merges past {@link MAX_MERGES}, or with a reference that does not parse. */
  readonly dropped: number;
  /** Merges skipped because their top-left cell lies inside an earlier merge. */
  readonly overlapping: number;
}

const mergeCache = new WeakMap<Element, MergeIndex>();

/** The position of the last range in `sorted` (by first column) that starts at or before `col`. */
function lastStartingBy(sorted: readonly CellRange[], col: number): number {
  let lo = 0;
  let hi = sorted.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (sorted[mid]!.firstCol <= col) lo = mid + 1;
    else hi = mid;
  }
  return lo - 1;
}

/**
 * Drop each merge whose top-left cell lies inside an earlier one (merges overlap only in an
 * invalid file). A sweep over the rows keeps the kept merges that reach the current row, sorted
 * by first column, so the cost depends on the merges, never on the cells they cover.
 */
function withoutOverlaps(sorted: readonly CellRange[]): { kept: CellRange[]; skipped: number } {
  const kept: CellRange[] = [];
  let active: CellRange[] = [];
  let minLast = Number.POSITIVE_INFINITY;
  let skipped = 0;
  for (const m of sorted) {
    if (minLast < m.firstRow) {
      active = active.filter((a) => a.lastRow >= m.firstRow);
      minLast = Number.POSITIVE_INFINITY;
      for (const a of active) minLast = Math.min(minLast, a.lastRow);
    }
    const at = lastStartingBy(active, m.firstCol);
    if (at >= 0 && active[at]!.lastCol >= m.firstCol) {
      skipped += 1;
      continue;
    }
    kept.push(m);
    active.splice(at + 1, 0, m);
    minLast = Math.min(minLast, m.lastRow);
  }
  return { kept, skipped };
}

/** The merged ranges of a worksheet root, sorted by first row. Cached per root. */
export function mergedRanges(worksheet: Element): MergeIndex {
  const hit = mergeCache.get(worksheet);
  if (hit) return hit;
  const ranges: CellRange[] = [];
  let dropped = 0;
  for (const node of child(worksheet, 'x:mergeCells')?.children ?? []) {
    if (node.type !== 'element' || node.local !== 'mergeCell') continue;
    const range = ranges.length < MAX_MERGES ? parseRange(attr(node, 'ref') ?? '') : undefined;
    if (range && (range.lastRow > range.firstRow || range.lastCol > range.firstCol))
      ranges.push(range);
    else dropped += 1;
  }
  ranges.sort((a, b) => a.firstRow - b.firstRow || a.firstCol - b.firstCol);
  const { kept, skipped } = withoutOverlaps(ranges);
  const index = { ranges: kept, dropped, overlapping: skipped };
  mergeCache.set(worksheet, index);
  return index;
}

const anchorCache = new WeakMap<MergeIndex, Map<number, CellRange>>();

/**
 * The merged range whose top-left cell is at `row` and `col`, or undefined. The map of anchors
 * is built once for each merge index, which is cached for each worksheet root, so a lookup
 * costs O(1) however many merges the sheet has.
 */
export function mergeAnchoredAt(
  index: MergeIndex,
  row: number,
  col: number
): CellRange | undefined {
  let anchors = anchorCache.get(index);
  if (!anchors) {
    anchors = new Map();
    for (const m of index.ranges) anchors.set(cellKey(m.firstRow, m.firstCol), m);
    anchorCache.set(index, anchors);
  }
  return anchors.get(cellKey(row, col));
}

/** The count of ranges whose first row is at most `row`. */
function startingBy(index: MergeIndex, row: number): number {
  let lo = 0;
  let hi = index.ranges.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (index.ranges[mid]!.firstRow <= row) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

/**
 * The merged ranges that meet `area`, sorted by first row. `work`, for tests, counts the ranges
 * that the search reads.
 */
export function mergesIn(
  index: MergeIndex,
  area: CellRange,
  work?: { mergesRead: number }
): CellRange[] {
  const out: CellRange[] = [];
  const end = startingBy(index, area.lastRow);
  if (work) work.mergesRead += end;
  for (let i = 0; i < end; i += 1) {
    const m = index.ranges[i]!;
    if (m.lastRow >= area.firstRow && m.firstCol <= area.lastCol && m.lastCol >= area.firstCol)
      out.push(m);
  }
  return out;
}

/** The merged range whose top-left cell is (`row`, `col`), or the range that covers it. */
export function mergeAt(index: MergeIndex, row: number, col: number): CellRange | undefined {
  const end = startingBy(index, row);
  for (let i = 0; i < end; i += 1) {
    const m = index.ranges[i]!;
    if (row <= m.lastRow && col >= m.firstCol && col <= m.lastCol) return m;
  }
  return undefined;
}

/**
 * Answers whether a cell lies inside one of `merges` (sorted by first row), one row at a time
 * in ascending order. Each row keeps the merges that reach it, sorted by first column, with the
 * running maximum of their last columns, so a test is a binary search: the cost depends on the
 * merges, never on the cells they cover.
 */
export function mergeCoverage(merges: readonly CellRange[]): (row: number, col: number) => boolean {
  let next = 0;
  let row = -1;
  let active: CellRange[] = [];
  let reach: number[] = [];
  // The smallest last row of the active merges: no merge ends before a row past it.
  let minLast = Number.POSITIVE_INFINITY;
  const advance = (r: number) => {
    if (r === row) return;
    if (r < row) {
      next = 0;
      active = [];
      minLast = Number.POSITIVE_INFINITY;
    }
    let entered = false;
    while (next < merges.length && merges[next]!.firstRow <= r) {
      active.push(merges[next]!);
      next += 1;
      entered = true;
    }
    row = r;
    // Filter only when a merge can have ended, or when merges entered (one that starts before a
    // skipped row can end before `r`). Sort only when merges entered: the filter keeps the
    // order. A row with no merge that starts or ends costs nothing, however many reach it.
    if (!entered && r <= minLast) return;
    active = active.filter((m) => m.lastRow >= r);
    if (entered) active.sort((a, b) => a.firstCol - b.firstCol);
    reach = [];
    let max = -1;
    minLast = Number.POSITIVE_INFINITY;
    for (const m of active) {
      reach.push((max = Math.max(max, m.lastCol)));
      minLast = Math.min(minLast, m.lastRow);
    }
  };
  return (r, c) => {
    advance(r);
    const at = lastStartingBy(active, c);
    return at >= 0 && reach[at]! >= c;
  };
}

/** The indexes that a span of an axis shows: from `start` points, `length` points long. */
export function visibleSpan(
  axis: AxisIndex,
  start: number,
  length: number
): { first: number; last: number } {
  const first = axis.indexAt(start);
  const last = axis.indexAt(start + Math.max(0, length) - 1e-6);
  return { first, last: Math.max(first, last) };
}

/** One pane of the window: a cell range and the offset that its top-left cell draws at. */
export interface Pane {
  readonly range: CellRange;
  /** The point position of the range's top-left corner inside the viewport. */
  readonly x: number;
  readonly y: number;
}

/**
 * The panes that a viewport of `width` by `height` points shows, scrolled by `scrollX` and
 * `scrollY` points in the scrolling pane. Frozen rows and columns draw first, at the top and
 * left, and do not scroll. The scroll position starts after the frozen rows and columns. The
 * rows and columns before a frozen band do not show.
 */
export function windowPanes(
  rows: AxisIndex,
  columns: AxisIndex,
  frozen: FrozenPanes,
  viewport: {
    readonly width: number;
    readonly height: number;
    readonly scrollX: number;
    readonly scrollY: number;
  }
): Pane[] {
  const r = frozenSpan(frozen, 'rows');
  const c = frozenSpan(frozen, 'cols');
  const frozenHeight = rows.offsetOf(r.end) - rows.offsetOf(r.first);
  const frozenWidth = columns.offsetOf(c.end) - columns.offsetOf(c.first);
  const bodyY = rows.offsetOf(r.end) + Math.max(0, viewport.scrollY);
  const bodyX = columns.offsetOf(c.end) + Math.max(0, viewport.scrollX);
  const bodyRows = visibleSpan(rows, bodyY, viewport.height - frozenHeight);
  const bodyCols = visibleSpan(columns, bodyX, viewport.width - frozenWidth);
  const panes: Pane[] = [];
  const pane = (
    r: { first: number; last: number },
    c: { first: number; last: number },
    x: number,
    y: number
  ) =>
    panes.push({
      range: { firstRow: r.first, firstCol: c.first, lastRow: r.last, lastCol: c.last },
      x,
      y,
    });
  const topRows = { first: r.first, last: r.end - 1 };
  const leftCols = { first: c.first, last: c.end - 1 };
  const shiftX = frozenWidth - (bodyX - columns.offsetOf(bodyCols.first));
  const shiftY = frozenHeight - (bodyY - rows.offsetOf(bodyRows.first));
  if (frozen.rows > 0 && frozen.cols > 0) pane(topRows, leftCols, 0, 0);
  if (frozen.rows > 0) pane(topRows, bodyCols, shiftX, 0);
  if (frozen.cols > 0) pane(bodyRows, leftCols, 0, shiftY);
  pane(bodyRows, bodyCols, shiftX, shiftY);
  return panes;
}
