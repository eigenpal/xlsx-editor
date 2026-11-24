// The map between the painted window and the grid: which cell a point of the window shows, and
// which rectangles of the window a range covers. The window has up to four panes (frozen rows,
// frozen columns, both, and the scrolling body), as `layoutSheet` paints it. Points throughout.
//
// The frozen band of an axis is the span `first` to `end - 1` (`frozenSpan`): it starts at the
// line of `x:sheetView/@topLeftCell` (`CT_SheetView`, ECMA-376 Part 1, 18.3.1.87), and holds
// `x:pane/@ySplit` rows or `@xSplit` columns (`CT_Pane`, 18.3.1.66), as `layout/panes.ts` reads
// them (approximation "Frozen band origin"). The band draws from the start of the window. The
// lines before the band do not show. The body draws after the band, from line `end` at scroll 0.

import type { GridPoint, GridRange } from '../contracts/index.ts';
import type { AxisIndex, FrozenPanes, Rect, SheetGeometry } from '../layout/index.ts';
import { frozenSpan, type FrozenSpan } from '../layout/panes.ts';

/** What the map needs: the geometry, the frozen panes, and the scroll of the body. */
export interface ViewState {
  readonly geometry: SheetGeometry;
  readonly frozen: FrozenPanes;
  /** The scroll of the body pane past the frozen panes, in points. */
  readonly scrollX: number;
  readonly scrollY: number;
  /** The size of the window, in points. */
  readonly width: number;
  readonly height: number;
}

/**
 * The view state of a laid-out sheet: its geometry and frozen panes, the scroll of a sheet view,
 * and a window of `width` by `height` points.
 */
export function viewStateOf(
  layout: { readonly geometry: SheetGeometry; readonly frozen: FrozenPanes },
  view: { readonly scrollX: number; readonly scrollY: number },
  width: number,
  height: number
): ViewState {
  const { geometry, frozen } = layout;
  return { geometry, frozen, scrollX: view.scrollX, scrollY: view.scrollY, width, height };
}

/**
 * Scroll `view` so that `cell` shows in a stage of `box` (CSS pixels, `scale` for each point),
 * with the headings outside the grid area.
 */
export function scrollViewToShow(
  layout: { readonly geometry: SheetGeometry; readonly frozen: FrozenPanes },
  view: { scrollX: number; scrollY: number },
  box: {
    readonly scale: number;
    readonly width: number;
    readonly height: number;
    readonly headWidth: number;
    readonly headHeight: number;
  },
  cell: GridPoint
): void {
  const width = (box.width - box.headWidth) / box.scale;
  const next = scrollToShow(
    viewStateOf(layout, view, width, (box.height - box.headHeight) / box.scale),
    cell
  );
  view.scrollX = next.x;
  view.scrollY = next.y;
}

/** The size of the frozen band of one axis, in points. */
const bandSize = (axis: AxisIndex, band: FrozenSpan): number =>
  axis.offsetOf(band.end) - axis.offsetOf(band.first);

/** The index at `offset` of the window on one axis. */
function indexAt(axis: AxisIndex, band: FrozenSpan, scroll: number, offset: number): number {
  const frozenSize = bandSize(axis, band);
  if (offset < frozenSize) return axis.indexAt(offset + axis.offsetOf(band.first));
  return axis.indexAt(offset - frozenSize + axis.offsetOf(band.end) + Math.max(0, scroll));
}

/** The cell at a point of the window. */
export function cellAtPoint(view: ViewState, x: number, y: number): GridPoint {
  return {
    row: indexAt(view.geometry.rows, frozenSpan(view.frozen, 'rows'), view.scrollY, y),
    col: indexAt(view.geometry.columns, frozenSpan(view.frozen, 'cols'), view.scrollX, x),
  };
}

/**
 * The intervals of the window that indexes `first` to `last` cover on one axis. The indexes
 * before the frozen band cover nothing.
 */
export function axisSpans(
  axis: AxisIndex,
  band: FrozenSpan,
  scroll: number,
  size: number,
  first: number,
  last: number
): [number, number][] {
  const out: [number, number][] = [];
  const frozenSize = bandSize(axis, band);
  const bandStart = axis.offsetOf(band.first);
  if (first < band.end) {
    const a = axis.offsetOf(Math.max(first, band.first)) - bandStart;
    const b = axis.offsetOf(Math.min(last, band.end - 1) + 1) - bandStart;
    if (b > a) out.push([a, Math.min(b, size)]);
  }
  if (last >= band.end) {
    const origin = frozenSize - axis.offsetOf(band.end) - Math.max(0, scroll);
    const a = Math.max(frozenSize, origin + axis.offsetOf(Math.max(first, band.end)));
    const b = Math.min(size, origin + axis.offsetOf(last + 1));
    if (b > a) out.push([a, b]);
  }
  return out;
}

/** The rectangles of the window that a range covers (one for each pane it shows in). */
export function rangeRects(view: ViewState, range: GridRange): Rect[] {
  const xs = axisSpans(
    view.geometry.columns,
    frozenSpan(view.frozen, 'cols'),
    view.scrollX,
    view.width,
    range.start.col,
    range.end.col
  );
  const ys = axisSpans(
    view.geometry.rows,
    frozenSpan(view.frozen, 'rows'),
    view.scrollY,
    view.height,
    range.start.row,
    range.end.row
  );
  const rects: Rect[] = [];
  for (const [x0, x1] of xs)
    for (const [y0, y1] of ys) rects.push({ x: x0, y: y0, width: x1 - x0, height: y1 - y0 });
  return rects;
}

/** One edge of the outline of a range in the window, in points. */
export interface OutlineEdge {
  readonly side: 'top' | 'right' | 'bottom' | 'left';
  /** The left or top end of the edge, on the outer line of the range. */
  readonly x: number;
  readonly y: number;
  readonly length: number;
}

/** Where the first and the last line of indexes `first` to `last` show, or null for neither. */
function axisEnds(
  axis: AxisIndex,
  band: FrozenSpan,
  scroll: number,
  size: number,
  first: number,
  last: number
): { readonly start: number | null; readonly end: number | null } {
  const frozenSize = bandSize(axis, band);
  const inBand = (i: number) => i >= band.first && i < band.end;
  // An index before the band does not show, so neither does its line.
  const before = (i: number) => i < band.first;
  const at = (line: number, frozen: boolean) =>
    axis.offsetOf(line) - axis.offsetOf(band.first) - (frozen ? 0 : Math.max(0, scroll));
  const start = at(first, inBand(first));
  const end = at(last + 1, inBand(last));
  const startShows = !before(first) && start < size && (inBand(first) || start >= frozenSize);
  const endShows = !before(last) && end <= size && (inBand(last) || end > frozenSize);
  return { start: startShows ? start : null, end: endShows ? end : null };
}

/**
 * The edges of the outline of a range that show in the window. An edge that a pane seam or the
 * end of the window cuts stays open: only the true edges of the range have a line.
 */
export function rangeOutline(view: ViewState, range: GridRange): OutlineEdge[] {
  const { columns, rows } = view.geometry;
  const { start, end } = range;
  const bandCols = frozenSpan(view.frozen, 'cols');
  const bandRows = frozenSpan(view.frozen, 'rows');
  const xs = axisSpans(columns, bandCols, view.scrollX, view.width, start.col, end.col);
  const ys = axisSpans(rows, bandRows, view.scrollY, view.height, start.row, end.row);
  if (xs.length === 0 || ys.length === 0) return [];
  const across = axisEnds(columns, bandCols, view.scrollX, view.width, start.col, end.col);
  const down = axisEnds(rows, bandRows, view.scrollY, view.height, start.row, end.row);
  const edges: OutlineEdge[] = [];
  for (const [x0, x1] of xs) {
    const length = x1 - x0;
    if (down.start !== null) edges.push({ side: 'top', x: x0, y: down.start, length });
    if (down.end !== null) edges.push({ side: 'bottom', x: x0, y: down.end, length });
  }
  for (const [y0, y1] of ys) {
    const length = y1 - y0;
    if (across.start !== null) edges.push({ side: 'left', x: across.start, y: y0, length });
    if (across.end !== null) edges.push({ side: 'right', x: across.end, y: y0, length });
  }
  return edges;
}

/**
 * The scroll offsets, in points, that put `topLeft` at the top-left corner of the scrolling
 * pane: the opening scroll of a sheet (`x:pane/@topLeftCell` or `x:sheetView/@topLeftCell`;
 * `CT_Pane` and `CT_SheetView`, sml.xsd). The scrolling pane starts after the frozen band, so a
 * cell before the end of the band gives a scroll of 0.
 */
export function scrollToTopLeft(
  geometry: SheetGeometry,
  frozen: FrozenPanes,
  topLeft: GridPoint
): { x: number; y: number } {
  const { columns, rows } = geometry;
  const endCol = frozenSpan(frozen, 'cols').end;
  const endRow = frozenSpan(frozen, 'rows').end;
  return {
    x: Math.max(0, columns.offsetOf(topLeft.col) - columns.offsetOf(endCol)),
    y: Math.max(0, rows.offsetOf(topLeft.row) - rows.offsetOf(endRow)),
  };
}

/**
 * The scroll that shows a cell in the body pane, or the current scroll when it shows already.
 * A cell in a frozen pane never scrolls. A cell before the frozen band cannot show, so it keeps
 * the scroll too.
 */
export function scrollToShow(view: ViewState, cell: GridPoint): { x: number; y: number } {
  const axis = (a: AxisIndex, band: FrozenSpan, scroll: number, size: number, i: number) => {
    if (i < band.end) return scroll;
    const frozenSize = bandSize(a, band);
    const start = a.offsetOf(i) - a.offsetOf(band.end);
    const end = start + a.sizeOf(i);
    const room = Math.max(0, size - frozenSize);
    if (start < scroll) return start;
    if (end > scroll + room) return Math.max(0, Math.min(start, end - room));
    return scroll;
  };
  const { columns, rows } = view.geometry;
  return {
    x: axis(columns, frozenSpan(view.frozen, 'cols'), view.scrollX, view.width, cell.col),
    y: axis(rows, frozenSpan(view.frozen, 'rows'), view.scrollY, view.height, cell.row),
  };
}

/** Where indexes `first` to `last` are on one axis of the window, not cut to the window. */
function axisPlace(
  axis: AxisIndex,
  band: FrozenSpan,
  scroll: number,
  first: number,
  last: number
): [number, number] {
  const size = axis.offsetOf(last + 1) - axis.offsetOf(first);
  if (first < band.end) return [axis.offsetOf(first) - axis.offsetOf(band.first), size];
  const body = bandSize(axis, band) - axis.offsetOf(band.end) - Math.max(0, scroll);
  return [body + axis.offsetOf(first), size];
}

/**
 * The whole rectangle of a range in the window, also where it is partly or wholly out of the
 * window: the cell editor covers the whole cell. A range that starts in a frozen pane is placed
 * by its frozen start.
 */
export function rangeRect(view: ViewState, range: GridRange): Rect {
  const [x, width] = axisPlace(
    view.geometry.columns,
    frozenSpan(view.frozen, 'cols'),
    view.scrollX,
    range.start.col,
    range.end.col
  );
  const [y, height] = axisPlace(
    view.geometry.rows,
    frozenSpan(view.frozen, 'rows'),
    view.scrollY,
    range.start.row,
    range.end.row
  );
  return { x, y, width, height };
}
