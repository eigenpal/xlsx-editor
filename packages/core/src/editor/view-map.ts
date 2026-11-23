// The map between the painted window and the grid: which cell a point of the window shows, and
// which rectangles of the window a range covers. The window has up to four panes (frozen rows,
// frozen columns, both, and the scrolling body), as `layoutSheet` paints it. Points throughout.

import type { GridPoint, GridRange } from '../contracts/index.ts';
import type { AxisIndex, FrozenPanes, Rect, SheetGeometry } from '../layout/index.ts';

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

/** The index at `offset` of the window on one axis. */
function indexAt(axis: AxisIndex, frozen: number, scroll: number, offset: number): number {
  const frozenSize = axis.offsetOf(frozen);
  if (offset < frozenSize) return axis.indexAt(offset);
  return axis.indexAt(offset - frozenSize + axis.offsetOf(frozen) + Math.max(0, scroll));
}

/** The cell at a point of the window. */
export function cellAtPoint(view: ViewState, x: number, y: number): GridPoint {
  return {
    row: indexAt(view.geometry.rows, view.frozen.rows, view.scrollY, y),
    col: indexAt(view.geometry.columns, view.frozen.cols, view.scrollX, x),
  };
}

/** The intervals of the window that indexes `first` to `last` cover on one axis. */
export function axisSpans(
  axis: AxisIndex,
  frozen: number,
  scroll: number,
  size: number,
  first: number,
  last: number
): [number, number][] {
  const out: [number, number][] = [];
  const frozenSize = axis.offsetOf(frozen);
  if (first < frozen) {
    const a = axis.offsetOf(first);
    const b = axis.offsetOf(Math.min(last, frozen - 1) + 1);
    if (b > a) out.push([a, Math.min(b, size)]);
  }
  if (last >= frozen) {
    const origin = frozenSize - axis.offsetOf(frozen) - Math.max(0, scroll);
    const a = Math.max(frozenSize, origin + axis.offsetOf(Math.max(first, frozen)));
    const b = Math.min(size, origin + axis.offsetOf(last + 1));
    if (b > a) out.push([a, b]);
  }
  return out;
}

/** The rectangles of the window that a range covers (one for each pane it shows in). */
export function rangeRects(view: ViewState, range: GridRange): Rect[] {
  const xs = axisSpans(
    view.geometry.columns,
    view.frozen.cols,
    view.scrollX,
    view.width,
    range.start.col,
    range.end.col
  );
  const ys = axisSpans(
    view.geometry.rows,
    view.frozen.rows,
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
  frozen: number,
  scroll: number,
  size: number,
  first: number,
  last: number
): { readonly start: number | null; readonly end: number | null } {
  const frozenSize = axis.offsetOf(frozen);
  const at = (line: number, inFrozen: boolean) =>
    inFrozen ? axis.offsetOf(line) : axis.offsetOf(line) - Math.max(0, scroll);
  const start = at(first, first < frozen);
  const end = at(last + 1, last < frozen);
  const startShows = start < size && (first < frozen || start >= frozenSize);
  const endShows = end <= size && (last < frozen || end > frozenSize);
  return { start: startShows ? start : null, end: endShows ? end : null };
}

/**
 * The edges of the outline of a range that show in the window. An edge that a pane seam or the
 * end of the window cuts stays open: only the true edges of the range have a line.
 */
export function rangeOutline(view: ViewState, range: GridRange): OutlineEdge[] {
  const { columns, rows } = view.geometry;
  const { start, end } = range;
  const xs = axisSpans(columns, view.frozen.cols, view.scrollX, view.width, start.col, end.col);
  const ys = axisSpans(rows, view.frozen.rows, view.scrollY, view.height, start.row, end.row);
  if (xs.length === 0 || ys.length === 0) return [];
  const across = axisEnds(columns, view.frozen.cols, view.scrollX, view.width, start.col, end.col);
  const down = axisEnds(rows, view.frozen.rows, view.scrollY, view.height, start.row, end.row);
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
 * `CT_Pane` and `CT_SheetView`, sml.xsd).
 */
export function scrollToTopLeft(
  geometry: SheetGeometry,
  frozen: FrozenPanes,
  topLeft: GridPoint
): { x: number; y: number } {
  return {
    x: Math.max(0, geometry.columns.offsetOf(topLeft.col) - geometry.columns.offsetOf(frozen.cols)),
    y: Math.max(0, geometry.rows.offsetOf(topLeft.row) - geometry.rows.offsetOf(frozen.rows)),
  };
}

/**
 * The scroll that shows a cell in the body pane, or the current scroll when it shows already.
 * A cell in a frozen pane never scrolls.
 */
export function scrollToShow(view: ViewState, cell: GridPoint): { x: number; y: number } {
  const axis = (a: AxisIndex, frozen: number, scroll: number, size: number, i: number): number => {
    if (i < frozen) return scroll;
    const frozenSize = a.offsetOf(frozen);
    const start = a.offsetOf(i) - a.offsetOf(frozen);
    const end = start + a.sizeOf(i);
    const room = Math.max(0, size - frozenSize);
    if (start < scroll) return start;
    if (end > scroll + room) return Math.max(0, Math.min(start, end - room));
    return scroll;
  };
  return {
    x: axis(view.geometry.columns, view.frozen.cols, view.scrollX, view.width, cell.col),
    y: axis(view.geometry.rows, view.frozen.rows, view.scrollY, view.height, cell.row),
  };
}

/** Where indexes `first` to `last` are on one axis of the window, not cut to the window. */
function axisPlace(
  axis: AxisIndex,
  frozen: number,
  scroll: number,
  first: number,
  last: number
): [number, number] {
  const size = axis.offsetOf(last + 1) - axis.offsetOf(first);
  if (first < frozen) return [axis.offsetOf(first), size];
  const frozenSize = axis.offsetOf(frozen);
  return [frozenSize + axis.offsetOf(first) - axis.offsetOf(frozen) - Math.max(0, scroll), size];
}

/**
 * The whole rectangle of a range in the window, also where it is partly or wholly out of the
 * window: the cell editor covers the whole cell. A range that starts in a frozen pane is placed
 * by its frozen start.
 */
export function rangeRect(view: ViewState, range: GridRange): Rect {
  const [x, width] = axisPlace(
    view.geometry.columns,
    view.frozen.cols,
    view.scrollX,
    range.start.col,
    range.end.col
  );
  const [y, height] = axisPlace(
    view.geometry.rows,
    view.frozen.rows,
    view.scrollY,
    range.start.row,
    range.end.row
  );
  return { x, y, width, height };
}
