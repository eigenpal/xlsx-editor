// The resize of columns and rows from the headings (design D3 of the change
// `workbook-structure-editing`): the hit band on a heading edge, the drag with its guide line,
// and the commit as one transaction of `setColumnWidth` or `setRowHeight` ops.
//
// The drag follows the table resize of DOCX Editor (`eigenpal/docx-editor` at f3116761,
// `packages/core/src/editor/surface-table-interaction.ts`): a band of 3 CSS pixels on each side
// of the edge (the 6 px handles of :413-452), pointer capture on the press (:574), Escape and
// pointer cancel cancel the drag (:869-877), a drag with no movement commits nothing
// (:633-640), and pointer up commits one op (`table-command-plan.ts:665-674`). The size has a
// floor of 0, as a column or a row of size 0 is valid, where DOCX Editor keeps a minimum row
// height (:156-161), and a ceiling at the largest size that the store accepts (`axis-size.ts`).
//
// The new edge converts back to the file with the inverse of the layout conversion: whole
// pixels of the layout grid (`metrics.pixelsPerInch`, which is CSS pixels at the zoom) become
// a width in characters of the maximum digit width (`columnWidthForPixels`, 18.3.1.13) or a
// height in points after the row scale (`rowHeightForPixels`). Every element comes from the
// document of the stage at use time (`host-realm.ts`).
//
// An edge that a column or row of size 0 shares with the next one hits the size-0 one when the
// pointer is right of (or below) the edge, and the one before it otherwise, so a column of
// width 0 can be dragged wider again. A hidden column or row (`@hidden`) is not hit this way:
// a new size does not show it.
//
// A press takes the painted state before the open cell edit commits: the commit clears the
// layout until the next paint, and the drag keeps the geometry of the press.
//
// A drag belongs to the sheet and the store revision of its press: when the shown sheet or the
// revision changes (a sheet tab, an undo key, or another edit), the drag is cancelled and writes
// nothing, so it never resizes another sheet. A zoom or a switch to view mode cancels it too, and
// its guide goes at once. In view mode, no heading edge is hit, and neither is the edge of a
// column or a row whose size the protection of the sheet locks (`@formatColumns`,
// `@formatRows`). A scroll during the drag moves the guide with the edge of the dragged line
// while the line shows.

import { selectionRanges } from '../automation/structure-commands.ts';
import type { GridSelection } from '../contracts/index.ts';
import {
  attr,
  children,
  sheetIndex,
  MAX_COLS,
  MAX_COLUMN_WIDTH,
  MAX_ROW_HEIGHT,
  MAX_ROWS,
  type XmlPart,
} from '../store/index.ts';
import {
  columnPixels,
  columnWidthForPixels,
  rowHeightForPixels,
  type FrozenPanes,
  type GridMetrics,
  type SheetGeometry,
} from '../layout/index.ts';
import type { CellEditing } from './cell-editing.ts';
import type { EditResult } from './cell-input.ts';
import type { EditorSnapshot } from './editor-model.ts';
import type { StageBox } from './grid-surface.ts';
import { windowScale } from './host-frame.ts';
import { axisSpans, cellAtPoint, viewStateOf, type ViewState } from './view-map.ts';

/** CSS pixels on each side of a heading edge that hit the resize band. */
export const RESIZE_BAND_PX = 3;

/** The axis of a resize. */
export type ResizeAxis = 'columns' | 'rows';

/** A heading edge: the column or row (0-based) whose right or bottom edge it is. */
export interface ResizeEdge {
  readonly axis: ResizeAxis;
  readonly index: number;
}

/**
 * Run one resize as one transaction and one undo step: `first` to `last` are 0-based, as in the
 * editor instance, and `size` is a width in characters or a height in points. The store refuses a
 * size or a range outside its limits (`axis-size.ts`).
 */
export function runResize(
  edits: CellEditing,
  sheet: string | undefined,
  axis: ResizeAxis,
  first: number,
  last: number,
  size: number
): EditResult {
  if (sheet === undefined) return edits.refuse({ code: 'not-editable' });
  const range = { sheet, first: first + 1, last: last + 1 };
  return edits.run([
    axis === 'columns'
      ? { op: 'setColumnWidth', ...range, width: size }
      : { op: 'setRowHeight', ...range, height: size },
  ]);
}

/** Whether the column or row at a 0-based index has `@hidden`. */
export type HiddenLine = (axis: ResizeAxis, index: number) => boolean;

const on = (v: string | undefined) => v === '1' || v === 'true';

/**
 * The `@hidden` flags of the columns (`x:col`) and rows (`x:row`) of a worksheet part. A part
 * can hold more than one `x:cols` element: each is read.
 */
export function hiddenLines(part: XmlPart | undefined): HiddenLine {
  return (axis, index) => {
    if (!part) return false;
    if (axis === 'rows') return on(attr(sheetIndex(part).rows.get(index), 'hidden'));
    for (const col of children(part.root, 'x:cols').flatMap((cols) => cols.children)) {
      if (col.type !== 'element' || col.local !== 'col') continue;
      const min = Number(attr(col, 'min'));
      const max = Number(attr(col, 'max'));
      if (index + 1 >= min && index + 1 <= max) return on(attr(col, 'hidden'));
    }
    return false;
  };
}

/** The edge at the window point `at` (points) of one axis, or null. */
function edgeOnAxis(
  map: ViewState,
  axis: ResizeAxis,
  at: number,
  band: number,
  hidden?: HiddenLine
): number | null {
  const columns = axis === 'columns';
  const sizes = columns ? map.geometry.columns : map.geometry.rows;
  const index = (p: number) => {
    const cell = cellAtPoint(map, columns ? p : 0, columns ? 0 : p);
    return columns ? cell.col : cell.row;
  };
  const span = (i: number) =>
    axisSpans(
      columns ? map.geometry.columns : map.geometry.rows,
      columns ? map.frozen.cols : map.frozen.rows,
      columns ? map.scrollX : map.scrollY,
      columns ? map.width : map.height,
      i,
      i
    ).find(([a, b]) => at >= a - band && at <= b + band);
  const i = index(at);
  const s = span(i);
  if (!s) return null;
  if (s[1] - at <= band) return i;
  if (at - s[0] > band) return null;
  // Right of (below) an edge that a line of size 0 shares: that line, so it can widen again.
  if (at >= s[0] && i > 0 && sizes.sizeOf(i - 1) === 0 && !hidden?.(axis, i - 1)) return i - 1;
  if (s[0] <= 0) return null;
  // The left or top edge of a column or row is the far edge of the one shown before it.
  const before = index(s[0] - 1e-6);
  return before !== i && span(before) ? before : null;
}

/** The heading edge under a stage point (CSS pixels), or null. */
export function resizeEdgeAt(
  map: ViewState,
  box: StageBox,
  x: number,
  y: number,
  hidden?: HiddenLine
): ResizeEdge | null {
  const band = RESIZE_BAND_PX / box.scale;
  if (y < box.headHeight && x >= box.headWidth) {
    const index = edgeOnAxis(map, 'columns', (x - box.headWidth) / box.scale, band, hidden);
    return index === null ? null : { axis: 'columns', index };
  }
  if (x < box.headWidth && y >= box.headHeight) {
    const index = edgeOnAxis(map, 'rows', (y - box.headHeight) / box.scale, band, hidden);
    return index === null ? null : { axis: 'rows', index };
  }
  return null;
}

/**
 * The columns or rows that a drag of `edge` resizes: every column (row) of a selected band of
 * whole columns (rows) that holds it, or the one column (row). In a selection of several ranges,
 * it is the band that holds the edge, so the result is one band of adjacent indexes.
 */
export function resizeBand(edge: ResizeEdge, selection: GridSelection): [number, number] {
  for (const { start, end } of selectionRanges(selection)) {
    const whole =
      edge.axis === 'columns'
        ? start.row === 0 &&
          end.row === MAX_ROWS - 1 &&
          edge.index >= start.col &&
          edge.index <= end.col
        : start.col === 0 &&
          end.col === MAX_COLS - 1 &&
          edge.index >= start.row &&
          edge.index <= end.row;
    if (whole) return edge.axis === 'columns' ? [start.col, end.col] : [start.row, end.row];
  }
  return [edge.index, edge.index];
}

/** The painted state that a resize reads: the stage, its box, the view, and the metrics. */
export interface PaintedGrid {
  readonly stage: HTMLElement;
  readonly box: StageBox;
  readonly map: ViewState;
  readonly metrics: GridMetrics;
  /** Whether a column or row is hidden; absent means none is. */
  readonly hidden?: HiddenLine;
}

/**
 * The painted window of the stage, from the last layout and the view of the shown sheet, with
 * the hidden lines of its part, or null before a paint.
 */
export function paintedGrid(
  els: { readonly stage: HTMLElement } | null,
  box: StageBox | null,
  layout: {
    readonly geometry: SheetGeometry;
    readonly frozen: FrozenPanes;
    readonly metrics: GridMetrics;
  } | null,
  view: { readonly scrollX: number; readonly scrollY: number } | undefined,
  part: XmlPart | undefined
): PaintedGrid | null {
  if (!els || !box || !layout || !view) return null;
  const w = (box.width - box.headWidth) / box.scale;
  const h = (box.height - box.headHeight) / box.scale;
  const map = viewStateOf(layout, view, w, h);
  return { stage: els.stage, box, map, metrics: layout.metrics, hidden: hiddenLines(part) };
}

/**
 * Whether a heading edge of `axis` can start a resize in an editor state: not in view mode, and
 * not on a protected sheet that locks that size (`allowFormatColumns`, `allowFormatRows`).
 */
export function resizable(
  state: Pick<EditorSnapshot, 'mode' | 'protection'>,
  axis: ResizeAxis
): boolean {
  const p = state.protection;
  if (state.mode === 'view') return false;
  return !p || (axis === 'columns' ? p.allowFormatColumns : p.allowFormatRows);
}

/** What a resize drag reads and calls. */
export interface ResizeDragHost {
  painted(): PaintedGrid | null;
  /** A key of the shown sheet and the store revision: a drag ends when it changes. */
  state(): string;
  selection(): GridSelection | null;
  /** Commit the size of a band, as `runResize`. */
  resize(axis: ResizeAxis, first: number, last: number, size: number): void;
  /**
   * Whether a heading edge of `axis` can start a resize: false in view mode, and false on a
   * protected sheet that does not allow that size. Then no edge is hit, so no drag starts.
   */
  resizable(axis: ResizeAxis): boolean;
}

interface Drag {
  readonly pointerId: number;
  readonly edge: ResizeEdge;
  readonly stage: HTMLElement;
  readonly guide: HTMLElement;
  /** The box of the stage at the press: the box until the next paint. */
  readonly box: StageBox;
  /** The client coordinate of the press, on the axis of the drag. */
  readonly from: number;
  /** Window pixels for one stage pixel on the axis of the drag, inside a host that scales. */
  readonly scale: number;
  /** The size of the column or row at the press, and its far edge in stage pixels. */
  readonly startPx: number;
  edgePx: number;
  readonly metrics: GridMetrics;
  /** The `ResizeDragHost.state` at the press. */
  readonly state: string;
  /** The new size in CSS pixels. */
  sizePx: number;
  readonly offKey: () => void;
}

/**
 * The largest size in CSS pixels of the layout grid that the store accepts: the widest column
 * (`MAX_COLUMN_WIDTH` characters) or the tallest row (`MAX_ROW_HEIGHT` points).
 */
export function maxResizePixels(axis: ResizeAxis, metrics: GridMetrics): number {
  if (axis === 'columns') return columnPixels(MAX_COLUMN_WIDTH, metrics.maxDigitWidth);
  return Math.floor(((MAX_ROW_HEIGHT * metrics.pixelsPerInch) / 72) * metrics.rowScale + 1e-9);
}

/**
 * The far edge of a column or row in stage pixels, in the painted window, or undefined when
 * the line is not in the window.
 */
function edgePixels(edge: ResizeEdge, painted: PaintedGrid): number | undefined {
  const { map, box } = painted;
  const columns = edge.axis === 'columns';
  const span = axisSpans(
    columns ? map.geometry.columns : map.geometry.rows,
    columns ? map.frozen.cols : map.frozen.rows,
    columns ? map.scrollX : map.scrollY,
    columns ? map.width : map.height,
    edge.index,
    edge.index
  ).at(-1);
  return span ? (columns ? box.headWidth : box.headHeight) + span[1] * box.scale : undefined;
}

/** The resize drag of one grid surface. */
export class ResizeDrag {
  #drag: Drag | null = null;
  readonly #host: ResizeDragHost;

  constructor(host: ResizeDragHost) {
    this.#host = host;
  }

  /** Whether a drag is in progress. */
  get active(): boolean {
    return this.#drag !== null;
  }

  /** The heading edge under a pointer event, or null. */
  edgeAt(e: MouseEvent, painted = this.#host.painted()): ResizeEdge | null {
    if (!painted) return null;
    // Window pixels to the stage's layout pixels, inside a host that scales the editor.
    const r = painted.stage.getBoundingClientRect();
    const scale = windowScale(painted.stage);
    const { map, box, hidden } = painted;
    const x = (e.clientX - r.left) / scale.x;
    const edge = resizeEdgeAt(map, box, x, (e.clientY - r.top) / scale.y, hidden);
    // In view mode, and on a protected sheet that locks the size, a heading edge shows no
    // resize cursor and starts no drag.
    return edge && this.#host.resizable(edge.axis) ? edge : null;
  }

  /** Show the resize cursor over a heading edge. */
  hover(e: PointerEvent): void {
    const stage = this.#host.painted()?.stage;
    if (!stage || this.#drag) return;
    const edge = this.edgeAt(e);
    if (edge) stage.dataset.resize = edge.axis;
    else delete stage.dataset.resize;
  }

  /**
   * Start a drag of `edge`: a guide line, pointer capture, and Escape to cancel. `painted` is the
   * state at the press, read before the press committed an open cell edit.
   */
  down(e: PointerEvent, edge: ResizeEdge, painted = this.#host.painted()): void {
    if (!painted) return;
    const { stage, box, map, metrics } = painted;
    const columns = edge.axis === 'columns';
    const axis = columns ? map.geometry.columns : map.geometry.rows;
    const startPx = axis.sizeOf(edge.index) * box.scale;
    const edgePx = edgePixels(edge, painted) ?? (columns ? box.headWidth : box.headHeight);
    const doc = stage.ownerDocument;
    const guide = doc.createElement('div');
    guide.className = 'xlsx-resize-guide';
    guide.setAttribute('aria-hidden', 'true');
    stage.append(guide);
    const onKey = (k: KeyboardEvent) => {
      if (k.key !== 'Escape') return;
      k.preventDefault();
      k.stopPropagation();
      this.cancel();
    };
    doc.addEventListener('keydown', onKey, true);
    this.#drag = {
      pointerId: e.pointerId,
      edge,
      stage,
      guide,
      box,
      from: columns ? e.clientX : e.clientY,
      scale: columns ? windowScale(stage).x : windowScale(stage).y,
      startPx,
      edgePx,
      metrics,
      state: this.#host.state(),
      sizePx: startPx,
      offKey: () => doc.removeEventListener('keydown', onKey, true),
    };
    stage.dataset.resize = edge.axis;
    this.#place(this.#drag, box);
    try {
      stage.setPointerCapture(e.pointerId);
    } catch {
      // A synthetic pointer has no capture.
    }
  }

  /**
   * Follow a change of the view during the drag: a change of the shown sheet, the store revision,
   * the zoom, or the mode cancels the drag and removes its guide, and a scroll moves the guide with
   * the edge. False when no drag goes on after it.
   */
  sync(): boolean {
    const drag = this.#drag;
    if (!drag) return false;
    if (this.#host.state() !== drag.state) {
      this.cancel();
      return false;
    }
    const painted = this.#host.painted();
    if (painted) {
      drag.edgePx = edgePixels(drag.edge, painted) ?? drag.edgePx;
      this.#place(drag, painted.box);
    }
    return true;
  }

  /** Follow the pointer: the guide line shows the new edge. */
  move(e: PointerEvent): void {
    const drag = this.#drag;
    if (!drag || drag.pointerId !== e.pointerId || !this.sync()) return;
    const box = this.#host.painted()?.box ?? drag.box;
    const at = drag.edge.axis === 'columns' ? e.clientX : e.clientY;
    // The guide line stops at the largest size that the store accepts.
    const max = maxResizePixels(drag.edge.axis, drag.metrics);
    drag.sizePx = Math.min(max, Math.max(0, drag.startPx + (at - drag.from) / drag.scale));
    this.#place(drag, box);
  }

  /**
   * Commit the new size as one transaction. A drag that ends at the size it started with
   * commits nothing, also after it moved away and back (DOCX Editor
   * `surface-table-interaction.ts:633-640` compares the end with the start).
   */
  up(e: PointerEvent): void {
    const drag = this.#drag;
    if (!drag || drag.pointerId !== e.pointerId) return;
    this.move(e);
    // A drag that `sync` cancelled writes nothing.
    if (this.#drag !== drag) return;
    this.#end();
    const selection = this.#host.selection();
    if (drag.sizePx === drag.startPx || !selection) return;
    const [first, last] = resizeBand(drag.edge, selection);
    const columns = drag.edge.axis === 'columns';
    // A drag to the ceiling writes the largest size, not the smallest one with those pixels.
    const size =
      drag.sizePx >= maxResizePixels(drag.edge.axis, drag.metrics)
        ? columns
          ? MAX_COLUMN_WIDTH
          : MAX_ROW_HEIGHT
        : columns
          ? columnWidthForPixels(drag.sizePx, drag.metrics.maxDigitWidth)
          : rowHeightForPixels(drag.sizePx, drag.metrics);
    this.#host.resize(drag.edge.axis, first, last, size);
  }

  /** Cancel the drag: nothing changes. */
  cancel(): void {
    this.#end();
  }

  #end(): void {
    const drag = this.#drag;
    if (!drag) return;
    this.#drag = null;
    drag.offKey();
    drag.guide.remove();
    delete drag.stage.dataset.resize;
    try {
      drag.stage.releasePointerCapture(drag.pointerId);
    } catch {
      // A synthetic pointer has no capture.
    }
  }

  #place(drag: Drag, box: StageBox): void {
    const at = drag.edgePx + drag.sizePx - drag.startPx;
    const s = drag.guide.style;
    if (drag.edge.axis === 'columns') {
      s.left = `${Math.round(at) - 1}px`;
      s.top = '0px';
      s.width = '2px';
      s.height = `${box.height}px`;
    } else {
      s.left = '0px';
      s.top = `${Math.round(at) - 1}px`;
      s.width = `${box.width}px`;
      s.height = '2px';
    }
  }
}
