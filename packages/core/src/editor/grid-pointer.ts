// The pointer input of the grid surface: a press on a cell selects it, a press on a heading
// selects whole rows or columns, a press on the corner selects the sheet, Shift extends, and a
// drag extends the selection until the pointer goes up. A press commits the open cell edit
// first; a refused commit keeps the editor and the selection. A press on the edge of a heading
// starts a resize drag instead (`grid-resize.ts`), and selects nothing. A double-click on that
// edge fits the columns or rows to their content (`grid-fit.ts`).
//
// The press focuses the grid before it selects. When that focus change ends a chrome command
// that the store refuses (a blur commits a sheet rename), its refusal shows after the new
// selection, which would otherwise clear it: a press never hides the reason of its own commit.

import type { GridPoint, GridSelection } from '../contracts/index.ts';
import { windowScale } from './host-frame.ts';
import { inRange, selectBand } from './selection.ts';
import { cellAtPoint } from './view-map.ts';
import {
  ResizeDrag,
  resizable,
  resizeBand,
  type PaintedGrid,
  type ResizeAxis,
  type ResizeEdge,
} from './grid-resize.ts';
import type { EditorSnapshot } from './editor-model.ts';
import type { EditRefusal } from './cell-input.ts';

/** What the pointer input reads and calls on the editor. */
export interface GridPointerHost {
  /** The stage, its box, and the view of the painted window, or null before a paint. */
  painted(): PaintedGrid | null;
  /**
   * A key of the open, the shown sheet, the revision, the zoom, and the mode: a resize drag
   * ends when it changes.
   */
  state(): string;
  /** The shown sheet's selection and its moving corner, or null when no sheet shows. */
  current(): {
    readonly sheet: number;
    readonly selection: GridSelection;
    readonly focus: GridPoint;
  } | null;
  /** Set the selection and its moving corner, and repaint; `reveal` scrolls to that cell. */
  setSelection(selection: GridSelection, focus: GridPoint, reveal?: GridPoint): void;
  select(anchor: GridPoint, focus?: GridPoint): void;
  /** Extend the selection from its anchor to `to` (`extendTo`), and reveal `to`. */
  extend(to: GridPoint): void;
  /** Commit the open cell edit; false when the store refused it. */
  commitOpenEdit(): boolean;
  focus(): void;
  /** The refusal that the editor shows now, or null. */
  refusal(): EditRefusal | null;
  /** Show a refusal again, through the refusal channel. */
  showRefusal(refusal: EditRefusal): void;
  /**
   * Set the size of columns or rows (0-based), as the editor's `setColumnWidth`, or fit them to
   * their content (`'fit'`, `grid-fit.ts`).
   */
  resize(axis: ResizeAxis, first: number, last: number, size: number | 'fit'): void;
  /**
   * The mode and the protection of the shown sheet: a heading edge starts no resize in view
   * mode, nor on a protected sheet that does not allow that size (`resizable`).
   */
  sizeState(): Pick<EditorSnapshot, 'mode' | 'protection'>;
}

/**
 * A band extended from the selection `from`: its anchor stays, and its active cell stays while
 * the band holds it, else the anchor becomes the active cell.
 */
function keepActive(band: GridSelection, from: GridSelection): GridSelection {
  const active = inRange(band.range, from.active) ? from.active : from.anchor;
  return { ...band, active, anchor: from.anchor };
}

/** How far a touch can move, in CSS pixels, and still be a tap. */
const TAP_SLOP = 10;
/** How far, in CSS pixels, a double-click on a heading edge may be from the press before it. */
const PRESS_SLOP = 2;

type Hit = { kind: 'corner' } | { kind: 'columns' | 'rows' | 'cells'; cell: GridPoint };

export class GridPointer {
  #drag: { pointerId: number; kind: 'cells' | 'columns' | 'rows' } | null = null;
  /** A touch that has not moved: a tap selects on release; a move pans the sheet instead. */
  #tap: { pointerId: number; x: number; y: number } | null = null;
  readonly #host: GridPointerHost;
  readonly #resize: ResizeDrag;
  /** The heading edge of the last press, and where it was: a double-click fits that edge. */
  #edgePress: { readonly edge: ResizeEdge; readonly x: number; readonly y: number } | null = null;

  constructor(host: GridPointerHost) {
    this.#host = host;
    this.#resize = new ResizeDrag({
      painted: () => host.painted(),
      state: () => host.state(),
      selection: () => host.current()?.selection ?? null,
      resize: (axis, first, last, size) => host.resize(axis, first, last, size),
      resizable: (axis) => resizable(host.sizeState(), axis),
    });
  }

  /**
   * Fit the columns or rows of the heading edge under a double-click to their content: every
   * column or row of a selection of whole columns (rows) that holds the edge, or the one under
   * the edge, as a drag resizes them. False when no edge that can resize is under the press.
   */
  fit(e: MouseEvent): boolean {
    // A press that commits an open cell edit clears the paint until the next frame: the edge of
    // that press stands for the double-click at the same point.
    const press = this.#edgePress;
    const same =
      press &&
      Math.abs(press.x - e.clientX) <= PRESS_SLOP &&
      Math.abs(press.y - e.clientY) <= PRESS_SLOP;
    const painted = this.#host.painted() !== null;
    const edge = painted ? this.#resize.edgeAt(e) : same ? press.edge : null;
    this.#edgePress = null;
    const selection = this.#host.current()?.selection;
    if (!edge || !selection) return false;
    const [first, last] = resizeBand(edge, selection);
    this.#host.resize(edge.axis, first, last, 'fit');
    return true;
  }

  /** Where a pointer event is: a heading, the corner, or a cell. */
  hit(e: MouseEvent): Hit | null {
    const painted = this.#host.painted();
    if (!painted) return null;
    const { stage, box, map } = painted;
    // Window pixels to the stage's layout pixels, inside a host that scales the editor.
    const r = stage.getBoundingClientRect();
    const scale = windowScale(stage);
    const x = (e.clientX - r.left) / scale.x;
    const y = (e.clientY - r.top) / scale.y;
    const cell = cellAtPoint(
      map,
      Math.max(0, x - box.headWidth) / box.scale,
      Math.max(0, y - box.headHeight) / box.scale
    );
    if (x < box.headWidth && y < box.headHeight) return { kind: 'corner' };
    if (y < box.headHeight) return { kind: 'columns', cell };
    if (x < box.headWidth) return { kind: 'rows', cell };
    return { kind: 'cells', cell };
  }

  down(e: PointerEvent): void {
    if (e.button !== 0) return;
    // A touch pans the sheet (`touch-action: pan-x pan-y`): it selects only on a tap, at release.
    if (e.pointerType === 'touch') {
      this.#tap = { pointerId: e.pointerId, x: e.clientX, y: e.clientY };
      return;
    }
    if (this.#press(e)) e.preventDefault();
  }

  /** Select what a press hits, as a mouse press or a tap does. Returns whether it acted. */
  #press(e: PointerEvent): boolean {
    const at = this.hit(e);
    const cur = this.#host.current();
    if (!at || !cur) return false;
    if (e.pointerType !== 'touch') e.preventDefault();
    // The painted state before the commit: a commit clears the layout until the next paint. A
    // tap never starts a resize.
    const painted = this.#host.painted();
    const heading = at.kind === 'columns' || at.kind === 'rows';
    const edge = heading && e.pointerType !== 'touch' ? this.#resize.edgeAt(e, painted) : null;
    if (!this.#host.commitOpenEdit()) return true;
    const before = this.#host.refusal();
    this.#host.focus();
    const made = this.#host.refusal();
    this.#select(e, at, cur, edge && painted ? { edge, painted } : null);
    // A refusal that the focus change made stays in view after the selection.
    if (made && made !== before && this.#host.refusal() !== made) this.#host.showRefusal(made);
    return true;
  }

  #select(
    e: PointerEvent,
    at: Hit,
    cur: NonNullable<ReturnType<GridPointerHost['current']>>,
    resize: { readonly edge: ResizeEdge; readonly painted: PaintedGrid } | null
  ): void {
    this.#edgePress = resize ? { edge: resize.edge, x: e.clientX, y: e.clientY } : null;
    if (resize) {
      this.#resize.down(e, resize.edge, resize.painted);
      return;
    }
    if (at.kind === 'corner') {
      const all = selectBand(cur.sheet, 'all', 0, 0);
      this.#host.setSelection(all, all.active);
      return;
    }
    if (at.kind === 'columns' || at.kind === 'rows') {
      const i = at.kind === 'columns' ? at.cell.col : at.cell.row;
      const anchor = e.shiftKey
        ? at.kind === 'columns'
          ? cur.selection.anchor.col
          : cur.selection.anchor.row
        : i;
      let selection = selectBand(cur.sheet, at.kind, anchor, i);
      // Shift extends the band from the anchor, and the active cell stays while the band holds
      // it, as `extendTo` keeps it.
      if (e.shiftKey) selection = keepActive(selection, cur.selection);
      this.#host.setSelection(selection, at.cell, at.cell);
    } else if (e.shiftKey) this.#host.extend(at.cell);
    else this.#host.select(at.cell);
    if (e.pointerType === 'touch') return;
    this.#drag = { pointerId: e.pointerId, kind: at.kind };
    try {
      this.#host.painted()?.stage.setPointerCapture(e.pointerId);
    } catch {
      // A synthetic pointer has no capture.
    }
  }

  move(e: PointerEvent): void {
    if (this.#resize.active) {
      this.#resize.move(e);
      return;
    }
    const tap = this.#tap;
    // A touch that moves past the slop is a pan, not a tap.
    if (
      tap?.pointerId === e.pointerId &&
      Math.hypot(e.clientX - tap.x, e.clientY - tap.y) > TAP_SLOP
    )
      this.#tap = null;
    const drag = this.#drag;
    if (!drag) this.#resize.hover(e);
    const cur = this.#host.current();
    if (!drag || drag.pointerId !== e.pointerId || !cur) return;
    const at = this.hit(e);
    if (!at || at.kind === 'corner') return;
    if (drag.kind === 'cells') {
      if (at.cell.row !== cur.focus.row || at.cell.col !== cur.focus.col)
        this.#host.extend(at.cell);
      return;
    }
    const anchor = drag.kind === 'columns' ? cur.selection.anchor.col : cur.selection.anchor.row;
    const i = drag.kind === 'columns' ? at.cell.col : at.cell.row;
    const band = selectBand(cur.sheet, drag.kind, anchor, i);
    this.#host.setSelection(keepActive(band, cur.selection), at.cell, at.cell);
  }

  up(e: PointerEvent): void {
    if (this.#resize.active) this.#resize.up(e);
    if (this.#drag?.pointerId === e.pointerId) this.#drag = null;
    if (this.#tap?.pointerId !== e.pointerId) return;
    this.#tap = null;
    this.#press(e);
  }

  /**
   * Follow a scroll, a sheet change, or a store change during a resize drag: the guide moves
   * with its edge, and a new sheet or revision cancels the drag (`ResizeDrag.sync`).
   */
  syncResize(): void {
    this.#resize.sync();
  }

  /** Cancel a resize drag: the workbook closed, another opened, or the surface unmounted. */
  cancelResize(): void {
    this.#resize.cancel();
  }

  /**
   * The browser took the pointer, for example to pan: no tap and no drag. A selection drag ends
   * where it is, and a resize drag is cancelled.
   */
  cancel(e: PointerEvent): void {
    if (this.#resize.active) this.#resize.cancel();
    if (this.#drag?.pointerId === e.pointerId) this.#drag = null;
    if (this.#tap?.pointerId === e.pointerId) this.#tap = null;
  }
}
