// The pointer input of the grid surface: a press on a cell selects it, a press on a heading
// selects whole rows or columns, a press on the corner selects the sheet, Shift extends, and a
// drag extends the selection until the pointer goes up. A press with the platform modifier
// (Command on Apple platforms, Control elsewhere) adds the cell or the heading band as a new
// range of the selection (`addRange`), and its drag extends that range; Shift then extends the
// range that holds the active cell. A plain press starts a selection of one range again. A press
// commits the open cell edit
// first; a refused commit keeps the editor and the selection. A press on the edge of a heading
// starts a resize drag instead (`grid-resize.ts`), and selects nothing. A double-click on that
// edge fits the columns or rows to their content (`grid-fit.ts`).
//
// The press focuses the grid before it selects. When that focus change ends a chrome command
// that the store refuses (a blur commits a sheet rename), its refusal shows after the new
// selection, which would otherwise clear it: a press never hides the reason of its own commit.

import type { GridPoint, GridSelection } from '../contracts/index.ts';
import { windowScale } from './host-frame.ts';
import { hostWindowOf } from './host-realm.ts';
import { isApplePlatform } from './chrome-menus.ts';
import { selectionRanges } from '../automation/structure-commands.ts';
import {
  addRange,
  inRange,
  isMultiRange,
  selectBand,
  selectCell,
  withActiveRange,
  type Merges,
} from './selection.ts';
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
  /** The merged ranges of the shown sheet, which a press that adds a range covers. */
  merges(): Merges;
  /** Commit the open cell edit; false when the store refused it. */
  commitOpenEdit(): boolean;
  focus(): void;
  /** The refusal that the editor shows now, or null. */
  refusal(): EditRefusal | null;
  /** Show a refusal again, through the refusal channel. */
  showRefusal(refusal: EditRefusal): void;
  /** Whether a host listens to `contextMenu`, so a press of its gesture selects nothing. */
  hasContextMenu(): boolean;
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
  // In a selection of several ranges, the band takes the place of the range of the active cell.
  if (isMultiRange(from)) return { ...withActiveRange(from, band.range), active };
  return { ...band, active, anchor: from.anchor };
}

/**
 * Whether a press adds a range to the selection: the platform modifier is down, Command on an
 * Apple platform, else Control. On an Apple platform, Control and a press opens a context menu.
 */
function addsRange(e: PointerEvent, target: Element | undefined): boolean {
  if (e.shiftKey || e.altKey) return false;
  const apple = isApplePlatform(target ? hostWindowOf(target).navigator : undefined);
  return apple ? e.metaKey && !e.ctrlKey : e.ctrlKey && !e.metaKey;
}

/**
 * Whether a press is the context-menu gesture of an Apple platform: Control and a primary press.
 * It selects nothing; the `contextmenu` event that follows keeps a selection that holds the cell.
 */
function appleContextPress(e: PointerEvent, target: Element | undefined): boolean {
  if (!e.ctrlKey || e.metaKey || e.button !== 0 || e.pointerType === 'touch') return false;
  return isApplePlatform(target ? hostWindowOf(target).navigator : undefined);
}

/** How far a touch can move, in CSS pixels, and still be a tap. */
const TAP_SLOP = 10;
/** How far, in CSS pixels, a double-click on a heading edge may be from the press before it. */
const PRESS_SLOP = 2;

export type Hit = { kind: 'corner' } | { kind: 'columns' | 'rows' | 'cells'; cell: GridPoint };

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

  /**
   * A secondary press that opens the grid context menu (`grid-context-menu.ts`): commit the open
   * cell edit, focus the grid, and select what the press hits when the selection does not hold
   * it, so the menu acts on the cells under the pointer. A press inside the selection keeps it.
   * Returns what the press hit and the selection after it, or null when it hit nothing or the
   * store refused the commit.
   */
  contextPress(e: MouseEvent): { readonly hit: Hit; readonly selection: GridSelection } | null {
    const at = this.hit(e);
    const cur = this.#host.current();
    if (!at || !cur || !this.#host.commitOpenEdit()) return null;
    this.#host.focus();
    const target =
      at.kind === 'cells'
        ? null
        : at.kind === 'corner'
          ? selectBand(cur.sheet, 'all', 0, 0)
          : selectBand(
              cur.sheet,
              at.kind,
              at.kind === 'columns' ? at.cell.col : at.cell.row,
              at.kind === 'columns' ? at.cell.col : at.cell.row
            );
    // A press inside any range of the selection keeps every range, so a command of the menu
    // applies to all of them.
    const holds = selectionRanges(cur.selection).some((range) =>
      target
        ? inRange(range, target.range.start) && inRange(range, target.range.end)
        : inRange(range, (at as { cell: GridPoint }).cell)
    );
    if (!holds) {
      if (target) this.#host.setSelection(target, target.active, target.active);
      else this.#host.select((at as { cell: GridPoint }).cell);
    }
    return { hit: at, selection: this.#host.current()?.selection ?? cur.selection };
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
    // The context-menu gesture of an Apple platform starts no resize either.
    // With no context menu to open, the press selects, as before the gesture had a menu.
    const context =
      this.#host.hasContextMenu() && appleContextPress(e, this.#host.painted()?.stage);
    this.#edgePress = resize && !context ? { edge: resize.edge, x: e.clientX, y: e.clientY } : null;
    if (context) return;
    if (resize) {
      this.#resize.down(e, resize.edge, resize.painted);
      return;
    }
    if (at.kind === 'corner') {
      const all = selectBand(cur.sheet, 'all', 0, 0);
      this.#host.setSelection(all, all.active);
      return;
    }
    const adds = e.pointerType !== 'touch' && addsRange(e, this.#host.painted()?.stage);
    if (adds) {
      // The new range is the cell (or its merged range), or the band of the heading.
      const one =
        at.kind === 'cells'
          ? selectCell(cur.sheet, at.cell, this.#host.merges())
          : selectBand(
              cur.sheet,
              at.kind,
              at.kind === 'columns' ? at.cell.col : at.cell.row,
              at.kind === 'columns' ? at.cell.col : at.cell.row
            );
      const next = addRange(cur.selection, one.range, one.active);
      // At `MAX_SELECTION_RANGES` no range is added: the press refuses with `too-many-ranges`,
      // which the live region announces, and starts no drag.
      if (next === cur.selection) return this.#host.showRefusal({ code: 'too-many-ranges' });
      this.#host.setSelection(next, at.kind === 'cells' ? one.active : at.cell, at.cell);
    } else if (at.kind === 'columns' || at.kind === 'rows') {
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
      // `extendTo` extends the range of the active cell: the added range of a modifier drag.
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
