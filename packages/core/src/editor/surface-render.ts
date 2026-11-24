// The grid surface of `xlsx-editor.ts`, DOM side. One paint runs in two steps: the layout of
// the visible window, with a row heading wide enough for its last row, and the paint of that
// window into the stage (the canvas, its accessible rows, the headings, and the selection). The
// editor keeps the view state; these functions read it and change only the scroll of a pending
// top-left cell. `bindSurface` listens to the input events of a mounted surface.

import type { GridPoint, GridRange, GridSelection } from '../contracts/index.ts';
import type { XlsxPackage } from '../store/index.ts';
import type { NumberLocale } from '../cascade/index.ts';
import {
  type GridProfile,
  layoutSheet,
  type SheetLayout,
  type TextMeasurer,
} from '../layout/index.ts';
import {
  accessibleGrid,
  paintSheet,
  type AccessibleGridLabels,
  type PictureStore,
} from '../output/index.ts';
import { hostTimers, hostWindowOf } from './host-realm.ts';
import { TaskMark } from './task-mark.ts';
import type { SheetOpening } from './sheet-opening.ts';
import {
  drawHeadings,
  drawSelection,
  rowHeadWidth,
  showExitCue,
  listener,
  type StageBox,
  type SurfaceElements,
} from './grid-surface.ts';
import type { CellEditMode } from './cell-editor.ts';
import { clipboardHandlers, type ClipboardEventHost } from './clipboard-events.ts';
import type { GridPointer } from './grid-pointer.ts';
import type { LiveAnnouncer } from './live-announcer.ts';
import { handleGridKey, type GridKeyHost } from './grid-keys.ts';
import { GridExit } from './grid-exit.ts';
import { bindGridContextMenu, type GridContextMenuInput } from './grid-context-menu.ts';
import { scrollToTopLeft, type ViewState } from './view-map.ts';

/** The view state of one sheet: zoom, scroll in points, and selection. */
export interface SheetView {
  zoom: number;
  scrollX: number;
  scrollY: number;
  selection: GridSelection;
  /** The moving corner of an extended selection. */
  focus: GridPoint;
  headings: boolean;
  /** The opening scroll is applied at the first paint. */
  pendingTopLeft: GridPoint | null;
}

/** The view of a sheet when it first shows: what the file states, and its first selection. */
export function openingView(opening: SheetOpening, selection: GridSelection): SheetView {
  const { topLeft } = opening;
  return {
    zoom: opening.zoom,
    scrollX: 0,
    scrollY: 0,
    selection,
    focus: selection.active,
    headings: opening.headings,
    pendingTopLeft: topLeft.row > 0 || topLeft.col > 0 ? topLeft : null,
  };
}

/** What `layoutWindow` reads besides the view. */
export interface WindowInput {
  readonly pkg: XlsxPackage;
  readonly sheetName: string;
  /** The layout of the previous paint, for a pending top-left cell. */
  readonly previous: SheetLayout | null;
  /** The row heading width of the previous paint. */
  readonly headWidth: number | undefined;
  /** The size of the scroll container, in CSS pixels. */
  readonly width: number;
  readonly height: number;
  readonly measurer: TextMeasurer;
  readonly profile: GridProfile;
  readonly locale: NumberLocale;
}

/**
 * Lay out the window of `view`. A second pass runs when the last shown row needs a wider row
 * heading. A pending top-left cell sets the scroll of `view` from the previous layout.
 */
export function layoutWindow(
  view: SheetView,
  input: WindowInput
): { readonly layout: SheetLayout; readonly box: StageBox } | null {
  const { zoom } = view;
  const scale = input.profile.pixelsPerInch / 72;
  const { width, height } = input;
  const headHeight = view.headings ? Math.round(20 * zoom) : 0;
  let headWidth = view.headings ? (input.headWidth ?? rowHeadWidth(0, zoom)) : 0;
  let previous = input.previous;
  let layout: SheetLayout | undefined;
  for (let pass = 0; pass < 2; pass += 1) {
    const gridW = Math.max(0, width - headWidth) / scale;
    const gridH = Math.max(0, height - headHeight) / scale;
    if (view.pendingTopLeft && previous) {
      const at = scrollToTopLeft(previous.geometry, previous.frozen, view.pendingTopLeft);
      [view.scrollX, view.scrollY, view.pendingTopLeft] = [at.x, at.y, null];
    }
    layout = layoutSheet(
      input.pkg,
      input.sheetName,
      { width: gridW, height: gridH, scrollX: view.scrollX, scrollY: view.scrollY },
      { measurer: input.measurer, profile: input.profile, locale: input.locale }
    );
    if (!layout) return null;
    previous = layout;
    if (!view.headings) break;
    let lastRow = 0;
    for (const p of layout.panes) for (const r of p.rows) lastRow = Math.max(lastRow, r.index);
    const need = rowHeadWidth(lastRow, zoom);
    if (need === headWidth) break;
    headWidth = need;
  }
  return layout ? { layout, box: { scale, headWidth, headHeight, width, height } } : null;
}

/** What `paintWindow` draws besides the layout. */
export interface WindowPaint {
  readonly box: StageBox;
  /** The view of the grid area, without the headings. */
  readonly map: ViewState;
  readonly pictures: PictureStore;
  readonly labels: AccessibleGridLabels;
  /** Every range of the selection, in the order of selection (`selectionRanges`). */
  readonly selection: readonly GridRange[];
  /** The range of the active cell: its merged range, or the cell. */
  readonly active: GridRange;
  /** The source range of a pending copy or cut on this sheet, or null. */
  readonly pending: GridRange | null;
  readonly headings: boolean;
}

/** Paint a laid-out window into the stage: the canvas, its accessible rows, and the overlays. */
export function paintWindow(els: SurfaceElements, layout: SheetLayout, p: WindowPaint): void {
  const { box } = p;
  els.stage.style.width = `${box.width}px`;
  els.stage.style.height = `${box.height}px`;
  const doc = els.grid.ownerDocument;
  let canvas = els.grid.querySelector('canvas');
  if (!canvas) {
    canvas = doc.createElement('canvas');
    canvas.setAttribute('aria-hidden', 'true');
  }
  paintSheet(canvas, layout, {
    scale: box.scale,
    width: p.map.width,
    height: p.map.height,
    gridlines: layout.showGridLines,
    frozenLines: true,
    pictures: p.pictures,
    devicePixelRatio: hostWindowOf(els.grid).devicePixelRatio || 1,
  });
  els.grid.style.left = `${box.headWidth}px`;
  els.grid.style.top = `${box.headHeight}px`;
  els.grid.replaceChildren(canvas, accessibleGrid(doc, layout, p.labels));
  drawHeadings(els, layout, box, p.selection, p.headings);
  drawSelection(els, box, p.map, p.selection, p.active, p.pending);
}

/** What the input events of the surface call (`bindSurface`). */
export interface SurfaceInput {
  /** The editor, as the key table and the clipboard events read it. */
  readonly editor: GridKeyHost & ClipboardEventHost;
  readonly pointer: GridPointer;
  /** The cell editor's input, when one is open. */
  input(): HTMLElement | null;
  beginEdit(text: string | undefined, mode: CellEditMode): void;
  onScroll(): void;
  onWheel(e: WheelEvent): void;
  /** The words that say how to leave the grid, when Escape arms its exit (`grid-exit.ts`). */
  exitHint(): string;
  /** The live region of the grid. */
  readonly live: Pick<LiveAnnouncer, 'say' | 'entered'>;
  /** The handlers of the grid context menu (`grid-context-menu.ts`). */
  readonly menus: GridContextMenuInput['menus'];
}

/**
 * Listen to the input events of a mounted surface: scroll, keys, the wheel, the clipboard
 * events, the exit from the grid, the focus that enters the grid, the pointer, a double-click
 * that edits a cell in place or fits a heading edge, and an input method that starts on the
 * grid. Each listener adds its cleanup to `cleanups`.
 */
export function bindSurface(cleanups: (() => void)[], els: SurfaceElements, s: SurfaceInput): void {
  const on = listener(cleanups);
  on(els.scroller, 'scroll', () => s.onScroll());
  // Escape, then Tab or Shift+Tab, leaves the grid from any cell (`grid-exit.ts`).
  const exit = new GridExit((armed) => {
    const hint = armed ? s.exitHint() : null;
    showExitCue(els, hint);
    if (hint !== null) s.live.say(hint);
  });
  on(els.scroller, 'keydown', (e) => {
    if (exit.beforeKeyTable(e, els.scroller)) return;
    handleGridKey(s.editor, e, els.scroller);
    exit.afterKeyTable(e, els.scroller);
  });
  on(els.scroller, 'focusout', () => exit.reset());
  // A press anywhere disarms the exit, also on the chrome, which keeps the focus on the grid.
  on(els.scroller.ownerDocument.documentElement, 'pointerdown', () => exit.reset(), {
    capture: true,
  });
  on(els.scroller, 'wheel', (e) => s.onWheel(e), { passive: false });
  const clip = clipboardHandlers(s.editor, els.scroller);
  on(els.scroller, 'copy', clip.onCopy);
  on(els.scroller, 'cut', clip.onCut);
  on(els.scroller, 'paste', clip.onPaste);
  on(els.scroller, 'focus', clip.onFocus);
  bindEntry(on, els, s);
  on(els.scroller, 'pointerdown', clip.onPress);
  on(els.stage, 'pointermove', (e) => s.pointer.move(e));
  on(els.stage, 'pointercancel', (e) => s.pointer.cancel(e));
  // A double-click on a heading edge fits its columns or rows to their content, and a
  // double-click on a cell edits it in place, as F2 does.
  on(els.stage, 'dblclick', (e) => {
    if (e.target === s.input()) return;
    if (s.pointer.fit(e)) {
      e.preventDefault();
      return;
    }
    if (s.pointer.hit(e)?.kind !== 'cells') return;
    e.preventDefault();
    s.beginEdit(undefined, 'edit');
  });
  // A secondary press, Shift+F10, or the Menu key asks for the grid context menu.
  bindGridContextMenu(on, els, {
    pointer: s.pointer,
    input: s.input,
    selection: () => s.editor.selection,
    menus: s.menus,
  });
  // An input method that starts on the grid composes in the cell editor (`grid-keys.ts`).
  on(els.scroller, 'compositionstart', (e) => {
    if (e.target === els.scroller) s.beginEdit('', 'enter');
  });
}

/**
 * The focus that enters the grid from outside it says the selection (`LiveAnnouncer.entered`),
 * unless the step that moved the focus said its own words. A focus from inside the grid, such
 * as from the cell editor, is no entry. A browser that blurs the hidden cell editor before the
 * grid takes the focus gives no `relatedTarget`, so a blur in the grid to no element marks the
 * task as one that left from inside (`TaskMark`). A press on the stage, or the release of a
 * tap, focuses the grid before it selects (`GridPointer`): the entry waits for the end of the
 * press, and is said only when the press said nothing, such as a press on the active cell.
 */
function bindEntry(on: ReturnType<typeof listener>, els: SurfaceElements, s: SurfaceInput): void {
  const leftInside = new TaskMark(hostTimers(() => els.scroller));
  let pressing = false;
  let entry = false;
  const press = (run: () => void) => {
    pressing = true;
    let entered = false;
    try {
      run();
    } finally {
      [pressing, entered, entry] = [false, entry, false];
    }
    if (entered) s.live.entered();
  };
  on(els.stage, 'pointerdown', (e) => press(() => s.pointer.down(e)));
  on(els.stage, 'pointerup', (e) => press(() => s.pointer.up(e)));
  on(els.scroller, 'focusout', (e) => {
    if (!e.relatedTarget) leftInside.note();
  });
  on(els.scroller, 'focus', (e) => {
    const from = e.relatedTarget as Node | null;
    if (from ? els.scroller.contains(from) : leftInside.held) return;
    if (pressing) entry = true;
    else s.live.entered();
  });
}
