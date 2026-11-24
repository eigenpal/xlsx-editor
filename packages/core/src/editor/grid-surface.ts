// The DOM of the grid surface: the scroll container, the sticky stage with the row and column
// headings, the painted window, and the selection overlay. The workbook editor owns the state;
// this module builds and updates elements only. Every element comes from the host document of
// the surface (`host-realm.ts`), and every text goes through `textContent`.

import type { GridRange } from '../contracts/index.ts';
import { MAX_COLS, MAX_ROWS, columnLetters } from '../store/index.ts';
import type { AxisEntry, PaneLayout, SheetLayout } from '../layout/index.ts';
import { rangeOutline, rangeRects, type ViewState } from './view-map.ts';

/** The elements of a mounted surface. */
export interface SurfaceElements {
  readonly scroller: HTMLDivElement;
  readonly extent: HTMLDivElement;
  readonly stage: HTMLDivElement;
  readonly corner: HTMLDivElement;
  readonly columnHeads: HTMLDivElement;
  readonly rowHeads: HTMLDivElement;
  readonly grid: HTMLDivElement;
  readonly overlay: HTMLDivElement;
  readonly live: HTMLDivElement;
  /** The cue of the armed exit from the grid (`grid-exit.ts`). */
  readonly exitCue: HTMLDivElement;
}

function div(doc: Document, className: string): HTMLDivElement {
  const el = doc.createElement('div');
  el.className = className;
  return el;
}

/** Build the surface in `host`. The scroll container takes the keyboard focus. */
export function buildSurface(host: HTMLElement): SurfaceElements {
  const doc = host.ownerDocument;
  const root = div(doc, 'xlsx-surface');
  // The page selection moves into the grid (`clipboard-events.ts`). Inside editable content of
  // the host page, a caret there would let editing commands change the surface. The cell editor
  // is a `textarea`, so it still takes text.
  root.setAttribute('contenteditable', 'false');
  const scroller = div(doc, 'xlsx-scroller');
  scroller.tabIndex = 0;
  scroller.setAttribute('role', 'grid');
  scroller.setAttribute('aria-readonly', 'false');
  scroller.setAttribute('aria-multiselectable', 'true');
  // The grid paints only the visible cells, so assistive technology reads the sheet size here.
  scroller.setAttribute('aria-rowcount', String(MAX_ROWS));
  scroller.setAttribute('aria-colcount', String(MAX_COLS));
  const extent = div(doc, 'xlsx-extent');
  const stage = div(doc, 'xlsx-stage');
  const corner = div(doc, 'xlsx-corner');
  const columnHeads = div(doc, 'xlsx-heads xlsx-column-heads');
  const rowHeads = div(doc, 'xlsx-heads xlsx-row-heads');
  const grid = div(doc, 'xlsx-grid');
  const overlay = div(doc, 'xlsx-overlay');
  overlay.setAttribute('aria-hidden', 'true');
  columnHeads.setAttribute('aria-hidden', 'true');
  rowHeads.setAttribute('aria-hidden', 'true');
  stage.append(grid, overlay, columnHeads, rowHeads, corner);
  extent.append(stage);
  scroller.append(extent);
  const live = div(doc, 'xlsx-live');
  live.setAttribute('aria-live', 'polite');
  live.setAttribute('aria-atomic', 'true');
  // Outside the stage, so it never paints into the workbook window. The live region says the
  // same words, so assistive technology skips the cue.
  const exitCue = div(doc, 'xlsx-exit-cue');
  exitCue.setAttribute('aria-hidden', 'true');
  root.append(scroller, live, exitCue);
  trackPointerFocus(root);
  host.append(root);
  return { scroller, extent, stage, corner, columnHeads, rowHeads, grid, overlay, live, exitCue };
}

/**
 * Mark the surface with `data-no-focus-ring` from a press on it until the focus leaves it. The
 * grid focus ring shows only without this mark (`editor.css`). A press focuses the grid from
 * script after `preventDefault()` (`grid-pointer.ts`), so `:focus-visible` alone takes that focus
 * for a keyboard focus. The listeners are on the surface itself, so they go with it.
 */
function trackPointerFocus(root: HTMLElement): void {
  root.addEventListener('pointerdown', () => root.setAttribute('data-no-focus-ring', ''), true);
  const leave = () => {
    if (!root.contains(root.ownerDocument.activeElement))
      root.removeAttribute('data-no-focus-ring');
  };
  root.addEventListener('focusout', (e) => {
    // A focus event's related target is an element or null. No `instanceof`: the surface can
    // live in another window (`host-realm.ts`).
    const next = e.relatedTarget as Node | null;
    if (next) {
      if (!root.contains(next)) root.removeAttribute('data-no-focus-ring');
      return;
    }
    // No related target: the focused cell editor closed, and the grid takes the focus back from
    // script after this event. Look again when the event is done.
    queueMicrotask(leave);
  });
}

/**
 * Focus `target`, an element of a surface, from script. The focus takes the ring of the element
 * that had the focus before, as a browser decides `:focus-visible` for a focus that script
 * moves: after a key on a toolbar control, the grid shows its ring; after the open of a workbook
 * or a press on the chrome, it shows none. A browser shows the ring for a script focus before
 * any user input, so the mark (`data-no-focus-ring`) decides here. A focus inside the surface
 * keeps the mark as it is.
 */
export function focusSurface(target: HTMLElement | null | undefined): void {
  if (!target) return;
  const root = target.closest('.xlsx-surface');
  const before = target.ownerDocument.activeElement;
  if (root && !(before && root.contains(before)) && !hasFocusRing(before))
    root.setAttribute('data-no-focus-ring', '');
  target.focus({ preventScroll: true });
}

/** Whether `el` shows a keyboard focus ring. A DOM without `:focus-visible` gives false. */
function hasFocusRing(el: Element | null): boolean {
  if (!el || el === el.ownerDocument.body) return false;
  try {
    return el.matches(':focus-visible');
  } catch {
    return false;
  }
}

/** Show the cue of the armed exit from the grid with `text`, or hide it with null. */
export function showExitCue(els: SurfaceElements, text: string | null): void {
  if (text === null) return els.exitCue.removeAttribute('data-shown');
  els.exitCue.textContent = text;
  els.exitCue.setAttribute('data-shown', '');
}

const px = (v: number): string => `${Math.round(v * 100) / 100}px`;

function place(el: HTMLElement, x: number, y: number, w: number, h: number): void {
  el.style.left = px(x);
  el.style.top = px(y);
  el.style.width = px(Math.max(0, w));
  el.style.height = px(Math.max(0, h));
}

/** The geometry of the stage, in CSS pixels. */
export interface StageBox {
  /** CSS pixels for each point. */
  readonly scale: number;
  readonly headWidth: number;
  readonly headHeight: number;
  readonly width: number;
  readonly height: number;
}

/** An axis entry with the band of its pane on that axis, in viewport points. */
interface HeadEntry extends AxisEntry {
  readonly band: readonly [number, number];
}

/**
 * Unique axis entries of the panes, by index, in viewport order, each with the band of its pane
 * on that axis (`PaneLayout.clip`). A row that the scrolled pane shows in part starts above the
 * pane, under the frozen rows; its heading is clipped to the band, as the pane clips its cells.
 */
function entries(panes: readonly PaneLayout[], axis: 'rows' | 'columns'): readonly HeadEntry[] {
  const seen = new Map<number, HeadEntry>();
  for (const p of panes) {
    const band: [number, number] =
      axis === 'rows' ? [p.clip.y, p.clip.y + p.clip.height] : [p.clip.x, p.clip.x + p.clip.width];
    for (const e of p[axis]) if (!seen.has(e.index)) seen.set(e.index, { ...e, band });
  }
  return [...seen.values()].sort((a, b) => a.start - b.start);
}

/**
 * Clip a heading to the band of its pane: the part outside the band does not show, so a heading
 * of the scrolled pane never paints over a heading of the frozen pane. `vertical` is true for a
 * row heading. Sizes are in CSS pixels.
 */
function clipHead(el: HTMLElement, e: HeadEntry, scale: number, vertical: boolean): void {
  const before = Math.max(0, e.band[0] - e.start) * scale;
  const after = Math.max(0, e.start + e.size - e.band[1]) * scale;
  if (before <= 0 && after <= 0) return;
  el.style.clipPath = vertical
    ? `inset(${px(before)} 0 ${px(after)} 0)`
    : `inset(0 ${px(after)} 0 ${px(before)})`;
}

/**
 * Draw the row and column headings of a layout, marking the ones that a range of `selected`
 * covers. It reads the shown headings only, once for each range.
 */
export function drawHeadings(
  els: SurfaceElements,
  layout: SheetLayout,
  box: StageBox,
  selected: readonly GridRange[],
  show: boolean
): void {
  const doc = els.stage.ownerDocument;
  for (const el of [els.corner, els.columnHeads, els.rowHeads]) el.hidden = !show;
  if (!show) {
    els.columnHeads.replaceChildren();
    els.rowHeads.replaceChildren();
    return;
  }
  place(els.corner, 0, 0, box.headWidth, box.headHeight);
  place(els.columnHeads, box.headWidth, 0, box.width - box.headWidth, box.headHeight);
  place(els.rowHeads, 0, box.headHeight, box.headWidth, box.height - box.headHeight);
  const fontSize = px(Math.max(6, Math.min(24, box.headHeight * 0.55)));
  els.columnHeads.style.fontSize = fontSize;
  els.rowHeads.style.fontSize = fontSize;
  const cols = entries(layout.panes, 'columns').map((e) => {
    const el = div(doc, 'xlsx-head');
    el.textContent = columnLetters(e.index + 1);
    el.dataset.col = String(e.index);
    if (selected.some((r) => e.index >= r.start.col && e.index <= r.end.col))
      el.classList.add('xlsx-head-selected');
    place(el, e.start * box.scale, 0, e.size * box.scale, box.headHeight);
    clipHead(el, e, box.scale, false);
    return el;
  });
  const rows = entries(layout.panes, 'rows').map((e) => {
    const el = div(doc, 'xlsx-head');
    el.textContent = String(e.index + 1);
    el.dataset.row = String(e.index);
    if (selected.some((r) => e.index >= r.start.row && e.index <= r.end.row))
      el.classList.add('xlsx-head-selected');
    place(el, 0, e.start * box.scale, box.headWidth, e.size * box.scale);
    clipHead(el, e, box.scale, true);
    return el;
  });
  els.columnHeads.replaceChildren(...cols);
  els.rowHeads.replaceChildren(...rows);
}

/** Empty the surface: no workbook is shown. */
export function clearSurface(els: SurfaceElements): void {
  els.grid.replaceChildren();
  els.overlay.replaceChildren();
  els.columnHeads.replaceChildren();
  els.rowHeads.replaceChildren();
  els.extent.style.width = '0px';
  els.extent.style.height = '0px';
}

/** Whether two ranges have the same corners. */
const sameRange = (a: GridRange, b: GridRange) =>
  a.start.row === b.start.row &&
  a.start.col === b.start.col &&
  a.end.row === b.end.row &&
  a.end.col === b.end.col;

/** The width of the outline of a pending copy or cut, in CSS pixels. */
const PENDING_EDGE = 2;

/**
 * Draw the selection: a tint over each range of `ranges`, and an outline around `activeRange`,
 * the active cell or the merged range that holds it. A selection of one range that is one cell,
 * or that is the active range, gets the outline only. In a selection of several ranges, each
 * range gets the tint and its outline, as one cell does too, so every range shows. Each range is
 * clipped to the window (`rangeRects`), so a range outside it draws nothing. A dashed outline,
 * above them, marks the source range of a pending copy or cut. It has a strip for each edge that
 * shows, inside the range, so a pane seam or the edge of the window adds no line. It lives in this
 * overlay only, never in the canvas of the workbook.
 */
export function drawSelection(
  els: SurfaceElements,
  box: StageBox,
  view: ViewState,
  ranges: readonly GridRange[],
  activeRange: GridRange,
  pending: GridRange | null = null
): void {
  const range = ranges[0];
  const single =
    ranges.length === 1 &&
    range !== undefined &&
    ((range.start.row === range.end.row && range.start.col === range.end.col) ||
      sameRange(activeRange, range));
  const activeRects = rangeRects(view, activeRange);
  const doc = els.overlay.ownerDocument;
  place(
    els.overlay,
    box.headWidth,
    box.headHeight,
    box.width - box.headWidth,
    box.height - box.headHeight
  );
  const out: HTMLElement[] = [];
  if (!single)
    for (const each of ranges)
      for (const r of rangeRects(view, each)) {
        const el = div(doc, 'xlsx-selection-range');
        place(el, r.x * box.scale, r.y * box.scale, r.width * box.scale, r.height * box.scale);
        out.push(el);
      }
  for (const r of activeRects) {
    const el = div(doc, 'xlsx-active-cell');
    place(el, r.x * box.scale, r.y * box.scale, r.width * box.scale, r.height * box.scale);
    out.push(el);
  }
  for (const edge of pending ? rangeOutline(view, pending) : []) {
    const el = div(doc, 'xlsx-pending-edge');
    el.dataset.side = edge.side;
    const across = edge.side === 'top' || edge.side === 'bottom';
    const length = edge.length * box.scale;
    place(
      el,
      edge.x * box.scale - (edge.side === 'right' ? PENDING_EDGE : 0),
      edge.y * box.scale - (edge.side === 'bottom' ? PENDING_EDGE : 0),
      across ? length : PENDING_EDGE,
      across ? PENDING_EDGE : length
    );
    out.push(el);
  }
  els.overlay.replaceChildren(...out);
}

/** The width of the row headings for the last row number shown, in CSS pixels. */
export function rowHeadWidth(lastRow: number, zoom: number): number {
  const digits = Math.max(3, String(lastRow + 1).length);
  return Math.round((digits * 7 + 12) * zoom);
}

/** A function that adds an event listener and pushes its removal to `cleanups`. */
export function listener(cleanups: (() => void)[]) {
  return <K extends keyof HTMLElementEventMap>(
    target: HTMLElement,
    type: K,
    fn: (e: HTMLElementEventMap[K]) => void,
    opts?: AddEventListenerOptions
  ): void => {
    target.addEventListener(type, fn, opts);
    cleanups.push(() => target.removeEventListener(type, fn, opts));
  };
}
