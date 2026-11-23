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
  host.append(root);
  return { scroller, extent, stage, corner, columnHeads, rowHeads, grid, overlay, live, exitCue };
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

/** Unique axis entries of the panes, by index, in viewport order. */
function entries(panes: readonly PaneLayout[], pick: (p: PaneLayout) => readonly AxisEntry[]) {
  const seen = new Map<number, AxisEntry>();
  for (const p of panes) for (const e of pick(p)) if (!seen.has(e.index)) seen.set(e.index, e);
  return [...seen.values()].sort((a, b) => a.start - b.start);
}

/** Draw the row and column headings of a layout, marking the selected ones. */
export function drawHeadings(
  els: SurfaceElements,
  layout: SheetLayout,
  box: StageBox,
  selected: GridRange | null,
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
  const cols = entries(layout.panes, (p) => p.columns).map((e) => {
    const el = div(doc, 'xlsx-head');
    el.textContent = columnLetters(e.index + 1);
    el.dataset.col = String(e.index);
    if (selected && e.index >= selected.start.col && e.index <= selected.end.col)
      el.classList.add('xlsx-head-selected');
    place(el, e.start * box.scale, 0, e.size * box.scale, box.headHeight);
    return el;
  });
  const rows = entries(layout.panes, (p) => p.rows).map((e) => {
    const el = div(doc, 'xlsx-head');
    el.textContent = String(e.index + 1);
    el.dataset.row = String(e.index);
    if (selected && e.index >= selected.start.row && e.index <= selected.end.row)
      el.classList.add('xlsx-head-selected');
    place(el, 0, e.start * box.scale, box.headWidth, e.size * box.scale);
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
 * Draw the selection: a tint over `range`, and an outline around `activeRange`, the active cell
 * or the merged range that holds it. A range of one cell, or one that is the active range, gets
 * the outline only. A dashed outline, above them, marks the source range of a pending copy or
 * cut. It has a strip for each edge that shows, inside the range, so a pane seam or the edge of
 * the window adds no line. It lives in this overlay only, never in the canvas of the workbook.
 */
export function drawSelection(
  els: SurfaceElements,
  box: StageBox,
  view: ViewState,
  range: GridRange,
  activeRange: GridRange,
  pending: GridRange | null = null
): void {
  const single =
    (range.start.row === range.end.row && range.start.col === range.end.col) ||
    sameRange(activeRange, range);
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
    for (const r of rangeRects(view, range)) {
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
