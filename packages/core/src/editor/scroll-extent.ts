// The scroll extent of the grid surface in CSS pixels: the extent in points of `scrollExtent`
// (`sheet-facts.ts`) at the scale of the stage box. A sheet whose extent passes `MAX_EXTENT_PX`
// scrolls in steps (`factor`), so the container stays inside the size that a browser can
// scroll. The workbook editor sets the sizes on its extent element.

import type { GridPoint } from '../contracts/index.ts';
import type { SheetLayout } from '../layout/index.ts';
import type { StageBox } from './grid-surface.ts';
import { scrollExtent } from './sheet-facts.ts';

/** The largest scroll extent of the surface, in CSS pixels; a larger sheet scrolls in steps. */
export const MAX_EXTENT_PX = 15_000_000;

/** The scroll extent: CSS pixels of the extent element, and the scroll step of each axis. */
export interface SurfaceExtent {
  readonly width: number;
  readonly height: number;
  readonly factor: { readonly x: number; readonly y: number };
}

/**
 * The extent for a layout. `facts` holds the last row and column with content (0-based, -1 for
 * none), and `focus` the moving corner of the selection.
 *
 * A container scrolls `extent - client` CSS pixels, and the view scrolls `total - window`. At
 * step 1 the two are equal. With a larger step, only the part past the window is in steps, so
 * the largest offset of the container maps to the end of the sheet: the extent is the client
 * size plus `(total - window) / step`. An extent of `head + total / step` would stop the
 * container about one window short of the last row, so a jump to the last row of the sheet
 * would be cut back to a scroll that does not show it.
 */
export function surfaceExtent(
  layout: SheetLayout,
  facts: { readonly lastRow: number; readonly lastCol: number },
  focus: GridPoint,
  box: StageBox
): SurfaceExtent {
  const size = scrollExtent(layout, facts, focus);
  const totalW = size.width * box.scale;
  const totalH = size.height * box.scale;
  const factor = {
    x: Math.max(1, totalW / MAX_EXTENT_PX),
    y: Math.max(1, totalH / MAX_EXTENT_PX),
  };
  const axis = (total: number, head: number, client: number, step: number) =>
    step === 1 ? head + total : client + Math.max(0, total - (client - head)) / step;
  return {
    width: axis(totalW, box.headWidth, box.width, factor.x),
    height: axis(totalH, box.headHeight, box.height, factor.y),
    factor,
  };
}

/**
 * The extent that reaches both the focus and `cell`, for a scroll that reveals `cell`. A scroll
 * past the old extent is cut to its end, and the container's scroll event then gives the cut
 * offset back to the view, so a jump to a far column, such as from the name box, would stop
 * short of it.
 */
export function extentReaching(
  layout: SheetLayout,
  facts: { readonly lastRow: number; readonly lastCol: number },
  focus: GridPoint,
  cell: GridPoint,
  box: StageBox
): SurfaceExtent {
  const reach = { row: Math.max(focus.row, cell.row), col: Math.max(focus.col, cell.col) };
  return surfaceExtent(layout, facts, reach, box);
}

/** What `SurfaceScroll` reads from the editor. */
export interface SurfaceScrollHost {
  /** The scroll container, or null while no surface is mounted. */
  scroller(): HTMLElement | null;
  /** The scroll of the shown sheet, in points, or undefined when no sheet shows. */
  view(): { scrollX: number; scrollY: number; readonly zoom: number } | undefined;
  /** CSS pixels for each point at a zoom. */
  scale(zoom: number): number;
  /** The user scrolled the container: the view took its scroll. */
  scrolled(): void;
}

/**
 * The scroll of the surface: the view keeps its scroll in points, and the container scrolls in
 * CSS pixels, in steps (`SurfaceExtent.factor`) on a sheet past `MAX_EXTENT_PX`.
 */
export class SurfaceScroll {
  #factor = { x: 1, y: 1 };
  #applying = false;
  readonly #host: SurfaceScrollHost;

  constructor(host: SurfaceScrollHost) {
    this.#host = host;
  }

  /** Size the extent element to `extent`, then put the container where the view says. */
  size(extent: SurfaceExtent, element: HTMLElement): void {
    this.#factor = extent.factor;
    element.style.width = `${extent.width}px`;
    element.style.height = `${extent.height}px`;
    this.restore();
  }

  /** Put the scroll container where the view's scroll says. */
  restore(): void {
    const scroller = this.#host.scroller();
    const view = this.#host.view();
    if (!scroller || !view) return;
    const scale = this.#host.scale(view.zoom);
    const left = (view.scrollX * scale) / this.#factor.x;
    const top = (view.scrollY * scale) / this.#factor.y;
    if (Math.abs(scroller.scrollLeft - left) < 1 && Math.abs(scroller.scrollTop - top) < 1) return;
    this.#applying = true;
    scroller.scrollLeft = left;
    scroller.scrollTop = top;
    this.#applying = false;
  }

  /** A scroll of the container: the view takes it, unless `restore` set it. */
  onScroll(): void {
    if (this.#applying) return;
    const scroller = this.#host.scroller();
    const view = this.#host.view();
    if (!scroller || !view) return;
    const scale = this.#host.scale(view.zoom);
    view.scrollX = (scroller.scrollLeft * this.#factor.x) / scale;
    view.scrollY = (scroller.scrollTop * this.#factor.y) / scale;
    this.#host.scrolled();
  }
}
