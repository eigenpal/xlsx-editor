// A preview of a sheet for a host element, such as the card of a workbook in a chat or a file
// list: the top-left window of the sheet, painted with the layout and paint of the grid
// (`layoutSheet`, `paintSheet`), so the preview shows what the grid shows at scroll 0, with no
// headings, no selection, and no frozen pane lines. After `renderThumbnail` of
// pptx-editor/packages/core/src/editor/editor-instance.ts at cfb7c234 (EigenPal). Differences:
// a sheet has no fixed size, so the host gives the size of the window and the zoom; and the call
// returns whether it painted.
//
// Layout and paint cost depend on the cells of the window, never on the sheet size (AGENTS.md,
// "Layout"). The values are the values that the file stores and the edits wrote, as in the
// grid: the preview evaluates no formula.

import { EN_US, type NumberLocale } from '../cascade/index.ts';
import { layoutSheet, zoomProfile, type TextMeasurer } from '../layout/index.ts';
import { paintSheet } from '../output/index.ts';
import { hostWindowOf } from './host-realm.ts';
import type { Opened } from './opened-workbook.ts';

/**
 * The options of {@link XlsxEditorInstance.renderPreview}.
 *
 * @public
 */
export interface SheetPreviewOptions {
  /** The position of the sheet in `snapshot().sheets`. The shown sheet when absent. */
  readonly sheet?: number;
  /** The width of the preview in CSS pixels, from 1 to 2048. The element's width when absent. */
  readonly width?: number;
  /** The height of the preview in CSS pixels, from 1 to 2048. The element's height when absent. */
  readonly height?: number;
  /** The zoom, from 0.1 to 4, where 1 is 100%. The default is 0.6. */
  readonly zoom?: number;
}

/** The default zoom of a preview: a card shows about the first eight columns. */
export const PREVIEW_ZOOM = 0.6;
/** The largest side of a preview, in CSS pixels. */
const MAX_SIDE = 2048;

const clamp = (value: number, low: number, high: number): number =>
  Math.min(high, Math.max(low, value));

/** A finite number in `[low, high]`, else `fallback`. */
const within = (value: number | undefined, low: number, high: number, fallback: number) =>
  value !== undefined && Number.isFinite(value) ? clamp(value, low, high) : fallback;

/** What the preview reads of the editor. @internal */
export interface PreviewSource {
  /** The open workbook, or null while none is ready. */
  readonly opened: Opened | null;
  readonly measurer: TextMeasurer | null;
  readonly locale: NumberLocale | undefined;
  /** The position of the shown sheet. */
  readonly active: number;
}

/**
 * Paint the preview of a sheet into `element`, in place of its children: one `canvas`, hidden
 * from assistive technology. Returns false, and changes nothing, when no workbook is ready, when
 * the sheet is not a worksheet, or when the element has no size and the options give none.
 * @internal
 */
export function renderSheetPreview(
  element: HTMLElement,
  source: PreviewSource,
  options: SheetPreviewOptions = {}
): boolean {
  const { opened, measurer } = source;
  if (!opened || !measurer) return false;
  const at = options.sheet ?? source.active;
  const sheet = Number.isInteger(at) ? opened.sheets[at] : undefined;
  const entry = Number.isInteger(at) ? opened.index.sheets[at] : undefined;
  if (!sheet?.worksheet || !entry) return false;
  const width = Math.round(within(options.width ?? element.clientWidth, 0, MAX_SIDE, 0));
  const height = Math.round(within(options.height ?? element.clientHeight, 0, MAX_SIDE, 0));
  if (width < 1 || height < 1) return false;
  const zoom = within(options.zoom, 0.1, 4, PREVIEW_ZOOM);
  const profile = zoomProfile(zoom);
  const scale = profile.pixelsPerInch / 72;
  const viewport = { width: width / scale, height: height / scale, scrollX: 0, scrollY: 0 };
  const layout = layoutSheet(opened.pkg, entry.name, viewport, {
    measurer,
    profile,
    locale: source.locale ?? EN_US,
  });
  if (!layout) return false;
  const own = element.childElementCount === 1 ? element.firstElementChild : null;
  const canvas =
    own?.tagName === 'CANVAS'
      ? (own as HTMLCanvasElement)
      : element.ownerDocument.createElement('canvas');
  canvas.setAttribute('aria-hidden', 'true');
  paintSheet(canvas, layout, {
    scale,
    width: viewport.width,
    height: viewport.height,
    gridlines: layout.showGridLines,
    frozenLines: false,
    pictures: opened.pictures,
    devicePixelRatio: hostWindowOf(element).devicePixelRatio || 1,
  });
  if (canvas !== own) element.replaceChildren(canvas);
  return true;
}

/**
 * The previews that an editor painted. A font face or a picture that finishes loading, and a
 * change of the fonts or the number locale, change the paint with no edit: the editor then calls
 * `repaint` (with the grid's paint), which paints each preview again with the editor's current
 * state. A preview that left the document is dropped at the next repaint. With no grid, the editor
 * registers the faces in the document of the preview, so the preview has the file's fonts.
 * @internal
 */
export class LivePreviews {
  readonly #items = new Map<HTMLElement, SheetPreviewOptions | undefined>();
  readonly #source: () => PreviewSource;

  constructor(source: () => PreviewSource) {
    this.#source = source;
  }

  render(element: HTMLElement, options: SheetPreviewOptions | undefined): boolean {
    const painted = renderSheetPreview(element, this.#source(), options);
    if (painted) this.#items.set(element, options);
    else this.#items.delete(element);
    return painted;
  }

  repaint(): void {
    for (const [element, options] of [...this.#items])
      if (element.isConnected) this.render(element, options);
      else this.#items.delete(element);
  }

  clear(): void {
    this.#items.clear();
  }
}
