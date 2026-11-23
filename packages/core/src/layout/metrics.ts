// Font metrics on the pixel grid, the grid profile, and the grid metrics of a workbook
// (design D6). Measured on the reference PDFs and, where marked, in the screen profile
// (reference/quick-ref/spreadsheetml.md, "Column widths and row heights"), which no
// committed reference shows. ECMA-376 Part 1, 18.3.1.13, defines the maximum digit width, but
// not the pixel density, the line height, or the row padding; those rules are measured:
//
// - The grid has 72 pixels for each inch, so a pixel is a point.
// - The maximum digit width (MDW) is the widest digit advance of the default font at its size,
//   rounded to whole pixels: Carlito 11 is 6 pixels (5.58 before rounding), Carlito 20 is 10.
// - A glyph advances by whole pixels: each advance rounds on its own. Every glyph of the
//   reference PDFs starts on a whole point (for example `Shared` advances 5, 6, 5, 4, 5, 6).
// - The line height of a font is its ascent plus its descent (the OS/2 win values that the
//   measurer selects), each rounded to whole pixels: Carlito 11 is 10 + 3 = 13 pixels.
// - A row without `@ht` is the line height of its font plus 2 pixels on screen: 11-point text is
//   15 points, 20-point text is 26. No committed reference shows the screen geometry; the
//   approximation "Screen geometry" names it.
// - The reference PDFs print every row shorter: `floor(height * 14 / 15)` whole points, where
//   the height is `@ht` or the screen height of the row's font. grid-structure.pdf prints `ht`
//   10, 20, and 33.75 at 9, 18, and 31 points, and every default Carlito 11 row at 14. The print
//   profile is what the fidelity harness compares with.

import type { XlsxPackage } from '../store/package/package.ts';
import type { WorkbookIndex } from '../store/workbook/workbook-index.ts';
import { defaultFont, type ResolvedFont } from '../cascade/cell-format.ts';
import type { GridMetrics } from './grid-geometry.ts';
import type { ResolvedFace, TextMeasurer } from './text/measurer.ts';

/** How a target rounds the grid: its pixel density and the padding of a row. */
export interface GridProfile {
  /** Pixels for each inch of the grid. */
  readonly pixelsPerInch: number;
  /** Pixels that a row without `@ht` adds to the line height of its font. */
  readonly rowPadding: number;
  /** The factor of every row height before it rounds down to whole pixels. */
  readonly rowScale: number;
}

/** The printed page of the reference PDFs: the profile that the fidelity harness measures. */
export const PRINT_PROFILE: GridProfile = Object.freeze({
  pixelsPerInch: 72,
  rowPadding: 2,
  rowScale: 14 / 15,
});
/**
 * The screen: rows of the line height plus 2 pixels (15 points for Carlito 11), not scaled. No
 * committed reference shows them: the approximation "Screen geometry".
 */
export const SCREEN_PROFILE: GridProfile = Object.freeze({
  pixelsPerInch: 72,
  rowPadding: 2,
  rowScale: 1,
});

/**
 * The screen profile at a zoom: 96 CSS pixels to the inch at zoom 1. Grid pixels are CSS
 * pixels, so gridlines stay sharp.
 */
export const zoomProfile = (zoom: number): GridProfile => ({
  ...SCREEN_PROFILE,
  pixelsPerInch: 96 * zoom,
});

/** The font facts that layout needs, in whole pixels of a profile. */
export interface FontPixels {
  readonly face: ResolvedFace;
  /** The font size in pixels (not rounded). */
  readonly sizePx: number;
  readonly ascent: number;
  readonly descent: number;
  /** `ascent + descent`. */
  readonly lineHeight: number;
  /** The widest digit advance, rounded. */
  readonly maxDigitWidth: number;
}

/** The font properties that select a face and its size. */
export type FontKey = Pick<ResolvedFont, 'family' | 'bold' | 'italic' | 'sizePt'>;

const DIGITS = '0123456789';
/** Font facts kept for each measurer; a full cache starts again. */
const FONT_CACHE_LIMIT = 4096;
const fontCaches = new WeakMap<TextMeasurer, Map<string, FontPixels>>();

/** Font facts of one font in a profile. Cached for each measurer. */
export function fontPixels(
  measurer: TextMeasurer,
  font: FontKey,
  profile: GridProfile
): FontPixels {
  let cache = fontCaches.get(measurer);
  if (!cache) {
    cache = new Map();
    fontCaches.set(measurer, cache);
  }
  const key = `${font.family}\u0000${font.bold}|${font.italic}|${font.sizePt}|${profile.pixelsPerInch}`;
  const hit = cache.get(key);
  if (hit) return hit;
  if (cache.size >= FONT_CACHE_LIMIT) cache.clear();
  const facts = measureFont(measurer, font, profile);
  cache.set(key, facts);
  return facts;
}

function measureFont(measurer: TextMeasurer, font: FontKey, profile: GridProfile): FontPixels {
  const face = measurer.resolve({ family: font.family, bold: font.bold, italic: font.italic });
  const sizePx = (font.sizePt * profile.pixelsPerInch) / 72;
  const ascent = Math.round(face.metrics.ascent * sizePx);
  const descent = Math.round(face.metrics.descent * sizePx);
  let widest = 0;
  for (const advance of measurer.advances(face, DIGITS)) widest = Math.max(widest, advance);
  return {
    face,
    sizePx,
    ascent,
    descent,
    lineHeight: ascent + descent,
    maxDigitWidth: Math.max(1, Math.round(widest * sizePx)),
  };
}

/**
 * The width of text in whole pixels: each advance rounds on its own. `advances` are em for each
 * UTF-16 code unit, as the measurer returns them.
 */
export function pixelWidth(advances: ArrayLike<number>, sizePx: number): number {
  let width = 0;
  for (let i = 0; i < advances.length; i += 1) {
    const a = advances[i]!;
    if (a !== 0) width += Math.round(a * sizePx);
  }
  return width;
}

/**
 * The height in points of a row without `@ht` whose tallest font has these facts, before the
 * row scale of the profile: the height that `@ht` would state.
 */
export function autoRowHeight(font: FontPixels, profile: GridProfile): number {
  return ((font.lineHeight + profile.rowPadding) * 72) / profile.pixelsPerInch;
}

/**
 * A row height in points (`x:row/@ht`, `CT_Row`, or an auto height) as the profile lays it out.
 * The snap to whole pixels, rounded down, is measured.
 */
export function rowSize(
  height: number,
  profile: Pick<GridProfile, 'pixelsPerInch' | 'rowScale'>
): number {
  const ptPerPx = 72 / profile.pixelsPerInch;
  return Math.floor((height / ptPerPx) * profile.rowScale + 1e-9) * ptPerPx;
}

/** The grid metrics of a workbook in a profile: from its default font. */
export function gridMetrics(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  measurer: TextMeasurer,
  profile: GridProfile
): GridMetrics {
  const font = fontPixels(measurer, defaultFont(pkg, index), profile);
  return {
    maxDigitWidth: font.maxDigitWidth,
    pixelsPerInch: profile.pixelsPerInch,
    rowScale: profile.rowScale,
    defaultRowHeight: autoRowHeight(font, profile),
  };
}
