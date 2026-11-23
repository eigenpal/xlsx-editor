// SpreadsheetML colors (`CT_Color`; the attribute table of `fgColor`, ECMA-376 Part 1, 18.8.19,
// defines them): `rgb`, `theme`, `indexed`, and `auto`, with `tint` applied in the HLS color
// space. See reference/quick-ref/themes-colors.md.
//
// Measured on fixture colors-fonts (fixtures/measured.json):
// - Theme indexes 0 to 3 are `lt1`, `dk1`, `lt2`, `dk2`, then the accents, as the quick
//   reference table states.
// - The alpha byte of `rgb` is ignored: `00FF0000` shows `FF0000`.
// - `auto` on a font or a border is black, the system foreground.
// - `auto` on a fill is white, the system background: the reference PDF paints a white
//   rectangle for a solid fill with `fgColor auto="1"` (colors-fonts.pdf, cell B23).
// - `indexed="64"` on a border is black, the system foreground.
// - Tint: the formula of the quick reference, computed in floating point, is within 2 of 255 on
//   each channel of the measured results (approximation `tint-rounding`).

import type { Element } from '../store/package/tree.ts';
import { attr } from '../store/package/query.ts';

/** A resolved color: `RRGGBB`, upper case. */
export type Rgb = string;

/** The default indexed palette, Part 1, 18.8.27. 64 and 65 are the system colors. */
export const DEFAULT_INDEXED: readonly Rgb[] = Object.freeze([
  '000000',
  'FFFFFF',
  'FF0000',
  '00FF00',
  '0000FF',
  'FFFF00',
  'FF00FF',
  '00FFFF',
  '000000',
  'FFFFFF',
  'FF0000',
  '00FF00',
  '0000FF',
  'FFFF00',
  'FF00FF',
  '00FFFF',
  '800000',
  '008000',
  '000080',
  '808000',
  '800080',
  '008080',
  'C0C0C0',
  '808080',
  '9999FF',
  '993366',
  'FFFFCC',
  'CCFFFF',
  '660066',
  'FF8080',
  '0066CC',
  'CCCCFF',
  '000080',
  'FF00FF',
  'FFFF00',
  '00FFFF',
  '800080',
  '800000',
  '008080',
  '0000FF',
  '00CCFF',
  'CCFFFF',
  'CCFFCC',
  'FFFF99',
  '99CCFF',
  'FF99CC',
  'CC99FF',
  'FFCC99',
  '3366FF',
  '33CCCC',
  '99CC00',
  'FFCC00',
  'FF9900',
  'FF6600',
  '666699',
  '969696',
  '003366',
  '339966',
  '003300',
  '333300',
  '993300',
  '993366',
  '333399',
  '333333',
]);

/** The system foreground and background colors, indexes 64 and 65. */
export const SYSTEM_FOREGROUND: Rgb = '000000';
export const SYSTEM_BACKGROUND: Rgb = 'FFFFFF';

/** The theme color scheme, in theme index order: lt1, dk1, lt2, dk2, accent1-6, hlink, folHlink. */
export type ThemePalette = readonly (Rgb | undefined)[];

/** What a color resolves against. */
export interface ColorContext {
  readonly theme: ThemePalette;
  /**
   * `x:colors/x:indexedColors/x:rgbColor/@rgb` (`CT_IndexedColors`) of the styles part, when the
   * file redefines the palette.
   */
  readonly indexed: readonly string[] | undefined;
}

/** Where a color is used: `auto` means black text and borders, and a white fill. */
export type ColorUse = 'font' | 'border' | 'fill';

const HEX = '0123456789ABCDEF';

/**
 * `RRGGBB` from an `ARGB` hex attribute (`CT_Color/@rgb`, `ST_UnsignedIntHex`) or an `RGB` one
 * (`a:srgbClr/@val`, `ST_HexColorRGB`), or undefined. The alpha byte is ignored.
 */
export function parseRgb(raw: string | undefined): Rgb | undefined {
  if (raw === undefined || (raw.length !== 8 && raw.length !== 6)) return undefined;
  const hex = raw.slice(raw.length - 6).toUpperCase();
  for (let i = 0; i < 6; i += 1) if (!HEX.includes(hex[i]!)) return undefined;
  return hex;
}

/** A small non-negative integer attribute, or undefined. */
function index(raw: string | undefined, max: number): number | undefined {
  if (raw === undefined || raw.length === 0 || raw.length > 4) return undefined;
  let n = 0;
  for (let i = 0; i < raw.length; i += 1) {
    const c = raw.charCodeAt(i);
    if (c < 48 || c > 57) return undefined;
    n = n * 10 + (c - 48);
  }
  return n <= max ? n : undefined;
}

/**
 * Resolve a `CT_Color` element: `@auto`, then `@rgb`, `@theme`, and `@indexed`, with `@tint`
 * (`xsd:double`, default 0) on the result. Returns undefined for a missing element and for a
 * color that names nothing; the caller then uses its default.
 */
export function resolveColor(
  el: Element | undefined,
  context: ColorContext,
  use: ColorUse
): Rgb | undefined {
  if (!el) return undefined;
  let base: Rgb | undefined;
  if (attr(el, 'auto') === '1' || attr(el, 'auto') === 'true')
    base = use === 'fill' ? SYSTEM_BACKGROUND : SYSTEM_FOREGROUND;
  else if (attr(el, 'rgb') !== undefined) base = parseRgb(attr(el, 'rgb'));
  else if (attr(el, 'theme') !== undefined) {
    const t = index(attr(el, 'theme'), 11);
    base = t === undefined ? undefined : context.theme[t];
  } else if (attr(el, 'indexed') !== undefined) base = indexedColor(attr(el, 'indexed'), context);
  if (base === undefined) return undefined;
  const tint = Number(attr(el, 'tint') ?? '0');
  return Number.isFinite(tint) && tint !== 0
    ? applyTint(base, Math.max(-1, Math.min(1, tint)))
    : base;
}

function indexedColor(raw: string | undefined, context: ColorContext): Rgb | undefined {
  const i = index(raw, 65);
  if (i === undefined) return undefined;
  if (i === 64) return SYSTEM_FOREGROUND;
  if (i === 65) return SYSTEM_BACKGROUND;
  const custom = context.indexed?.[i];
  return custom !== undefined ? parseRgb(custom) : DEFAULT_INDEXED[i];
}

/** `CT_Color/@tint`: lightness in HLS, `L' = L(1+t)` or `L(1-t) + t`, with L from 0 to 1. */
export function applyTint(rgb: Rgb, tint: number): Rgb {
  const [h, l, s] = rgbToHls(rgb);
  const lightness = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  return hlsToRgb(h, lightness, s);
}

function rgbToHls(rgb: Rgb): [number, number, number] {
  const r = parseInt(rgb.slice(0, 2), 16) / 255;
  const g = parseInt(rgb.slice(2, 4), 16) / 255;
  const b = parseInt(rgb.slice(4, 6), 16) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, l, 0];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  h /= 6;
  return [h, l, s];
}

function hueChannel(p: number, q: number, t: number): number {
  let x = t;
  if (x < 0) x += 1;
  if (x > 1) x -= 1;
  if (x < 1 / 6) return p + (q - p) * 6 * x;
  if (x < 1 / 2) return q;
  if (x < 2 / 3) return p + (q - p) * (2 / 3 - x) * 6;
  return p;
}

function hlsToRgb(h: number, l: number, s: number): Rgb {
  const channels =
    s === 0
      ? [l, l, l]
      : (() => {
          const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
          const p = 2 * l - q;
          return [hueChannel(p, q, h + 1 / 3), hueChannel(p, q, h), hueChannel(p, q, h - 1 / 3)];
        })();
  let out = '';
  for (const v of channels) {
    const n = Math.max(0, Math.min(255, Math.round(v * 255)));
    out += HEX[n >> 4]! + HEX[n & 15]!;
  }
  return out;
}
