// The theme of a workbook as the cell format cascade reads it: the color scheme in theme index
// order, and the major and minor Latin typefaces (reference/quick-ref/themes-colors.md). The
// theme part root is `a:theme` (`CT_OfficeStyleSheet`, dml-main.xsd). Its `a:clrScheme` is
// `CT_ColorScheme`, its `a:fontScheme` is `CT_FontScheme`, `a:majorFont` and `a:minorFont` are
// `CT_FontCollection`, and `a:latin` is `CT_TextFont`.
//
// A workbook with no theme part, or a theme that lacks an entry, leaves that entry undefined;
// a color that names it resolves to the caller's default.

import type { Element } from '../store/package/tree.ts';
import type { XlsxPackage } from '../store/package/package.ts';
import { attr, child, path } from '../store/package/query.ts';
import { themeRoot, type WorkbookIndex } from '../store/workbook/workbook-index.ts';
import { parseRgb, type Rgb, type ThemePalette } from './color.ts';

/** The theme facts the cascade uses. */
export interface WorkbookTheme {
  readonly palette: ThemePalette;
  /** `a:majorFont/a:latin/@typeface`. */
  readonly majorLatin: string | undefined;
  /** `a:minorFont/a:latin/@typeface`. */
  readonly minorLatin: string | undefined;
}

/** Theme index order: index 0 is `lt1` and index 1 is `dk1` (measured on fixture colors-fonts). */
const THEME_ORDER = [
  'lt1',
  'dk1',
  'lt2',
  'dk2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
] as const;

/**
 * The color of one `a:clrScheme` entry: `a:srgbClr/@val` (`CT_SRgbColor`), or
 * `a:sysClr/@lastClr` (`CT_SystemColor`, an optional attribute).
 */
function schemeColor(entry: Element | undefined): Rgb | undefined {
  if (!entry) return undefined;
  const srgb = child(entry, 'a:srgbClr');
  if (srgb) return parseRgb(attr(srgb, 'val'));
  const sys = child(entry, 'a:sysClr');
  return sys ? parseRgb(attr(sys, 'lastClr')) : undefined;
}

const cache = new WeakMap<Element, WorkbookTheme>();
const EMPTY: WorkbookTheme = Object.freeze({
  palette: Object.freeze([]),
  majorLatin: undefined,
  minorLatin: undefined,
});

/** The theme of a workbook. Cached per theme root. */
export function workbookTheme(pkg: XlsxPackage, index: WorkbookIndex): WorkbookTheme {
  const root = themeRoot(pkg, index);
  if (!root) return EMPTY;
  const hit = cache.get(root);
  if (hit) return hit;
  const scheme = path(root, 'a:themeElements', 'a:clrScheme');
  const fonts = path(root, 'a:themeElements', 'a:fontScheme');
  const typeface = (name: string) => {
    const value = attr(path(fonts, name, 'a:latin'), 'typeface');
    return value === undefined || value === '' ? undefined : value;
  };
  const theme: WorkbookTheme = {
    palette: Object.freeze(THEME_ORDER.map((name) => schemeColor(child(scheme, `a:${name}`)))),
    majorLatin: typeface('a:majorFont'),
    minorLatin: typeface('a:minorFont'),
  };
  cache.set(root, theme);
  return theme;
}
