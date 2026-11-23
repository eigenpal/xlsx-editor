// The cell format cascade: `x:c/@s` through `x:cellXfs` to the number format, font, fill,
// border, alignment, and protection of a cell (ECMA-376 Part 1, 18.8.45 and 18.8.10). The
// schema types (transitional `sml.xsd`): `x:xf` is `CT_Xf`, `x:cellXfs` is `CT_CellXfs`,
// `x:font` is `CT_Font`, `x:fill` is `CT_Fill`, `x:border` is `CT_Border`, `x:alignment` is
// `CT_CellAlignment`, and `x:protection` is `CT_CellProtection`.
//
// Measured on fixture colors-fonts (fixtures/measured.json): a cell format whose
// `applyFont` is absent still shows its font (B29 is bold and red), so the cascade reads the
// format's own `fontId`, `fillId`, `borderId`, and `numFmtId`, whatever the `apply*` attributes
// say. A font with `x:scheme` `minor` or `major` shows the theme typeface, not `x:name` (B27).
//
// An index out of range falls back to the default (index 0) and returns a problem; the caller
// records it with the cell reference. Results are cached per styles part revision and index.
//
// File strings are sanitized here, at the cascade boundary: alignment, border style, and pattern
// values outside their ECMA-376 enumerations (transitional `sml.xsd`) fall back to the default,
// and a font family is capped at 255 characters.

import type { Element } from '../store/package/tree.ts';
import type { XlsxPackage } from '../store/package/package.ts';
import { attr, child } from '../store/package/query.ts';
import {
  styleLists,
  type StyleLists,
  type WorkbookIndex,
} from '../store/workbook/workbook-index.ts';
import {
  resolveColor,
  SYSTEM_BACKGROUND,
  SYSTEM_FOREGROUND,
  type ColorContext,
  type Rgb,
} from './color.ts';
import type { CascadeProblem } from './problems.ts';
import { workbookTheme, type WorkbookTheme } from './theme.ts';
import { runPropsOf, type TextRun } from './values.ts';

/** `x:u/@val` (`ST_UnderlineValues`). */
export type UnderlineStyle = 'none' | 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';

/** A resolved `x:font` (`CT_Font`) or run properties `x:rPr` (`CT_RPrElt`). */
export interface ResolvedFont {
  readonly family: string;
  /** Font size in points. */
  readonly sizePt: number;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly underline: UnderlineStyle;
  readonly strike: boolean;
  readonly color: Rgb;
  /** `x:vertAlign/@val` (`ST_VerticalAlignRun`, shared-commonSimpleTypes.xsd). */
  readonly vertAlign: 'baseline' | 'superscript' | 'subscript';
}

/**
 * A resolved `x:fill` (`CT_Fill`): `x:patternFill` (`CT_PatternFill`, `@patternType` of
 * `ST_PatternType`) or `x:gradientFill` (`CT_GradientFill`, stops of `CT_GradientStop`).
 */
export type ResolvedFill =
  | { readonly kind: 'none' }
  | { readonly kind: 'solid'; readonly color: Rgb }
  | {
      readonly kind: 'pattern';
      readonly pattern: string;
      readonly foreground: Rgb;
      readonly background: Rgb | undefined;
    }
  | {
      readonly kind: 'gradient';
      /** `@type` (`ST_GradientType`, default `linear`). */
      readonly type: 'linear' | 'path';
      /** `@degree` (`xsd:double`, default 0). */
      readonly degree: number;
      /** The inner rectangle of a `path` gradient; all 0 for a `linear` one. */
      readonly rect: GradientRect;
      /** `x:stop/@position` and `x:stop/x:color` (`CT_GradientStop`). */
      readonly stops: readonly { readonly position: number; readonly color: Rgb }[];
    };

/**
 * The inner rectangle of a `path` gradient (`@left`, `@right`, `@top`, `@bottom`): the positions
 * of its edges as fractions of the width and height, from the left and top edges. Paint applies
 * them to the box of the cell: the cell and its right and bottom gridline pixel. Measured on
 * fixture `fills-patterns`: `left="0.2" right="0.4"` paints the first stop from 0.2 to
 * 0.4 of the width. Each value lies in 0 to 1, with `left <= right` and `top <= bottom`.
 */
export interface GradientRect {
  readonly left: number;
  readonly right: number;
  readonly top: number;
  readonly bottom: number;
}

/** One edge of a border (`CT_BorderPr`). */
export interface BorderEdge {
  /** `ST_BorderStyle`; `none` when the edge has no line. */
  readonly style: string;
  readonly color: Rgb;
}

/** A resolved `x:border` (`CT_Border`), with `@diagonalUp` and `@diagonalDown`. */
export interface ResolvedBorder {
  readonly left: BorderEdge;
  readonly right: BorderEdge;
  readonly top: BorderEdge;
  readonly bottom: BorderEdge;
  readonly diagonal: BorderEdge;
  readonly diagonalUp: boolean;
  readonly diagonalDown: boolean;
}

/** A resolved `x:alignment` (`CT_CellAlignment`). */
export interface ResolvedAlignment {
  /** `ST_HorizontalAlignment`; `general` aligns by value type. */
  readonly horizontal: string;
  /** `ST_VerticalAlignment`; the schema default is `bottom`. */
  readonly vertical: string;
  readonly wrapText: boolean;
  readonly shrinkToFit: boolean;
  readonly indent: number;
  /** `ST_TextRotation`: 0 to 180 degrees, or 255 for vertical stacked text. */
  readonly textRotation: number;
}

/** A resolved `x:protection` (`CT_CellProtection`): `@locked` and `@hidden`. */
export interface ResolvedProtection {
  readonly locked: boolean;
  readonly hidden: boolean;
}

export interface ResolvedCellFormat {
  /** The cell format index that was used: the cell's `@s`, or 0 after a fallback. */
  readonly xf: number;
  readonly numFmtId: number;
  /** The custom format code of `numFmtId`, when the styles part defines one. */
  readonly formatCode: string | undefined;
  readonly font: ResolvedFont;
  readonly fill: ResolvedFill;
  readonly border: ResolvedBorder;
  readonly alignment: ResolvedAlignment;
  readonly protection: ResolvedProtection;
}

export interface CellFormatResult {
  readonly format: ResolvedCellFormat;
  /** Why the format fell back, for the caller to record with the cell reference. */
  readonly problems: readonly CascadeProblem[];
}

/** One run of a rich string with its resolved font. */
export interface RichTextRun {
  readonly text: string;
  readonly font: ResolvedFont;
}

/** `ST_HorizontalAlignment`, Part 1, 18.18.40. */
const HORIZONTAL = new Set([
  'general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed',
]); // prettier-ignore
/** `ST_VerticalAlignment`, Part 1, 18.18.88. */
const VERTICAL = new Set(['top', 'center', 'bottom', 'justify', 'distributed']);
/** `ST_BorderStyle`, Part 1, 18.18.3. */
const BORDER_STYLES = new Set([
  'none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed',
  'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot',
]); // prettier-ignore
/** `ST_PatternType`, Part 1, 18.18.55. */
const PATTERNS = new Set([
  'none', 'solid', 'mediumGray', 'darkGray', 'lightGray', 'darkHorizontal', 'darkVertical',
  'darkDown', 'darkUp', 'darkGrid', 'darkTrellis', 'lightHorizontal', 'lightVertical',
  'lightDown', 'lightUp', 'lightGrid', 'lightTrellis', 'gray125', 'gray0625',
]); // prettier-ignore
/** `ST_UnderlineValues`. */
const UNDERLINES = new Set(['single', 'double', 'singleAccounting', 'doubleAccounting', 'none']);

/** The longest font family the cascade keeps. */
const MAX_FAMILY_LENGTH = 255;

/** The font of a cell format with no font element. */
const DEFAULT_FONT: ResolvedFont = Object.freeze({
  family: 'Calibri',
  sizePt: 11,
  bold: false,
  italic: false,
  underline: 'none',
  strike: false,
  color: SYSTEM_FOREGROUND,
  vertAlign: 'baseline',
});

/** `value` when it is in `allowed`, else `fallback`. */
const member = (allowed: ReadonlySet<string>, value: string | undefined, fallback: string) =>
  value !== undefined && allowed.has(value) ? value : fallback;

const WHITE_FILL: ResolvedFill = Object.freeze({ kind: 'solid', color: SYSTEM_BACKGROUND });

const NO_EDGE: BorderEdge = Object.freeze({ style: 'none', color: SYSTEM_FOREGROUND });

/** A non-negative decimal integer of up to 9 digits, or undefined. */
function nonNegative(raw: string | undefined): number | undefined {
  if (raw === undefined || raw.length === 0 || raw.length > 9) return undefined;
  let n = 0;
  for (let i = 0; i < raw.length; i += 1) {
    const c = raw.charCodeAt(i);
    if (c < 48 || c > 57) return undefined;
    n = n * 10 + (c - 48);
  }
  return n;
}

/** A `CT_BooleanProperty`: undefined when absent, else its value (absent `val` is true). */
const flag = (el: Element | undefined): boolean | undefined => {
  if (!el) return undefined;
  const v = attr(el, 'val');
  return v === undefined || v === '1' || v === 'true';
};

/** A font family name: empty is absent, and the name is capped at 255 characters. */
function familyName(raw: string | undefined): string | undefined {
  return raw === undefined || raw === '' ? undefined : raw.slice(0, MAX_FAMILY_LENGTH);
}

/**
 * A `CT_Font` (`nameTag` `x:name`) or a run's `CT_RPrElt` (`nameTag` `x:rFont`). Each absent
 * property takes the value of `base`.
 */
function resolveFont(
  el: Element | undefined,
  base: ResolvedFont,
  nameTag: 'x:name' | 'x:rFont',
  theme: WorkbookTheme,
  colors: ColorContext
): ResolvedFont {
  const scheme = attr(child(el, 'x:scheme'), 'val');
  const themed =
    scheme === 'major' ? theme.majorLatin : scheme === 'minor' ? theme.minorLatin : undefined;
  const szRaw = attr(child(el, 'x:sz'), 'val');
  const size = szRaw === undefined ? NaN : Number(szRaw);
  const u = child(el, 'x:u');
  const underline = u ? member(UNDERLINES, attr(u, 'val') ?? 'single', 'single') : base.underline;
  const vertEl = child(el, 'x:vertAlign');
  const vert = attr(vertEl, 'val');
  return {
    family: familyName(themed) ?? familyName(attr(child(el, nameTag), 'val')) ?? base.family,
    sizePt: Number.isFinite(size) && size >= 1 && size <= 409 ? size : base.sizePt,
    bold: flag(child(el, 'x:b')) ?? base.bold,
    italic: flag(child(el, 'x:i')) ?? base.italic,
    underline: underline as UnderlineStyle,
    strike: flag(child(el, 'x:strike')) ?? base.strike,
    color: resolveColor(child(el, 'x:color'), colors, 'font') ?? base.color,
    vertAlign:
      vert === 'superscript' || vert === 'subscript' || vert === 'baseline'
        ? vert
        : vertEl
          ? 'baseline'
          : base.vertAlign,
  };
}

/** The gradient stops of `x:gradientFill`. At most 16 are read. */
function gradient(el: Element, colors: ColorContext, problems: CascadeProblem[]): ResolvedFill {
  const stops: { position: number; color: Rgb }[] = [];
  for (const node of el.children) {
    if (stops.length >= 16) break;
    if (node.type !== 'element' || node.local !== 'stop') continue;
    const position = Number(attr(node, 'position'));
    const color = resolveColor(child(node, 'x:color'), colors, 'font');
    if (Number.isFinite(position) && color) stops.push({ position, color });
  }
  const degree = Number(attr(el, 'degree') ?? '0');
  const type = attr(el, 'type') === 'path' ? 'path' : 'linear';
  return {
    kind: 'gradient',
    type,
    degree: Number.isFinite(degree) ? degree : 0,
    rect: type === 'path' ? gradientRect(el, problems) : NO_RECT,
    stops,
  };
}

const NO_RECT: GradientRect = Object.freeze({ left: 0, right: 0, top: 0, bottom: 0 });

/** A finite `xsd:double`, with no catastrophic backtracking. */
const XSD_DOUBLE = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;

/**
 * The inner rectangle of a `path` gradient. An absent edge is 0, the schema default. A value
 * outside 0 to 1 is limited, a value that is not a number is 0, and a pair in reverse order is put
 * in order, with the problem `gradient-rectangle-out-of-range`.
 */
function gradientRect(el: Element, problems: CascadeProblem[]): GradientRect {
  let changed: string | undefined;
  const edge = (name: string): number => {
    const raw = attr(el, name);
    if (raw === undefined) return 0;
    // The lexical space of `xsd:double`: `INF` and `-INF` limit to 1 and 0, and any other text,
    // such as `0x1`, is not a number.
    const text = raw.trim();
    const value =
      text === 'INF' || text === '+INF'
        ? Infinity
        : text === '-INF'
          ? -Infinity
          : XSD_DOUBLE.test(text)
            ? Number(text)
            : NaN;
    if (value >= 0 && value <= 1) return value + 0;
    changed ??= `${name}=${raw.slice(0, 16)}`;
    return Number.isNaN(value) ? 0 : Math.min(1, Math.max(0, value));
  };
  let left = edge('left');
  let right = edge('right');
  let top = edge('top');
  let bottom = edge('bottom');
  if (left > right) {
    changed ??= 'left>right';
    [left, right] = [right, left];
  }
  if (top > bottom) {
    changed ??= 'top>bottom';
    [top, bottom] = [bottom, top];
  }
  if (changed !== undefined)
    problems.push({ code: 'gradient-rectangle-out-of-range', detail: changed });
  return { left, right, top, bottom };
}

function resolveFill(
  el: Element | undefined,
  colors: ColorContext,
  problems: CascadeProblem[]
): ResolvedFill {
  const grad = child(el, 'x:gradientFill');
  if (grad) return gradient(grad, colors, problems);
  const pattern = child(el, 'x:patternFill');
  const type = member(PATTERNS, attr(pattern, 'patternType') ?? 'solid', 'none');
  if (!pattern || type === 'none') return { kind: 'none' };
  const fg = resolveColor(child(pattern, 'x:fgColor'), colors, 'fill');
  const bg = resolveColor(child(pattern, 'x:bgColor'), colors, 'fill');
  if (type === 'solid') return fg ? { kind: 'solid', color: fg } : { kind: 'none' };
  return { kind: 'pattern', pattern: type, foreground: fg ?? SYSTEM_FOREGROUND, background: bg };
}

function edge(el: Element | undefined, colors: ColorContext): BorderEdge {
  const style = member(BORDER_STYLES, attr(el, 'style'), 'none');
  if (!el || style === 'none') return NO_EDGE;
  return {
    style,
    color: resolveColor(child(el, 'x:color'), colors, 'border') ?? SYSTEM_FOREGROUND,
  };
}

function resolveBorder(el: Element | undefined, colors: ColorContext): ResolvedBorder {
  // `start` and `end` are the leading and trailing edges (`CT_Border`); a left-to-right sheet
  // reads them as left and right when those are absent.
  return {
    left: edge(child(el, 'x:left') ?? child(el, 'x:start'), colors),
    right: edge(child(el, 'x:right') ?? child(el, 'x:end'), colors),
    top: edge(child(el, 'x:top'), colors),
    bottom: edge(child(el, 'x:bottom'), colors),
    diagonal: edge(child(el, 'x:diagonal'), colors),
    diagonalUp: attr(el, 'diagonalUp') === '1' || attr(el, 'diagonalUp') === 'true',
    diagonalDown: attr(el, 'diagonalDown') === '1' || attr(el, 'diagonalDown') === 'true',
  };
}

function resolveAlignment(el: Element | undefined): ResolvedAlignment {
  const on = (name: string) => attr(el, name) === '1' || attr(el, name) === 'true';
  const indent = nonNegative(attr(el, 'indent')) ?? 0;
  const rotation = nonNegative(attr(el, 'textRotation')) ?? 0;
  return {
    horizontal: member(HORIZONTAL, attr(el, 'horizontal'), 'general'),
    vertical: member(VERTICAL, attr(el, 'vertical'), 'bottom'),
    wrapText: on('wrapText'),
    shrinkToFit: on('shrinkToFit'),
    indent: Math.min(indent, 250),
    textRotation: rotation <= 180 || rotation === 255 ? rotation : 0,
  };
}

function resolveProtection(el: Element | undefined): ResolvedProtection {
  return {
    locked: attr(el, 'locked') !== '0' && attr(el, 'locked') !== 'false',
    hidden: attr(el, 'hidden') === '1' || attr(el, 'hidden') === 'true',
  };
}

interface FormatCache {
  readonly theme: WorkbookTheme;
  readonly byIndex: Map<number, CellFormatResult>;
}

const caches = new WeakMap<StyleLists, FormatCache>();

/** At most this many distinct cell formats are cached per styles part. */
const CACHE_CAP = 65_536;

/**
 * The resolved format of a cell format index (`x:c/@s`, as the file wrote it, or undefined for
 * the default). Out-of-range indexes fall back to index 0 and return a problem.
 */
export function cellFormat(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  s: string | undefined
): CellFormatResult {
  const lists = styleLists(pkg, index);
  const theme = workbookTheme(pkg, index);
  let cache = caches.get(lists);
  // The theme is a separate part: a new theme revision starts a new cache.
  if (!cache || cache.theme !== theme) {
    cache = { theme, byIndex: new Map() };
    caches.set(lists, cache);
  }
  const problems: CascadeProblem[] = [];
  let xf = s === undefined ? 0 : nonNegative(s);
  if (xf === undefined || (xf > 0 && xf >= lists.cellXfs.length)) {
    problems.push({ code: 'cell-format-out-of-range', detail: (s ?? '').slice(0, 16) });
    xf = 0;
  }
  const hit = cache.byIndex.get(xf);
  if (hit)
    return problems.length ? { format: hit.format, problems: [...problems, ...hit.problems] } : hit;
  const result = resolveXf(lists, cache.theme, xf);
  if (cache.byIndex.size < CACHE_CAP) cache.byIndex.set(xf, result);
  return problems.length
    ? { format: result.format, problems: [...problems, ...result.problems] }
    : result;
}

function pick(
  list: readonly Element[],
  raw: string | undefined,
  code: CascadeProblem['code'],
  problems: CascadeProblem[]
): Element | undefined {
  const i = raw === undefined ? 0 : nonNegative(raw);
  if (i !== undefined && i < list.length) return list[i];
  // An index out of range is reported also when the list is empty.
  if (raw !== undefined) problems.push({ code, detail: raw.slice(0, 16) });
  return list[0];
}

function resolveXf(lists: StyleLists, theme: WorkbookTheme, xf: number): CellFormatResult {
  const el = lists.cellXfs[xf];
  const colors: ColorContext = { theme: theme.palette, indexed: lists.indexedColors };
  const problems: CascadeProblem[] = [];
  const rawNumFmtId = attr(el, 'numFmtId');
  let numFmtId = rawNumFmtId === undefined ? 0 : nonNegative(rawNumFmtId);
  if (numFmtId === undefined) {
    problems.push({ code: 'number-format-id-malformed', detail: rawNumFmtId!.slice(0, 16) });
    numFmtId = 0;
  }
  const format: ResolvedCellFormat = {
    xf,
    numFmtId,
    formatCode: lists.numFmts.get(numFmtId),
    font: resolveFont(
      pick(lists.fonts, attr(el, 'fontId'), 'font-out-of-range', problems),
      DEFAULT_FONT,
      'x:name',
      theme,
      colors
    ),
    fill: resolveFill(
      pick(lists.fills, attr(el, 'fillId'), 'fill-out-of-range', problems),
      colors,
      problems
    ),
    border: resolveBorder(
      pick(lists.borders, attr(el, 'borderId'), 'border-out-of-range', problems),
      colors
    ),
    alignment: resolveAlignment(child(el, 'x:alignment')),
    protection: resolveProtection(child(el, 'x:protection')),
  };
  return { format, problems };
}

/** The resolved runs of one list of runs, for each cell font, with the styles they came from. */
const resolvedRuns = new WeakMap<
  readonly TextRun[],
  WeakMap<ResolvedFont, { lists: StyleLists; theme: WorkbookTheme; runs: readonly RichTextRun[] }>
>();

/**
 * The runs of a rich string with resolved fonts. Each run reads its `x:rPr` through the font
 * rules of the cascade, and a property that the run does not set comes from `cellFont`. The
 * sender controls the number of runs, so the same runs resolve once for each cell font.
 */
export function richTextRuns(
  runs: readonly TextRun[],
  cellFont: ResolvedFont,
  pkg: XlsxPackage,
  index: WorkbookIndex
): readonly RichTextRun[] {
  const lists = styleLists(pkg, index);
  const theme = workbookTheme(pkg, index);
  let fonts = resolvedRuns.get(runs);
  const hit = fonts?.get(cellFont);
  if (hit && hit.lists === lists && hit.theme === theme) return hit.runs;
  const colors: ColorContext = { theme: theme.palette, indexed: lists.indexedColors };
  const resolved = runs.map((run) => {
    const props = runPropsOf(run);
    return {
      text: run.text,
      font: props ? resolveFont(props, cellFont, 'x:rFont', theme, colors) : cellFont,
    };
  });
  if (!fonts) resolvedRuns.set(runs, (fonts = new WeakMap()));
  fonts.set(cellFont, { lists, theme, runs: resolved });
  return resolved;
}

const defaultFonts = new WeakMap<StyleLists, { theme: WorkbookTheme; font: ResolvedFont }>();

/**
 * The default font of a workbook: the font of the Normal cell style (`x:cellStyle` with
 * `builtinId="0"`, then its `xfId` in `x:cellStyleXfs`, then that format's `fontId`), or the
 * font of `x:cellXfs[0]` when the styles part names no Normal style (quick reference, "Styles").
 * It sets the maximum digit width of column widths and the default row height.
 */
export function defaultFont(pkg: XlsxPackage, index: WorkbookIndex): ResolvedFont {
  const lists = styleLists(pkg, index);
  const theme = workbookTheme(pkg, index);
  const hit = defaultFonts.get(lists);
  if (hit && hit.theme === theme) return hit.font;
  const colors: ColorContext = { theme: theme.palette, indexed: lists.indexedColors };
  const normal = lists.cellStyles.find((el) => attr(el, 'builtinId') === '0');
  const styleXf = normal ? nonNegative(attr(normal, 'xfId')) : undefined;
  const xf =
    styleXf !== undefined && styleXf < lists.cellStyleXfs.length
      ? lists.cellStyleXfs[styleXf]
      : lists.cellXfs[0];
  const id = nonNegative(attr(xf, 'fontId')) ?? 0;
  const font = resolveFont(
    lists.fonts[id] ?? lists.fonts[0],
    DEFAULT_FONT,
    'x:name',
    theme,
    colors
  );
  defaultFonts.set(lists, { theme, font });
  return font;
}

/**
 * `format` with the differential format `dxf` (`x:dxf`, `CT_Dxf`, Part 1, 18.8.14) laid over
 * it, for a true conditional formatting rule. The font properties that `x:font` sets replace the
 * cell's, including the family and size; the row height does not change. A solid or untyped
 * `x:patternFill` paints its `x:bgColor`. Without `x:bgColor`, it keeps the fill of the cell
 * format, and paints white over a cell with no fill. Other differential fills keep the fill of
 * the cell format. An `x:numFmt` replaces the number format. Borders, alignment, and protection
 * do not apply.
 */
export function overlayDxf(
  format: ResolvedCellFormat,
  dxf: Element,
  theme: WorkbookTheme,
  indexedColors: readonly string[] | undefined
): ResolvedCellFormat {
  const colors: ColorContext = { theme: theme.palette, indexed: indexedColors };
  const fontEl = child(dxf, 'x:font');
  const font = fontEl ? resolveFont(fontEl, format.font, 'x:name', theme, colors) : format.font;
  const pattern = child(child(dxf, 'x:fill'), 'x:patternFill');
  // `ST_PatternType` is a token: white space around it does not change it.
  const type = attr(pattern, 'patternType')?.trim();
  const bg =
    pattern && (type === undefined || type === 'solid')
      ? resolveColor(child(pattern, 'x:bgColor'), colors, 'fill')
      : undefined;
  // Without `x:bgColor`, the cell keeps its own fill, and a cell with no fill paints the
  // background color, white (`conditional-rules.pdf`, B10, B23, and B24).
  const fill: ResolvedFill = bg
    ? { kind: 'solid', color: bg }
    : pattern && (type === undefined || type === 'solid') && format.fill.kind === 'none'
      ? WHITE_FILL
      : format.fill;
  const numFmt = child(dxf, 'x:numFmt');
  const id = nonNegative(attr(numFmt, 'numFmtId'));
  const code = attr(numFmt, 'formatCode');
  const number =
    id !== undefined && code !== undefined ? { numFmtId: id, formatCode: code } : undefined;
  if (font === format.font && fill === format.fill && !number) return format;
  return { ...format, ...number, font, fill };
}
