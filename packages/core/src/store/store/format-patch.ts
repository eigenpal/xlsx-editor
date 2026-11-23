// The format patch of the `applyFormat` op: the cell format properties that one command sets,
// as intent (design D2). The op never carries a cell format, font, fill, or border index; the
// store resolves or appends the style records when it applies the op (`style-records.ts`).
//
// Every value is checked here, before the store builds a record from it: a color is six hex
// digits, a size is a finite number of points in the range that the cascade reads, a font name
// is a short string of XML characters, and an alignment is a member of its ECMA-376
// enumeration. A custom number format code is checked for its length and its characters, and
// then for whether the number format parser reads it (`lowerNumberFormat`); a code that equals
// a built-in code of ECMA-376 Part 1, 18.8.30, is written as that built-in id.

import { builtinFormatId, formatCodeReadProblem } from '../number-format/builtins.ts';
import { isPlainXmlText } from './xml-text.ts';

/** The border lines that a patch draws (`ST_BorderStyle`, Part 1, 18.18.3). */
export const PATCH_BORDER_STYLES = [
  'thin',
  'medium',
  'thick',
  'dashed',
  'dotted',
  'double',
] as const;
export type PatchBorderStyle = (typeof PATCH_BORDER_STYLES)[number];

/** `ST_HorizontalAlignment` (Part 1, 18.18.40). */
export const PATCH_HORIZONTAL = [
  'general',
  'left',
  'center',
  'right',
  'fill',
  'justify',
  'centerContinuous',
  'distributed',
] as const;
export type PatchHorizontal = (typeof PATCH_HORIZONTAL)[number];

/** `ST_VerticalAlignment` (Part 1, 18.18.88). */
export const PATCH_VERTICAL = ['top', 'center', 'bottom', 'justify', 'distributed'] as const;
export type PatchVertical = (typeof PATCH_VERTICAL)[number];

/**
 * The cell format properties that one command sets. An absent property keeps each cell's own
 * value. Colors are `RRGGBB` hex digits.
 */
export interface FormatPatch {
  readonly bold?: boolean;
  readonly italic?: boolean;
  /** A single underline, or none. */
  readonly underline?: boolean;
  /** The font family. Setting it drops the theme font scheme of the cell's font. */
  readonly fontName?: string;
  /**
   * The font size in points, 1 to 409. The range is an engine limit: `x:sz/@val` (`CT_FontSize`,
   * sml.xsd) is an `xsd:double` with no bound.
   */
  readonly fontSize?: number;
  readonly fontColor?: string;
  /** A solid fill of one color, or no fill. */
  readonly fill?: { readonly type: 'solid'; readonly color: string } | { readonly type: 'none' };
  /**
   * Borders: `outline` draws the edges of the range, `all` draws every edge of every cell, and
   * `none` removes every edge and the diagonal. `style` is `thin` when absent; `color` is the
   * automatic color when absent.
   */
  readonly borders?: {
    readonly edges: 'outline' | 'all' | 'none';
    readonly style?: PatchBorderStyle;
    readonly color?: string;
  };
  readonly horizontal?: PatchHorizontal;
  readonly vertical?: PatchVertical;
  readonly wrapText?: boolean;
  /**
   * A built-in number format by id (0 to 163), or a custom format code. A custom code that a
   * number format of the styles part already holds reuses its id; a new code gets the lowest
   * unused id from 164 up. Give one of `id` and `code`. A patch that gives both is refused
   * with `bad-format`.
   */
  readonly numberFormat?:
    { readonly id: number; readonly code?: never } | { readonly code: string; readonly id?: never };
}

/**
 * The longest font name that a patch sets. Longer names are refused, not cut: 31 characters is
 * the longest face name that the font tables of common systems hold. This is an engine limit:
 * `x:name/@val` (`CT_FontName`, sml.xsd) is an `ST_Xstring` with no length limit.
 */
export const MAX_FONT_NAME_LENGTH = 31;
/**
 * The longest custom number format code that the number format parser reads. This is an engine
 * limit: `x:numFmt/@formatCode` is an `ST_Xstring` with no length limit.
 *
 * @public
 */
export const MAX_FORMAT_CODE_LENGTH = 255;
/** The custom number format ids start here (Part 1, 18.8.30). */
export const FIRST_CUSTOM_NUMFMT_ID = 164;

const HEX6 = /^[0-9A-Fa-f]{6}$/;

const isColor = (v: unknown): v is string => typeof v === 'string' && HEX6.test(v);
const oneOf = (list: readonly string[], v: unknown): boolean =>
  typeof v === 'string' && list.includes(v);
const isBool = (v: unknown): boolean => v === undefined || typeof v === 'boolean';

/** The properties that a patch can carry. */
const PATCH_KEYS: ReadonlySet<string> = new Set([
  'bold',
  'italic',
  'underline',
  'fontName',
  'fontSize',
  'fontColor',
  'fill',
  'borders',
  'horizontal',
  'vertical',
  'wrapText',
  'numberFormat',
]);

/** Why a font or fill value of a patch cannot be applied, or undefined. */
function fontFillProblem(patch: FormatPatch): string | undefined {
  if (!isBool(patch.bold) || !isBool(patch.italic) || !isBool(patch.underline))
    return 'bold, italic, and underline are true or false';
  const name = patch.fontName;
  if (
    name !== undefined &&
    (typeof name !== 'string' ||
      name.trim().length === 0 ||
      name.length > MAX_FONT_NAME_LENGTH ||
      !isPlainXmlText(name))
  )
    return `a font name is 1 to ${MAX_FONT_NAME_LENGTH} characters of text`;
  const size = patch.fontSize;
  if (size !== undefined && !(typeof size === 'number' && size >= 1 && size <= 409))
    return 'a font size is a number of points from 1 to 409';
  if (patch.fontColor !== undefined && !isColor(patch.fontColor))
    return 'a font color is six hex digits';
  const fill = patch.fill;
  if (fill !== undefined) {
    if (typeof fill !== 'object' || fill === null) return 'a fill is solid or none';
    if (fill.type === 'solid') {
      if (!isColor(fill.color)) return 'a fill color is six hex digits';
    } else if (fill.type !== 'none') return 'a fill is solid or none';
  }
  return undefined;
}

/**
 * Why a patch cannot be applied, or undefined. A patch that sets nothing is refused, so a
 * command never writes an empty transaction that looks like an edit.
 */
export function patchProblem(patch: FormatPatch): string | undefined {
  if (typeof patch !== 'object' || patch === null) return 'the format patch is not an object';
  let any = false;
  for (const key of Object.keys(patch)) {
    if (!PATCH_KEYS.has(key)) return `the format patch has an unknown property ${key}`;
    if ((patch as Readonly<Record<string, unknown>>)[key] !== undefined) any = true;
  }
  if (!any) return 'the format patch sets nothing';
  const fontFill = fontFillProblem(patch);
  if (fontFill) return fontFill;
  if (!isBool(patch.wrapText)) return 'wrapText is true or false';
  const borders = patch.borders;
  if (borders !== undefined) {
    if (
      typeof borders !== 'object' ||
      borders === null ||
      !oneOf(['outline', 'all', 'none'], borders.edges)
    )
      return 'borders are outline, all, or none';
    if (borders.style !== undefined && !oneOf(PATCH_BORDER_STYLES, borders.style))
      return `a border style is one of ${PATCH_BORDER_STYLES.join(', ')}`;
    if (borders.color !== undefined && !isColor(borders.color))
      return 'a border color is six hex digits';
  }
  if (patch.horizontal !== undefined && !oneOf(PATCH_HORIZONTAL, patch.horizontal))
    return 'the horizontal alignment is not an ECMA-376 value';
  if (patch.vertical !== undefined && !oneOf(PATCH_VERTICAL, patch.vertical))
    return 'the vertical alignment is not an ECMA-376 value';
  const nf = patch.numberFormat;
  if (nf !== undefined) {
    if (typeof nf !== 'object' || nf === null) return 'a number format is an id or a code';
    // Both would be silently read as the id: the patch names one of them.
    if (nf.id !== undefined && nf.code !== undefined)
      return 'a number format is an id or a code, not both';
    if (nf.id === undefined && nf.code === undefined) return 'a number format is an id or a code';
    if (nf.id !== undefined) {
      if (!Number.isInteger(nf.id) || nf.id < 0 || nf.id >= FIRST_CUSTOM_NUMFMT_ID)
        return `a built-in number format id is 0 to ${FIRST_CUSTOM_NUMFMT_ID - 1}`;
    } else if (
      typeof nf.code !== 'string' ||
      nf.code.length === 0 ||
      nf.code.length > MAX_FORMAT_CODE_LENGTH ||
      !isPlainXmlText(nf.code)
    )
      return `a number format code is 1 to ${MAX_FORMAT_CODE_LENGTH} characters of text`;
  }
  return undefined;
}

/** Which records of a cell format a patch changes. */
export interface PatchGroups {
  readonly font: boolean;
  readonly fill: boolean;
  readonly border: boolean;
  readonly alignment: boolean;
  readonly numFmt: boolean;
}

export function patchGroups(patch: FormatPatch): PatchGroups {
  return {
    font:
      patch.bold !== undefined ||
      patch.italic !== undefined ||
      patch.underline !== undefined ||
      patch.fontName !== undefined ||
      patch.fontSize !== undefined ||
      patch.fontColor !== undefined,
    fill: patch.fill !== undefined,
    border: patch.borders !== undefined,
    alignment:
      patch.horizontal !== undefined ||
      patch.vertical !== undefined ||
      patch.wrapText !== undefined,
    numFmt: patch.numberFormat !== undefined,
  };
}

/**
 * The patch with its custom number format code lowered: a code that equals a built-in code
 * becomes that built-in id, so the styles part gets no custom format for it. `defined` holds the
 * `x:numFmt` codes of the workbook by id: when it gives the built-in id another code, the code
 * stays a code, because the id would show that other code. A code that the number format
 * parser cannot read is a problem (`unreadable-number-format`). Other values of the patch pass
 * unchecked: `patchProblem` checks them.
 */
export function lowerNumberFormat(
  patch: FormatPatch,
  defined: ReadonlyMap<number, string> = new Map()
): { readonly patch: FormatPatch } | { readonly problem: string } {
  const nf = patch.numberFormat;
  if (typeof nf !== 'object' || nf === null || !('code' in nf) || typeof nf.code !== 'string')
    return { patch };
  const builtin = builtinFormatId(nf.code);
  const redefined = builtin !== undefined && (defined.get(builtin) ?? nf.code) !== nf.code;
  if (builtin !== undefined && !redefined)
    return { patch: { ...patch, numberFormat: { id: builtin } } };
  const problem = formatCodeReadProblem(nf.code);
  if (problem) return { problem: `${problem.code}${problem.detail ? `: ${problem.detail}` : ''}` };
  return { patch };
}
