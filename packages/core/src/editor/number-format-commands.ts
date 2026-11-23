// The number format commands of the toolbar, DOM-free: the entries of the number format list
// with the code that each one writes, and the decimal step of the increase decimals and
// decrease decimals commands.
//
// A list entry writes a format code, never a bare id. The store writes a code that equals a
// built-in code of ECMA-376 Part 1, 18.8.30, as that id, but only when the workbook gives the id
// no other code (`lowerNumberFormat`). So an entry never writes an id that shows another code.
//
// The decimal step reads a format code that the file supplies, so it is bounded and linear: the
// engine's number format parser (`parseFormatCode`, at most 255 characters) classifies each
// section, and one pass over the characters finds the digit placeholders that the parser reads.
// No regular expression reads the code. The step changes one place in each number section and
// keeps every other character: the sections, the colors, the conditions, the literals, and the
// exponent part. A code that the step cannot change in that way is refused with a reason.

import { builtinFormatCode, EN_US, formatValue, type NumberLocale } from '../cascade/index.ts';
// Lane-internal paths: the parser and the symbol check are not part of the public entries.
import { currencySymbolOf } from '../cascade/number-format/builtins.ts';
import { parseFormatCode, type Section as FormatSection } from '../store/number-format/parse.ts';
import { FIRST_CUSTOM_NUMFMT_ID, MAX_FORMAT_CODE_LENGTH } from '../store/index.ts';
import { isPlainXmlText } from '../store/store/xml-text.ts';

/** The value of an entry of the number format list. @public */
export type NumberFormatChoiceValue =
  | 'general'
  | 'number'
  | 'currency'
  | 'accounting'
  | 'percent'
  | 'scientific'
  | 'fraction'
  | 'shortDate'
  | 'longDate'
  | 'time'
  | 'text';

/**
 * One entry of the number format list: its value, the built-in id that it names (null for an
 * entry with no built-in id), and the catalog key of its name.
 *
 * @public
 */
export interface NumberFormatChoice {
  readonly value: NumberFormatChoiceValue;
  /** The built-in id that it names (ECMA-376 Part 1, 18.8.30), or null for a custom code. */
  readonly id: number | null;
  readonly labelKey: `numberFormat.${NumberFormatChoiceValue}`;
}

/**
 * The entries of the number format list. An entry with an `id` writes that built-in id of
 * ECMA-376 Part 1, 18.8.30, or the code of the id in the number locale when the workbook gives
 * the id another code. An entry with a null `id` writes a code that the format id table has no
 * fixed id for: currency and accounting with the currency symbol of the number locale
 * (`NumberLocale.currencySymbol`), and the long date. `numberFormatChoiceCode` gives the code
 * of each entry.
 *
 * The list shows `numberFormatChoices(locale)`: it leaves out the entries that have no code in
 * the number locale.
 *
 * @internal
 */
export const NUMBER_FORMAT_CHOICES = [
  { value: 'general', id: 0, labelKey: 'numberFormat.general' },
  { value: 'number', id: 2, labelKey: 'numberFormat.number' },
  { value: 'currency', id: null, labelKey: 'numberFormat.currency' },
  { value: 'accounting', id: null, labelKey: 'numberFormat.accounting' },
  { value: 'percent', id: 10, labelKey: 'numberFormat.percent' },
  { value: 'scientific', id: 11, labelKey: 'numberFormat.scientific' },
  { value: 'fraction', id: 12, labelKey: 'numberFormat.fraction' },
  { value: 'shortDate', id: 14, labelKey: 'numberFormat.shortDate' },
  { value: 'longDate', id: null, labelKey: 'numberFormat.longDate' },
  { value: 'time', id: 21, labelKey: 'numberFormat.time' },
  { value: 'text', id: 49, labelKey: 'numberFormat.text' },
] as const satisfies readonly NumberFormatChoice[];

/** Whether a value is the value of an entry of the number format list. @internal */
export function isNumberFormatChoice(value: string): value is NumberFormatChoiceValue {
  return NUMBER_FORMAT_CHOICES.some((c) => c.value === value);
}

/** The code of the long date entry: weekday, month name, day, and four-digit year. */
const LONG_DATE_CODE = 'dddd, mmmm d, yyyy';

/**
 * A literal of a format code: the text in quotation marks, or each character after a
 * backslash when the text holds a quotation mark (ECMA-376 Part 1, 18.8.31).
 */
function literal(text: string): string {
  if (!text.includes('"')) return `"${text}"`;
  let out = '';
  for (const ch of text) out += `\\${ch}`;
  return out;
}

/**
 * The format code that an entry of the number format list shows in a number locale: the code
 * of its built-in id in the locale (`builtins` of the locale, or the ECMA-376 table), or the
 * code that it writes. Currency and accounting put the currency symbol of the locale before the
 * number, with two decimal places. Undefined for currency and accounting when the locale's
 * currency symbol is not valid: the list then leaves the two entries out.
 *
 * @public
 */
export function numberFormatChoiceCode(
  value: NumberFormatChoiceValue,
  locale: NumberLocale = EN_US
): string | undefined {
  const choice = NUMBER_FORMAT_CHOICES.find((c) => c.value === value);
  if (!choice) return undefined;
  if (choice.id !== null) return builtinFormatCode(choice.id, locale);
  if (choice.value === 'longDate') return LONG_DATE_CODE;
  const symbol = currencySymbolOf(locale);
  if (symbol === undefined) return undefined;
  const q = literal(symbol);
  return choice.value === 'currency'
    ? `${q}#,##0.00`
    : `_(${q}* #,##0.00_);_(${q}* \\(#,##0.00\\);_(${q}* "-"??_);_(@_)`;
}

/**
 * The entries of the number format list that have a code in a number locale: all of them,
 * unless the locale's currency symbol is not valid, which leaves out currency and accounting.
 * The list never writes another symbol in place of the locale's.
 *
 * @public
 */
export function numberFormatChoices(locale: NumberLocale = EN_US): readonly NumberFormatChoice[] {
  return NUMBER_FORMAT_CHOICES.filter((c) => numberFormatChoiceCode(c.value, locale) !== undefined);
}

/**
 * The number format to write for a built-in id and its code in the locale: the id, unless the
 * workbook gives the id another code (`defined`, the `x:numFmt` codes by id), because the cell
 * would then show that other code. Then the code, which the store writes as a custom code.
 */
export function builtinOrCode(
  id: number,
  code: string,
  defined: ReadonlyMap<number, string>
): { readonly id: number } | { readonly code: string } {
  const own = defined.get(id);
  return own === undefined || own === code ? { id } : { code };
}

/**
 * The prefix of a number format value that is no choice: `custom:` and the format code, so a
 * code that reads like a choice (`percent`) never shows as that choice.
 *
 * @public
 */
export const CUSTOM_NUMBER_FORMAT_PREFIX = 'custom:' as const;

/**
 * The entry of the number format list that a format shows, or `custom:` and its code when it
 * shows no entry. An entry with an id matches the cell's code in the locale (`builtins` of the
 * locale, or the ECMA-376 table), with that id or with a custom id. A workbook that gives a
 * built-in id another code shows that code, so it matches no entry with the id. Currency,
 * accounting, and the long date match their code. A format whose `hasCode` is false matches
 * no entry: it shows `General`, but its id names a code that the workbook or the locale does
 * not give.
 *
 * @public
 */
export function numberFormatValue(
  format: { readonly id: number; readonly code: string; readonly hasCode?: boolean },
  locale: NumberLocale = EN_US
): NumberFormatChoiceValue | `${typeof CUSTOM_NUMBER_FORMAT_PREFIX}${string}` {
  if (format.hasCode === false) return `${CUSTOM_NUMBER_FORMAT_PREFIX}${format.code}`;
  for (const c of NUMBER_FORMAT_CHOICES) {
    if (c.id === null) {
      if (format.code === numberFormatChoiceCode(c.value, locale)) return c.value;
    } else if (
      (format.id === c.id || format.id >= FIRST_CUSTOM_NUMFMT_ID) &&
      format.code === builtinFormatCode(c.id, locale)
    )
      return c.value;
  }
  return `${CUSTOM_NUMBER_FORMAT_PREFIX}${format.code}`;
}

/** The most decimal places that the increase decimals command writes in one section. @internal */
export const MAX_STEP_DECIMALS = 30;

/**
 * Why the decimal step cannot change a format code:
 *
 * - `noNumber`: a decrease on `General` with no number (`value` is null). For the rule of a
 *   selection, see `DecimalSelectionProblem`.
 * - `generalSection`: a `General` section with a color, a condition, a literal, or other
 *   sections, as in `[Red]General` or `General;-0`.
 * - `date`: a date or time section.
 * - `text`: the code shows text in place of a number: the value as text (`@`), or only
 *   literals.
 * - `fraction`: a fraction section, as in `# ?/?`.
 * - `noDecimals`: decrease on a code with no decimal place to remove.
 * - `maxDecimals`: increase on a section with `MAX_STEP_DECIMALS` places.
 * - `tooLong`: the new code is longer than 255 characters.
 * - `unreadable`: the number format parser cannot read the code, the step would make a comma
 *   scale the number or add or remove a group separator (`0.00,0` does not step down to
 *   `0.00,`), or the code holds a control character that a saved code cannot hold.
 * - `noCode`: the cell format names a built-in id that has no code in the number locale.
 * - `missingCode`: the cell format names a custom id that the workbook gives no code.
 *
 * @internal
 */
export type DecimalStepProblem =
  | 'noNumber'
  | 'generalSection'
  | 'date'
  | 'text'
  | 'fraction'
  | 'noDecimals'
  | 'maxDecimals'
  | 'tooLong'
  | 'unreadable'
  | 'noCode'
  | 'missingCode';

/** The result of a decimal step: the new code, or why the step cannot change the code. @internal */
export type DecimalStepResult =
  | { readonly ok: true; readonly code: string }
  | { readonly ok: false; readonly problem: DecimalStepProblem };

/** One character of a code that no quote, bracket, escape, pad, or fill holds. */
interface PlainChar {
  readonly at: number;
  readonly ch: string;
}

/** The plain characters of one section, its first quoted literals, and its fill. */
interface SectionChars {
  readonly plain: PlainChar[];
  /** The index after the closing quote of the first `"..."` literal. */
  quoteEnd: number | undefined;
  /** The index after the last dash literal before the first `?`: `"-"`, `\-`, or `-`. */
  dashEnd: number | undefined;
  /** The section has a fill (`*x`). */
  fill: boolean;
  /** The index of the `*` of the fill, if any. */
  fillAt: number | undefined;
}

/**
 * The plain characters of each section of a code, by the lexical rules of `parseFormatCode`:
 * `"..."` and `[...]` are skipped whole, and `\x`, `_x`, and `*x` are skipped as two
 * characters. Linear in the code length. The caller parses the code first, so every quote and
 * bracket is closed. `dashEnd` is the last dash before the first `?` of the section, so a dash
 * currency symbol before the fill, as in `_("-"* "-"??_)`, is not taken for the zero literal.
 */
function sectionChars(code: string): SectionChars[] {
  const empty = (): SectionChars => ({
    plain: [],
    quoteEnd: undefined,
    dashEnd: undefined,
    fill: false,
    fillAt: undefined,
  });
  const sections: SectionChars[] = [empty()];
  let padded = false;
  let i = 0;
  while (i < code.length) {
    const ch = code[i]!;
    const section = sections[sections.length - 1]!;
    if (ch === ';') {
      sections.push(empty());
      padded = false;
    } else if (ch === '"') {
      const start = i;
      i = code.indexOf('"', i + 1);
      section.quoteEnd ??= i + 1;
      if (!padded && code.slice(start, i + 1) === '"-"') section.dashEnd = i + 1;
    } else if (ch === '[') i = code.indexOf(']', i + 1);
    else if (ch === '\\' || ch === '_' || ch === '*') {
      if (!padded && ch === '\\' && code[i + 1] === '-') section.dashEnd = i + 2;
      if (ch === '*') {
        section.fill = true;
        section.fillAt = i;
      }
      i += 1;
    } else {
      if (!padded && ch === '-') section.dashEnd = i + 1;
      padded ||= ch === '?';
      section.plain.push({ at: i, ch });
    }
    i += 1;
  }
  return sections;
}

const isDigit = (ch: string): boolean => ch === '0' || ch === '#' || ch === '?';

/** The digit placeholders and the decimal point of the mantissa of one number section. */
interface Places {
  readonly int: readonly PlainChar[];
  readonly frac: readonly PlainChar[];
  readonly point: PlainChar | undefined;
}

/**
 * A section of a quoted literal followed by one or more `?` with no point, as `"-"??`: padding.
 * Another literal can follow the run, as in `"-"??\ "€"`. A run of `?` before any quoted
 * literal, as `??" kg"`, shows a number with a unit, and is no padding. In the zero section
 * (`zero`), a dash as `\-` or `-` before the run is such a literal too, as in `\-??`: there a
 * zero shows no number, and in another section `-??` shows a negative number.
 */
function isPadding(p: Places, c: SectionChars, zero = false): boolean {
  const ends = [c.quoteEnd, zero ? c.dashEnd : undefined].filter((e) => e !== undefined);
  const first = p.int[0];
  return (
    !p.point &&
    p.frac.length === 0 &&
    first !== undefined &&
    ends.some((e) => e <= first.at) &&
    p.int.every((q) => q.ch === '?')
  );
}

/** The places of one section, as `analyzeNumber` of the parser reads them. */
function placesOf(code: string, plain: readonly PlainChar[]): Places {
  let end = plain.length;
  for (let k = 0; k < plain.length; k++) {
    const p = plain[k]!;
    const next = code[p.at + 1];
    if ((p.ch === 'e' || p.ch === 'E') && (next === '+' || next === '-')) {
      end = k;
      break;
    }
  }
  const int: PlainChar[] = [];
  const frac: PlainChar[] = [];
  let point: PlainChar | undefined;
  for (let k = 0; k < end; k++) {
    const p = plain[k]!;
    if (p.ch === '.' && !point) point = p;
    else if (isDigit(p.ch)) (point ? frac : int).push(p);
  }
  return { int, frac, point };
}

/** One change of the code: remove `remove` characters at `at`, then insert `insert`. */
interface Edit {
  readonly at: number;
  readonly remove: number;
  readonly insert: string;
}

/**
 * Where the zero section of a code with three or more sections keeps the room for its decimal
 * places, as `"-"??` of the accounting code does: `?` padding (`isPadding`) in one run. The `?`
 * only pad the literal, so the step adds or removes one `?` and never adds a decimal point.
 * With no `?`, as after a decrease to no places, the room is empty, after the first dash
 * (`"-"`, `\-`, or `-`) of a section with a fill, as in the accounting code. Null for a zero
 * section of literals and more than one run of `?`, as `"a"?"b"?`, which stays as it is, so the
 * runs never merge. Undefined for any other section: a zero section with another literal and
 * no `?`, as `"zero"* `, or a dash with no fill, as a bare `"-"`, stays as it is.
 */
function roomOf(
  index: number,
  sections: number,
  places: Places,
  chars: SectionChars
): { readonly marks: readonly PlainChar[]; readonly at: number } | null | undefined {
  if (index !== 2 || sections < 3 || places.point || places.frac.length > 0) return undefined;
  const { int } = places;
  if (int.length === 0) {
    const end = chars.dashEnd;
    if (!chars.fill || end === undefined) return undefined;
    // A fill right after the literal, as in `"-"* ??`, stays between the literal and its `?`.
    return { marks: [], at: chars.fillAt === end ? end + 2 : end };
  }
  if (!isPadding(places, chars, true)) return undefined;
  if (int.some((p, k) => k > 0 && p.at !== int[k - 1]!.at + 1)) return null;
  return { marks: int, at: int[int.length - 1]!.at + 1 };
}

/**
 * The edits of one number section, or undefined when the section stays as it is. A decrease of
 * the last place of a section with no integer digit, as `.0`, writes `0` in place of the point,
 * so the section keeps a digit placeholder.
 */
function sectionEdit(places: Places, step: 1 | -1): Edit[] | undefined {
  const { int, frac, point } = places;
  if (step > 0) {
    if (point) {
      const last = frac[frac.length - 1];
      return [{ at: (last ?? point).at + 1, remove: 0, insert: last?.ch ?? '0' }];
    }
    const last = int[int.length - 1];
    return last ? [{ at: last.at + 1, remove: 0, insert: '.0' }] : undefined;
  }
  const last = frac[frac.length - 1];
  if (!last || !point) return undefined;
  if (frac.length === 1 && int.length === 0)
    return [
      { at: last.at, remove: 1, insert: '' },
      { at: point.at, remove: 1, insert: '0' },
    ];
  const edits: Edit[] = [{ at: last.at, remove: 1, insert: '' }];
  if (frac.length === 1) edits.push({ at: point.at, remove: 1, insert: '' });
  return edits;
}

const isGeneral = (s: FormatSection): boolean =>
  s.kind === 'general' &&
  s.color === undefined &&
  s.cond === undefined &&
  s.tokens.length === 1 &&
  s.tokens[0]!.kind === 'general';

/** Whether a code is the plain `General` code, whose places come from the number it shows. */
export function isGeneralCode(code: string): boolean {
  const parsed = parseFormatCode(code);
  return parsed.ok && parsed.format.sections.length === 1 && isGeneral(parsed.format.sections[0]!);
}

/** The text that `General` shows for a number, in the default number locale. */
const generalText = (value: number): string =>
  formatValue(value, 'General', { locale: EN_US, date1904: false }).text;

/** Whether `General` shows a number in exponent form, at full width. */
export const generalExponent = (value: number): boolean => generalText(value).includes('E');

/** The decimal places that `General` shows for a number, before any exponent. */
export function generalDecimals(value: number): number {
  const text = generalText(value);
  const point = text.indexOf('.');
  if (point < 0) return 0;
  let n = 0;
  while (point + 1 + n < text.length && text[point + 1 + n]! >= '0' && text[point + 1 + n]! <= '9')
    n += 1;
  return n;
}

/** A code with `places` decimal places and no group separator: `0`, `0.0`, `0.00`, ... */
const fixedCode = (places: number): string => (places === 0 ? '0' : `0.${'0'.repeat(places)}`);

/** A scientific code with `places` decimal places in the mantissa: `0E+00`, `0.0E+00`, ... */
const scientificCode = (places: number): string => `${fixedCode(places)}E+00`;

/**
 * The decimal step of `General`, with the number that the cell shows, if any. The places are
 * those that `General` shows before any exponent. When `General` shows the number in exponent
 * form, by the rule of the `General` formatter, the step keeps that form: `1e20` shows `1E+20`,
 * so an increase gives `0.0E+00`, and `1.5e-10` shows `1.5E-10`, so a decrease gives `0E+00`.
 */
function stepGeneral(step: 1 | -1, value: number | null): DecimalStepResult {
  if (value === null)
    return step > 0 ? { ok: true, code: '0.0' } : { ok: false, problem: 'noNumber' };
  return stepGeneralText(step, generalText(value), '.') ?? { ok: false, problem: 'unreadable' };
}

/**
 * The decimal step of `General` from the text that it shows, with the decimal separator of
 * that text: the places of the text before any exponent, one more or one less, in exponent form
 * when the text is. The grid fits `General` to its column, so the shown text can have fewer
 * places than the number (`3.141593` for 3.14159265358 in a narrow column). Undefined for a
 * text that shows no number, such as `#` that fill a column that is too narrow.
 */
export function stepGeneralText(
  step: 1 | -1,
  text: string,
  decimal: string
): DecimalStepResult | undefined {
  const match = /^-?\d+(?:(\D)(\d+))?(E[+-]\d+)?$/.exec(text);
  if (!match || (match[1] !== undefined && match[1] !== decimal)) return undefined;
  const shown = match[2]?.length ?? 0;
  const code = match[3] ? scientificCode : fixedCode;
  if (step > 0)
    return shown >= MAX_STEP_DECIMALS
      ? { ok: false, problem: 'maxDecimals' }
      : { ok: true, code: code(shown + 1) };
  return shown === 0 ? { ok: false, problem: 'noDecimals' } : { ok: true, code: code(shown - 1) };
}

/**
 * Add (`step` 1) or remove (`step` -1) one decimal place in each number section of a format
 * code. The step keeps every other character of the code: the sections, the colors, the
 * conditions, the literals, the text section, and the exponent part. A number section with no
 * digit placeholder, as `.` in `0.00;.`, stays as it is. Increase repeats the last decimal
 * placeholder (`0`, `#`, or `?`), or adds `0`. Decrease removes the decimal point with the last
 * decimal place, so `0.` steps up to `0.0` and back to `0`. A decrease of a section with one
 * place and no integer digit, as `.0`, gives `0`, so the section keeps a digit.
 *
 * `General` takes the places that it shows for `value`, the number of the cell, one
 * more or one less (`0.00` for 1.5 on an increase). A number that `General` shows in exponent
 * form keeps it (`0.0E+00` for 1e20 on an increase). With no number, an increase gives `0.0`.
 *
 * The rules of `?` padding beside a literal and of places that a literal splits are in the
 * TSDoc of `decimalStepOf`.
 */
export function stepDecimals(
  code: string,
  step: 1 | -1,
  value: number | null = null
): DecimalStepResult {
  // A code with a control character, as a tab from `&#9;` in the file, is one that the store
  // refuses to write (`patchProblem`): the step keeps it, so it refuses the code.
  if (!isPlainXmlText(code)) return { ok: false, problem: 'unreadable' };
  const parsed = parseFormatCode(code);
  if (!parsed.ok) return { ok: false, problem: 'unreadable' };
  const sections = parsed.format.sections;
  if (sections.length === 1 && isGeneral(sections[0]!)) return stepGeneral(step, value);
  if (sections.some((s) => s.kind === 'general')) return { ok: false, problem: 'generalSection' };
  if (sections.some((s) => s.kind === 'date')) return { ok: false, problem: 'date' };
  if (sections.some((s) => s.kind === 'number' && s.fraction))
    return { ok: false, problem: 'fraction' };
  const chars = sectionChars(code);
  if (chars.length !== sections.length) return { ok: false, problem: 'unreadable' };
  // A section with a digit placeholder outside `?` padding (`isPadding`), other than the padding
  // of the zero section (`roomOf`): such padding then keeps its run.
  const zeroAt = sections.length >= 3 ? 2 : -1;
  const placed = sections.some((s, k) => {
    if (s.kind !== 'number') return false;
    const p = placesOf(code, chars[k]!.plain);
    if (k === zeroAt && roomOf(k, sections.length, p, chars[k]!) !== undefined) return false;
    return p.int.length + p.frac.length > 0 && !isPadding(p, chars[k]!);
  });
  const edits: Edit[] = [];
  const changed = new Map<number, number>();
  let decimals = 0;
  let digits = false;
  // The zero section's room for its places (`roomOf`), and whether the step changes it.
  let roomed = false;
  let roomEdit = false;
  // A decrease that empties a room that no increase finds again (`roomOf` needs a fill and `"-"`).
  let roomLost = false;
  for (let k = 0; k < sections.length; k++) {
    const section = sections[k]!;
    if (section.kind !== 'number') continue;
    const places = placesOf(code, chars[k]!.plain);
    // The scan reads the places that the parser reads, or the step refuses the code.
    if (places.frac.length !== section.frac.length) return { ok: false, problem: 'unreadable' };
    const room = roomOf(k, sections.length, places, chars[k]!);
    if (room === null) continue;
    if (room) {
      roomed = true;
      const last = room.marks[room.marks.length - 1];
      if (step < 0 && last) edits.push({ at: last.at, remove: 1, insert: '' });
      roomLost ||=
        step < 0 && room.marks.length === 1 && !(chars[k]!.fill && chars[k]!.dashEnd !== undefined);
      // An empty room gets its `?` back only after a decrease to no places: when the number
      // sections before it had no places either. Else it stays empty, so the steps return.
      const grow = step > 0 && (room.marks.length > 0 || decimals === 0);
      if (grow) edits.push({ at: room.at, remove: 0, insert: '?' });
      roomEdit ||= grow || last !== undefined;
      continue;
    }
    // `?` padding after a quoted literal, as `"-"??` in `0.00E+00;"-"??` or `"-"??;0.00`, keeps
    // its run when another section has a digit placeholder outside such padding (`placed`): an
    // increase would add a point that no decrease removes. A section with no digit placeholder,
    // as `.` in `0.00;.`, stays as it is.
    if (placed && isPadding(places, chars[k]!)) continue;
    if (places.int.length + places.frac.length === 0) continue;
    decimals = Math.max(decimals, places.frac.length);
    digits ||= places.int.length + places.frac.length > 0;
    if (step > 0 && places.frac.length >= MAX_STEP_DECIMALS)
      return { ok: false, problem: 'maxDecimals' };
    const edit = sectionEdit(places, step);
    if (!edit) continue;
    edits.push(...edit);
    changed.set(k, places.frac.length + step);
  }
  // A code whose only number is the room of its zero section, as `"pos";"neg";"-"??`, steps that
  // room alone. It keeps its last `?` when no increase could add it back.
  const roomOnly = !digits && roomEdit;
  if (!digits && !roomed) return { ok: false, problem: 'text' };
  if (step < 0 && (roomOnly ? roomLost : decimals === 0))
    return { ok: false, problem: 'noDecimals' };
  if (changed.size === 0 && !roomOnly) return { ok: false, problem: 'unreadable' };
  let next = code;
  for (const e of edits.sort((a, b) => b.at - a.at))
    next = next.slice(0, e.at) + e.insert + next.slice(e.at + e.remove);
  if (next.length > MAX_FORMAT_CODE_LENGTH) return { ok: false, problem: 'tooLong' };
  // The new code reads as the old one with one place more or less in each changed section.
  const reread = parseFormatCode(next);
  if (!reread.ok || reread.format.sections.length !== sections.length)
    return { ok: false, problem: 'unreadable' };
  // A comma that would scale the number, as `,` after the removed place of `0.00,0`, or a
  // thousands separator that would come or go, changes the number that the section shows.
  for (const [k, places] of changed) {
    const [was, now] = [sections[k]!, reread.format.sections[k]!];
    if (now.frac.length !== places || now.shift !== was.shift || now.thousands !== was.thousands)
      return { ok: false, problem: 'unreadable' };
  }
  return { ok: true, code: next };
}

/**
 * The decimal step of a cell's number format: the step of its code, with the number of
 * the cell for `General`, which steps from the full text of `General`. The grid can show fewer
 * places in a narrow column, and the command steps from the shown text
 * (`ShownGeneralText`). A format whose `hasCode`
 * is false shows `General`, but its code is not known: the step refuses it, as `noCode` for a
 * built-in id that has no code in the locale (ids 5 to 8, 23 to 36, 41 to 44, and 50 to 163 of
 * ECMA-376 Part 1, 18.8.30, unless the locale or the workbook gives one), and as `missingCode`
 * for a custom id that the workbook gives no code. So the step never writes a code that the
 * file did not mean.
 *
 * A step that would make a comma scale the number, or would add or remove a group separator,
 * gives `unreadable`: `0.00,0` does not step down to `0.00,`. So does a code with a control
 * character, which the store does not write.
 *
 * In the zero section of a code with three or more sections, a run of `?` with no point after
 * a quoted literal or a dash, as in `"-"??` or `\-??`, pads the literal to the places of the
 * number. It gets or loses one `?`, and gets no decimal point. When a section with a fill has a
 * dash and no `?`, an increase adds one `?` after the first dash only when the number sections
 * before it have no places, as after a decrease to no places. Else the section stays as it is,
 * so `"-"` with no `?` beside `#,##0.00` steps up and down with no change. Another literal can
 * follow, as in `"-"\ "€"`. The `?` go after a fill that follows the literal, so the accounting
 * code keeps its zero literal in place. Beside another literal, or with no fill, a decrease to
 * no places removes the last `?`, and an increase does not add it back: `0.00;-0.00;"x"??`
 * goes down to `0;-0;"x"` and up again to `0.00;-0.00;"x"`. That is the known approximation
 * `decimal-padding`. Two runs of `?` in the zero section, as `"a"?"b"?`, stay as they are.
 *
 * The step adds or removes the last decimal placeholder of a section. When a literal or a
 * second point splits the places, as in `0.00" "0` or `0.00%0`, a decrease and an increase can
 * move a place to the other side of the literal: `0.00" "0` goes down to `0.00" "` and up again
 * to `0.000" "`. That is the known approximation `decimal-split-places`.
 *
 * A section of a quoted literal followed by `?` only, with no point, as `"-"??`, is padding too.
 * Such padding stays as it is when another section has a digit placeholder outside such
 * padding, as in `0.00E+00;"-"??` or `"-"??;0.00`: an increase would add a point that no
 * decrease removes. The padding of the zero section does not count, but a zero section with
 * other digit placeholders does.
 *
 * In a code whose only digit placeholders are this padding, the padding of the first two
 * sections takes a decimal point, as a number does. The padding of the zero section takes one
 * `?`: `"$"??;"$"??;"-"??` steps up to `"$"??.0;"$"??.0;"-"???`. When only the zero section has
 * padding, as `"pos";"neg";"-"??`, the step changes only that padding. A decrease keeps the
 * last `?` when no increase could add it back: in a section with no fill or no `"-"`.
 *
 * A run of `?` before any quoted literal, as `??" kg"`, shows a number with a unit. It steps as
 * a number section, also in the zero section.
 *
 * @internal
 */
export function decimalStepOf(
  format: { readonly id: number; readonly code: string; readonly hasCode?: boolean },
  step: 1 | -1,
  value: number | null = null
): DecimalStepResult {
  if (format.hasCode === false)
    return {
      ok: false,
      problem: format.id < FIRST_CUSTOM_NUMFMT_ID ? 'noCode' : 'missingCode',
    };
  return stepDecimals(format.code, step, value);
}

/**
 * A code with every decimal place removed, by decrease steps (at most `MAX_STEP_DECIMALS`), and
 * the number of steps. Undefined when the steps do not lead back to `code` by increases: then a
 * decimal placeholder is `#` or `?`, or a place is not one that a decimal step writes.
 */
function withoutPlaces(
  code: string
): { readonly base: string; readonly places: number } | undefined {
  let base = code;
  let places = 0;
  for (; places < MAX_STEP_DECIMALS; places++) {
    const r = stepDecimals(base, -1);
    if (!r.ok) break;
    base = r.code;
  }
  let back = base;
  for (let k = 0; k < places; k++) {
    const r = stepDecimals(back, 1);
    if (!r.ok) return undefined;
    back = r.code;
  }
  return back === code ? { base, places } : undefined;
}

/**
 * The list entry of a code that differs from it only in its decimal places, and those places.
 *
 * @internal
 */
export interface NumberFormatFamily {
  /** The list entry. */
  readonly value: NumberFormatChoiceValue;
  /** The decimal places of each section that has decimal places, which a step keeps equal. */
  readonly places: number;
}

/**
 * The decimal places of a code that a decimal step can change: the places of each section that
 * has them, which a step keeps equal. Undefined for a code with no step (`General`, a date,
 * text, or a fraction), or with places that a step does not write, such as `0.#%`.
 *
 * @internal
 */
export function numberFormatPlaces(code: string): number | undefined {
  // `General` shows as many places as the number needs: it has no count of its own.
  if (isGeneralCode(code)) return undefined;
  const own = withoutPlaces(code);
  // A code with no step at all (a date, text, or `General`) has no places to name.
  if (!own || (own.places === 0 && !stepDecimals(code, 1).ok)) return undefined;
  return own.places;
}

/**
 * The entry of the number format list whose code differs from `code` only in the number of
 * decimal places of its sections, as after a decimal step, and those places: `0.0%` and `0%`
 * are Percentage. `numberFormatListView` names the code with them, such as "Currency, 3 decimal
 * places". Only places that a decimal step writes count: `0.#%` and `0.??` belong to no entry.
 * The check reads each code with at most `MAX_STEP_DECIMALS` (30) decrease steps and as many
 * increase steps, so a code with more places belongs to no entry. The exact code of an entry gives undefined: use `numberFormatValue` for that case.
 *
 * @internal
 */
export function numberFormatFamilyOf(
  code: string,
  locale: NumberLocale = EN_US
): NumberFormatFamily | undefined {
  const own = withoutPlaces(code);
  if (!own || numberFormatPlaces(code) === undefined) return undefined;
  for (const choice of numberFormatChoices(locale)) {
    const entry = numberFormatChoiceCode(choice.value, locale);
    if (entry === undefined || entry === code) continue;
    if (withoutPlaces(entry)?.base === own.base) return { value: choice.value, places: own.places };
  }
  return undefined;
}
