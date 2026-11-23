// The typed number forms of the cell editor, DOM-free: text such as `12%`, `$5.50`,
// `1,234.5`, `1 1/2`, `2026-11-12`, or `13:45` read as a number and the number format of its
// form (ECMA-376 Part 1, 18.8.30), and the inverse: the edit text of a stored number with a
// date, time, or percent format, in a form that reads back to the same value.
//
// Each function uses no regular expression and takes time linear in its input: a typed text
// of at most `MAX_ENTRY_LENGTH` characters, or a format code. No count from the input bounds
// a loop.

import {
  builtinFormatCode,
  dateToSerial,
  isDateFormat,
  parseDouble,
  serialToDate,
  subsecondDigits,
  type NumberLocale,
} from '../cascade/index.ts';
import type { FormatPatch } from '../store/index.ts';

/** The longest text that reads as a typed number form. Longer text stays a string. */
export const MAX_ENTRY_LENGTH = 64;

/**
 * How typed text reads: the workbook's date system and the year of a date typed with no year.
 *
 * @public
 */
export interface CellEntryOptions {
  /**
   * True when the workbook uses the 1904 date system (`x:workbookPr/@date1904`, Part 1,
   * 18.17.4.1). False when absent.
   */
  readonly date1904?: boolean;
  /**
   * The year of a date typed with no year, such as `3/15`. The year of the host clock when
   * absent. A value that is not an integer, or a year that the date system cannot hold, leaves
   * such a date as text.
   */
  readonly year?: number;
}

/** The currency symbols that a typed amount can carry, before or after the number. */
const CURRENCY_SYMBOLS = ['$', '€', '£', '¥'] as const;

/** A number format of a patch: a built-in id or a custom code. */
export type NumberFormatRef = NonNullable<FormatPatch['numberFormat']>;

/** A typed number form: the value, and the format that the form implies. */
export interface NumberForm {
  readonly value: number;
  readonly format?: NumberFormatRef;
}

/** Built-in format ids of ECMA-376 Part 1, 18.8.30. */
const FMT = {
  grouped: 3,
  grouped2: 4,
  percent: 9,
  percent2: 10,
  fraction1: 12,
  fraction2: 13,
  date: 14,
  dayMonth: 16,
  time12: 18,
  time12s: 19,
  time: 20,
  times: 21,
  dateTime: 22,
  elapsedSeconds: 46,
} as const;

/**
 * A typed number form, or undefined for text that is no number form or is longer than
 * `MAX_ENTRY_LENGTH`. A plain decimal is not a form here: the caller reads it first.
 */
export function readNumberForm(
  text: string,
  locale: NumberLocale,
  options: CellEntryOptions = {}
): NumberForm | undefined {
  if (text.length > MAX_ENTRY_LENGTH) return undefined;
  const form = readFormOf(text, locale, options);
  return form && Number.isFinite(form.value) ? form : undefined;
}

function readFormOf(
  text: string,
  locale: NumberLocale,
  options: CellEntryOptions
): NumberForm | undefined {
  return (
    readAmount(text, locale) ??
    readFraction(text) ??
    readDateTime(text, locale, options.date1904 === true, options.year)
  );
}

const isDigitCode = (c: number): boolean => c >= 48 && c <= 57;

/** Whether `text` is one or more ASCII digits. */
function allDigits(text: string): boolean {
  if (text.length === 0) return false;
  for (let i = 0; i < text.length; i += 1) if (!isDigitCode(text.charCodeAt(i))) return false;
  return true;
}

/** The spaces that can stand between a number and a currency symbol. */
const isSpace = (ch: string | undefined): boolean =>
  ch === ' ' || ch === '\u00a0' || ch === '\u202f';

/**
 * A number with a sign or parentheses, a currency symbol, a percent sign, or group separators.
 * A form that has none of these is a plain number, which `parseDouble` reads.
 */
function readAmount(text: string, locale: NumberLocale): NumberForm | undefined {
  let s = text;
  let negative = false;
  if (s.startsWith('(') && s.endsWith(')') && s.length > 2) {
    negative = true;
    s = s.slice(1, -1);
  } else if (s.startsWith('-')) {
    negative = true;
    s = s.slice(1);
  }
  let currency: string | undefined;
  let after = false;
  for (const sym of CURRENCY_SYMBOLS) {
    if (s.startsWith(sym)) {
      currency = sym;
      s = s.slice(sym.length);
      break;
    }
    if (s.endsWith(sym)) {
      currency = sym;
      after = true;
      s = s.slice(0, -sym.length);
      if (isSpace(s[s.length - 1])) s = s.slice(0, -1);
      break;
    }
  }
  // A sign after a leading currency symbol: `$-5`.
  if (currency !== undefined && !after && !negative && s.startsWith('-')) {
    negative = true;
    s = s.slice(1);
  }
  const percent = currency === undefined && s.endsWith('%');
  if (percent) s = s.slice(0, -1);
  const core = readGrouped(s, locale);
  if (!core) return undefined;
  const marked = negative || currency !== undefined || percent || core.grouped;
  if (!marked) return undefined;
  const sign = negative ? -1 : 1;
  if (percent)
    return {
      // The decimal point moves in the text, so `1.1%` is the double nearest to 0.011: a
      // division of the parsed double by 100 can give 0.011000000000000001.
      value: Number(`${negative ? '-' : ''}${core.text}e-2`),
      format: { id: core.decimals ? FMT.percent2 : FMT.percent },
    };
  if (currency !== undefined) {
    const digits = core.decimals ? '#,##0.00' : '#,##0';
    const code = after ? `${digits} "${currency}"` : `"${currency}"${digits}`;
    return { value: sign * core.value, format: { code } };
  }
  if (core.grouped)
    return {
      value: sign * core.value,
      format: { id: core.decimals ? FMT.grouped2 : FMT.grouped },
    };
  return { value: sign * core.value };
}

/**
 * Digits with optional group separators and an optional decimal part, in the locale's
 * separators: `1,234.5`, `1234`, `.5`. A group after the first holds exactly three digits.
 */
function readGrouped(
  text: string,
  locale: NumberLocale
): { value: number; text: string; grouped: boolean; decimals: boolean } | undefined {
  const { decimal, group } = locale;
  const at = text.indexOf(decimal);
  const whole = at < 0 ? text : text.slice(0, at);
  const frac = at < 0 ? '' : text.slice(at + decimal.length);
  if (at >= 0 && !allDigits(frac)) return undefined;
  let digits = whole;
  let grouped = false;
  if (group !== decimal && group.length > 0 && whole.includes(group)) {
    const parts = whole.split(group);
    const first = parts[0]!;
    // A first group that starts with `0` (`0,123`) is no grouped number: it can be a decimal
    // in another locale.
    if (first.length < 1 || first.length > 3 || !allDigits(first) || first.startsWith('0'))
      return undefined;
    for (let i = 1; i < parts.length; i += 1)
      if (parts[i]!.length !== 3 || !allDigits(parts[i]!)) return undefined;
    digits = parts.join('');
    grouped = true;
  }
  if (digits.length > 0 && !allDigits(digits)) return undefined;
  if (digits.length === 0 && frac.length === 0) return undefined;
  const plain = `${digits || '0'}${frac ? `.${frac}` : ''}`;
  const value = parseDouble(plain);
  return value === undefined
    ? undefined
    : { value, text: plain, grouped, decimals: frac.length > 0 };
}

/**
 * A fraction with a whole part, as format 12 or 13 shows it: `1 1/2`, `-0 3/16`. A fraction
 * with no whole part (`1/2`) reads as a date, as the locale orders it.
 */
function readFraction(text: string): NumberForm | undefined {
  const negative = text.startsWith('-');
  const s = negative ? text.slice(1) : text;
  const space = s.indexOf(' ');
  const slash = s.indexOf('/');
  if (space < 1 || slash < space) return undefined;
  const whole = s.slice(0, space);
  // Format 13 pads the numerator with spaces (`1  1/16`), so more than one space can come here.
  let start = space + 1;
  while (s[start] === ' ') start += 1;
  const num = s.slice(start, slash);
  const den = s.slice(slash + 1);
  if (!allDigits(whole) || !allDigits(num) || !allDigits(den) || den.length > 2) return undefined;
  const d = Number(den);
  // A proper fraction only, as format 12 or 13 writes it: `1 5/2` is not read.
  if (Number(num) >= d) return undefined;
  const value = Number(whole) + Number(num) / d;
  if (!Number.isFinite(value)) return undefined;
  return {
    value: negative ? -value : value,
    format: { id: den.length === 1 ? FMT.fraction1 : FMT.fraction2 },
  };
}

/**
 * A time as a fraction of a day, and whether it has seconds, a 12-hour clock, or more than 23
 * hours (an elapsed time, such as `36:00`).
 */
interface TypedTime {
  readonly fraction: number;
  readonly seconds: boolean;
  readonly twelve: boolean;
  readonly elapsed: boolean;
}

/**
 * `H:MM`, `H:MM:SS`, each with an optional `AM` or `PM`, with or without a space before it.
 * Without `AM` or `PM`, the hours can be 24 to 9999: an elapsed time.
 */
function readTime(text: string): TypedTime | undefined {
  let s = text;
  let twelve = false;
  let pm = false;
  const tail = s.slice(-2).toUpperCase();
  if (tail === 'AM' || tail === 'PM') {
    twelve = true;
    pm = tail === 'PM';
    s = s.slice(0, -2);
    if (s.endsWith(' ')) s = s.slice(0, -1);
  }
  const parts = s.split(':');
  if (parts.length < 2 || parts.length > 3) return undefined;
  const [h, m, sec] = parts;
  if (!allDigits(h!) || h!.length > 4 || !allDigits(m!) || m!.length !== 2) return undefined;
  if (sec !== undefined && (!allDigits(sec) || sec.length !== 2)) return undefined;
  let hours = Number(h);
  const minutes = Number(m);
  const seconds = sec === undefined ? 0 : Number(sec);
  if (minutes > 59 || seconds > 59) return undefined;
  if (twelve) {
    if (hours < 1 || hours > 12) return undefined;
    hours = (hours % 12) + (pm ? 12 : 0);
  }
  return {
    fraction: (hours * 3600 + minutes * 60 + seconds) / 86400,
    seconds: sec !== undefined,
    twelve,
    elapsed: hours > 23,
  };
}

/** A calendar date as its parts. */
interface TypedDate {
  readonly year: number;
  readonly month: number;
  readonly day: number;
  /** Whether the text gave the year. */
  readonly withYear: boolean;
}

/**
 * How the locale's short date (built-in format 14) orders and separates its parts, or
 * undefined when that code is not a plain date (`plainSeparator`) that `readDate` reads back.
 */
function dateShape(locale: NumberLocale): CodeShape | undefined {
  const code = builtinFormatCode(FMT.date, locale);
  if (code === undefined || !plainSeparator(withoutTags(code))) return undefined;
  const shape = codeShape(code);
  return shape.order.length === 3 ? shape : undefined;
}

/**
 * A date: ISO 8601 `YYYY-MM-DD` (also with `/`), or two or three numbers in the locale's order,
 * separated by `/` or `-`, or by `.` when the locale's short date uses `.` (and only with a
 * year). A two-digit year from 00 to 29 is in the 2000s, from 30 to 99 in the 1900s.
 */
function readDate(text: string, locale: NumberLocale, year: number): TypedDate | undefined {
  let sep: string | undefined;
  for (const ch of ['/', '-', '.'])
    if (text.includes(ch)) {
      if (sep !== undefined) return undefined;
      sep = ch;
    }
  if (sep === undefined) return undefined;
  const parts = text.split(sep);
  if (parts.length < 2 || parts.length > 3) return undefined;
  if (!parts.every((p) => allDigits(p) && p.length <= 4)) return undefined;
  if (parts.length === 3 && parts[0]!.length === 4 && sep !== '.') {
    if (parts[1]!.length > 2 || parts[2]!.length > 2) return undefined;
    return {
      year: Number(parts[0]),
      month: Number(parts[1]),
      day: Number(parts[2]),
      withYear: true,
    };
  }
  const shape = dateShape(locale);
  if (shape === undefined || (sep === '.' && !shape.dot)) return undefined;
  const { order } = shape;
  if (parts.length === 2) {
    if (sep === '.') return undefined;
    const dm = order.replace('y', '');
    if (parts.some((p) => p.length > 2)) return undefined;
    const [a, b] = [Number(parts[0]), Number(parts[1])];
    return dm === 'dm'
      ? { year, month: b, day: a, withYear: false }
      : { year, month: a, day: b, withYear: false };
  }
  const by: Record<string, string> = {};
  for (let i = 0; i < 3; i += 1) by[order[i]!] = parts[i]!;
  const y = by.y!;
  if (by.d!.length > 2 || by.m!.length > 2 || (y.length !== 2 && y.length !== 4)) return undefined;
  const yy = Number(y);
  return {
    year: y.length === 4 ? yy : yy < 30 ? 2000 + yy : 1900 + yy,
    month: Number(by.m),
    day: Number(by.d),
    withYear: true,
  };
}

/** A date, a time, or a date and a time separated by one space. */
function readDateTime(
  text: string,
  locale: NumberLocale,
  date1904: boolean,
  year: number | undefined
): NumberForm | undefined {
  const time = readTime(text);
  // An elapsed time with no seconds keeps the form it is typed in: `[h]:mm`, which no built-in
  // format of Part 1, 18.8.30 is.
  if (time?.elapsed && !time.seconds) return { value: time.fraction, format: { code: '[h]:mm' } };
  if (time)
    return {
      value: time.fraction,
      format: {
        id: time.elapsed
          ? FMT.elapsedSeconds
          : time.twelve
            ? time.seconds
              ? FMT.time12s
              : FMT.time12
            : time.seconds
              ? FMT.times
              : FMT.time,
      },
    };
  const space = text.indexOf(' ');
  const datePart = space < 0 ? text : text.slice(0, space);
  const timePart = space < 0 ? undefined : readTime(text.slice(space + 1));
  if (space >= 0 && (!timePart || timePart.elapsed)) return undefined;
  const date = readDate(datePart, locale, year ?? new Date().getFullYear());
  if (!date) return undefined;
  const serial = dateToSerial(isoDate(date.year, date.month, date.day), date1904);
  if (serial === undefined) return undefined;
  if (timePart)
    return { value: serial + timePart.fraction, format: dateTimeFormat(locale, timePart) };
  // With no date order in the locale, only ISO 8601 reads back as a date.
  if (!dateShape(locale)) return { value: serial, format: { code: 'yyyy-mm-dd' } };
  return { value: serial, format: { id: date.withYear ? FMT.date : FMT.dayMonth } };
}

/**
 * The format of a date and a time, in a form that the locale reads back. Built-in format 22
 * when it is a plain date (day, month, and year with one separator, in the order of the short
 * date, format 14) and one clock (` h:mm` or ` hh:mm`), for a 24-hour time with no seconds.
 * Typed seconds add `:ss` to that code, and a 12-hour time takes its date and the typed clock,
 * so the cell looks as format 22 shows it. Otherwise the short date code and the typed clock, or
 * `yyyy-mm-dd` and the clock when the locale has no date order.
 */
function dateTimeFormat(locale: NumberLocale, time: TypedTime): NumberFormatRef {
  const clock = `${time.seconds ? 'h:mm:ss' : 'h:mm'}${time.twelve ? ' AM/PM' : ''}`;
  const shape = dateShape(locale);
  // With no date order in the locale, only ISO 8601 reads back as a date.
  if (!shape) return { code: `yyyy-mm-dd ${clock}` };
  const code22 = builtinFormatCode(FMT.dateTime, locale);
  const lower = code22?.toLowerCase();
  const tail = lower && CLOCK_TAILS.find((t) => lower.endsWith(t));
  const date = tail ? code22!.slice(0, -tail.length) : undefined;
  if (date !== undefined && plainDate(date, shape)) {
    if (!time.twelve && !time.seconds) return { id: FMT.dateTime };
    if (time.twelve) return { code: `${date} ${clock}` };
    // The seconds follow the minutes of format 22, also before its AM/PM.
    const at = code22!.length - tail!.length + tail!.indexOf(':mm') + 3;
    return { code: `${code22!.slice(0, at)}:ss${code22!.slice(at)}` };
  }
  return { code: `${builtinFormatCode(FMT.date, locale)!} ${clock}` };
}

/**
 * A code with no leading locale tags (`[$-409]`, `[$-x-sysdate]`), which show nothing. Another
 * tag, such as an elapsed time (`[h]`), stays, so the code is not a plain date.
 */
function withoutTags(code: string): string {
  let at = 0;
  while (code[at] === '[' && code[at + 1] === '$') {
    const end = code.indexOf(']', at);
    if (end < 0) return code;
    at = end + 1;
  }
  return code.slice(at);
}

/** The clocks that can end format 22, in lower case: one clock, also on a 12-hour clock. */
const CLOCK_TAILS = [' hh:mm am/pm', ' h:mm am/pm', ' hh:mm', ' h:mm'];

/** The tokens of a plain date code. */
const DATE_TOKENS = new Set(['d', 'dd', 'm', 'mm', 'yy', 'yyyy']);

/**
 * The separator of a plain date code: three day, month, or year tokens (`dd/mm/yyyy`) with one
 * separator, `/`, `.`, or `-`. Undefined for any other code.
 */
function plainSeparator(code: string): string | undefined {
  let sep: string | undefined;
  for (const ch of code) {
    if (ch !== 'd' && ch !== 'm' && ch !== 'y') {
      sep = ch;
      break;
    }
  }
  if (sep !== '/' && sep !== '.' && sep !== '-') return undefined;
  const parts = code.split(sep);
  return parts.length === 3 && parts.every((part) => DATE_TOKENS.has(part)) ? sep : undefined;
}

/**
 * Whether a date code reads back in the locale: three plain tokens in the order of the short
 * date, with `/` or `-`, or `.` when the short date uses `.` (`readDate`).
 */
function plainDate(code: string, shape: CodeShape): boolean {
  const sep = plainSeparator(withoutTags(code));
  if (!sep || (sep === '.' && !shape.dot)) return false;
  return codeShape(code).order === shape.order;
}

// ------------------------------------------------------------------ edit text

/**
 * The text that the cell editor shows for a stored number, in a form that `readNumberForm`
 * reads back to the same value. A date format gives an ISO 8601 date, a time on a 24-hour
 * clock, or both (`2026-11-12 18:00`); an elapsed-time format (`[h]:mm`) gives the hours
 * (`36:00`); a format with `%` gives a percent (`12%`). The form rounds what the cell does not
 * show: a time to whole seconds, a percent to 15 significant digits. It must read back to the
 * value within that rounding, else the text is the number with the locale decimal separator:
 * serial 60 of the 1900 system, a value past 9999-12-31, or a value with a fraction of a second
 * in a format that shows it (`mm:ss.0`). The editor writes nothing for an unchanged edit text,
 * so the rounding changes no value. Only the first section of the code counts, as for a
 * positive value.
 */
export function numberEditText(
  value: number,
  code: string,
  locale: NumberLocale,
  date1904: boolean
): string {
  const plain = String(value).replace('.', locale.decimal);
  const shape = codeShape(code);
  // A format that shows fractions of a second (`mm:ss.0`) shows more than whole seconds: when
  // the shown fraction is not zero, no form reads back to it, so the text is the plain number.
  // The value rounds to the shown digits as the cell rounds it (`renderDate`).
  const unit = 10 ** subsecondDigits(code);
  if (unit > 1 && Math.round(Number((value * 86400 * unit).toPrecision(15))) % unit !== 0)
    return plain;
  const form = isDateFormat(code)
    ? shape.elapsed
      ? elapsedText(value)
      : dateTimeText(value, date1904, !shape.order.includes('d') && !shape.order.includes('y'))
    : shape.percent
      ? `${String(Number((value * 100).toPrecision(15))).replace('.', locale.decimal)}%`
      : undefined;
  if (form === undefined) return plain;
  const back = readNumberForm(form, locale, { date1904 });
  if (!back) return plain;
  // Half a second for a time, else one unit in the 15th significant digit.
  const tolerance = shape.percent ? Math.abs(value) * 1e-14 : 0.5 / 86400;
  return Math.abs(back.value - value) <= tolerance ? form : plain;
}

/**
 * The plain text of a copied date: the shown text when it reads back to the same day, else the
 * edit text (`numberEditText`). A day and a month with no year (`2-Jan`), or a two-digit year
 * outside 1930 to 2029, does not read back. A time or an elapsed time keeps its shown text.
 */
export function copiedDateText(
  shown: string,
  value: number,
  code: string,
  locale: NumberLocale,
  date1904: boolean
): string {
  if (!isDateFormat(code)) return shown;
  const shape = codeShape(code);
  if (shape.elapsed || (!shape.order.includes('d') && !shape.order.includes('y'))) return shown;
  const day = (serial: number): number => Math.floor(Math.round(serial * 86400) / 86400);
  const back = readNumberForm(shown, locale, { date1904 });
  if (back && day(back.value) === day(value)) return shown;
  return numberEditText(value, code, locale, date1904);
}

const pad2 = (n: number): string => String(n).padStart(2, '0');

/** A date as ISO 8601 `YYYY-MM-DD`. */
const isoDate = (year: number, month: number, day: number): string =>
  `${String(year).padStart(4, '0')}-${pad2(month)}-${pad2(day)}`;

/** Whole seconds as `H:MM` or `H:MM:SS`. */
function clockText(seconds: number): string {
  const h = Math.floor(seconds / 3600);
  const m = Math.floor(seconds / 60) % 60;
  const sec = seconds % 60;
  return `${h}:${pad2(m)}${sec ? `:${pad2(sec)}` : ''}`;
}

/**
 * A serial as `YYYY-MM-DD`, `H:MM[:SS]`, or both, or undefined outside the date range. A value
 * under 1 in a format with no day and no year (`timeOnly`) is a time of day.
 */
function dateTimeText(serial: number, date1904: boolean, timeOnly: boolean): string | undefined {
  if (!Number.isFinite(serial) || serial < 0) return undefined;
  // Round the whole value to seconds first, so a value just under midnight is the next day.
  const total = Math.round(serial * 86400);
  const day = Math.floor(total / 86400);
  const seconds = total - day * 86400;
  // A time format shows a value that rounds to a whole day as `0:00`; `24:00` reads back to it.
  if (timeOnly && total === 86400) return '24:00';
  const time = clockText(seconds);
  // Serial 0 is no date in the 1900 system (1900-01-00), so its edit text is a time. In the
  // 1904 system it is 1904-01-01, so only a time format shows it as a time.
  if (day === 0 && (!date1904 || timeOnly)) return time;
  const date = serialToDate(day, date1904);
  if (!date) return undefined;
  const text = isoDate(date.year, date.month, date.day);
  return seconds ? `${text} ${time}` : text;
}

/** An elapsed time as `H:MM[:SS]` with the whole hours, or undefined past 9999 hours. */
function elapsedText(value: number): string | undefined {
  if (!Number.isFinite(value) || value < 0) return undefined;
  const seconds = Math.round(value * 86400);
  return seconds < 10000 * 3600 ? clockText(seconds) : undefined;
}

/** What the first section of a number format code shows. */
interface CodeShape {
  /**
   * The day, month, and year letters in order of their first use, such as `mdy`. An `m` can
   * also be the minute of `h:mm`, so a caller reads only `d` and `y`, or the order of a date.
   */
  readonly order: string;
  /** Whether the section has a `.` as a literal, such as the separator of `dd.mm.yyyy`. */
  readonly dot: boolean;
  readonly percent: boolean;
  /** Whether the section has an elapsed time bracket: `[h]`, `[mm]`, or `[ss]`. */
  readonly elapsed: boolean;
}

/**
 * The shape of the first section of a code. Quoted text, bracketed sections (`[$-409]`), and
 * the character after `\`, `_` (a space of its width), or `*` (a fill) are literals: they name
 * no part. One pass over the code.
 */
function codeShape(code: string): CodeShape {
  let order = '';
  let dot = false;
  let percent = false;
  let elapsed = false;
  let quoted = false;
  let bracket = '';
  let inBracket = false;
  let escaped = false;
  for (const ch of code.toLowerCase()) {
    if (escaped) escaped = false;
    else if (quoted) quoted = ch !== '"';
    else if (inBracket) {
      if (ch !== ']') bracket += ch;
      else {
        inBracket = false;
        const first = bracket[0];
        if (
          (first === 'h' || first === 'm' || first === 's') &&
          [...bracket].every((c) => c === first)
        )
          elapsed = true;
      }
    } else if (ch === ';') break;
    else if (ch === '"') quoted = true;
    else if (ch === '[') {
      inBracket = true;
      bracket = '';
    } else if (ch === '\\' || ch === '_' || ch === '*') escaped = true;
    else if (ch === '.') dot = true;
    else if (ch === '%') percent = true;
    else if ((ch === 'd' || ch === 'm' || ch === 'y') && !order.includes(ch)) order += ch;
  }
  return { order, dot, percent, elapsed };
}
