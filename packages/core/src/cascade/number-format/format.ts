// Formats a cell value through a number format code (ECMA-376 Part 1, 18.8.31): the
// `x:numFmt/@formatCode` (`CT_NumFmt`, sml.xsd) of a custom id, or a built-in code. It parses the
// code once (cached by code, with a size cap), selects the section for the value by sign or by
// condition, and renders it. A code that the parser refuses falls back to `General` and
// returns the problem. A date section with a value that is not a date fills the cell with `#`.
//
// Colors: the reference PDFs paint `[Red]` text red (1 0 0) and `[Blue]` text blue (0 0 1),
// measured on `number-formats.pdf`. A section color is the `RRGGBB` of its token.

import { formatGeneral } from './digits.ts';
import { renderDate } from './dates.ts';
import type { NumberLocale } from './builtins.ts';
import type { CascadeProblem } from '../problems.ts';
import {
  MAX_CODE_LENGTH,
  type ParsedFormat,
  parseFormatCode,
  type Section,
} from '../../store/number-format/parse.ts';
import { renderNumber, type Pad } from './render-number.ts';

export interface FormatOptions {
  readonly locale: NumberLocale;
  /** `x:workbookPr/@date1904`. */
  readonly date1904: boolean;
}

export interface FormattedValue {
  /** The shown text, without the fill run. */
  readonly text: string;
  /** `RRGGBB` of the section's color token. */
  readonly color?: string;
  /** A character repeated to fill the column width, inserted at `index` of `text`. */
  readonly fill?: { readonly char: string; readonly index: number };
  /**
   * The spaces of `text` that are as wide as another character: a `?` placeholder is as wide as
   * a digit, and `_x` as wide as `x` (ECMA-376 Part 1, 18.8.31). Indexes skip the fill run.
   */
  readonly pads?: readonly Pad[];
  /**
   * A number shown through a text section (`@`): `general` alignment treats it as text
   * (number-formats.pdf draws `42` through `@` at the left of its cell).
   */
  readonly alignAsText?: boolean;
  /** Why the code fell back to `General`. */
  readonly problem?: CascadeProblem;
}

const CACHE_LIMIT = 512;
const cache = new Map<string, ParsedFormat | CascadeProblem>();

const TOO_LONG = (code: string): CascadeProblem => ({
  code: 'format-code-too-long',
  detail: code.slice(0, 16),
});

const isProblem = (v: ParsedFormat | CascadeProblem): v is CascadeProblem => 'code' in v;

/** The parsed code, or its problem. Cached by code; a code over the length cap is not cached. */
function parsed(code: string): ParsedFormat | CascadeProblem {
  if (code.length > MAX_CODE_LENGTH) return TOO_LONG(code);
  const hit = cache.get(code);
  if (hit !== undefined) return hit;
  const result = parseFormatCode(code);
  const value = result.ok ? result.format : result.problem;
  if (cache.size >= CACHE_LIMIT) {
    const oldest = cache.keys().next();
    if (!oldest.done) cache.delete(oldest.value);
  }
  cache.set(code, value);
  return value;
}

/** Formats a number, text, or boolean value through a format code. */
export function formatValue(
  value: number | string | boolean,
  code: string,
  options: FormatOptions
): FormattedValue {
  if (typeof value === 'boolean') return { text: value ? 'TRUE' : 'FALSE' };
  if (typeof value === 'number' && !Number.isFinite(value)) return { text: '#NUM!' };
  const format = parsed(code);
  const sep = { decimal: options.locale.decimal, group: options.locale.group };
  if (isProblem(format)) {
    const text =
      typeof value === 'string'
        ? value
        : signed(value, formatGeneral(Math.abs(value), sep.decimal));
    return { text, problem: format };
  }
  if (typeof value === 'string') return formatText(value, format);
  const chosen = selectSection(format, value);
  if (!chosen) return withColor(errorFill(), undefined);
  const { section, absolute } = chosen;
  const abs = Math.abs(value);
  const negative = value < 0 && !absolute;
  if (section.kind === 'date') {
    const out = negative ? undefined : renderDate(abs, section, sep, options.date1904);
    return withColor(out ?? errorFill(), section.color);
  }
  const rendered = renderNumber(abs, section, sep);
  const out: Rendered =
    section.kind === 'text' ? { ...rendered, alignAsText: true } : { ...rendered };
  if (!negative) return withColor(out, section.color);
  const shifted: Rendered = { ...out, text: '-' + out.text };
  if (out.fill) shifted.fill = { char: out.fill.char, index: out.fill.index + 1 };
  if (out.pads) shifted.pads = out.pads.map((p) => ({ index: p.index + 1, char: p.char }));
  return withColor(shifted, section.color);
}

interface Rendered {
  text: string;
  fill?: { char: string; index: number };
  pads?: Pad[];
  alignAsText?: boolean;
}

const signed = (value: number, text: string): string => (value < 0 ? '-' + text : text);

function errorFill(): { text: string; fill: { char: string; index: number } } {
  return { text: '', fill: { char: '#', index: 0 } };
}

function withColor(out: Rendered, color: string | undefined): FormattedValue {
  return color ? { ...out, color } : out;
}

/** Text goes through the fourth section, or through a first section that holds `@`. */
function formatText(value: string, format: ParsedFormat): FormattedValue {
  const sections = format.sections;
  const section =
    sections.length >= 4 ? sections[3]! : sections[0]!.kind === 'text' ? sections[0]! : undefined;
  if (!section) return { text: value };
  let text = '';
  let fill: { char: string; index: number } | undefined;
  const pads: Pad[] = [];
  for (const t of section.tokens) {
    if (t.kind === 'lit') {
      if (t.pad !== undefined) pads.push({ index: text.length, char: t.pad });
      text += t.text;
    } else if (t.kind === 'text') text += value;
    else if (t.kind === 'fill' && !fill) fill = { char: t.char, index: text.length };
  }
  const out: Rendered = fill ? { text, fill } : { text };
  if (pads.length) out.pads = pads;
  return withColor(out, section.color);
}

function matches(cond: { op: string; value: number }, v: number): boolean {
  switch (cond.op) {
    case '<':
      return v < cond.value;
    case '<=':
      return v <= cond.value;
    case '>':
      return v > cond.value;
    case '>=':
      return v >= cond.value;
    case '=':
      return v === cond.value;
    default:
      return v !== cond.value;
  }
}

/** A condition that only negative values meet: its section shows the value without a sign. */
function negativeOnly(cond: { op: string; value: number }): boolean {
  return (cond.op === '<' && cond.value <= 0) || (cond.op === '<=' && cond.value < 0);
}

/** The section for a number, and whether it shows the absolute value. */
function selectSection(
  format: ParsedFormat,
  v: number
): { section: Section; absolute: boolean } | undefined {
  const all = format.sections;
  const numeric = all.length >= 4 ? all.slice(0, 3) : all;
  if (numeric[0]!.cond || numeric[1]?.cond) {
    for (let k = 0; k < numeric.length && k < 3; k++) {
      const s = numeric[k]!;
      if (s.cond ? matches(s.cond, v) : true)
        return {
          section: s,
          absolute: s.cond ? negativeOnly(s.cond) : k === 1 && !numeric[0]!.cond,
        };
    }
    return undefined;
  }
  if (numeric.length === 1) return { section: numeric[0]!, absolute: false };
  if (v < 0) return { section: numeric[1]!, absolute: true };
  if (v === 0 && numeric.length >= 3) return { section: numeric[2]!, absolute: false };
  return { section: numeric[0]!, absolute: false };
}

/**
 * True when the first section of the code formats a date or time: the section that a
 * non-negative value, such as a converted `t="d"` date, uses.
 */
export function isDateFormat(code: string): boolean {
  const format = parsed(code);
  return !isProblem(format) && format.sections[0]!.kind === 'date';
}

/**
 * The digits of fractions of a second that the first section of a date code shows, as in
 * `mm:ss.0`, or 0 for a code that shows none or is not valid.
 */
export function subsecondDigits(code: string): number {
  const format = parsed(code);
  return isProblem(format) ? 0 : format.sections[0]!.subsec;
}

/** The problem of a code that falls back to `General`, or undefined for a valid code. */
export function formatCodeProblem(code: string): CascadeProblem | undefined {
  const format = parsed(code);
  return isProblem(format) ? format : undefined;
}
