// Renders a non-negative number through one number section: integer and fraction digits,
// grouping, scaling, percent, scientific notation, and fractions. The sign is the caller's job.
// Each digit placeholder renders as a string: `0` shows a zero where no digit is, `#` shows
// nothing, and `?` shows a space as wide as a digit. Integer digits fill from the right; extra
// digits go to the leftmost placeholder. `pads` names each space that stands for the width of
// another character: a `?` placeholder (a digit, `0`) and `_x` (the character `x`).

import {
  type Decimal,
  formatGeneral,
  roundDecimal,
  roundSignificant,
  shiftDecimal,
  toDecimal,
} from './digits.ts';
import type { DigitChar, Section } from '../../store/number-format/parse.ts';

export interface Separators {
  readonly decimal: string;
  readonly group: string;
}

/** A space of `text` at `index` that is as wide as the character `char`. */
export interface Pad {
  readonly index: number;
  readonly char: string;
}

export interface Rendered {
  text: string;
  fill?: { char: string; index: number };
  pads?: Pad[];
}

/** Stands for the space of a `?` placeholder inside the digit parts, until `emit` writes it. */
const DIGIT_PAD = '\u2007';

const pad = (c: DigitChar): string => (c === '0' ? '0' : c === '?' ? DIGIT_PAD : '');

/** Integer digits right-aligned into placeholders, with group separators when `group` is set. */
function placeInteger(digits: string, slots: readonly DigitChar[], group: string): string[] {
  const out = slots.map(() => '');
  const n = slots.length;
  if (n === 0) return out;
  const total = Math.max(n, digits.length);
  for (let r = total - 1; r >= 0; r--) {
    const slot = r >= n ? 0 : n - 1 - r;
    const isDigit = r < digits.length;
    let cell = isDigit ? digits[digits.length - 1 - r]! : r < n ? pad(slots[slot]!) : '';
    if (group && r > 0 && r % 3 === 0 && (isDigit || cell === '0')) cell += group;
    out[slot] += cell;
  }
  return out;
}

/** Fraction digits left-aligned; trailing zeros on `#` and `?` placeholders drop. */
function placeFraction(frac: string, slots: readonly DigitChar[]): string[] {
  let last = frac.length - 1;
  while (last >= 0 && frac[last] === '0') last--;
  return slots.map((c, k) => (k <= last ? frac[k]! : pad(c)));
}

/** Denominator digits left-aligned and padded on the right. */
function placeLeft(digits: string, slots: readonly DigitChar[]): string[] {
  const out = slots.map((c, k) => (k < digits.length ? digits[k]! : pad(c)));
  if (digits.length > slots.length && slots.length > 0)
    out[slots.length - 1] += digits.slice(slots.length);
  return out;
}

interface Parts {
  int: string[];
  frac: string[];
  exp: string[];
  num: string[];
  den: string[];
  /** Integer digits with no integer placeholder, shown before the point. */
  loose: string;
  expSign: string;
  /** The fraction is zero: the numerator, slash, and denominator show as spaces. */
  blankFraction: boolean;
}

/** Renders `abs` (finite, non-negative) through a number, text, or general section. */
export function renderNumber(abs: number, section: Section, sep: Separators): Rendered {
  if (section.kind === 'text' || section.kind === 'general')
    return emit(section, sep, emptyParts(), formatGeneral(abs, sep.decimal));
  const parts = section.fraction
    ? fractionParts(abs, section)
    : section.scientific
      ? scientificParts(shiftDecimal(toDecimal(abs), section.shift), section)
      : fixedParts(shiftDecimal(toDecimal(abs), section.shift), section, sep.group);
  return emit(section, sep, parts, '');
}

function emptyParts(): Parts {
  return {
    int: [],
    frac: [],
    exp: [],
    num: [],
    den: [],
    loose: '',
    expSign: '',
    blankFraction: false,
  };
}

function fixedParts(d: Decimal, section: Section, group: string): Parts {
  const r = roundDecimal(d, section.frac.length);
  const parts = emptyParts();
  parts.int = placeInteger(r.int, section.int, section.thousands ? group : '');
  if (section.int.length === 0) parts.loose = r.int;
  parts.frac = placeFraction(r.frac, section.frac);
  return parts;
}

function scientificParts(d: Decimal, section: Section): Parts {
  const n = Math.max(1, section.int.length);
  const engineering = section.int.length > 1 && section.int[0] === '#';
  let value = d;
  let e = 0;
  let mantissa = { int: '', frac: ''.padEnd(section.frac.length, '0') };
  if (value.digits) {
    // A carry in rounding can add a digit; then the exponent moves once more.
    for (let pass = 0; pass < 2; pass++) {
      const e0 = value.exp - 1;
      e = engineering ? Math.floor(e0 / n) * n : e0 - (n - 1);
      const intDigits = e0 - e + 1;
      value = roundSignificant(value, intDigits + section.frac.length);
      if (value.exp - 1 !== e0) continue;
      mantissa = roundDecimal(shiftDecimal(value, -e), section.frac.length);
      break;
    }
  }
  const parts = emptyParts();
  parts.int = placeInteger(mantissa.int, section.int, '');
  parts.frac = placeFraction(mantissa.frac, section.frac);
  parts.exp = placeInteger(String(Math.abs(e)), section.exp, '');
  parts.expSign = e < 0 ? '-' : '';
  return parts;
}

function fractionParts(abs: number, section: Section): Parts {
  const value = abs * 10 ** section.shift;
  const hasInt = section.int.length > 0;
  let whole = hasInt ? Math.floor(value) : 0;
  const rest = value - whole;
  let num: number;
  let den: number;
  if (section.fixedDen !== undefined && section.fixedDen > 0) {
    den = section.fixedDen;
    num = Math.floor(rest * den + 0.5);
  } else {
    const max = 10 ** Math.max(1, section.den.length) - 1;
    [num, den] = approximate(rest, max);
  }
  if (hasInt && num === den) {
    whole += 1;
    num = 0;
  }
  const parts = emptyParts();
  const wholeDigits = whole > 0 ? decimalDigits(whole) : '';
  parts.blankFraction = num === 0 && hasInt;
  const intDigits = parts.blankFraction && !wholeDigits ? '0' : wholeDigits;
  parts.int = placeInteger(intDigits, section.int, '');
  if (parts.blankFraction) {
    parts.num = section.num.map(() => ' ');
    parts.den = section.den.map(() => ' ');
  } else {
    parts.num = placeInteger(decimalDigits(num), section.num, '');
    parts.den = placeLeft(decimalDigits(den), section.den);
  }
  return parts;
}

/** The decimal digits of a non-negative whole number, never in exponent notation (1e21 and up). */
const decimalDigits = (n: number): string => roundDecimal(toDecimal(n), 0).int || '0';

/** The best fraction with a denominator up to `max`, by continued fractions. */
export function approximate(x: number, max: number): [number, number] {
  let p0 = 0;
  let q0 = 1;
  let p1 = 1;
  let q1 = 0;
  let b = x;
  let p = 0;
  let q = 1;
  for (let step = 0; step < 64; step++) {
    const a = Math.floor(b);
    const pn = a * p1 + p0;
    const qn = a * q1 + q0;
    if (qn > max) break;
    p = pn;
    q = qn;
    if (b - a < 1e-9) break;
    b = 1 / (b - a);
    p0 = p1;
    q0 = q1;
    p1 = pn;
    q1 = qn;
  }
  return [p, q];
}

/** Writes the tokens of a section with the rendered parts. */
function emit(section: Section, sep: Separators, parts: Parts, general: string): Rendered {
  let text = '';
  let fill: Rendered['fill'];
  const pads: Pad[] = [];
  for (const t of section.tokens) {
    switch (t.kind) {
      case 'lit':
        if (t.pad !== undefined) pads.push({ index: text.length, char: t.pad });
        text += t.text;
        break;
      case 'fill':
        if (!fill) fill = { char: t.char, index: text.length };
        break;
      case 'digit':
        for (const ch of parts[t.role][t.index] ?? '') {
          if (ch === DIGIT_PAD) pads.push({ index: text.length, char: '0' });
          text += ch === DIGIT_PAD ? ' ' : ch;
        }
        break;
      case 'point':
        text += parts.loose + sep.decimal;
        break;
      case 'percent':
        text += '%';
        break;
      case 'exp':
        text += t.source[0] + (parts.expSign || (t.plus ? '+' : ''));
        break;
      case 'slash':
        text += parts.blankFraction ? ' ' : '/';
        break;
      case 'den':
        text += parts.blankFraction ? ''.padEnd(String(t.value).length, ' ') : String(t.value);
        break;
      case 'text':
      case 'general':
        text += general;
        break;
      default:
        break;
    }
  }
  const out: Rendered = fill ? { text, fill } : { text };
  if (pads.length) out.pads = pads;
  return out;
}
