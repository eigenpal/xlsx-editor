// Decimal digits of a double for number formats. A value is read as its 15 significant digits
// (the precision that a cell value shows), and rounding works on that digit string, half away
// from zero, so `2.25` with one decimal gives `2.3` and `1.005` with two gives `1.01`. Also the
// `General` format: the shortest form that fits 11 characters, scientific beyond that.

/** A non-negative value as `0.digits × 10^exp`. `digits` has no leading or trailing zeros. */
export interface Decimal {
  readonly digits: string;
  readonly exp: number;
}

export const ZERO: Decimal = { digits: '', exp: 0 };

/** The 15 significant digits of a finite non-negative value. */
export function toDecimal(abs: number): Decimal {
  if (abs === 0 || !Number.isFinite(abs)) return ZERO;
  const text = abs.toExponential(14);
  const e = text.indexOf('e');
  let digits = text[0]! + text.slice(2, e);
  let end = digits.length;
  while (end > 0 && digits[end - 1] === '0') end--;
  digits = digits.slice(0, end);
  return { digits, exp: Number(text.slice(e + 1)) + 1 };
}

/** Multiplies by `10^shift`. */
export function shiftDecimal(d: Decimal, shift: number): Decimal {
  return d.digits ? { digits: d.digits, exp: d.exp + shift } : ZERO;
}

const zeros = (count: number): string => ''.padEnd(Math.max(0, count), '0');

/** Adds one to the last digit of a digit string; a carry out of the first digit prepends `1`. */
function increment(digits: string): string {
  let k = digits.length - 1;
  while (k >= 0 && digits[k] === '9') k--;
  if (k < 0) return '1' + zeros(digits.length);
  return digits.slice(0, k) + String(Number(digits[k]) + 1) + zeros(digits.length - k - 1);
}

/**
 * Rounds to `decimals` places, half away from zero. `int` has no leading zeros (empty for zero);
 * `frac` has exactly `decimals` digits.
 */
export function roundDecimal(d: Decimal, decimals: number): { int: string; frac: string } {
  if (!d.digits) return { int: '', frac: zeros(decimals) };
  const intText = d.exp > 0 ? (d.digits + zeros(d.exp - d.digits.length)).slice(0, d.exp) : '';
  const fracText = d.exp >= 0 ? d.digits.slice(d.exp) : zeros(-d.exp) + d.digits;
  let combined = intText + (fracText + zeros(decimals)).slice(0, decimals);
  if ((fracText[decimals] ?? '0') >= '5') combined = increment(combined || '0');
  const cut = combined.length - decimals;
  let int = combined.slice(0, cut);
  let start = 0;
  while (start < int.length && int[start] === '0') start++;
  int = int.slice(start);
  return { int, frac: combined.slice(cut) };
}

/** `count` significant digits, half away from zero; a carry can raise the exponent. */
export function roundSignificant(d: Decimal, count: number): Decimal {
  if (!d.digits || d.digits.length <= count) return d;
  let kept = d.digits.slice(0, count);
  let exp = d.exp;
  if (d.digits[count]! >= '5') {
    const next = increment(kept);
    if (next.length > kept.length) exp++;
    kept = next.slice(0, count);
  }
  let end = kept.length;
  while (end > 0 && kept[end - 1] === '0') end--;
  return { digits: kept.slice(0, end), exp };
}

/**
 * The engine's width of the `General` format (built-in id 0, Part 1, 18.8.30), in characters.
 * The standard does not state this width. The `General` cases of fixture `number-formats`
 * (scripts/fixtures/cases.ts) check it against the reference PDF in measured.test.ts: for
 * example, 1234567.891 shows in full and 123456789012 shows as 1.23457E+11.
 */
const GENERAL_WIDTH = 11;

/** The `General` form of a non-negative value, with the locale's decimal separator. */
export function formatGeneral(abs: number, decimal: string): string {
  if (!Number.isFinite(abs)) return abs > 0 ? 'Infinity' : 'NaN';
  const d = toDecimal(abs);
  if (!d.digits) return '0';
  const e = d.exp - 1;
  if (e >= GENERAL_WIDTH) return generalScientific(d, decimal);
  if (e >= 0) {
    const r = roundDecimal(d, Math.max(0, GENERAL_WIDTH - 1 - (e + 1)));
    if (r.int.length > GENERAL_WIDTH) return generalScientific(d, decimal);
    return joinTrimmed(r.int, r.frac, decimal);
  }
  const r = roundDecimal(d, GENERAL_WIDTH - 2);
  const fixed = joinTrimmed(r.int, r.frac, decimal);
  const shown = significantCount(r.int + r.frac);
  const sci = roundSignificant(d, 6);
  if (shown === 0 || shown < sci.digits.length) return generalScientific(d, decimal);
  return fixed;
}

function significantCount(digits: string): number {
  let start = 0;
  while (start < digits.length && digits[start] === '0') start++;
  let end = digits.length;
  while (end > start && digits[end - 1] === '0') end--;
  return end - start;
}

function joinTrimmed(int: string, frac: string, decimal: string): string {
  let end = frac.length;
  while (end > 0 && frac[end - 1] === '0') end--;
  const head = int || '0';
  return end > 0 ? head + decimal + frac.slice(0, end) : head;
}

function generalScientific(d: Decimal, decimal: string): string {
  let e = d.exp - 1;
  const expDigits = Math.abs(e) >= 100 ? 3 : 2;
  // Mantissa digit, point, decimals, `E`, sign, exponent digits: at most 11 characters.
  const decimals = GENERAL_WIDTH - 4 - expDigits;
  const r = roundSignificant(d, decimals + 1);
  e = r.exp - 1;
  const mantissa = r.digits.padEnd(1, '0');
  const tail = mantissa.slice(1);
  const sign = e < 0 ? '-' : '+';
  const exp = String(Math.abs(e)).padStart(2, '0');
  return `${mantissa[0]}${tail ? decimal + tail : ''}E${sign}${exp}`;
}
