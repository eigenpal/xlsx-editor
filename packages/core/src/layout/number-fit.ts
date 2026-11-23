// A number that does not fit the text area of its cell (quick reference, "Text in cells").
//
// Measured on basic-values.pdf and text-layout.pdf. Numbers never show across neighbor cells.
// - A General number shows the longest rounding narrower than the text area: 1234567.891 in a
//   48 point area shows `1234568` (42 points), 98765.4321 in 43 points shows `98765.4` (39; the
//   45 point `98765.43` does not show), and 123456789012 shows `1.2E+11`, not the 43 point
//   `1.23E+11`. A rounding that reaches zero shows `0`.
// - A formatted number that does not fit, and a General number that fits in no form, fills the
//   text area with `#`: five 5 point `#` in a 25 point area.

/** At most this many `#` fill one cell. */
const MAX_HASHES = 1000;
const GENERAL_SCIENTIFIC_DIGITS = 5;

const trimZeros = (s: string): string =>
  s.includes('.') ? s.replace(/0+$/, '').replace(/\.$/, '') : s;

// Text widths and text areas are whole pixels in points, so two widths that differ differ by at
// least a fraction of a point. A text area comes from absolute column offsets in points, and its
// floating-point noise depends on the offset. The tolerance keeps that noise out of the test.
const NOISE = 1e-6;

/**
 * Whether text `width` points wide fits `available` points. General needs room to spare; a
 * formatted number or text may fill the area exactly.
 */
export const fitsWidth = (width: number, available: number, general: boolean): boolean =>
  general ? width < available - NOISE : width <= available + NOISE;

/** The text of a number that fits `available` points, given a width function. */
export function fitNumber(
  value: number,
  text: string,
  general: boolean,
  decimal: string,
  available: number,
  width: (text: string) => number
): string {
  const fits = (candidate: string) => fitsWidth(width(candidate), available, general);
  if (fits(text)) return text;
  if (general && Number.isFinite(value)) {
    const abs = Math.abs(value);
    const sign = value < 0 ? '-' : '';
    const local = (s: string) => sign + s.replace('.', decimal);
    const fixedDecimals = Math.min(20, Math.max(0, (text.split(decimal)[1] ?? '').length - 1));
    if (abs >= 1e-4 && abs < 1e11)
      for (let d = fixedDecimals; d >= 0; d -= 1) {
        const fixed = trimZeros(abs.toFixed(d));
        // A rounding that reaches zero shows `0` (text-layout.pdf: 0.000123456 in 30 points).
        if (Number(fixed) === 0) {
          if (fits('0')) return '0';
          break;
        }
        const candidate = local(fixed);
        if (fits(candidate)) return candidate;
      }
    for (let d = GENERAL_SCIENTIFIC_DIGITS; d >= 0; d -= 1) {
      const [mantissa = '0', exp = '0'] = abs.toExponential(d).split('e');
      const e = Number(exp);
      const candidate =
        local(trimZeros(mantissa)) + `E${e < 0 ? '-' : '+'}${String(Math.abs(e)).padStart(2, '0')}`;
      if (fits(candidate)) return candidate;
    }
  }
  const hash = width('#');
  const count = hash > 0 ? Math.min(MAX_HASHES, Math.floor(available / hash)) : 0;
  return '#'.repeat(Math.max(0, count));
}
