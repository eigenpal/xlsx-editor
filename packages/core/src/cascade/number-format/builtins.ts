// Number format locales and the built-in format table (ECMA-376 Part 1, 18.8.30). A locale
// gives the decimal and group separators and can replace built-in codes, because the built-in
// date and negative-number codes follow the locale that shows them. The ECMA-376 table itself
// lives in the store lane, which writes a code that equals a built-in code as its id.

import { BUILTIN_FORMATS } from '../../store/number-format/builtins.ts';
import { isPlainXmlText } from '../../store/store/xml-text.ts';

export { BUILTIN_FORMATS };

/**
 * The separators of a number locale, and its built-in format codes, for the shown values.
 *
 * @public
 */
export interface NumberLocale {
  /** The decimal separator. */
  readonly decimal: string;
  /** The thousands group separator. */
  readonly group: string;
  /** Built-in codes that replace the ECMA-376 table, by format id. */
  readonly builtins?: Readonly<Record<number, string>>;
  /**
   * The currency symbol of the currency and accounting entries of the number format list, as
   * the text before the number. `$` when absent. A symbol is 1 to 8 UTF-16 code units of
   * plain text: no control code, no U+FFFE or U+FFFF, and no lone surrogate. With any other
   * symbol, the list leaves out the two entries: the editor never writes another symbol in its
   * place. The symbol changes only the code that a list entry writes, never a shown value.
   */
  readonly currencySymbol?: string;
}

/** The currency symbol of a number locale that gives none. @internal */
export const DEFAULT_CURRENCY_SYMBOL = '$';

/** The longest currency symbol that a number locale can give, in UTF-16 code units. @internal */
export const MAX_CURRENCY_SYMBOL_LENGTH = 8;

/**
 * The currency symbol of a number locale: its own symbol, `$` when it gives none, or undefined
 * when its symbol is not valid (`NumberLocale.currencySymbol`).
 */
export function currencySymbolOf(locale: NumberLocale): string | undefined {
  const symbol = locale.currencySymbol;
  if (symbol === undefined) return DEFAULT_CURRENCY_SYMBOL;
  return typeof symbol === 'string' &&
    symbol.length > 0 &&
    symbol.length <= MAX_CURRENCY_SYMBOL_LENGTH &&
    isPlainXmlText(symbol)
    ? symbol
    : undefined;
}

/**
 * Whether two number locales format the same: equal separators, equal built-in codes, and the
 * same currency symbol. A host can pass a new object with each render; only a change of value
 * applies.
 */
export function sameNumberLocale(
  a: NumberLocale | undefined,
  b: NumberLocale | undefined
): boolean {
  if (a === b) return true;
  const x = a ?? EN_US;
  const y = b ?? EN_US;
  if (x.decimal !== y.decimal || x.group !== y.group) return false;
  if (x.currencySymbol !== y.currencySymbol && currencySymbolOf(x) !== currencySymbolOf(y))
    return false;
  const xb = x.builtins ?? {};
  const yb = y.builtins ?? {};
  const keys = Object.keys(xb);
  if (keys.length !== Object.keys(yb).length) return false;
  return keys.every((k) => Object.hasOwn(yb, k) && xb[Number(k)] === yb[Number(k)]);
}

/** The ECMA-376 built-in codes with `.` and `,`. */
export const EN_US: NumberLocale = { decimal: '.', group: ',' };

/**
 * The code of a built-in format id in a locale, or undefined for an id that has no known code
 * (the locale-dependent ids 5 to 8, 23 to 36, 41 to 44, and 50 to 163).
 */
export function builtinFormatCode(id: number, locale: NumberLocale): string | undefined {
  if (!Number.isInteger(id)) return undefined;
  if (locale.builtins && Object.hasOwn(locale.builtins, id)) return locale.builtins[id];
  return Object.hasOwn(BUILTIN_FORMATS, id) ? BUILTIN_FORMATS[id] : undefined;
}
