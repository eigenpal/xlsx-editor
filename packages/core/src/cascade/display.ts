// The display text of a cell: its stored value (values.ts) through the number format of its
// cell format (cell-format.ts), in a number locale. Errors show their code, and an empty cell
// or a stale formula with no cached value shows nothing.

import type { ResolvedCellFormat } from './cell-format.ts';
import type { CellValue } from './values.ts';
import {
  builtinFormatCode,
  dateToSerial,
  formatValue,
  type FormattedValue,
  type NumberLocale,
} from './number-format/index.ts';

export interface DisplayOptions {
  readonly locale: NumberLocale;
  readonly date1904: boolean;
}

/** The format code of a cell format: its custom code, the built-in code, or `General`. */
export function formatCodeOf(format: ResolvedCellFormat, locale: NumberLocale): string {
  return format.formatCode ?? builtinFormatCode(format.numFmtId, locale) ?? 'General';
}

/** The text that a cell shows, with the color and fill of its number format. */
export function displayValue(
  value: CellValue,
  format: ResolvedCellFormat,
  options: DisplayOptions
): FormattedValue {
  const code = formatCodeOf(format, options.locale);
  switch (value.kind) {
    case 'empty':
      return { text: '' };
    case 'error':
      return { text: value.code };
    case 'boolean':
      return { text: value.value ? 'TRUE' : 'FALSE' };
    case 'string':
      return formatValue(value.text, code, options);
    case 'number':
      return formatValue(value.value, code, options);
    case 'date': {
      const serial = dateToSerial(value.iso, options.date1904);
      return serial === undefined ? { text: value.iso } : formatValue(serial, code, options);
    }
  }
}

/**
 * The text that a cell shows before layout fits it to its box: rich text runs as stored, every
 * other value through {@link displayValue}, as layout's cell text does. The editor's
 * announcement reads it, so assistive technology hears the text that the grid paints.
 */
export function shownText(
  value: CellValue,
  format: ResolvedCellFormat,
  options: DisplayOptions
): string {
  if (value.kind === 'string' && value.runs && value.runs.length > 0) return value.text;
  return displayValue(value, format, options).text;
}
