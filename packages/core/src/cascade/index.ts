// The `cascade` lane: cell values, the cell format cascade, number formats, theme colors, and
// theme fonts. DOM-free. It reads the store; it never writes it.
//
// The public list is explicit: internal tables, caps, and parsers stay inside the lane.

export type { CascadeProblem, CascadeProblemCode } from './problems.ts';
export {
  type CellValue,
  cellValue,
  /** The `xsd:double` reader, shared with the cell editor that reads typed numbers. */
  parseDouble,
  type ResolvedValue,
  sheetStale,
  stringItem,
  type StringItem,
  type TextRun,
} from './values.ts';
export { resolveColor, type ColorContext, type ColorUse, type Rgb } from './color.ts';
export { workbookTheme, type WorkbookTheme } from './theme.ts';
export {
  type BorderEdge,
  cellFormat,
  defaultFont,
  type CellFormatResult,
  type ResolvedAlignment,
  type ResolvedBorder,
  type ResolvedCellFormat,
  type GradientRect,
  type ResolvedFill,
  type ResolvedFont,
  type ResolvedProtection,
  richTextRuns,
  type RichTextRun,
  type UnderlineStyle,
} from './cell-format.ts';
export {
  type CheckBudget,
  conditionalFormat,
  sheetHasHiddenRules,
  rowHasShownRules,
  LAYOUT_CHECK_BUDGET,
  type ConditionalRule,
  type NotShownReason,
  sheetRules,
  type SheetRules,
} from './conditional.ts';
export {
  BUILTIN_FORMATS,
  builtinFormatCode,
  type CalendarDate,
  dateToSerial,
  EN_US,
  type FormatOptions,
  type FormattedValue,
  /** Why the number format parser cannot read a code, for the format commands. */
  formatCodeProblem,
  formatValue,
  isDateFormat,
  type NumberLocale,
  sameNumberLocale,
  serialToDate,
  subsecondDigits,
} from './number-format/index.ts';
export { type DisplayOptions, displayValue, formatCodeOf, shownText } from './display.ts';
