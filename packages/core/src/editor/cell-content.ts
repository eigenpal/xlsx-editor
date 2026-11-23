// The content of a cell as the formula bar shows it and the cell editor opens with it: the
// formula with `=`, or the value in a form that reads back. DOM-free.

import type { GridPoint } from '../contracts/index.ts';
import {
  attr,
  child,
  cellKey,
  decodeXstring,
  sharedFormulaText,
  sheetIndex,
  wrappedFormula,
  xmlPartOf,
  type WorkbookIndex,
  type XlsxPackage,
} from '../store/index.ts';
import { cellFormat, cellValue, formatCodeOf, type NumberLocale } from '../cascade/index.ts';
import { parseCellInput } from './cell-input.ts';
import { numberEditText } from './typed-forms.ts';

/**
 * What a cell holds, as a formula bar shows it: `=` and the formula, or the value. A shared formula
 * child shows its master's formula moved to the child (design D4); a `t="d"` date shows its ISO
 * 8601 text. A formula in an `mc:AlternateContent` child shows the text of its `x:f`
 * (`wrappedFormula`); a shared child there shows the formula of its direct master. A shared child
 * whose master text cannot be read shows its value, as does an empty `x:f`. A number with a date,
 * time, or percent format shows in a form that reads back (`numberEditText`): an ISO 8601 date, a
 * 24-hour time, or a percent. A text that would read back as another type shows a leading `'`. In a
 * cell with the Text format, only a text that starts with `'`, or with `=` and is not a lone `=`,
 * takes it. The `_xHHHH_` escapes of formula text read as their characters.
 */
export function cellContent(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  partName: string | undefined,
  at: GridPoint,
  locale: NumberLocale
): string {
  const part = xmlPartOf(pkg, partName);
  if (!part) return '';
  const entry = sheetIndex(part).cells.get(cellKey(at.row, at.col));
  if (!entry) return '';
  // A formula in a wrapper child shows its own text, as the grid marks the cell as a formula
  // cell. A wrapped shared child shows the formula of its direct master, moved to the cell.
  const f = child(entry.el, 'x:f') ?? wrappedFormula(entry.el);
  const text = f && sharedFormulaText(part, entry, f);
  // `x:f` is an `ST_Formula`, a restriction of `ST_Xstring` (sml.xsd): its `_xHHHH_` escapes
  // read as the characters they stand for.
  if (text) return `=${decodeXstring(text)}`;
  const v = cellValue(pkg, index, entry.el, false).value;
  const code = formatCodeOf(cellFormat(pkg, index, attr(entry.el, 's')).format, locale);
  switch (v.kind) {
    case 'empty':
      return '';
    case 'string': {
      // In a cell with the Text format, an entry stays text (`typedInput`). A formula and a
      // text with its own leading `'` are the exceptions, so only those take a leading `'`.
      if (code === '@')
        return v.text.startsWith("'") || (v.text.startsWith('=') && v.text.length > 1)
          ? `'${v.text}`
          : v.text;
      // A text that would read back as another type (a number, a formula, a boolean, or a
      // text with its own leading `'`) takes a leading `'`, so an edit keeps it text.
      const back = parseCellInput(v.text, locale, { date1904: index.date1904 });
      return back.type === 'string' && back.value === v.text ? v.text : `'${v.text}`;
    }
    case 'boolean':
      return v.value ? 'TRUE' : 'FALSE';
    case 'error':
      return v.code;
    case 'date':
      return v.iso;
    case 'number':
      // A date, a time, or a percent shows in a form that reads back (`numberEditText`).
      return numberEditText(v.value, code, locale, index.date1904);
  }
}
