// The built-in number format codes of ECMA-376 Part 1, 18.8.30, and the format code checks
// that the store applies before it writes a custom code to the styles part. The id is an
// `ST_NumFmtId` (sml.xsd), on `x:xf/@numFmtId` (`CT_Xf`) and `x:numFmt/@numFmtId` (`CT_NumFmt`).

import { parseFormatCode, type FormatCodeProblem } from './parse.ts';

/**
 * The built-in format codes of ECMA-376 Part 1, 18.8.30. Id 47 is `mm:ss.0`: the printed table
 * gives `mmss.0`, and the measured `number-formats` fixture shows the minutes and seconds
 * separated by `:`.
 */
export const BUILTIN_FORMATS: Readonly<Record<number, string>> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'mm-dd-yy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yy h:mm',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mm:ss.0',
  48: '##0.0E+0',
  49: '@',
};

/** The built-in id of a format code, by the ECMA-376 table. */
const BUILTIN_IDS: ReadonlyMap<string, number> = (() => {
  const out = new Map<string, number>();
  for (const [id, code] of Object.entries(BUILTIN_FORMATS))
    if (!out.has(code)) out.set(code, Number(id));
  return out;
})();

/** The built-in id whose ECMA-376 code is exactly `code`, or undefined. */
export function builtinFormatId(code: string): number | undefined {
  return BUILTIN_IDS.get(code);
}

/** Why the number format parser cannot read a custom format code, or undefined. */
export function formatCodeReadProblem(code: string): FormatCodeProblem | undefined {
  const r = parseFormatCode(code);
  return r.ok ? undefined : r.problem;
}
