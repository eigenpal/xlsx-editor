// The fixture `conditional-rules`: `cellIs` rules with literal operands and the parts of a
// differential format that the overlay sets (design D3 and D4 of
// `workbook-conditional-formatting`). Column A names each case; column B holds the value that
// its rule tests.

import { c, row, type FixtureWorkbook } from './builder.ts';

/** The differential formats, by index. */
const DXFS = [
  '<dxf><font><b/></font></dxf>',
  '<dxf><fill><patternFill><bgColor rgb="FFFFFF00"/></patternFill></fill></dxf>',
  '<dxf><font><i/><color rgb="FFFF0000"/></font></dxf>',
  '<dxf><numFmt numFmtId="200" formatCode="0.00"/></dxf>',
  '<dxf><font><strike/></font></dxf>',
  '<dxf><font><i/></font></dxf>',
  '<dxf><font><b/></font><fill><patternFill><bgColor rgb="FF00B050"/></patternFill></fill></dxf>',
  '<dxf><fill><patternFill><fgColor rgb="FF0070C0"/></patternFill></fill></dxf>',
  '<dxf><font><u/></font></dxf>',
  '<dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFFC000"/></patternFill></fill></dxf>',
  '<dxf><font><color theme="5"/></font><fill><patternFill><bgColor theme="4" tint="0.59999389629810485"/></patternFill></fill></dxf>',
  '<dxf><fill><patternFill patternType="solid"><fgColor rgb="FF0070C0"/></patternFill></fill></dxf>',
];

const rule = (
  ref: string,
  operator: string,
  dxf: number,
  priority: number,
  formulas: readonly string[],
  stop = false
) =>
  `<conditionalFormatting sqref="${ref}"><cfRule type="cellIs" dxfId="${dxf}" priority="${priority}" operator="${operator}"${stop ? ' stopIfTrue="1"' : ''}>` +
  formulas.map((f) => `<formula>${f.replaceAll('"', '&quot;')}</formula>`).join('') +
  '</cfRule></conditionalFormatting>';

/** Each case: its label, its value cell, and its rules. */
const CASES: readonly (readonly [string, string, string])[] = [
  ['greaterThan 3: fill', c('B1', 5), rule('B1', 'greaterThan', 1, 1, ['3'])],
  ['lessThan 3: red italic', c('B2', 2), rule('B2', 'lessThan', 2, 2, ['3'])],
  ['between 10 and 5: 0.00', c('B3', 7), rule('B3', 'between', 3, 3, ['10', '5'])],
  ['notEqual 4: bold', c('B4', 4), rule('B4', 'notEqual', 0, 4, ['4'])],
  ['text, greaterThan 3: bold', c('B5', 'abc'), rule('B5', 'greaterThan', 0, 5, ['3'])],
  ['ABC equal "abc": strike', c('B6', 'ABC'), rule('B6', 'equal', 4, 6, ['"abc"'])],
  ['TRUE greaterThan "z": bold', c('B7', true), rule('B7', 'greaterThan', 0, 7, ['"z"'])],
  [
    'two true rules',
    c('B8', 9),
    rule('B8', 'greaterThan', 5, 8, ['1']) + rule('B8', 'greaterThan', 6, 9, ['1']),
  ],
  [
    'stopIfTrue',
    c('B9', 9),
    rule('B9', 'greaterThan', 5, 10, ['1'], true) + rule('B9', 'greaterThan', 6, 11, ['1']),
  ],
  ['fgColor only', c('B10', 6), rule('B10', 'greaterThan', 7, 12, ['1'])],
  ['error, notEqual 5', c('B11', '#DIV/0!', { t: 'e' }), rule('B11', 'notEqual', 0, 13, ['5'])],
  ['equal 8: bold', c('B12', 8), rule('B12', 'equal', 0, 14, ['8'])],
  ['lessThanOrEqual 8: bold', c('B13', 8), rule('B13', 'lessThanOrEqual', 0, 15, ['8'])],
  ['greaterThanOrEqual 9: bold', c('B14', 8), rule('B14', 'greaterThanOrEqual', 0, 16, ['9'])],
  ['notBetween 1 and 2: bold', c('B15', 5), rule('B15', 'notBetween', 0, 17, ['1', '2'])],
  ['empty x:c, lessThan 5: fill', '<c r="B16"/>', rule('B16', 'lessThan', 1, 18, ['5'])],
  ['no x:c, equal 0: fill', '', rule('B17', 'equal', 1, 19, ['0'])],
  ['empty x:c, equal "": fill', '<c r="B18"/>', rule('B18', 'equal', 1, 20, ['""'])],
  ['greaterThan 1: underline', c('B19', 3), rule('B19', 'greaterThan', 8, 21, ['1'])],
  ['solid bgColor: orange fill', c('B20', 3), rule('B20', 'greaterThan', 9, 22, ['1'])],
  ['theme colors', c('B21', 3), rule('B21', 'greaterThan', 10, 23, ['1'])],
  ['own fill, yellow rule', c('B22', 3, { s: 1 }), rule('B22', 'greaterThan', 1, 24, ['1'])],
  ['own fill, fgColor only', c('B23', 3, { s: 1 }), rule('B23', 'greaterThan', 7, 25, ['1'])],
  [
    'own fill, solid fgColor only',
    c('B24', 3, { s: 1 }),
    rule('B24', 'greaterThan', 11, 26, ['1']),
  ],
  ['between 1 and 10: bold', c('B25', 5), rule('B25', 'between', 0, 27, ['1', '10'])],
  ['notEqual 4 on 5: bold', c('B26', 5), rule('B26', 'notEqual', 0, 28, ['4'])],
];

export function conditionalRulesWorkbook(): FixtureWorkbook {
  return {
    sheets: [
      {
        name: 'Rules',
        cols: '<cols><col min="1" max="1" width="28" customWidth="1"/></cols>',
        rows: CASES.map(([label, value], i) => row(i + 1, [c(`A${i + 1}`, label), value])).join(''),
        afterMerges: CASES.map(([, , rules]) => rules).join(''),
      },
    ],
    styles: {
      fills: [
        '<fill><patternFill patternType="solid"><fgColor rgb="FF9999FF"/></patternFill></fill>',
      ],
      cellXfs: ['<xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"/>'],
      dxfs: DXFS,
    },
  };
}
