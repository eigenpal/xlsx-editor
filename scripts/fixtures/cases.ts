// The measurement cases of section 4 of the active change: number formats, dates, colors, and
// fonts. Each case is one cell of a fixture. The fixture catalog builds the cells, a private
// measurement records the shown result of each one in `fixtures/measured.json`, and the cascade tests compare the engine's result with that record.
//
// A case names its cell, so the record and the tests find it without a second list.

import { c, row } from './builder.ts';

/** One number format case: a value shown through a built-in format id or a custom code. */
export interface FormatCase {
  readonly value: number | string;
  /** The built-in format id, when the case uses one. */
  readonly id?: number;
  /** The custom format code, when the case uses one. */
  readonly code?: string;
  /** `t="d"`: the value is an ISO 8601 date. */
  readonly isoDate?: boolean;
}

/** Cases of the 1900 date system. */
export const FORMAT_CASES: readonly FormatCase[] = [
  // General: the shortest form that fits 11 characters.
  { id: 0, value: 0 },
  { id: 0, value: 1234.5 },
  { id: 0, value: -3.5 },
  { id: 0, value: 0.1 },
  { id: 0, value: 1234567.891 },
  { id: 0, value: 123456789012 },
  { id: 0, value: 12345678901234 },
  { id: 0, value: 0.000012345 },
  { id: 0, value: 1 / 3 },
  { id: 0, value: 2 / 3 },
  { id: 0, value: -0.000001 },
  // Built-in number formats.
  { id: 1, value: 2.5 },
  { id: 1, value: -2.5 },
  { id: 2, value: 3.14159 },
  { id: 3, value: 1234567 },
  { id: 4, value: 1234.5 },
  { id: 4, value: -1234.5 },
  { id: 9, value: 0.256 },
  { id: 10, value: 0.25678 },
  { id: 11, value: 12345 },
  { id: 12, value: 1.75 },
  { id: 13, value: 0.3333 },
  // Built-in date and time formats.
  { id: 14, value: 45352 },
  { id: 15, value: 45352 },
  { id: 16, value: 45352 },
  { id: 17, value: 45352 },
  { id: 18, value: 0.75 },
  { id: 19, value: 0.75 },
  { id: 20, value: 0.75 },
  { id: 21, value: 0.5004 },
  { id: 22, value: 45352.75 },
  { id: 37, value: -1234 },
  { id: 38, value: -1234 },
  { id: 39, value: -1234.5 },
  { id: 40, value: -1234.5 },
  { id: 45, value: 0.0123 },
  { id: 46, value: 1.5 },
  { id: 47, value: 0.0123 },
  { id: 48, value: 12345 },
  { id: 49, value: 42 },
  // Custom codes: digits, separators, scaling, and literals.
  { code: '0.000', value: 3.14159 },
  { code: '0.0', value: 2.25 },
  { code: '0.##', value: 1.5 },
  { code: '0.##', value: 2 },
  { code: '#', value: 0 },
  { code: '00000', value: 42 },
  { code: '?.??', value: 1.5 },
  { code: '#,##0,', value: 1234567 },
  { code: '0.0,,"M"', value: 1234567 },
  { code: '"$"#,##0', value: 1234.5 },
  { code: '\\$0', value: 5 },
  { code: '0%', value: 0.256 },
  { code: '0.00E+00', value: 12345 },
  { code: '##0.0E+0', value: 12345 },
  { code: '# ?/?', value: 1.75 },
  { code: '# ??/??', value: -2.4 },
  // Sections, conditions, colors, and text.
  { code: '#,##0.00;[Red]-#,##0.00', value: -1234.5 },
  { code: '0;-0;"zero";"text:"@', value: 0 },
  { code: '0;-0;"zero";"text:"@', value: -5 },
  { code: '0;-0;"zero";"text:"@', value: 'abc' },
  { code: '[>=100]0;[<0]-0.0;0.00', value: 150 },
  { code: '[>=100]0;[<0]-0.0;0.00', value: -2 },
  { code: '[>=100]0;[<0]-0.0;0.00', value: 5 },
  { code: '[Blue]0', value: 7 },
  { code: '@', value: 42 },
  { code: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)', value: 0 },
  // Dates and times.
  { code: 'yyyy-mm-dd', value: 45352 },
  { code: 'dddd, mmmm d, yyyy', value: 45352 },
  { code: 'mmm-yy', value: 45352 },
  { code: '[$-409]mmmm', value: 45352 },
  { code: 'h:mm AM/PM', value: 0.75 },
  { code: '[h]:mm:ss', value: 1.5 },
  { code: 'mm:ss.0', value: 0.00123 },
  { code: 'h:mm:ss', value: 0.999999 },
  { code: 'yyyy-mm-dd hh:mm', value: 59 },
  { code: 'yyyy-mm-dd hh:mm', value: 60 },
  { code: 'yyyy-mm-dd hh:mm', value: 61 },
  { code: 'yyyy-mm-dd', value: '2024-03-01', isoDate: true },
];

/** More cases of the 1900 date system: hours below 10, positive values of 37 to 40, serials that are no date, and a numbered color. */
export const FORMAT_CASES_MORE: readonly FormatCase[] = [
  { id: 20, value: 0.3 },
  { id: 21, value: 0.3 },
  { id: 22, value: 45352.3 },
  { id: 37, value: 1234 },
  { id: 38, value: 1234 },
  { id: 39, value: 1234.5 },
  { id: 40, value: 1234.5 },
  { code: 'yyyy-mm-dd', value: -1 },
  { code: 'yyyy-mm-dd', value: 3000000 },
  { code: '[h]:mm', value: 3000000 },
  { code: '[Color10]0', value: 8 },
];

/** Cases of the 1904 date system. */
export const FORMAT_CASES_1904: readonly FormatCase[] = [
  { id: 14, value: 0 },
  { id: 14, value: 1 },
  { code: 'yyyy-mm-dd', value: 1 },
  { code: 'yyyy-mm-dd', value: 45000 },
  { code: 'yyyy-mm-dd hh:mm', value: 1.5 },
  { code: 'yyyy-mm-dd', value: '2024-03-01', isoDate: true },
];

/** Rows that fit one Letter page at 15 points each, with the D8 margins. */
const ROWS_PER_BLOCK = 45;

/** The cell of case `i`: blocks of 45 rows, each block in its own column pair. */
export function caseCell(i: number): { label: string; ref: string } {
  const block = Math.floor(i / ROWS_PER_BLOCK);
  const r = (i % ROWS_PER_BLOCK) + 1;
  const label = String.fromCharCode(65 + block * 3);
  const value = String.fromCharCode(66 + block * 3);
  return { label: `${label}${r}`, ref: `${value}${r}` };
}

/** The styles part and rows of a sheet of format cases. */
export function formatSheet(cases: readonly FormatCase[]): {
  numFmts: string[];
  cellXfs: string[];
  rows: string;
  cols: string;
} {
  const codes = [...new Set(cases.flatMap((k) => (k.code === undefined ? [] : [k.code])))];
  const ids = [
    ...new Set(cases.map((k) => (k.code === undefined ? k.id! : 164 + codes.indexOf(k.code)))),
  ];
  const esc = (s: string) =>
    s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const numFmts = codes.map(
    (code, i) => `<numFmt numFmtId="${164 + i}" formatCode="${esc(code)}"/>`
  );
  const cellXfs = ids.map(
    (id) =>
      `<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`
  );
  const byRow = new Map<number, string[]>();
  cases.forEach((k, i) => {
    const id = k.code === undefined ? k.id! : 164 + codes.indexOf(k.code);
    const s = ids.indexOf(id) + 1;
    const { label, ref } = caseCell(i);
    const r = Number(ref.slice(1));
    const cells = byRow.get(r) ?? [];
    cells.push(c(label, k.code ?? `id ${k.id}`));
    if (k.isoDate) cells.push(c(ref, String(k.value), { t: 'd', s }));
    else if (typeof k.value === 'string') cells.push(c(ref, k.value, { s }));
    else cells.push(c(ref, k.value, { s }));
    byRow.set(r, cells);
  });
  const rows = [...byRow].map(([r, cells]) => row(r, cells)).join('');
  const blocks = Math.ceil(cases.length / ROWS_PER_BLOCK);
  const cols = Array.from({ length: blocks }, (_, b) => {
    const at = b * 3 + 1;
    return (
      `<col min="${at}" max="${at}" width="18" customWidth="1"/>` +
      `<col min="${at + 1}" max="${at + 1}" width="24" customWidth="1"/>` +
      `<col min="${at + 2}" max="${at + 2}" width="3" customWidth="1"/>`
    );
  }).join('');
  return { numFmts, cellXfs, rows, cols: `<cols>${cols}</cols>` };
}

/** One color or font case: a cell with a fill, a font, and a border. */
export interface StyleCase {
  readonly ref: string;
  /** What the case shows, written in the label cell. */
  readonly label: string;
  /** `x:fill` XML, or undefined for no fill. */
  readonly fill?: string;
  /** `x:font` XML, or undefined for the default font. */
  readonly font?: string;
  /** `x:border` XML, or undefined for no border. */
  readonly border?: string;
  /** Extra attributes of the `x:xf`. */
  readonly xf?: string;
}

const solid = (color: string) =>
  `<fill><patternFill patternType="solid"><fgColor ${color}/><bgColor indexed="64"/></patternFill></fill>`;
const fontWith = (inner: string) => `<font>${inner}<sz val="11"/><family val="2"/></font>`;

/** Fills, font colors, theme fonts, and borders, one cell each, in column B. */
export const STYLE_CASES: readonly StyleCase[] = [
  ...Array.from({ length: 10 }, (_, t) => ({
    ref: `B${t + 1}`,
    label: `fill theme ${t}`,
    fill: solid(`theme="${t}"`),
  })),
  ...[-0.5, -0.25, 0.25, 0.4, 0.5, 0.6, 0.8, 0.9].map((tint, i) => ({
    ref: `B${i + 11}`,
    label: `fill theme 4 tint ${tint}`,
    fill: solid(`theme="4" tint="${tint}"`),
  })),
  { ref: 'B19', label: 'fill indexed 10', fill: solid('indexed="10"') },
  { ref: 'B20', label: 'fill indexed 22', fill: solid('indexed="22"') },
  { ref: 'B21', label: 'fill rgb FF00B050', fill: solid('rgb="FF00B050"') },
  { ref: 'B22', label: 'fill rgb alpha 00', fill: solid('rgb="00FF0000"') },
  { ref: 'B23', label: 'fill auto', fill: solid('auto="1"') },
  {
    ref: 'B24',
    label: 'font color theme 4',
    font: fontWith('<color theme="4"/><name val="Carlito"/><scheme val="minor"/>'),
  },
  {
    ref: 'B25',
    label: 'font color indexed 10',
    font: fontWith('<color indexed="10"/><name val="Carlito"/><scheme val="minor"/>'),
  },
  {
    ref: 'B26',
    label: 'font color auto',
    font: fontWith('<color auto="1"/><name val="Carlito"/><scheme val="minor"/>'),
  },
  {
    ref: 'B27',
    label: 'font minor scheme, other name',
    font: fontWith('<color theme="1"/><name val="Courier New"/><scheme val="minor"/>'),
  },
  {
    ref: 'B28',
    label: 'font major scheme, other name',
    font: fontWith('<color theme="1"/><name val="Courier New"/><scheme val="major"/>'),
  },
  {
    ref: 'B29',
    label: 'bold font without applyFont',
    font: fontWith('<b/><color rgb="FFFF0000"/><name val="Carlito"/>'),
    xf: 'noApply',
  },
  ...[-0.5, 0.4, 0.8].map((tint, i) => ({
    ref: `B${i + 31}`,
    label: `fill theme 9 tint ${tint}`,
    fill: solid(`theme="9" tint="${tint}"`),
  })),
  { ref: 'B34', label: 'fill rgb 1F3864 tint 0.6', fill: solid('rgb="FF1F3864" tint="0.6"') },
  {
    ref: 'B30',
    label: 'border thin bottom indexed 64',
    border:
      '<border><left/><right/><top/><bottom style="thin"><color indexed="64"/></bottom><diagonal/></border>',
  },
];

/** The styles part and rows of the style case sheet. */
export function styleSheet(): {
  fonts: string[];
  fills: string[];
  borders: string[];
  cellXfs: string[];
  rows: string;
} {
  const fonts: string[] = [];
  const fills: string[] = [];
  const borders: string[] = [];
  const cellXfs: string[] = [];
  const rows: [number, string][] = [];
  STYLE_CASES.forEach((k, i) => {
    // The defaults: font 0, fills 0 and 1 (none and gray125), border 0.
    let fontId = 0;
    let fillId = 0;
    let borderId = 0;
    if (k.font) fontId = fonts.push(k.font);
    if (k.fill) fillId = fills.push(k.fill) + 1;
    if (k.border) borderId = borders.push(k.border);
    const apply =
      k.xf === 'noApply'
        ? ''
        : `${k.font ? ' applyFont="1"' : ''}${k.fill ? ' applyFill="1"' : ''}${k.border ? ' applyBorder="1"' : ''}`;
    cellXfs.push(
      `<xf numFmtId="0" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"${apply}/>`
    );
    const r = Number(k.ref.slice(1));
    rows.push([r, row(r, [c(`A${r}`, k.label), c(k.ref, 'Sample', { s: i + 1 })])]);
  });
  // Rows go in row order, whatever the order of the cases.
  rows.sort((a, b) => a[0] - b[0]);
  return { fonts, fills, borders, cellXfs, rows: rows.map(([, xml]) => xml).join('') };
}
