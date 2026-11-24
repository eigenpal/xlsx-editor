// The first sheets of the sample workbook: the welcome sheet, a sales table with formulas and
// frozen panes, and the number formats. All names and figures are made up for the sample.

import type { FixtureSheet } from '../fixtures/builder.ts';
import {
  COLORS,
  THIN,
  columns,
  frozenPane,
  sheetRow,
  type CellStyle,
  type StringTable,
  type StyleTable,
  cellWriter,
  dateSerial,
} from './showcase-kit.ts';

/** What each sheet shows, in tab order. The welcome sheet lists them. */
export const SHEET_GUIDE: readonly (readonly [string, string])[] = [
  ['Welcome', 'This page'],
  ['Sales', 'Formulas with cached results, frozen panes, borders, and fills'],
  ['Formats', 'Number, currency, percent, fraction, date, and time formats'],
  ['Text', 'Fonts, rich text, alignment, wrap, rotation, and merged cells'],
  ['Layout', 'Column widths, row heights, hidden rows and columns, a picture, and shapes'],
  ['Schedule', 'Dates, date formulas, a hidden row, and a long list to scroll'],
];

/** The tips of the Welcome sheet, from row 6. */
const TIPS = [
  'Select a cell and type to replace its value. Press Enter to keep the edit.',
  'Use the toolbar to change the font, fill, borders, alignment, and number format.',
  'Insert and delete rows and columns from the toolbar.',
  'Undo and redo each edit from the toolbar or the keyboard.',
  'Open the other sheets with the tabs under the grid.',
  'Select Save at the top of the page, or in the toolbar, to write the edited workbook to a file. If your browser cannot ask where to save, Save downloads the file.',
  'Formulas show stored results. After an edit, a notice says that they can be out of date.',
];

/** The first row of the sheet guide on the Welcome sheet, under its heading and header row. */
export const GUIDE_FIRST_ROW = 6 + TIPS.length + 3;

/** Welcome: what the sample holds and what to try. */
export function buildWelcomeSheet(styles: StyleTable, strings: StringTable): FixtureSheet {
  const c = cellWriter(strings);
  const title = styles.xf({ font: { size: 24, bold: true, color: COLORS.green } });
  const lead = styles.xf({ font: { size: 12, italic: true, color: COLORS.muted } });
  const heading = styles.xf({ font: { size: 13, bold: true } });
  const tip = styles.xf({ align: { indent: 1 } });
  const head: CellStyle = {
    font: { bold: true, color: COLORS.white },
    fill: COLORS.green,
    border: THIN,
    align: { vertical: 'center' },
  };
  const th = styles.xf(head);
  const td = styles.xf({ border: THIN, align: { vertical: 'center' } });
  const tdName = styles.xf({ border: THIN, font: { bold: true }, align: { vertical: 'center' } });
  const rows = [
    sheetRow(2, [c('B2', 'XLSX Editor', title), c('C2', undefined, title)], { height: 36 }),
    sheetRow(3, [
      c('B3', 'A sample workbook for the demo. Every name and figure in it is made up.', lead),
    ]),
    sheetRow(5, [c('B5', 'What to try', heading)], { height: 20 }),
    ...TIPS.map((t, i) => sheetRow(6 + i, [c(`B${6 + i}`, `${i + 1}.  ${t}`, tip)])),
    sheetRow(
      GUIDE_FIRST_ROW - 2,
      [c(`B${GUIDE_FIRST_ROW - 2}`, 'What each sheet shows', heading)],
      {
        height: 20,
      }
    ),
    sheetRow(
      GUIDE_FIRST_ROW - 1,
      [
        c(`B${GUIDE_FIRST_ROW - 1}`, 'Sheet', th),
        c(`C${GUIDE_FIRST_ROW - 1}`, 'What it shows', th),
      ],
      { height: 20 }
    ),
    ...SHEET_GUIDE.map(([name, text], i) =>
      sheetRow(
        GUIDE_FIRST_ROW + i,
        [c(`B${GUIDE_FIRST_ROW + i}`, name, tdName), c(`C${GUIDE_FIRST_ROW + i}`, text, td)],
        {
          height: 18,
        }
      )
    ),
  ];
  return {
    name: 'Welcome',
    cols: columns([
      [1, 1, 3],
      [2, 2, 16],
      [3, 3, 72],
    ]),
    rows: rows.join(''),
    // The title spans B:C, so the selection outline frames all of it.
    merges: ['B2:C2'],
    headings: true,
    // Open on the title, which spans B2:C2, so the selection frames all of it.
    viewChildren: '<selection activeCell="B2" sqref="B2:C2"/>',
  };
}

/** Synthetic quarterly figures for ten regions, in whole dollars. */
const REGIONS: readonly (readonly [string, number, number, number, number])[] = [
  ['North', 48_200, 51_900, 55_400, 60_100],
  ['South', 39_800, 37_600, 41_200, 44_900],
  ['East', 52_300, 54_000, 53_100, 57_800],
  ['West', 61_700, 64_200, 66_900, 71_300],
  ['Central', 28_400, 30_100, 29_700, 33_200],
  ['Coast', 35_900, 38_800, 42_600, 40_100],
  ['Highlands', 18_200, 19_600, 21_300, 22_800],
  ['Lakes', 24_700, 23_900, 25_100, 27_400],
  ['Plains', 31_500, 33_800, 30_200, 29_800],
  ['Valley', 22_100, 24_300, 26_800, 29_500],
];

/** The first and the last region row of the Sales sheet. The Total row follows. */
export const SALES_FIRST_ROW = 5;
export const SALES_LAST_ROW = SALES_FIRST_ROW + REGIONS.length - 1;

/** Sales: a table whose Total, Share, and Change columns and Total row are formulas. */
export function buildSalesSheet(styles: StyleTable, strings: StringTable): FixtureSheet {
  const c = cellWriter(strings);
  const money = '"$"#,##0';
  const title = styles.xf({
    font: { size: 16, bold: true, color: COLORS.green },
    fill: COLORS.greenLight,
    align: { vertical: 'center', indent: 1 },
  });
  // A white fill over the merged note hides the gridlines behind its text.
  const note = styles.xf({
    font: { italic: true, color: COLORS.muted },
    fill: COLORS.white,
    align: { indent: 1 },
  });
  const head = (horizontal: string) =>
    styles.xf({
      font: { bold: true, color: COLORS.white },
      fill: COLORS.green,
      border: THIN,
      align: { horizontal, vertical: 'center' },
    });
  const body = (band: boolean, extra: CellStyle = {}) =>
    styles.xf({ border: THIN, ...(band ? { fill: COLORS.band } : {}), ...extra });
  const total: CellStyle = {
    font: { bold: true },
    border: { left: 'thin', right: 'thin', top: 'thin', bottom: 'double' },
  };
  const first = SALES_FIRST_ROW;
  const last = SALES_LAST_ROW;
  const totalRow = last + 1;
  const grand = REGIONS.reduce((n, [, a, b, c2, d]) => n + a + b + c2 + d, 0);
  const rows: string[] = [
    sheetRow(
      1,
      [
        c('A1', 'Quarterly sales by region', title),
        ...['B', 'C', 'D', 'E', 'F', 'G', 'H'].map((col) => c(`${col}1`, undefined, title)),
      ],
      { height: 30 }
    ),
    sheetRow(2, [
      c('A2', 'Made-up figures in US dollars. Total, Share, and Change are formulas.', note),
      ...['B', 'C', 'D', 'E', 'F', 'G', 'H'].map((col) => c(`${col}2`, undefined, note)),
    ]),
    sheetRow(
      4,
      ['Region', 'Q1', 'Q2', 'Q3', 'Q4', 'Total', 'Share', 'Change'].map((text, i) =>
        c(`${'ABCDEFGH'[i]}4`, text, head(i === 0 ? 'left' : 'center'))
      ),
      { height: 20 }
    ),
  ];
  REGIONS.forEach(([name, ...q], i) => {
    const r = first + i;
    const band = i % 2 === 1;
    const sum = q.reduce((n, v) => n + v, 0);
    rows.push(
      sheetRow(r, [
        c(`A${r}`, name, body(band)),
        ...q.map((v, k) => c(`${'BCDE'[k]}${r}`, v, body(band, { numFmt: money }))),
        c(
          `F${r}`,
          { f: `SUM(B${r}:E${r})`, v: sum },
          body(band, { numFmt: money, font: { bold: true } })
        ),
        c(`G${r}`, { f: `F${r}/$F$${totalRow}`, v: sum / grand }, body(band, { numFmt: '0.0%' })),
        c(
          `H${r}`,
          { f: `(E${r}-B${r})/B${r}`, v: (q[3] - q[0]) / q[0] },
          body(band, { numFmt: '+0.0%;[Red]-0.0%' })
        ),
      ])
    );
  });
  const colTotal = (k: number) => REGIONS.reduce((n, row) => n + (row[k + 1] as number), 0);
  rows.push(
    sheetRow(totalRow, [
      c(`A${totalRow}`, 'Total', styles.xf(total)),
      ...[0, 1, 2, 3].map((k) =>
        c(
          `${'BCDE'[k]}${totalRow}`,
          { f: `SUM(${'BCDE'[k]}${first}:${'BCDE'[k]}${last})`, v: colTotal(k) },
          styles.xf({ ...total, numFmt: money })
        )
      ),
      c(
        `F${totalRow}`,
        { f: `SUM(F${first}:F${last})`, v: grand },
        styles.xf({ ...total, numFmt: money })
      ),
      c(
        `G${totalRow}`,
        {
          f: `SUM(G${first}:G${last})`,
          // The sum of the shares as the formula adds them, which can differ from 1 in the last bit.
          v: REGIONS.reduce((n, [, a, b, c2, d]) => n + (a + b + c2 + d) / grand, 0),
        },
        styles.xf({ ...total, numFmt: '0.0%' })
      ),
      c(
        `H${totalRow}`,
        {
          f: `(E${totalRow}-B${totalRow})/B${totalRow}`,
          v: (colTotal(3) - colTotal(0)) / colTotal(0),
        },
        styles.xf({ ...total, numFmt: '+0.0%;[Red]-0.0%' })
      ),
    ])
  );
  return {
    name: 'Sales',
    cols: columns([
      [1, 1, 16],
      [2, 5, 12],
      [6, 6, 13],
      [7, 8, 11],
    ]),
    rows: rows.join(''),
    // The note of row 2 sits in a merged range with a white fill, so no gridline crosses it.
    merges: ['A1:H1', 'A2:H2'],
    gridLines: true,
    headings: true,
    // Rows only: a frozen column would cut the merged title on a sideways scroll.
    viewChildren: frozenPane(4, 0),
  };
}

/** Formats: one value per row, the format code, and how the cell shows it. */
export function buildFormatsSheet(styles: StyleTable, strings: StringTable): FixtureSheet {
  const c = cellWriter(strings);
  const head = styles.xf({
    font: { bold: true, color: COLORS.white },
    fill: COLORS.green,
    border: { bottom: 'thin' },
  });
  const group = styles.xf({
    font: { bold: true, color: COLORS.green },
    border: { bottom: 'thin' },
  });
  // Centered, so that "Shows as" and "Format code" do not read as one phrase.
  const headCenter = styles.xf({
    font: { bold: true, color: COLORS.white },
    fill: COLORS.green,
    border: { bottom: 'thin' },
    align: { horizontal: 'center' },
  });
  const code = styles.xf({
    font: { name: 'Courier New', size: 10, color: COLORS.muted },
    align: { indent: 1 },
  });
  const date = dateSerial(2026, 3, 14);
  const time = 0.6875; // 16:30
  const sections: readonly (readonly [
    string,
    readonly (readonly [string, number | string, number | boolean | string | { error: string }])[],
  ])[] = [
    [
      'Numbers',
      [
        ['General', 'General', 1234.5678],
        ['Two decimals', '0.00', 1234.5678],
        ['Thousands separator', '#,##0', 1234567.891],
        ['Thousands and decimals', '#,##0.00', -1234567.891],
        ['Negative in red', '#,##0.00;[Red]-#,##0.00', -42.5],
        ['Leading zeros', '000000', 4217],
        ['Scientific', '0.00E+00', 0.000123456],
        ['Fraction', '# ?/?', 3.75],
        ['Fraction in sixteenths', '# ??/16', 2.3125],
        ['Thousands as K', '#,##0,"K"', 1234567],
        ['Condition', '[<1000]0" small";#,##0" large"', 250],
      ],
    ],
    [
      'Money and percent',
      [
        ['Currency', '"$"#,##0.00', 1999.5],
        ['Currency, negative in parentheses', '"$"#,##0.00_);("$"#,##0.00)', -1999.5],
        ['Euro after the value', '#,##0.00" €"', 875.25],
        ['Percent', '0%', 0.256],
        ['Percent with one decimal', '0.0%', 0.0425],
        ['Signed percent', '+0.0%;-0.0%;0.0%', -0.031],
      ],
    ],
    [
      'Dates and times',
      [
        ['Short date', 14, date],
        ['ISO date', 'yyyy-mm-dd', date],
        ['Long date', 'dddd, mmmm d, yyyy', date],
        ['Day and month', 'd mmm', date],
        ['Month and year', 'mmmm yyyy', date],
        ['Time', 'h:mm AM/PM', time],
        ['Time with seconds', 'hh:mm:ss', time + 15 / 86_400],
        ['Date and time', 'yyyy-mm-dd hh:mm', date + time],
        ['Elapsed hours', '[h]:mm', 1.5 + time],
      ],
    ],
    [
      'Other values',
      [
        ['Text format', '@', '00123'],
        ['Boolean', 'General', true],
        ['Error', 'General', { error: '#DIV/0!' }],
      ],
    ],
  ];
  const rows: string[] = [
    sheetRow(
      1,
      [c('A1', 'Format', head), c('B1', 'Shows as', headCenter), c('C1', 'Format code', head)],
      {
        height: 20,
      }
    ),
  ];
  // The next free row. A blank row comes before each group but the first.
  let r = 2;
  for (const [title, items] of sections) {
    if (r > 2) r += 1;
    rows.push(
      sheetRow(r, [
        c(`A${r}`, title, group),
        c(`B${r}`, undefined, group),
        c(`C${r}`, undefined, group),
      ])
    );
    r += 1;
    for (const [label, fmt, value] of items) {
      const s = styles.xf({ numFmt: fmt === 'General' ? 0 : fmt, align: { horizontal: 'right' } });
      const shown = fmt === 14 ? 'Built-in format 14' : String(fmt);
      rows.push(sheetRow(r, [c(`A${r}`, label), c(`B${r}`, value, s), c(`C${r}`, shown, code)]));
      r += 1;
    }
  }
  return {
    name: 'Formats',
    cols: columns([
      [1, 1, 32],
      [2, 2, 28],
      [3, 3, 40],
    ]),
    rows: rows.join(''),
    gridLines: true,
    headings: true,
    // Open on the first value, not on a group heading.
    viewChildren: frozenPane(1, 0, 'B3'),
  };
}
