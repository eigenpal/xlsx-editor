// Typed number forms: the reading of each form and each refusal (`readCellEntry`),
// and the commit through the editor: the value as `x:v`, the number format on a General cell
// in the same transaction, one undo step, the Text format, and save and reopen.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { createXlsxEditor, type XlsxEditorInstance } from '../index.ts';
import { parseCellInput, readCellEntry, type TypedEntry } from '../cell-input.ts';
import { numberEditText } from '../typed-forms.ts';
import { pastedFormatOps } from '../entry-ops.ts';
import { MAX_STYLE_RECORDS, readXlsx, type WorkbookOp } from '../../store/index.ts';
import { flushEditor } from './flush-editor.ts';
import type { FontSource } from '../../layout/index.ts';
import { builtinFormatCode, EN_US, type NumberLocale } from '../../cascade/index.ts';

const fontsDir = join(
  dirname(Bun.resolveSync('@docx-editor.dev/fonts/package.json', import.meta.dir)),
  'assets'
);
const FONTS: FontSource[] = [
  {
    family: 'Carlito',
    bold: false,
    italic: false,
    bytes: new Uint8Array(readFileSync(join(fontsDir, 'Carlito-Regular.ttf'))),
    origin: 'substitute',
  },
];

/** A locale with `,` as the decimal separator, `.` as the group, and a day-first short date. */
const DE: NumberLocale = { decimal: ',', group: '.', builtins: { 14: 'dd.mm.yyyy' } };
/** A locale whose short date names no day and month order. */
const NO_ORDER: NumberLocale = { decimal: '.', group: ',', builtins: { 14: 'General' } };

const read = (text: string, locale = EN_US, date1904 = false): TypedEntry =>
  readCellEntry(text, locale, { date1904, year: 2026 });
const num = (value: number, numberFormat?: TypedEntry['numberFormat']): TypedEntry =>
  numberFormat === undefined
    ? { value: { type: 'number', value } }
    : { value: { type: 'number', value }, numberFormat };
const text = (value: string): TypedEntry => ({ value: { type: 'string', value } });
const builtin = (id: number): string => builtinFormatCode(id, EN_US)!;

describe('readCellEntry', () => {
  test('a percent is the double nearest to the typed decimal over 100, with format 9 or 10', () => {
    expect(read('12%')).toEqual(num(0.12, { id: 9 }));
    expect(read('12.5%')).toEqual(num(0.125, { id: 10 }));
    expect(read('-3%')).toEqual(num(-0.03, { id: 9 }));
    expect(read('1,200%')).toEqual(num(12, { id: 9 }));
    expect(read('12,5%', DE)).toEqual(num(0.125, { id: 10 }));
  });

  test('a currency amount is a number with a currency code', () => {
    expect(read('$5.50')).toEqual(num(5.5, { code: '"$"#,##0.00' }));
    expect(read('$1,234')).toEqual(num(1234, { code: '"$"#,##0' }));
    expect(read('-$5')).toEqual(num(-5, { code: '"$"#,##0' }));
    expect(read('$-5')).toEqual(num(-5, { code: '"$"#,##0' }));
    expect(read('($5.00)')).toEqual(num(-5, { code: '"$"#,##0.00' }));
    expect(read('£3')).toEqual(num(3, { code: '"£"#,##0' }));
    expect(read('5,50 €', DE)).toEqual(num(5.5, { code: '#,##0.00 "€"' }));
    expect(read('1.234€', DE)).toEqual(num(1234, { code: '#,##0 "€"' }));
  });

  test('a grouped number has format 3 or 4, in the locale separators', () => {
    expect(read('1,234.5')).toEqual(num(1234.5, { id: 4 }));
    expect(read('1,234,567')).toEqual(num(1234567, { id: 3 }));
    expect(read('-1,234')).toEqual(num(-1234, { id: 3 }));
    expect(read('1.234,5', DE)).toEqual(num(1234.5, { id: 4 }));
  });

  test('a negative in parentheses is a negative number', () => {
    expect(read('(100)')).toEqual(num(-100));
    expect(read('(1,000)')).toEqual(num(-1000, { id: 3 }));
    expect(read('(12%)')).toEqual(num(-0.12, { id: 9 }));
  });

  test('a fraction with a whole part has format 12 or 13', () => {
    expect(read('1 1/2')).toEqual(num(1.5, { id: 12 }));
    expect(read('0 3/16')).toEqual(num(0.1875, { id: 13 }));
    expect(read('-2 1/4')).toEqual(num(-2.25, { id: 12 }));
    // The copied text of format 13 pads the numerator with a space.
    expect(read('1  1/16')).toEqual(num(1.0625, { id: 13 }));
  });

  test('a date is a serial in the date system of the workbook, with format 14', () => {
    expect(read('2026-11-12')).toEqual(num(46338, { id: 14 }));
    expect(read('2026/11/12')).toEqual(num(46338, { id: 14 }));
    expect(read('11/12/2026')).toEqual(num(46338, { id: 14 }));
    expect(read('11-12-26')).toEqual(num(46338, { id: 14 }));
    expect(read('12.11.2026', DE)).toEqual(num(46338, { id: 14 }));
    expect(read('12/11/2026', DE)).toEqual(num(46338, { id: 14 }));
    // The 1904 system counts from 1904-01-01 (Part 1, 18.17.4.1).
    expect(read('2026-11-12', EN_US, true)).toEqual(num(46338 - 1462, { id: 14 }));
    // Before 1900-03-01 the 1900 system counts one day less.
    expect(read('1900-01-01')).toEqual(num(1, { id: 14 }));
    expect(read('1/1/30')).toEqual(num(10959, { id: 14 }));
  });

  test('a day and a month with no year is a date in the given year, with format 16', () => {
    expect(read('11/12')).toEqual(num(46338, { id: 16 }));
    expect(read('1/2')).toEqual(num(46024, { id: 16 }));
    expect(read('12/11', DE)).toEqual(num(46338, { id: 16 }));
    expect(read('11/12/2026', DE)).toEqual(num(46367, { id: 14 }));
  });

  test('a time is a fraction of a day, with format 18 to 21', () => {
    expect(read('13:45')).toEqual(num(0.5729166666666666, { id: 20 }));
    expect(read('13:45:30')).toEqual(num(49530 / 86400, { id: 21 }));
    expect(read('1:45 PM')).toEqual(num(0.5729166666666666, { id: 18 }));
    expect(read('12:00am')).toEqual(num(0, { id: 18 }));
    expect(read('12:00:01 PM')).toEqual(num(43201 / 86400, { id: 19 }));
  });

  test('a date and a time is the serial and the fraction, with format 22', () => {
    expect(read('2026-11-12 18:00')).toEqual(num(46338.75, { id: 22 }));
    // A 12-hour clock keeps AM or PM, as a typed time alone does.
    expect(read('11/12/2026 6:00 PM')).toEqual(num(46338.75, { code: 'm/d/yy h:mm AM/PM' }));
    // A format 22 that is not a plain date and one clock takes the short date code, so its copy
    // reads back as a date.
    for (const odd of [
      'dd/mm/yyyy [h]:mm',
      'dd/mm/yyyy h:mm;[Red]dd/mm/yyyy h:mm',
      'yyyy/dd/mm',
      '[ss] dd/mm/yyyy h:mm',
      'dd/mm/yyyy "at" h:mm',
      'dd.mm.yyyy hh:mm',
      'dd/mm/yyyy "x h:mm',
      'dd/mm/yyyy AM/PM h:mm',
    ]) {
      const locale: NumberLocale = {
        decimal: '.',
        group: ',',
        builtins: { 14: 'dd/mm/yyyy', 22: odd },
      };
      expect(read('12/11/2026 18:00', locale).numberFormat).toEqual({ code: 'dd/mm/yyyy h:mm' });
      expect(read('12/11/2026 18:00:30', locale).numberFormat).toEqual({
        code: 'dd/mm/yyyy h:mm:ss',
      });
      expect(read('12/11/2026 6:00 PM', locale).numberFormat).toEqual({
        code: 'dd/mm/yyyy h:mm AM/PM',
      });
    }
    // Seconds follow the minutes of format 22, also on a 12-hour clock or with a capital `H`.
    for (const [code22, withSeconds] of [
      ['m/d/yyyy h:mm AM/PM', 'm/d/yyyy h:mm:ss AM/PM'],
      ['m/d/yyyy H:mm', 'm/d/yyyy H:mm:ss'],
    ]) {
      const locale: NumberLocale = {
        decimal: '.',
        group: ',',
        builtins: { 14: 'm/d/yyyy', 22: code22! },
      };
      expect(read('11/12/2026 18:00:30', locale).numberFormat).toEqual({ code: withSeconds });
    }
  });

  test('a date and a time in a day-first locale keeps the order of its short date', () => {
    // Format 22 of the locale is month first, so the cell takes the short date and `h:mm`.
    expect(read('12.11.2026 18:00', DE)).toEqual(num(46338.75, { code: 'dd.mm.yyyy h:mm' }));
    const dayFirst22 = { ...DE, builtins: { ...DE.builtins, 22: 'dd.mm.yyyy hh:mm' } };
    expect(read('12.11.2026 18:00', dayFirst22)).toEqual(num(46338.75, { id: 22 }));
  });

  test('a date and a time in a year-first locale keeps the year first', () => {
    const iso: NumberLocale = { decimal: ',', group: ' ', builtins: { 14: 'yyyy-mm-dd' } };
    expect(read('2026-11-12 18:00', iso)).toEqual(num(46338.75, { code: 'yyyy-mm-dd h:mm' }));
    const slash: NumberLocale = { decimal: '.', group: ',', builtins: { 14: 'yyyy/m/d' } };
    expect(read('2026/11/12 18:00', slash)).toEqual(num(46338.75, { code: 'yyyy/m/d h:mm' }));
  });

  test('a date and a time with seconds keeps them in its format', () => {
    expect(read('2026-11-12 18:00:30')).toEqual(
      num(46338.75 + 30 / 86400, { code: 'm/d/yy h:mm:ss' })
    );
    expect(read('12.11.2026 18:00:30', DE)).toEqual(
      num(46338.75 + 30 / 86400, { code: 'dd.mm.yyyy h:mm:ss' })
    );
    // A locale with no date order reads back only ISO 8601.
    expect(read('2026-11-12 18:00:30', NO_ORDER)).toEqual(
      num(46338.75 + 30 / 86400, { code: 'yyyy-mm-dd h:mm:ss' })
    );
  });

  test('plain numbers and booleans keep their reading, with no format', () => {
    expect(read('1e3')).toEqual(num(1000));
    expect(read('-2.5')).toEqual(num(-2.5));
    expect(read('TRUE')).toEqual({ value: { type: 'boolean', value: true } });
  });

  test('a leading apostrophe keeps the rest as text, and the apostrophe is not stored', () => {
    expect(read("'12%")).toEqual(text('12%'));
    expect(read("'007")).toEqual(text('007'));
    expect(read("'2026-11-12")).toEqual(text('2026-11-12'));
  });

  test('a form that does not read whole stays text', () => {
    for (const t of [
      '12%%',
      '%',
      '$',
      '$5$',
      '$5%',
      '1,23',
      '1,2345',
      ',123',
      '1.234.5',
      '(5',
      '()',
      '--5',
      '1 /2',
      '1 1/0',
      '1 1/100',
      '1 5/2',
      'Jan 1',
      '2026-0011-0012',
      '0,123',
      '01,234',
      '2026-13-01',
      '2026-02-30',
      '1900-02-29',
      '1899-12-31',
      '1/2/3/4',
      '1/2-3',
      '12:60',
      '13:00 PM',
      '0:00 AM',
      '1:5',
      '2026-11-12T18:00',
      '2026-11-12 25:00',
      '#N/A',
      'abc',
      ' 4',
      '4 ',
      '١٢%',
    ])
      expect(read(t)).toEqual(text(t));
  });

  test('a locale with no day and month order reads only ISO 8601 dates', () => {
    expect(read('11/12/2026', NO_ORDER)).toEqual(text('11/12/2026'));
    expect(read('11/12', NO_ORDER)).toEqual(text('11/12'));
    // With no date order, a date takes an ISO 8601 code, which reads back.
    expect(read('2026-11-12', NO_ORDER)).toEqual(num(46338, { code: 'yyyy-mm-dd' }));
    // A short date that is not a plain date has no order that reads back.
    const spaced: NumberLocale = { decimal: ',', group: ' ', builtins: { 14: 'yyyy. mm. dd.' } };
    expect(read('2026-11-12', spaced)).toEqual(num(46338, { code: 'yyyy-mm-dd' }));
    expect(read('2026-11-12 13:45', spaced).numberFormat).toEqual({ code: 'yyyy-mm-dd h:mm' });
  });

  test('a percent of up to 64 characters stores a finite number', () => {
    for (const t of [`${'1'.repeat(62)}%`, `-${'1'.repeat(61)}%`, `.${'1'.repeat(61)}%`]) {
      const v = read(t).value;
      expect(v.type).toBe('number');
      expect(Number.isFinite((v as { value: number }).value)).toBe(true);
    }
  });

  test('an elapsed time over 23 hours has `[h]:mm`, or format 46 with seconds', () => {
    expect(read('36:00')).toEqual(num(1.5, { code: '[h]:mm' }));
    expect(read('100:30:15')).toEqual(num((100 * 3600 + 30 * 60 + 15) / 86400, { id: 46 }));
    for (const t of ['13:00 PM', '10000:00', '2026-11-12 25:00']) expect(read(t)).toEqual(text(t));
  });

  test('the parse is bounded: text over 64 characters stays text', () => {
    const long = `${'1,'.repeat(5000)}000%`;
    expect(read(long)).toEqual(text(long));
    expect(read(`$${'9'.repeat(64)}`)).toEqual(text(`$${'9'.repeat(64)}`));
    expect(read(`${'1/'.repeat(5000)}1`)).toEqual(text(`${'1/'.repeat(5000)}1`));
  });
});

describe('readCellEntry: precision, separators, and defaults', () => {
  test('a percent with decimals is the double nearest to the typed decimal', () => {
    for (let i = 0; i <= 2000; i += 1) {
      const typed = (i / 10).toFixed(1);
      expect(read(`${typed}%`).value).toEqual({ type: 'number', value: Number(`${typed}e-2`) });
    }
    expect(read('1.1%')).toEqual(num(0.011, { id: 10 }));
    expect(read('0.07%')).toEqual(num(0.0007, { id: 10 }));
    expect(read('-1.1%')).toEqual(num(-0.011, { id: 10 }));
  });

  test('a `.` separates a date only in a locale whose short date uses `.`', () => {
    for (const t of ['1.2.10', '1.2.2024', '3.14.15']) expect(read(t)).toEqual(text(t));
    expect(read('1.2.10', DE)).toEqual(num(40210, { id: 14 }));
  });

  test('the short date order skips quoted, bracketed, and escaped letters', () => {
    const bracketed: NumberLocale = { ...EN_US, builtins: { 14: '[$-x-sysdate]dd/mm/yyyy' } };
    expect(read('12/11/2026', bracketed)).toEqual(num(46338, { id: 14 }));
    const elapsed: NumberLocale = { ...EN_US, builtins: { 14: '[h]dd/mm/yyyy' } };
    expect(read('12/11/2026', elapsed)).toEqual(text('12/11/2026'));
    // An escaped letter shows, so the short date does not read back: only ISO 8601 is a date.
    const escaped: NumberLocale = { ...EN_US, builtins: { 14: '\\y dd/mm/yyyy' } };
    expect(read('12/11/2026', escaped)).toEqual(text('12/11/2026'));
    expect(read('2026-11-12', escaped)).toEqual(num(46338, { code: 'yyyy-mm-dd' }));
  });

  test('a currency symbol can follow a no-break space, and ¥ is a currency', () => {
    expect(read('5,50 €', DE)).toEqual(num(5.5, { code: '#,##0.00 "€"' }));
    expect(read('5,50 €', DE)).toEqual(num(5.5, { code: '#,##0.00 "€"' }));
    expect(read('¥1,000')).toEqual(num(1000, { code: '"¥"#,##0' }));
  });

  test('a date with no year takes the year of the host clock by default', () => {
    // Read the year before and after, so a run across New Year passes.
    const before = new Date().getFullYear();
    const got = readCellEntry('3/15', EN_US);
    const after = new Date().getFullYear();
    const as = (year: number): TypedEntry => ({
      ...readCellEntry(`${year}-03-15`, EN_US),
      numberFormat: { id: 16 },
    });
    expect([as(before), as(after)]).toContainEqual(got);
  });
});

describe('numberEditText', () => {
  const edit = (value: number, code: string, locale = EN_US) =>
    numberEditText(value, code, locale, false);

  test('a date, a time, or a percent shows as it is typed, and reads back to its value', () => {
    for (const [typed, shown] of [
      ['2026-11-12', '2026-11-12'],
      ['11/12/2026', '2026-11-12'],
      ['13:45', '13:45'],
      ['1:45 PM', '13:45'],
      ['13:45:30', '13:45:30'],
      ['2026-11-12 18:00', '2026-11-12 18:00'],
      ['12%', '12%'],
      ['1.1%', '1.1%'],
      ['-3%', '-3%'],
    ] as const) {
      const entry = read(typed);
      const id = entry.numberFormat?.id;
      const code = id !== undefined ? builtin(id) : '';
      const value = (entry.value as { value: number }).value;
      expect(edit(value, code)).toBe(shown);
      expect(read(shown).value).toEqual(entry.value);
    }
    expect(edit(0.125, '0.00%', DE)).toBe('12,5%');
  });

  test('in the 1904 date system, serial 0 is a date', () => {
    expect(numberEditText(0, 'mm-dd-yy', EN_US, true)).toBe('1904-01-01');
    expect(numberEditText(0.5, 'm/d/yy h:mm', EN_US, true)).toBe('1904-01-01 12:00');
    expect(numberEditText(0.5, 'h:mm', EN_US, true)).toBe('12:00');
    expect(numberEditText(46338 - 1462, 'mm-dd-yy', EN_US, true)).toBe('2026-11-12');
  });

  test('an elapsed-time format shows the hours, and only the first section counts', () => {
    expect(edit(1.5, '[h]:mm')).toBe('36:00');
    expect(edit(1.5, '[mm]:ss')).toBe('36:00');
    expect(edit(0.5, '[h]:mm:ss')).toBe('12:00');
    expect(edit(5, '0;0%')).toBe('5');
    expect(edit(-0.05, '0;0%')).toBe('-0.05');
  });

  test('a value just under midnight is the next day, or 24:00 in a time format', () => {
    expect(edit(46338.99999999, 'mm-dd-yy')).toBe('2026-11-13');
    expect(edit(0.99999999, 'h:mm')).toBe('24:00');
  });

  test('the edit text rounds what the cell does not show', () => {
    expect(edit(0.30000000000000004, '0%')).toBe('30%');
    expect(edit(1 / 3, '0%')).toBe('33.3333333333333%');
    expect(edit(46338.123456, 'm/d/yy h:mm')).toBe('2026-11-12 2:57:47');
    expect(edit(0.5000001, 'h:mm')).toBe('12:00');
  });

  test('a format that shows fractions of a second keeps them: the plain number', () => {
    expect(edit(62.35 / 86400, 'h:mm:ss.00')).toBe(String(62.35 / 86400));
    expect(edit(62.55 / 86400, builtin(47))).toBe(String(62.55 / 86400));
    expect(edit(62.5 / 86400, '[ss].0')).toBe(String(62.5 / 86400));
    // The digits come from the code as the cell reads it: a literal can stand before them.
    expect(edit(62.4 / 86400, 'mm:ss .0')).toBe(String(62.4 / 86400));
    // A whole second far in the future keeps the time form, as the cell shows `.0`.
    expect(edit(2900000 + 1 / 86400, 'yyyy-mm-dd hh:mm:ss.0')).toBe('9839-12-05 0:00:01');
    // Whole seconds lose nothing, so they keep the time form.
    expect(edit(62 / 86400, builtin(47))).toBe('0:01:02');
    expect(edit(0.5, 'h:mm:ss.000')).toBe('12:00');
  });

  test('a value that would not read back shows the plain number', () => {
    expect(edit(60, 'mm-dd-yy')).toBe('60');
    expect(edit(1e7, 'mm-dd-yy')).toBe('10000000');
    expect(edit(46338.5, '0.00')).toBe('46338.5');
    expect(edit(1.5, 'General')).toBe('1.5');
    // A fraction that the plain number shows is not a fraction form.
    expect(edit(1.5, '# ?/?')).toBe('1.5');
    expect(edit(0.5, '"%"0')).toBe('0.5');
    // The character after `_` or `*` is a literal, not a part of the code.
    expect(edit(5, '0_%')).toBe('5');
    expect(edit(5, '0*%')).toBe('5');
  });
});

describe('pastedFormatOps', () => {
  const pkg = (() => {
    const r = readXlsx(book());
    if (!r.ok) throw new Error('read');
    return r.pkg;
  })();
  const cells = (refs: readonly [number, number][], id = 9) =>
    refs.map(([row, col]) => ({ at: { row, col }, numberFormat: { id } }));
  const refs = (ops: readonly WorkbookOp[]) => ops.map((o) => ('ref' in o ? o.ref : undefined));

  test('a block or a column of one format is one op', () => {
    const block: [number, number][] = [];
    for (let r = 0; r < 100; r += 1) for (let c = 0; c < 100; c += 1) block.push([r, c]);
    expect(refs(pastedFormatOps(pkg, 'S', cells(block)))).toEqual(['A1:CV100']);
    expect(
      refs(
        pastedFormatOps(
          pkg,
          'S',
          cells([
            [2, 1],
            [3, 1],
            [4, 1],
          ])
        )
      )
    ).toEqual(['B3:B5']);
  });

  test('cells of other formats or other spans are other ops', () => {
    const ops = pastedFormatOps(pkg, 'S', [
      ...cells([
        [0, 0],
        [1, 0],
        [0, 1],
      ]),
      ...cells([[1, 1]], 14),
    ]);
    expect(refs(ops)).toEqual(['A1:A2', 'B1', 'B2']);
  });
});

// ------------------------------------------------------------------ through the editor

let viewers: XlsxEditorInstance[] = [];
afterEach(() => {
  for (const v of viewers) v.destroy();
  viewers = [];
});

async function mount(bytes: Uint8Array) {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.appendChild(host);
  const viewer = createXlsxEditor({ fonts: FONTS, translate: (k) => k });
  viewers.push(viewer);
  viewer.attach(host);
  const scroller = host.querySelector('.xlsx-scroller') as HTMLElement;
  Object.defineProperty(scroller, 'clientWidth', { get: () => 800 });
  Object.defineProperty(scroller, 'clientHeight', { get: () => 400 });
  await viewer.load(bytes);
  flushEditor(viewer);
  return viewer;
}

/** A sheet: A1 General, B1 with format 2 (`0.00`), C1 with the Text format (`@`). */
const book = (date1904 = false) =>
  buildXlsx({
    sheets: [
      {
        name: 'S',
        rows: row(1, [c('A1', 1), '<c r="B1" s="1"><v>2</v></c>', '<c r="C1" s="2"/>']),
      },
    ],
    styles: {
      cellXfs: [
        '<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
        '<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
      ],
    },
    date1904,
  });

/** The saved `x:c` of `ref` and the number format of its cell format: id and code. */
function savedCell(bytes: Uint8Array, ref: string) {
  const files = unzipSync(bytes);
  const sheet = strFromU8(files['xl/worksheets/sheet1.xml']!);
  const styles = strFromU8(files['xl/styles.xml']!);
  const cell = new RegExp(`<c r="${ref}"[^>]*?(?:/>|>.*?</c>)`).exec(sheet)?.[0] ?? '';
  const s = Number(/ s="(\d+)"/.exec(cell)?.[1] ?? '0');
  const xfs = /<cellXfs[^>]*>(.*?)<\/cellXfs>/.exec(styles)?.[1] ?? '';
  const xf = xfs.match(/<xf [^>]*?(?:\/>|>.*?<\/xf>)/g)?.[s] ?? '';
  const id = Number(/numFmtId="(\d+)"/.exec(xf)?.[1] ?? '0');
  const code = new RegExp(`<numFmt numFmtId="${id}" formatCode="([^"]*)"`).exec(styles)?.[1];
  return { cell, id, code };
}

describe('typed entry in the editor', () => {
  test('an entry at another cell sets its format there, and undo keeps the selection', async () => {
    const viewer = await mount(book());
    viewer.select({ row: 0, col: 0 });
    expect(viewer.enterValue('12%', { row: 3, col: 0 }).ok).toBe(true);
    const saved = savedCell(viewer.save()!, 'A4');
    expect(saved.cell).toContain('<v>0.12</v>');
    expect(saved.id).toBe(9);
    expect(viewer.snapshot().selectionRef).toBe('A1');
    expect(viewer.undo().ok).toBe(true);
    expect(savedCell(viewer.save()!, 'A4').cell).toBe('');
    expect(viewer.snapshot().selectionRef).toBe('A1');
  });

  test('a percent in a General cell stores a number and format 9 in one undo step', async () => {
    const viewer = await mount(book());
    viewer.select({ row: 1, col: 0 });
    expect(viewer.enterValue('12%').ok).toBe(true);
    const saved = savedCell(viewer.save()!, 'A2');
    expect(saved.cell).toContain('<v>0.12</v>');
    expect(saved.cell).not.toContain('t="s"');
    expect(saved.id).toBe(9);
    // One undo restores the value and the format.
    expect(viewer.undo().ok).toBe(true);
    expect(viewer.snapshot().canUndo).toBe(false);
    expect(savedCell(viewer.save()!, 'A2').cell).toBe('');
  });

  test('a date takes format 14, and save and reopen keep the value and the format', async () => {
    const viewer = await mount(book());
    viewer.select({ row: 0, col: 0 });
    expect(viewer.enterValue('2026-11-12').ok).toBe(true);
    const bytes = viewer.save()!;
    expect(savedCell(bytes, 'A1').cell).toContain('<v>46338</v>');
    expect(savedCell(bytes, 'A1').id).toBe(14);
    const reopened = await mount(bytes);
    reopened.select({ row: 0, col: 0 });
    expect(reopened.snapshot().activeCell?.content).toBe('2026-11-12');
    const again = savedCell(reopened.save()!, 'A1');
    expect(again.cell).toContain('<v>46338</v>');
    expect(again.id).toBe(14);
  });

  test('a date uses the 1904 date system of the workbook', async () => {
    const viewer = await mount(book(true));
    viewer.select({ row: 0, col: 0 });
    expect(viewer.enterValue('2026-11-12').ok).toBe(true);
    expect(savedCell(viewer.save()!, 'A1').cell).toContain(`<v>${46338 - 1462}</v>`);
  });

  test('a currency amount writes its custom code', async () => {
    const viewer = await mount(book());
    viewer.select({ row: 0, col: 0 });
    expect(viewer.enterValue('$5.50').ok).toBe(true);
    const saved = savedCell(viewer.save()!, 'A1');
    expect(saved.cell).toContain('<v>5.5</v>');
    expect(saved.id).toBeGreaterThanOrEqual(164);
    expect(saved.code).toBe('&quot;$&quot;#,##0.00');
  });

  test('a cell with another number format keeps it', async () => {
    const viewer = await mount(book());
    viewer.select({ row: 0, col: 1 });
    expect(viewer.enterValue('12%').ok).toBe(true);
    const saved = savedCell(viewer.save()!, 'B1');
    expect(saved.cell).toContain('<v>0.12</v>');
    expect(saved.id).toBe(2);
  });

  test('a cell with the Text format keeps typed numbers, percents, and dates as text', async () => {
    const viewer = await mount(book());
    viewer.select({ row: 0, col: 2 });
    for (const typed of ['12%', '42', '2026-11-12', 'TRUE']) {
      expect(viewer.enterValue(typed).ok).toBe(true);
      const saved = savedCell(viewer.save()!, 'C1');
      expect(saved.cell).toContain('t="s"');
      expect(saved.id).toBe(49);
      expect(viewer.snapshot().activeCell?.content).toBe(typed);
    }
  });

  test('a leading apostrophe stores the text without the apostrophe', async () => {
    const viewer = await mount(book());
    viewer.select({ row: 0, col: 0 });
    expect(viewer.enterValue("'007").ok).toBe(true);
    const saved = savedCell(viewer.save()!, 'A1');
    expect(saved.cell).toContain('t="s"');
    expect(saved.id).toBe(0);
    expect(viewer.snapshot().activeCell?.content).toBe("'007");
  });

  test('a plain-text paste reads each cell as typed entry', async () => {
    const viewer = await mount(book());
    viewer.select({ row: 2, col: 0 });
    expect(viewer.paste({ text: '12%\t2026-11-12' }).ok).toBe(true);
    const bytes = viewer.save()!;
    expect(savedCell(bytes, 'A3').id).toBe(9);
    expect(savedCell(bytes, 'B3').id).toBe(14);
    expect(viewer.undo().ok).toBe(true);
    expect(savedCell(viewer.save()!, 'A3').cell).toBe('');
  });
});

/** A protected sheet whose A1 is an unlocked General cell. */
const protectedBook = () =>
  buildXlsx({
    sheets: [
      {
        name: 'S',
        rows: row(1, ['<c r="A1" s="1"/>']) + row(2, ['<c r="A2" s="1"/>']),
        beforeMerges: '<sheetProtection sheet="1"/>',
      },
    ],
    styles: {
      cellXfs: [
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1"><protection locked="0"/></xf>',
      ],
    },
  });

describe('typed entry: protection, Text cells, and re-edit', () => {
  test('an unlocked cell of a protected sheet stores the number and keeps its format', async () => {
    const viewer = await mount(protectedBook());
    viewer.select({ row: 0, col: 0 });
    expect(viewer.enterValue('12%')).toEqual({ ok: true });
    let saved = savedCell(viewer.save()!, 'A1');
    expect(saved.cell).toContain('<v>0.12</v>');
    expect(saved.id).toBe(0);
    viewer.select({ row: 1, col: 0 });
    expect(viewer.paste({ text: '2026-11-12' }).ok).toBe(true);
    saved = savedCell(viewer.save()!, 'A2');
    expect(saved.cell).toContain('<v>46338</v>');
    expect(saved.id).toBe(0);
  });

  test('a full styles part stores the number and keeps the cell format', async () => {
    // `MAX_STYLE_RECORDS` cell formats: format 0 and 63,999 more.
    const xf = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';
    const viewer = await mount(
      buildXlsx({
        sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }],
        styles: { cellXfs: Array.from({ length: MAX_STYLE_RECORDS - 1 }, () => xf) },
      })
    );
    viewer.select({ row: 0, col: 0 });
    expect(viewer.enterValue('12%')).toEqual({ ok: true });
    viewer.select({ row: 1, col: 0 });
    expect(viewer.paste({ text: '2026-11-12\t7' }).ok).toBe(true);
    const bytes = viewer.save()!;
    expect(savedCell(bytes, 'A1')).toMatchObject({ id: 0 });
    expect(savedCell(bytes, 'A1').cell).toContain('<v>0.12</v>');
    expect(savedCell(bytes, 'A2').cell).toContain('<v>46338</v>');
    expect(savedCell(bytes, 'B2').cell).toContain('<v>7</v>');
  });

  test('a styles part near the cap still takes the format when it fits', async () => {
    // Two records under the cap (format 0 and MAX - 3 more): the dry run decides, and each new
    // format fits.
    const xf = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';
    const viewer = await mount(
      buildXlsx({
        sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }],
        styles: { cellXfs: Array.from({ length: MAX_STYLE_RECORDS - 3 }, () => xf) },
      })
    );
    viewer.select({ row: 0, col: 0 });
    expect(viewer.enterValue('12%')).toEqual({ ok: true });
    viewer.select({ row: 1, col: 0 });
    expect(viewer.paste({ text: '2026-11-12' }).ok).toBe(true);
    const bytes = viewer.save()!;
    expect(savedCell(bytes, 'A1').id).toBe(9);
    expect(savedCell(bytes, 'A2').id).toBe(14);
  });

  test('a date and a time copies in the locale order and pastes back to the same serial', async () => {
    for (const [short, typed] of [
      ['yyyy-mm-dd', '2026-11-12 13:45:30'],
      ['dd.mm.yyyy', '2026-11-12 13:45:30'],
      ['yyyy/m/d', '2026-11-12 13:45:30'],
      ['mm-dd-yy', '2026-11-12 1:45 PM'],
      ['yyyy-mm-dd', '2026-11-12 1:45:30 PM'],
      ['General', '2026-11-12 13:45'],
      ['General', '2026-11-12'],
      ['yyyy. mm. dd.', '2026-11-12'],
      ['yyyy. mm. dd.', '2026-11-12 13:45'],
    ] as const) {
      const locale: NumberLocale = { decimal: ',', group: '.', builtins: { 14: short } };
      const viewer = await mount(book());
      viewer.setNumberLocale(locale);
      viewer.select({ row: 1, col: 0 });
      expect(viewer.enterValue(typed)).toEqual({ ok: true });
      const out = viewer.copy();
      if (!out.ok) throw new Error(out.refusal.code);
      viewer.select({ row: 2, col: 0 });
      // The copy keeps the typed clock.
      if (typed.endsWith('PM')) expect(out.content.text).toContain('PM');
      expect(viewer.paste({ text: out.content.text }).ok).toBe(true);
      const bytes = viewer.save()!;
      const v = (ref: string) => /<v>([^<]*)<\/v>/.exec(savedCell(bytes, ref).cell)?.[1];
      expect(v('A2')).toBeDefined();
      expect(v('A3')).toBe(v('A2'));
    }
  });

  test('a redefined built-in date id takes the locale code and copies back to the same date', async () => {
    const viewer = await mount(
      buildXlsx({
        sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }],
        styles: { numFmts: ['<numFmt numFmtId="14" formatCode="yyyy"/>'] },
      })
    );
    viewer.setNumberLocale(DE);
    viewer.select({ row: 0, col: 1 });
    expect(viewer.enterValue('12.11.2026')).toEqual({ ok: true });
    const out = viewer.copy();
    if (!out.ok) throw new Error(out.refusal.code);
    expect(out.content.text).toBe('12.11.2026');
    viewer.select({ row: 0, col: 2 });
    expect(viewer.paste({ text: out.content.text }).ok).toBe(true);
    const bytes = viewer.save()!;
    expect(savedCell(bytes, 'B1')).toMatchObject({ id: 164, code: 'dd.mm.yyyy' });
    expect(savedCell(bytes, 'C1').cell).toContain('<v>46338</v>');
  });

  test('a file code for a built-in id that is not the locale code also takes the locale code', async () => {
    // The file writes the default code of id 14, which wins over the code of the DE locale.
    const viewer = await mount(
      buildXlsx({
        sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }],
        styles: { numFmts: ['<numFmt numFmtId="14" formatCode="mm-dd-yy"/>'] },
      })
    );
    viewer.setNumberLocale(DE);
    viewer.select({ row: 0, col: 1 });
    expect(viewer.enterValue('12.11.2026')).toEqual({ ok: true });
    const out = viewer.copy();
    if (!out.ok) throw new Error(out.refusal.code);
    expect(out.content.text).toBe('12.11.2026');
    // In the US English locale the file code is the locale code, so the id stays.
    viewer.setNumberLocale(EN_US);
    viewer.select({ row: 0, col: 2 });
    expect(viewer.enterValue('11/12/2026')).toEqual({ ok: true });
    const bytes = viewer.save()!;
    expect(savedCell(bytes, 'B1')).toMatchObject({ id: 164, code: 'dd.mm.yyyy' });
    expect(savedCell(bytes, 'C1').id).toBe(14);
  });

  test('a locale code that the store cannot read stores the value with no format', async () => {
    // The file redefines format 9, so the entry takes the locale code, which the store cannot read.
    const bad: NumberLocale = { ...EN_US, builtins: { 9: '0%[' } };
    const viewer = await mount(
      buildXlsx({
        sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }],
        styles: { numFmts: ['<numFmt numFmtId="9" formatCode="0.0%"/>'] },
      })
    );
    viewer.setNumberLocale(bad);
    viewer.select({ row: 0, col: 1 });
    expect(viewer.enterValue('12%')).toEqual({ ok: true });
    viewer.select({ row: 1, col: 1 });
    expect(viewer.paste({ text: '15%' }).ok).toBe(true);
    const bytes = viewer.save()!;
    expect(savedCell(bytes, 'B1').cell).toContain('<v>0.12</v>');
    expect(savedCell(bytes, 'B1').cell).not.toContain(' s="');
    expect(savedCell(bytes, 'B2').cell).toContain('<v>0.15</v>');
  });

  test('a negative currency amount keeps its sign in copied text and pastes as a number', async () => {
    const from = await mount(book());
    from.select({ row: 1, col: 0 });
    expect(from.enterValue('-$1,234.50')).toEqual({ ok: true });
    const out = from.copy();
    if (!out.ok) throw new Error(out.refusal.code);
    expect(out.content.text).toBe('-$1,234.50');
    const to = await mount(book());
    to.select({ row: 1, col: 0 });
    expect(to.paste({ text: out.content.text }).ok).toBe(true);
    const saved = savedCell(to.save()!, 'A2');
    expect(saved.cell).toContain('<v>-1234.5</v>');
    expect(saved.code).toBe('&quot;$&quot;#,##0.00');
  });

  test('a built-in id that the workbook redefines gives the typed form a custom id', async () => {
    // The file defines ids 9 and 14 with other codes, so they do not show a percent or a date.
    const viewer = await mount(
      buildXlsx({
        sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }],
        styles: {
          numFmts: [
            '<numFmt numFmtId="9" formatCode="0.0"/>',
            '<numFmt numFmtId="14" formatCode="yyyy"/>',
          ],
        },
      })
    );
    viewer.select({ row: 0, col: 1 });
    expect(viewer.enterValue('12%')).toEqual({ ok: true });
    expect(viewer.snapshot().activeCell?.content).toBe('12%');
    viewer.select({ row: 1, col: 1 });
    expect(viewer.paste({ text: '2026-11-12\t13:45' }).ok).toBe(true);
    const bytes = viewer.save()!;
    expect(savedCell(bytes, 'B1')).toMatchObject({ id: 164, code: '0%' });
    expect(savedCell(bytes, 'B2')).toMatchObject({ id: 165, code: 'mm-dd-yy' });
    // Id 20 is not redefined, so the time keeps the built-in id.
    expect(savedCell(bytes, 'C2').id).toBe(20);
  });

  test('a styles part that the store cannot read stores the number with no format', async () => {
    const files = unzipSync(buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] }));
    files['xl/styles.xml'] = strToU8(
      '<foo xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>'
    );
    const viewer = await mount(zipSync(files));
    viewer.select({ row: 0, col: 1 });
    expect(viewer.enterValue('12%')).toEqual({ ok: true });
    viewer.select({ row: 1, col: 0 });
    expect(viewer.paste({ text: '1\t2026-11-12' }).ok).toBe(true);
    const sheet = strFromU8(unzipSync(viewer.save()!)['xl/worksheets/sheet1.xml']!);
    expect(sheet).toContain('<c r="B1"><v>0.12</v></c>');
    expect(sheet).toContain('<c r="B2"><v>46338</v></c>');
  });

  test('a full number format list, also with repeated ids, stores the number', async () => {
    const numFmt = '<numFmt numFmtId="164" formatCode="0.0"/>';
    const viewer = await mount(
      buildXlsx({
        sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }],
        styles: { numFmts: Array.from({ length: MAX_STYLE_RECORDS }, () => numFmt) },
      })
    );
    viewer.select({ row: 0, col: 0 });
    expect(viewer.enterValue('$5.50')).toEqual({ ok: true });
    viewer.select({ row: 1, col: 0 });
    expect(viewer.paste({ text: '36:00' }).ok).toBe(true);
    const bytes = viewer.save()!;
    expect(savedCell(bytes, 'A1').cell).toContain('<v>5.5</v>');
    expect(savedCell(bytes, 'A1').id).toBe(0);
    expect(savedCell(bytes, 'A2').cell).toContain('<v>1.5</v>');
  });

  test('a Text cell stores a formula, clears on an empty entry, and drops a leading apostrophe', async () => {
    const viewer = await mount(book());
    viewer.select({ row: 0, col: 2 });
    expect(viewer.enterValue('=1+1').ok).toBe(true);
    expect(savedCell(viewer.save()!, 'C1').cell).toContain('<f>1+1</f>');
    expect(viewer.enterValue("'=x").ok).toBe(true);
    expect(viewer.snapshot().activeCell?.content).toBe("'=x");
    // A Text cell keeps `007` as text, so its edit text needs no apostrophe.
    expect(viewer.enterValue("'007").ok).toBe(true);
    expect(viewer.snapshot().activeCell?.content).toBe('007');
    expect(viewer.enterValue('').ok).toBe(true);
    expect(savedCell(viewer.save()!, 'C1').cell).not.toContain('<v>');
    expect(savedCell(viewer.save()!, 'C1').id).toBe(49);
  });

  test('a paste into a Text cell keeps the text', async () => {
    const viewer = await mount(book());
    viewer.select({ row: 0, col: 2 });
    expect(viewer.paste({ text: '12%' }).ok).toBe(true);
    const saved = savedCell(viewer.save()!, 'C1');
    expect(saved.cell).toContain('t="s"');
    expect(saved.id).toBe(49);
  });

  test('a cell whose custom code is General takes the format of the typed form', async () => {
    const viewer = await mount(
      buildXlsx({
        sheets: [{ name: 'S', rows: row(1, ['<c r="A1" s="1"/>']) }],
        styles: {
          numFmts: ['<numFmt numFmtId="164" formatCode="General"/>'],
          cellXfs: [
            '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
          ],
        },
      })
    );
    viewer.select({ row: 0, col: 0 });
    expect(viewer.enterValue('12%').ok).toBe(true);
    expect(savedCell(viewer.save()!, 'A1').id).toBe(9);
  });

  test('re-editing a typed date, time, or percent shows it as typed', async () => {
    const viewer = await mount(book());
    for (const [i, typed] of ['2026-11-12', '13:45', '12%', '$5.50'].entries()) {
      viewer.select({ row: 4 + i, col: 0 });
      expect(viewer.enterValue(typed).ok).toBe(true);
      viewer.select({ row: 4 + i, col: 0 });
      const shown = viewer.snapshot().activeCell?.content;
      expect(shown).toBe(typed === '$5.50' ? '5.5' : typed);
      // F2 opens the cell editor with the same text.
      expect(viewer.beginEdit().ok).toBe(true);
      expect(viewer.snapshot().editText).toBe(shown!);
      viewer.cancelEdit();
      // Enter on the unchanged text writes nothing: one undo empties the cell, so the typed
      // entry is the only step of the cell.
      expect(viewer.enterValue(shown!).ok).toBe(true);
      expect(viewer.undo().ok).toBe(true);
      viewer.select({ row: 4 + i, col: 0 });
      expect(viewer.snapshot().activeCell?.content).toBe('');
      expect(viewer.redo().ok).toBe(true);
    }
  });
});

describe('typed entry: stored text that reads as a number form', () => {
  const forms = ['12%', '2026-11-12', '(100)', '$5', '1/2'];
  const textBook = (date1904 = false) =>
    buildXlsx({
      sheets: [{ name: 'S', rows: forms.map((t, i) => row(i + 1, [c(`A${i + 1}`, t)])).join('') }],
      date1904,
    });

  test('shows a leading apostrophe, and Enter on that text keeps the text and writes nothing', async () => {
    const viewer = await mount(textBook());
    for (const [i, stored] of forms.entries()) {
      viewer.select({ row: i, col: 0 });
      expect(viewer.snapshot().activeCell?.content).toBe(`'${stored}`);
      expect(viewer.enterValue(`'${stored}`).ok).toBe(true);
    }
    // No Enter wrote a step, so there is nothing to undo.
    expect(viewer.undo().ok).toBe(false);
    const bytes = viewer.save()!;
    for (const i of forms.keys()) {
      const saved = savedCell(bytes, `A${i + 1}`);
      expect(saved.cell).toMatch(/t="(?:s|inlineStr)"/);
      expect(saved.cell).not.toContain('<v>');
    }
  });

  test('in the 1904 date system, a date before 1904 is plain text', async () => {
    const viewer = await mount(
      buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', '1903-12-31')]) }], date1904: true })
    );
    viewer.select({ row: 0, col: 0 });
    expect(viewer.snapshot().activeCell?.content).toBe('1903-12-31');
  });
});

describe('parseCellInput', () => {
  test('reads a number form as a number, in the date system and year that it is given', () => {
    expect(parseCellInput('12%', EN_US)).toEqual({ type: 'number', value: 0.12 });
    expect(parseCellInput('2026-11-12', EN_US)).toEqual({ type: 'number', value: 46338 });
    expect(parseCellInput('2026-11-12', EN_US, { date1904: true })).toEqual({
      type: 'number',
      value: 46338 - 1462,
    });
    expect(parseCellInput('11/12', EN_US, { year: 2026 })).toEqual({
      type: 'number',
      value: 46338,
    });
  });
});

describe('typed entry and the built-in formats of the locale', () => {
  /** A1 with built-in 44 (a currency format), A2 with built-in 31 (a date format), no x:numFmt. */
  const localeFormats = () =>
    buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: row(1, ['<c r="A1" s="1"/>']) + row(2, ['<c r="A2" s="2"/>']),
        },
      ],
      styles: {
        cellXfs: [
          '<xf numFmtId="44" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
          '<xf numFmtId="31" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
        ],
      },
    });

  test('a cell with a built-in id that has no code in the locale keeps its format', async () => {
    const viewer = await mount(localeFormats());
    viewer.select({ row: 0, col: 0 });
    expect(viewer.enterValue('$5.50').ok).toBe(true);
    viewer.select({ row: 1, col: 0 });
    expect(viewer.paste({ text: '2026-11-13' }).ok).toBe(true);
    const bytes = viewer.save()!;
    expect(savedCell(bytes, 'A1')).toMatchObject({ id: 44 });
    expect(savedCell(bytes, 'A1').cell).toContain('<v>5.5</v>');
    expect(savedCell(bytes, 'A2')).toMatchObject({ id: 31 });
    expect(savedCell(bytes, 'A2').cell).toContain('<v>46339</v>');
  });

  test('a typed elapsed time saves its `[h]:mm` code', async () => {
    const viewer = await mount(book());
    viewer.select({ row: 0, col: 0 });
    expect(viewer.enterValue('36:00').ok).toBe(true);
    const saved = savedCell(viewer.save()!, 'A1');
    expect(saved.cell).toContain('<v>1.5</v>');
    expect(saved.code).toBe('[h]:mm');
  });
});
