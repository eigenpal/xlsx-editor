// Format edits seen through layout: each test applies a format patch through the store and reads
// the layout records that a painter draws, so a patch is checked where it shows, not only in
// the styles part. The geometry is the print profile of the fidelity harness, with the substitute
// fonts of the other layout tests.

import { beforeAll, describe, expect, test } from 'bun:test';
import {
  buildXlsx,
  c,
  row,
  type FixtureWorkbook,
} from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx, type XlsxPackage } from '../../store/package/package.ts';
import { saveXlsx } from '../../store/package-write/save.ts';
import { WorkbookDocumentStore } from '../../store/store/document-store.ts';
import type { FormatPatch } from '../../store/store/format-patch.ts';
import type { WorkbookOp } from '../../store/store/ops.ts';
import { layoutSheet, type CellTextBox, type SheetLayout } from '../sheet-layout.ts';
import { PRINT_PROFILE } from '../metrics.ts';
import type { TextMeasurer } from '../text/measurer.ts';
import { substituteMeasurer } from './test-fonts.ts';

let measurer: TextMeasurer;
beforeAll(async () => {
  measurer = await substituteMeasurer();
});

const VIEW = { width: 540, height: 720, scrollX: 0, scrollY: 0 };

function open(wb: FixtureWorkbook): XlsxPackage {
  const r = readXlsx(buildXlsx(wb));
  if (!r.ok) throw new Error('read');
  return r.pkg;
}

function layout(pkg: XlsxPackage): SheetLayout {
  const out = layoutSheet(pkg, 'S', VIEW, { measurer, profile: PRINT_PROFILE });
  if (!out) throw new Error('no sheet');
  return out;
}

const fmt = (ref: string, patch: FormatPatch): WorkbookOp => ({
  op: 'applyFormat',
  sheet: 'S',
  ref,
  patch,
});

/** The layout after one format patch, with the layout before it. */
function edited(wb: FixtureWorkbook, ...ops: WorkbookOp[]) {
  const store = new WorkbookDocumentStore(open(wb));
  const before = layout(store.package);
  const r = store.transact(ops);
  if (!r.ok) throw new Error(JSON.stringify(r.rejection));
  return { store, before, after: layout(store.package) };
}

const texts = (l: SheetLayout): CellTextBox[] => l.panes.flatMap((p) => p.texts);
const textAt = (l: SheetLayout, r: number, col: number) =>
  texts(l).find((t) => t.row === r && t.col === col);
const shown = (l: SheetLayout, r: number, col: number) =>
  textAt(l, r, col)
    ?.text.lines.map((line) => line.spans.map((s) => s.text).join(''))
    .join('\n');
const boxAt = (l: SheetLayout, r: number, col: number) =>
  l.panes.flatMap((p) => p.cells).find((b) => b.row === r && b.col === col);

const words = (): FixtureWorkbook => ({
  sheets: [
    {
      name: 'S',
      rows:
        row(1, [c('A1', 'Hello world')]) +
        '<row r="2" ht="15"><c r="A2" t="inlineStr"><is><t>Sized</t></is></c></row>' +
        '<row r="3" ht="15" customHeight="1"><c r="A3" t="inlineStr"><is><t>Kept</t></is></c></row>',
    },
  ],
});

describe('font', () => {
  test('bold changes the span style and widens the line', () => {
    const { before, after } = edited(words(), fmt('A1', { bold: true }));
    const span = textAt(after, 0, 0)!.text.lines[0]!.spans[0]!;
    expect(span.style.bold).toBe(true);
    expect(textAt(before, 0, 0)!.text.lines[0]!.spans[0]!.style.bold).toBe(false);
    expect(span.width).toBeGreaterThan(textAt(before, 0, 0)!.text.lines[0]!.spans[0]!.width);
  });

  test('a font size grows a fitted row and keeps a row that the user sized', () => {
    const { before, after } = edited(words(), fmt('A2:A3', { fontSize: 24 }));
    expect(before.geometry.rows.sizeOf(1)).toBe(14);
    expect(after.geometry.rows.sizeOf(1)).toBe(28);
    expect(after.geometry.rows.sizeOf(2)).toBe(before.geometry.rows.sizeOf(2));
    expect(textAt(after, 1, 0)!.text.lines[0]!.spans[0]!.style.sizePt).toBe(24);
  });
});

describe('fill, borders, and alignment', () => {
  test('a fill and borders add a cell box', () => {
    const { before, after } = edited(
      words(),
      fmt('A1:B2', {
        fill: { type: 'solid', color: 'FFFF00' },
        borders: { edges: 'outline', style: 'medium', color: '0000FF' },
      })
    );
    expect(boxAt(before, 0, 0)).toBeUndefined();
    const a1 = boxAt(after, 0, 0)!;
    expect(a1.fill).toEqual({ kind: 'solid', color: 'FFFF00' } as never);
    expect([a1.border.left.style, a1.border.top.style, a1.border.right.style]).toEqual([
      'medium',
      'medium',
      'none',
    ]);
    const b2 = boxAt(after, 1, 1)!;
    expect([b2.border.right.style, b2.border.bottom.style, b2.border.left.style]).toEqual([
      'medium',
      'medium',
      'none',
    ]);
  });

  test('a right alignment moves the text to the right edge of the cell', () => {
    const { before, after } = edited(words(), fmt('A2', { horizontal: 'right' }));
    const x0 = textAt(before, 1, 0)!.text.lines[0]!;
    const x1 = textAt(after, 1, 0)!.text.lines[0]!;
    const cell = after.geometry.columns.sizeOf(0);
    expect(x1.x).toBeGreaterThan(x0.x);
    expect(x1.x + x1.width).toBe(cell - 2);
  });
});

describe('wrapping', () => {
  test('wrapped text breaks into lines and its row grows to hold them', () => {
    const long = 'one two three four five six seven eight nine ten';
    const { before, after } = edited(
      { sheets: [{ name: 'S', rows: row(1, [c('A1', long)]) }] },
      fmt('A1', { wrapText: true })
    );
    expect(textAt(before, 0, 0)!.text.lines).toHaveLength(1);
    const lines = textAt(after, 0, 0)!.text.lines.length;
    expect(lines).toBeGreaterThan(2);
    // 14 points for each Calibri 11 line in print.
    expect(after.geometry.rows.sizeOf(0)).toBe(14 * lines);
    // Every line sits inside the row.
    const last = textAt(after, 0, 0)!.text.lines.at(-1)!;
    expect(last.baseline).toBeLessThan(after.geometry.rows.sizeOf(0));
  });

  test('wrapped text in a merged range leaves the row height alone', () => {
    const long = 'one two three four five six seven eight nine ten';
    const { after } = edited(
      {
        sheets: [
          {
            name: 'S',
            rows: row(1, [c('A1', long)]),
            merges: ['A1:B1'],
          },
        ],
      },
      fmt('A1', { wrapText: true })
    );
    expect(after.geometry.rows.sizeOf(0)).toBe(14);
  });
});

describe('number formats', () => {
  const values = (): FixtureWorkbook => ({
    sheets: [
      {
        name: 'S',
        rows: row(1, [c('A1', 1.5), c('B1', 0.125), c('C1', 45352), c('D1', 2), c('E1', -2)]),
      },
    ],
  });

  test('the displayed text follows the number format of the patch', () => {
    const { before, after } = edited(
      values(),
      fmt('A1', { numberFormat: { code: '0.00' } }),
      fmt('B1', { numberFormat: { code: '0.0%' } }),
      fmt('C1', { numberFormat: { id: 14 } }),
      fmt('D1:E1', { numberFormat: { code: '[Red]0.0;[Blue]-0.0' } })
    );
    expect(shown(before, 0, 0)).toBe('1.5');
    expect(shown(after, 0, 0)).toBe('1.50');
    expect(shown(after, 0, 1)).toBe('12.5%');
    expect(shown(after, 0, 2)).toBe('03-01-24');
    expect(shown(after, 0, 3)).toBe('2.0');
    expect(textAt(after, 0, 3)!.text.lines[0]!.spans[0]!.style.color).toEqual('FF0000' as never);
    expect(shown(after, 0, 4)).toBe('-2.0');
    expect(textAt(after, 0, 4)!.text.lines[0]!.spans[0]!.style.color).toEqual('0000FF' as never);
  });

  test('a pad after a `*x` fill keeps its character after the fill run', () => {
    const { after } = edited(
      values(),
      fmt('A1', { numberFormat: { code: '*x0_W' } }),
      fmt('B1', { numberFormat: { code: '0_W' } })
    );
    const span = (col: number) => textAt(after, 0, col)!.text.lines[0]!.spans[0]!;
    const advance = (col: number, i: number) => {
      const s = span(col);
      return (s.positions[i + 1] ?? s.x + s.width) - s.positions[i]!;
    };
    const filled = span(0);
    const last = filled.text.length - 1;
    expect(filled.text).toMatch(/^x{2,}2 $/);
    // The last space is as wide as `W`, as without a fill; every fill `x` keeps its own width.
    expect(advance(0, last)).toBe(advance(1, span(1).text.length - 1));
    expect(advance(0, last)).toBeGreaterThan(advance(0, 0));
    for (let i = 1; i < last - 1; i += 1) expect(advance(0, i)).toBe(advance(0, 0));
  });
});

describe('row and column formats', () => {
  const styled = (): FixtureWorkbook => ({
    sheets: [
      {
        name: 'S',
        cols: '<cols><col min="4" max="4" width="9" style="1" customWidth="1"/></cols>',
        rows: '<row r="3" s="1" customFormat="1"><c r="A3"><v>1</v></c></row>',
      },
    ],
    styles: {
      fills: [
        '<fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill></fill>',
      ],
      cellXfs: ['<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>'],
    },
  });

  test('an empty cell paints the fill of its row or its column format', () => {
    const l = layout(open(styled()));
    const green = { kind: 'solid', color: '00FF00' };
    // Row 3: every empty visible cell; A3 has a cell of its own with the default format.
    expect(boxAt(l, 2, 1)?.fill).toEqual(green as never);
    expect(boxAt(l, 2, 5)?.fill).toEqual(green as never);
    expect(boxAt(l, 2, 0)).toBeUndefined();
    // Column D: every empty visible cell.
    expect(boxAt(l, 0, 3)?.fill).toEqual(green as never);
    expect(boxAt(l, 30, 3)?.fill).toEqual(green as never);
    expect(boxAt(l, 0, 2)).toBeUndefined();
  });

  test('a format on an empty cell of a styled row starts from the row format', () => {
    const { after } = edited(styled(), fmt('B3', { bold: true }));
    expect(boxAt(after, 2, 1)?.fill).toEqual({ kind: 'solid', color: '00FF00' } as never);
  });

  test('a cleared empty cell in a styled row stops showing the row format', () => {
    const { after } = edited(styled(), fmt('B3', { bold: true }), {
      op: 'clearFormats',
      sheet: 'S',
      ref: 'B3',
    });
    expect(boxAt(after, 2, 1)).toBeUndefined();
  });
});

describe('clear and save', () => {
  const patch: FormatPatch = {
    bold: true,
    fontSize: 14,
    fill: { type: 'solid', color: 'ABCDEF' },
    borders: { edges: 'all' },
    horizontal: 'center',
    wrapText: true,
    numberFormat: { code: '0.000' },
  };
  const wb = (): FixtureWorkbook => ({
    sheets: [{ name: 'S', rows: row(1, [c('A1', 1.5), c('B1', 'Some text that wraps')]) }],
  });

  test('clearFormats gives back the layout of the workbook before the format', () => {
    const { store, before } = edited(wb(), fmt('A1:C3', patch));
    store.transact([{ op: 'clearFormats', sheet: 'S', ref: 'A1:C3' }]);
    expect(layout(store.package)).toEqual(before);
  });

  test('a saved and reopened edit lays out as the edit in memory', () => {
    const { store, after } = edited(wb(), fmt('A1:C3', patch));
    const reopened = readXlsx(saveXlsx(store.package));
    if (!reopened.ok) throw new Error('reopen');
    expect(layout(reopened.pkg)).toEqual(after);
  });
});

describe('conditional formatting through edits and formulas', () => {
  const rule = (ref: string) =>
    `<conditionalFormatting sqref="${ref}"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>3</formula></cfRule></conditionalFormatting>`;
  const bold = (l: SheetLayout, r: number, col: number) =>
    textAt(l, r, col)?.text.lines[0]?.spans[0]?.style.bold;
  const book = (cells: string, calcPr?: string): FixtureWorkbook => ({
    sheets: [{ name: 'S', rows: row(1, [cells]), afterMerges: rule('A1:B1') }],
    styles: { dxfs: ['<dxf><font><b/></font></dxf>'] },
    ...(calcPr ? { calcPr } : {}),
  });

  test('a value edit changes the result of a rule', () => {
    const { before, after } = edited(book(c('A1', 5)), {
      op: 'setCellValue',
      sheet: 'S',
      ref: 'A1',
      value: { type: 'number', value: 2 },
    });
    expect([bold(before, 0, 0), bold(after, 0, 0)]).toEqual([true, false]);
  });

  test('a rule over an out-of-date formula result follows the cached value, and the text keeps its mark', () => {
    const l = layout(open(book(c('A1', 5, { f: '2+3' }), 'fullCalcOnLoad="1"')));
    const t = textAt(l, 0, 0)!;
    expect([t.text.lines[0]!.spans[0]!.style.bold, t.stale]).toEqual([true, true]);
  });

  test('a formula without a cached value does not paint its rule', () => {
    const fill = '<dxf><fill><patternFill><bgColor rgb="FFFFFF00"/></patternFill></fill></dxf>';
    const wb: FixtureWorkbook = {
      sheets: [
        {
          name: 'S',
          rows: row(1, ['<c r="A1"><f>B9</f></c>']),
          afterMerges:
            '<conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="0" priority="1" operator="lessThan"><formula>5</formula></cfRule></conditionalFormatting>',
        },
      ],
      styles: { dxfs: [fill] },
    };
    const cells = layout(open(wb)).panes.flatMap((p) => p.cells);
    expect(cells.filter((cell) => cell.row === 0 && cell.col === 0)).toEqual([]);
  });
});

describe('wrapText on a number', () => {
  test('a General number fits its cell on one line, and a formatted one fills with #', () => {
    const { after } = edited(
      {
        sheets: [
          {
            name: 'S',
            cols: '<cols><col min="1" max="2" width="8" customWidth="1"/></cols>',
            rows: row(1, [c('A1', 1234567.891), c('B1', 1234567.891)]),
          },
        ],
      },
      fmt('A1:B1', { wrapText: true }),
      fmt('B1', { numberFormat: { code: '#,##0.000' } })
    );
    const a1 = textAt(after, 0, 0)!.text;
    expect(a1.lines.length).toBe(1);
    expect(a1.lines[0]!.spans.map((s) => s.text).join('')).toBe('1234568');
    const b1 = textAt(after, 0, 1)!.text;
    expect(b1.lines.length).toBe(1);
    expect(b1.lines[0]!.spans.map((s) => s.text).join('')).toMatch(/^#+$/);
  });
});
