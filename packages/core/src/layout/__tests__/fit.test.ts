// The fit of columns and rows to their content (`fit.ts`): a column takes the smallest width in
// characters whose pixels hold its widest shown text with the margins of the cell box, after the
// number format and in the cell's font; cells that wrap, rotate, or span merged columns do not
// count; empty and hidden columns have no width; a row takes the tallest automatic height of
// layout at the zooms of the fit.

import { beforeAll, describe, expect, test } from 'bun:test';
import {
  buildXlsx,
  c,
  row,
  type FixtureWorkbook,
} from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx, type XlsxPackage } from '../../store/package/package.ts';
import { MAX_COLS, MAX_ROWS } from '../../store/cell-ref.ts';
import { cellFormat } from '../../cascade/index.ts';
import { workbookIndex } from '../../store/workbook/workbook-index.ts';
import { textMargin, textWidth } from '../cell-text.ts';
import { fitColumnWidths, fitRowHeights } from '../fit.ts';
import { columnPixels } from '../grid-geometry.ts';
import { autoRowHeight, fontPixels, gridMetrics, zoomProfile } from '../metrics.ts';
import type { TextMeasurer } from '../text/measurer.ts';
import { layoutSheet } from '../sheet-layout.ts';
import { substituteMeasurer } from './test-fonts.ts';

let measurer: TextMeasurer;
beforeAll(async () => {
  measurer = await substituteMeasurer();
});

const profile = zoomProfile(1);

const STYLES: FixtureWorkbook['styles'] = {
  fonts: ['<font><sz val="20"/><name val="Calibri"/></font>'],
  cellXfs: [
    '<xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf>',
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment textRotation="90"/></xf>',
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left" indent="2"/></xf>',
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="centerContinuous"/></xf>',
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="justify"/></xf>',
  ],
  dxfs: ['<dxf><numFmt numFmtId="200" formatCode="0.000000"/></dxf>'],
};
/** The cell format indexes of `STYLES`. */
const NUMBER = 1;
const BIG = 2;
const WRAP = 3;
const ROTATED = 4;
const INDENT = 5;
const ACROSS = 6;
const JUSTIFY = 7;

function open(
  rows: string,
  extra: { cols?: string; merges?: string[]; afterMerges?: string; formatPr?: string } = {}
): XlsxPackage {
  const r = readXlsx(buildXlsx({ styles: STYLES, sheets: [{ name: 'S', rows, ...extra }] }));
  if (!r.ok) throw new Error('open');
  return r.pkg;
}

const columns = (pkg: XlsxPackage, first = 0, last = MAX_COLS - 1) =>
  fitColumnWidths(pkg, 'S', first, last, { measurer, profile })!;
const rows = (pkg: XlsxPackage, first = 0, last = MAX_ROWS - 1) =>
  fitRowHeights(pkg, 'S', first, last, { measurer, profile })!;

/** The pixels that a text in the format of a cell needs: the text, its margins, the gridline. */
function needed(pkg: XlsxPackage, text: string, s?: number): number {
  const index = workbookIndex(pkg)!;
  const font = cellFormat(pkg, index, s === undefined ? undefined : String(s)).format.font;
  const mdw = gridMetrics(pkg, index, measurer, profile).maxDigitWidth;
  const ctx = { measurer, profile, maxDigitWidth: mdw };
  const pad = textMargin(fontPixels(measurer, font, profile).maxDigitWidth);
  return Math.ceil((textWidth(ctx, text, font) * 96) / 72 - 1e-9) + 1 + 2 * pad;
}

const mdwOf = (pkg: XlsxPackage) =>
  gridMetrics(pkg, workbookIndex(pkg)!, measurer, profile).maxDigitWidth;

describe('the width of a column', () => {
  test('is the smallest width in 1/256 characters whose pixels hold the widest text', () => {
    const pkg = open(row(1, [c('A1', 'short'), c('B1', 'x')]) + row(2, [c('A2', 'a longer text')]));
    const widths = columns(pkg);
    const width = widths.get(0)!;
    const px = needed(pkg, 'a longer text');
    expect(columnPixels(width, mdwOf(pkg))).toBe(px);
    expect(columnPixels(width - 1 / 256, mdwOf(pkg))).toBeLessThan(px);
    expect(widths.get(1)! < width).toBe(true);
  });

  test('measures the text after the number format, in the font of the cell', () => {
    const pkg = open(row(1, [c('A1', 1234.5, { s: NUMBER }), c('B1', 'Wide', { s: BIG })]));
    const widths = columns(pkg);
    expect(columnPixels(widths.get(0)!, mdwOf(pkg))).toBe(needed(pkg, '1,234.50', NUMBER));
    expect(columnPixels(widths.get(1)!, mdwOf(pkg))).toBe(needed(pkg, 'Wide', BIG));
  });

  test('adds the indent of the cell, one maximum digit width for each level', () => {
    const pkg = open(row(1, [c('A1', 'text'), c('B1', 'text', { s: INDENT })]));
    const widths = columns(pkg);
    const mdw = mdwOf(pkg);
    expect(columnPixels(widths.get(1)!, mdw) - columnPixels(widths.get(0)!, mdw)).toBe(2 * mdw);
  });

  test('leaves out wrapped and rotated cells, a two-column merge, and empty and hidden columns', () => {
    const long = 'a text that is much wider than the default column';
    const pkg = open(
      row(1, [
        c('A1', long, { s: WRAP }),
        c('B1', long, { s: ROTATED }),
        c('C1', long),
        c('E1', long),
        c('F1', 'x'),
      ]) + row(2, [c('A2', 'ab')]),
      { merges: ['C1:D1'], cols: '<cols><col min="5" max="5" width="9" hidden="1"/></cols>' }
    );
    const widths = columns(pkg);
    expect(columnPixels(widths.get(0)!, mdwOf(pkg))).toBe(needed(pkg, 'ab'));
    expect([...widths.keys()].sort()).toEqual([0, 5]);
  });

  test('counts a number with wrapText or justify, which shows on one line', () => {
    const pkg = open(
      row(1, [
        c('A1', 1234567.891, { s: WRAP }),
        c('B1', 1234567.891, { s: JUSTIFY }),
        c('C1', 1234567.891),
      ])
    );
    const widths = columns(pkg);
    // The same width as the number in a cell with no alignment.
    expect(widths.get(2)).toBeGreaterThan(0);
    expect(widths.get(0)).toBe(widths.get(2)!);
    expect(widths.get(1)).toBe(widths.get(2)!);
  });

  test('leaves out cells that show no text, justified text, and text centered across cells', () => {
    const long = 'a text that is much wider than the default column';
    const pkg = open(
      row(1, [
        c('A1', '', { t: 'str', f: '""' }),
        c('B1', ''),
        c('C1', long, { s: ACROSS }),
        c('D1', long, { s: JUSTIFY }),
        c('E1', 'x'),
      ])
    );
    expect([...columns(pkg).keys()]).toEqual([4]);
  });

  test('leaves out the cells that a merged range covers', () => {
    const long = 'a text that is much wider than the default column';
    const pkg = open(row(1, [c('A1', 'ab')]) + row(2, [c('A2', long)]), { merges: ['A1:A2'] });
    expect(columnPixels(columns(pkg).get(0)!, mdwOf(pkg))).toBe(needed(pkg, 'ab'));
  });

  test('leaves out the cells of a hidden row', () => {
    const long = 'a text that is much wider than the default column';
    const pkg = open(row(1, [c('A1', long)], 'hidden="1"') + row(2, [c('A2', 'ab')]));
    expect(columnPixels(columns(pkg).get(0)!, mdwOf(pkg))).toBe(needed(pkg, 'ab'));
  });

  test('measures a formula with a cached value in the format of its conditional formatting', () => {
    const pkg = open(row(1, [c('A1', 1.5, { f: '1.5' }), c('B1', 1.5, { f: '1.5' })]), {
      afterMerges:
        '<conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>',
    });
    const widths = columns(pkg);
    expect(widths.get(0)!).toBeGreaterThan(widths.get(1)!);
  });

  test('takes the margins of rich text from its first run, also an empty one', () => {
    const r = readXlsx(
      buildXlsx({
        rawStrings: [
          '<si><r><rPr><sz val="40"/><rFont val="Calibri"/></rPr><t></t></r><r><rPr><sz val="8"/><rFont val="Calibri"/></rPr><t>small</t></r></si>',
          '<si><r><rPr><sz val="8"/><rFont val="Calibri"/></rPr><t>small</t></r></si>',
        ],
        sheets: [{ name: 'S', rows: row(1, [c('A1', 0, { t: 's' }), c('B1', 1, { t: 's' })]) }],
      })
    );
    if (!r.ok) throw new Error('open');
    const widths = columns(r.pkg);
    expect(widths.get(0)!).toBeGreaterThan(widths.get(1)!);
  });

  test('reads only the columns of the range', () => {
    const pkg = open(row(1, [c('A1', 'one'), c('B1', 'two'), c('C1', 'three')]));
    expect([...columns(pkg, 1, 1).keys()]).toEqual([1]);
  });

  test('keeps a column whose text measures nothing, such as a line break alone', () => {
    const pkg = open(row(1, [c('A1', '\n'), c('B1', '\u200b'), c('C1', '\r\n\n'), c('D1', 'x')]));
    expect([...columns(pkg).keys()]).toEqual([3]);
  });

  test('is undefined for a sheet that does not exist', () => {
    const pkg = open(row(1, [c('A1', 'one')]));
    expect(fitColumnWidths(pkg, 'Nope', 0, 0, { measurer, profile })).toBeUndefined();
  });

  test('does not join a rich cell with a plain cell whose text and font are the same', () => {
    // The empty first run of the rich cell sets its margin font; its text takes the cell font.
    const width = (order: 'plain first' | 'rich first') => {
      const plain = c(order === 'plain first' ? 'A1' : 'A2', 'Some text here');
      const rich = c(order === 'plain first' ? 'A2' : 'A1', 0, { t: 's' });
      const r = readXlsx(
        buildXlsx({
          rawStrings: [
            '<si><r><rPr><sz val="48"/><rFont val="Calibri"/></rPr><t></t></r><r><t>Some text here</t></r></si>',
          ],
          sheets: [
            {
              name: 'S',
              rows:
                row(1, [order === 'plain first' ? plain : rich]) +
                row(2, [order === 'plain first' ? rich : plain]),
            },
          ],
        })
      );
      if (!r.ok) throw new Error('open');
      return columns(r.pkg).get(0);
    };
    expect(width('plain first')).toBe(width('rich first')!);
  });

  test('keeps a column whose shown format the conditional formatting budget leaves unknown', () => {
    // 80 rules over 60,000 rows use up `LAYOUT_CHECK_BUDGET` before the last rows, whose rule
    // format shows more digits than the base format.
    const n = 60_000;
    let cells = '';
    for (let r = 1; r <= n; r += 1) cells += row(r, [c(`A${r}`, r > n - 1000 ? 123456789 : 1)]);
    let rules = '';
    for (let k = 1; k <= 80; k += 1)
      rules += `<conditionalFormatting sqref="A1:A${n}"><cfRule type="cellIs" dxfId="0" priority="${k}" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>`;
    const pkg = open(cells, { afterMerges: rules });
    expect([...columns(pkg, 0, 0).keys()]).toEqual([]);
  });

  test('reads many merged ranges in one sweep', () => {
    const n = 80_000;
    let cells = '';
    const merges: string[] = [];
    for (let r = 1; r <= n; r += 1) {
      cells += row(r, [c(`A${r}`, `x${r}`), c(`C${r}`, r)]);
      merges.push(`A${r}:B${r}`);
    }
    const pkg = open(cells, { merges });
    const started = performance.now();
    const widths = columns(pkg);
    // One sweep: a lookup for each cell over every merge took many seconds.
    expect(performance.now() - started).toBeLessThan(20_000);
    expect([...widths.keys()]).toEqual([2]);
  });
});

describe('the height of a row', () => {
  test('is the automatic height of its tallest font, and a row with no cell takes the default', () => {
    const pkg = open(
      row(1, [c('A1', 'Big', { s: BIG })], 'ht="40" customHeight="1"') +
        row(2, [], 'ht="30"') +
        row(3, [c('A3', 'x')], 'hidden="1"')
    );
    const heights = rows(pkg);
    const index = workbookIndex(pkg)!;
    const big = cellFormat(pkg, index, String(BIG)).format.font;
    expect(heights.get(0)).toBe(autoRowHeight(fontPixels(measurer, big, profile), profile));
    expect(heights.get(1)).toBe(gridMetrics(pkg, index, measurer, profile).defaultRowHeight);
    expect(heights.has(2)).toBe(false);
  });

  test('fits the content on a sheet with a custom default height, so no text is cut', () => {
    const pkg = open(
      row(1, [c('A1', 'Big', { s: BIG })], 'ht="12" customHeight="1"') +
        row(2, [c('A2', 'x')]) +
        row(3, []),
      { formatPr: 'defaultRowHeight="12" customHeight="1"' }
    );
    const index = workbookIndex(pkg)!;
    const big = cellFormat(pkg, index, String(BIG)).format.font;
    const plain = gridMetrics(pkg, index, measurer, profile).defaultRowHeight;
    const heights = rows(pkg);
    expect(heights.get(0)).toBe(autoRowHeight(fontPixels(measurer, big, profile), profile));
    // Row 2 shows the custom default height without `@ht`, and its text needs more.
    expect(heights.get(1)).toBe(plain);
    // Row 3 has no content and shows the custom default height already.
    expect(heights.has(2)).toBe(false);
  });

  test('has no entry for a row that shows its fitted height already', () => {
    const pkg = open(row(1, [c('A1', 'x')]) + row(2, [c('A2', 'x')], 'ht="40"'));
    const fitted = rows(pkg).get(1)!;
    const again = open(row(1, [c('A1', 'x')]) + row(2, [c('A2', 'x')], `ht="${fitted}"`));
    expect([...rows(pkg).keys()]).toEqual([1]);
    expect([...rows(again).keys()]).toEqual([]);
  });

  test('keeps the height of a row with rotated text', () => {
    const pkg = open(
      row(1, [c('A1', 'up', { s: ROTATED })], 'ht="80"') + row(2, [c('A2', 'x')], 'ht="40"')
    );
    expect([...rows(pkg).keys()]).toEqual([1]);
  });

  test('reads only the rows of the range that the file holds', () => {
    const pkg = open(row(1, [c('A1', 'a')], 'ht="40"') + row(5, [c('A5', 'b')], 'ht="40"'));
    expect([...rows(pkg, 1, 4).keys()]).toEqual([4]);
  });

  test('holds every wrapped line at each zoom of the fit', () => {
    // Line heights round to whole pixels at each zoom, so the height at zoom 1 alone can cut the
    // last wrapped line at 90%.
    const zooms = [0.5, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 2, 4];
    const text = 'The quick brown fox jumps over the lazy dog, again and again, until the end.';
    for (const width of [8, 10, 12, 14, 17]) {
      const cols = `<cols><col min="1" max="1" width="${width}" customWidth="1"/></cols>`;
      const sheet = (attrs: string) => open(row(1, [c('A1', text, { s: WRAP })], attrs), { cols });
      const fitted = fitRowHeights(sheet('ht="400" customHeight="1"'), 'S', 0, 0, {
        measurer,
        profile,
        zooms,
      })!.get(0)!;
      for (const zoom of zooms) {
        const view = { width: 2000, height: 2000, scrollX: 0, scrollY: 0 };
        const auto = layoutSheet(sheet(''), 'S', view, { measurer, profile: zoomProfile(zoom) })!;
        // One pixel of the zoom is the noise of the rounding, and cuts no line.
        const pixel = 72 / (96 * zoom);
        expect(fitted).toBeGreaterThanOrEqual(auto.geometry.rows.sizeOf(0) - pixel - 1e-6);
      }
    }
  });

  test('keeps a row with no wrapped cell at its height at zoom 1', () => {
    const zooms = [0.5, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 2, 4];
    for (const size of [11, 12, 14, 18, 20]) {
      const one = readXlsx(
        buildXlsx({
          styles: {
            fonts: [`<font><sz val="${size}"/><name val="Calibri"/></font>`],
            cellXfs: [
              '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
            ],
          },
          sheets: [
            { name: 'S', rows: row(1, [c('A1', 'Hg', { s: 1 })], 'ht="40" customHeight="1"') },
          ],
        })
      );
      if (!one.ok) throw new Error('open');
      const plain = fitRowHeights(one.pkg, 'S', 0, 0, { measurer, profile })!.get(0);
      expect(fitRowHeights(one.pkg, 'S', 0, 0, { measurer, profile, zooms })!.get(0)).toBe(plain);
    }
  });

  test('reads back a raised height as fitted already', () => {
    const zooms = [0.5, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 2, 4];
    const text = 'The quick brown fox jumps over the lazy dog, again and again, until the end.';
    const cols = '<cols><col min="1" max="1" width="12" customWidth="1"/></cols>';
    const sheet = (attrs: string) => open(row(1, [c('A1', text, { s: WRAP })], attrs), { cols });
    const options = { measurer, profile, zooms };
    const fitted = fitRowHeights(sheet('ht="400" customHeight="1"'), 'S', 0, 0, options)!.get(0)!;
    // Raised above zoom 1, and stored with the 8 decimals of the file.
    expect(fitted).toBeGreaterThan(
      fitRowHeights(sheet('ht="400"'), 'S', 0, 0, { measurer, profile })!.get(0)!
    );
    const stored = (Math.round(fitted * 1e8) / 1e8).toFixed(8).replace(/\.?0+$/, '');
    expect(fitRowHeights(sheet(`ht="${stored}"`), 'S', 0, 0, options)!.size).toBe(0);
  });
});

describe('a fitted column, laid out again', () => {
  const ZOOMS = [1, 0.85, 1.25];
  /** A sheet whose columns take `widths` (0-based column to characters). */
  const sheet = (widths?: Map<number, number>) => {
    const cols = widths
      ? `<cols>${[...widths]
          .map(
            ([col, w]) => `<col min="${col + 1}" max="${col + 1}" width="${w}" customWidth="1"/>`
          )
          .join('')}</cols>`
      : undefined;
    const r = readXlsx(
      buildXlsx({
        rawStrings: [
          '<si><r><rPr><sz val="24"/><rFont val="Calibri"/></rPr><t>Big</t></r><r><rPr><sz val="8"/><rFont val="Calibri"/></rPr><t> and small text</t></r></si>',
        ],
        styles: {
          numFmts: ['<numFmt numFmtId="164" formatCode="0.00"/>'],
          cellXfs: [
            '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
            '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
          ],
          dxfs: ['<dxf><numFmt numFmtId="200" formatCode="0.000000"/></dxf>'],
        },
        sheets: [
          {
            name: 'S',
            rows: row(1, [
              c('A1', 98765),
              c('B1', 123456789),
              c('C1', 12345.678, { s: 1 }),
              c('D1', 45000, { s: 2 }),
              c('E1', 0, { t: 's' }),
              c('F1', 1234567),
              c('G1', 'plain text'),
            ]),
            ...(cols ? { cols } : {}),
            afterMerges:
              '<conditionalFormatting sqref="F1"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>',
          },
        ],
      })
    );
    if (!r.ok) throw new Error('open');
    return r.pkg;
  };

  test('shows each value unchanged and inside its text area at each zoom of the fit', () => {
    const widths = fitColumnWidths(sheet(), 'S', 0, 6, { measurer, profile, zooms: ZOOMS })!;
    expect(widths.size).toBe(7);
    const fitted = sheet(widths);
    const plain = sheet(new Map([0, 1, 2, 3, 4, 5, 6].map((col) => [col, 100])));
    for (const zoom of ZOOMS) {
      const options = { measurer, profile: zoomProfile(zoom) };
      const viewport = { width: 4000, height: 200, scrollX: 0, scrollY: 0 };
      const shown = (pkg: XlsxPackage) =>
        layoutSheet(pkg, 'S', viewport, options)!.panes.flatMap((p) => p.texts);
      // The text that a wide column shows is the text that the fitted column must show.
      const wide = new Map(
        layoutSheet(plain, 'S', { ...viewport }, options)!
          .panes.flatMap((p) => p.texts)
          .map((t) => [
            t.col,
            t.text.lines.map((l) => l.spans.map((sp) => sp.text).join('')).join(''),
          ])
      );
      for (const t of shown(fitted)) {
        const text = t.text.lines.map((l) => l.spans.map((sp) => sp.text).join('')).join('');
        expect({ col: t.col, zoom, text }).toEqual({ col: t.col, zoom, text: wide.get(t.col)! });
        expect(t.text.right).toBeLessThanOrEqual(t.text.areaRight + 1e-6);
        expect(t.text.left).toBeGreaterThanOrEqual(t.text.areaLeft - 1e-6);
      }
    }
  });

  test('keeps a fitted shrink-to-fit text at its size wherever the column starts', () => {
    const texts = [
      'shrink me please',
      'Wide header text',
      'abc',
      'Total amount due',
      'x',
      'Il1 mmm WWW',
    ];
    const shrink = (widths?: Map<number, number>) => {
      const cols = widths
        ? `<cols>${[...widths]
            .map(([c, w]) => `<col min="${c + 1}" max="${c + 1}" width="${w}" customWidth="1"/>`)
            .join('')}</cols>`
        : undefined;
      const r = readXlsx(
        buildXlsx({
          styles: {
            cellXfs: [
              '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment shrinkToFit="1"/></xf>',
            ],
          },
          sheets: [
            {
              name: 'S',
              rows: row(
                1,
                texts.map((t, i) => c(`${String.fromCharCode(66 + i)}1`, t, { s: 1 }))
              ),
              ...(cols ? { cols } : {}),
            },
          ],
        })
      );
      if (!r.ok) throw new Error('open');
      return r.pkg;
    };
    const zoom = 1.25;
    const widths = fitColumnWidths(shrink(), 'S', 1, texts.length, {
      measurer,
      profile,
      // The zoom steps from 50% that the editor fits at (`FIT_ZOOMS`).
      zooms: [0.5, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4],
    })!;
    const fonts = (pkg: XlsxPackage) =>
      layoutSheet(
        pkg,
        'S',
        { width: 4000, height: 200, scrollX: 0, scrollY: 0 },
        { measurer, profile: zoomProfile(zoom) }
      )!
        .panes.flatMap((p) => p.texts)
        .map((t) => JSON.stringify(t.text.lines.map((l) => l.spans.map((sp) => sp.style))));
    const wide = fonts(shrink(new Map(texts.map((_, i) => [i + 1, 100]))));
    // Column A moves each fitted column through many offsets; the noise of an offset in points
    // must not shrink a text that fits.
    for (let a = 0.5; a <= 40; a += 0.13)
      expect(fonts(shrink(new Map([[0, a], ...widths])))).toEqual(wide);
  });

  test('keeps each value unchanged wherever the column starts', () => {
    const widths = fitColumnWidths(sheet(), 'S', 1, 6, { measurer, profile, zooms: ZOOMS })!;
    const viewport = { width: 4000, height: 200, scrollX: 0, scrollY: 0 };
    const textsOf = (pkg: XlsxPackage, zoom: number) =>
      layoutSheet(pkg, 'S', viewport, { measurer, profile: zoomProfile(zoom) })!
        .panes.flatMap((p) => p.texts)
        .filter((t) => t.col > 0)
        .map((t) => t.text.lines.map((l) => l.spans.map((sp) => sp.text).join('')).join(''));
    const wide = ZOOMS.map((zoom) =>
      textsOf(sheet(new Map([1, 2, 3, 4, 5, 6].map((col) => [col, 100]))), zoom)
    );
    // Column A moves the start of each fitted column through many offsets in points.
    for (let a = 1; a <= 30; a += 0.37)
      ZOOMS.forEach((zoom, z) =>
        expect(textsOf(sheet(new Map([[0, a], ...widths])), zoom)).toEqual(wide[z]!)
      );
  });
});
