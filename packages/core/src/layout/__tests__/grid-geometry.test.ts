import { describe, expect, test } from 'bun:test';
import { buildXlsx, c, row, type FixtureSheet } from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx, xmlPartOf } from '../../store/package/package.ts';
import { sheetByName, workbookIndex } from '../../store/workbook/workbook-index.ts';
import { sheetIndex } from '../../store/workbook/sheet-index.ts';
import {
  baseColumnWidth,
  columnPixels,
  columnWidthForPixels,
  rowHeightForPixels,
  sheetGeometry,
  type AutoRowHeight,
  type GridMetrics,
} from '../grid-geometry.ts';

/** The grid metrics of Calibri 11 on screen (the screen profile). */
const REFERENCE: GridMetrics = {
  maxDigitWidth: 6,
  pixelsPerInch: 72,
  defaultRowHeight: 15,
  rowScale: 1,
};

function geometry(
  sheet: Omit<FixtureSheet, 'name'>,
  metrics = REFERENCE,
  autoHeight?: AutoRowHeight
) {
  const read = readXlsx(buildXlsx({ sheets: [{ name: 'S', ...sheet }] }));
  if (!read.ok) throw new Error('read');
  const index = workbookIndex(read.pkg)!;
  const part = xmlPartOf(read.pkg, sheetByName(index, 'S')!.partName)!;
  return sheetGeometry(sheetIndex(part), metrics, autoHeight);
}

describe('column widths (Part 1, 18.3.1.13)', () => {
  test('the measured screen widths and the reference PDFs', () => {
    const widths = [1, 2, 5, 8.43, 10, 16, 30, 50];
    const cols = `<cols>${widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('')}</cols>`;
    const g = geometry({ rows: row(1, [c('A1', 1)]), cols });
    expect(widths.map((_, i) => g.columns.sizeOf(i))).toEqual([6, 12, 30, 51, 60, 96, 180, 300]);
    expect(g.columns.sizeOf(8)).toBe(53);
  });

  test('the spec scenario: width 10 at an MDW of 7 is 70 pixels', () => {
    expect(columnPixels(10, 7)).toBe(70);
  });

  test('default columns from baseColWidth', () => {
    expect(geometry({ rows: '' }).columns.sizeOf(0)).toBe(53);
    expect(
      geometry({ rows: '', formatPr: 'defaultRowHeight="15" baseColWidth="10"' }).columns.sizeOf(0)
    ).toBe(65);
    // Calibri 20: MDW 10, measured 87 points.
    expect(columnPixels(baseColumnWidth(8, 10), 10)).toBe(87);
  });

  test('a hidden column has zero width, and the next column starts where it starts', () => {
    const g = geometry({
      rows: '',
      cols: '<cols><col min="3" max="3" width="10" hidden="1"/></cols>',
    });
    expect(g.columns.sizeOf(2)).toBe(0);
    expect(g.columns.offsetOf(3)).toBe(g.columns.offsetOf(2));
  });

  test('a full-width column range is one segment', () => {
    const g = geometry({
      rows: '',
      cols: '<cols><col min="1" max="16384" width="2" customWidth="1"/></cols>',
    });
    expect(g.columns.total).toBe(12 * 16_384);
  });
});

describe('row heights', () => {
  test('measured custom heights keep whole pixels, rounded down', () => {
    const rows = [
      row(1, [c('A1', 1)]),
      row(2, [], 'ht="10" customHeight="1"'),
      row(3, [], 'ht="20" customHeight="1"'),
      row(4, [], 'ht="33.75" customHeight="1"'),
    ].join('');
    const g = geometry({ rows });
    expect([0, 1, 2, 3, 4].map((i) => g.rows.sizeOf(i))).toEqual([15, 10, 20, 33, 15]);
  });

  test('the default font sets the default row height unless the sheet marks its own', () => {
    expect(geometry({ rows: '' }, { ...REFERENCE, defaultRowHeight: 26 }).rows.sizeOf(0)).toBe(26);
    expect(
      geometry(
        { rows: '', formatPr: 'defaultRowHeight="18" customHeight="1"' },
        { ...REFERENCE, defaultRowHeight: 26 }
      ).rows.sizeOf(0)
    ).toBe(18);
  });

  test('a hidden row has zero height', () => {
    const g = geometry({ rows: row(2, [c('A2', 1)], 'hidden="1"') });
    expect(g.rows.sizeOf(1)).toBe(0);
    expect(g.rows.offsetOf(2)).toBe(15);
  });

  test('96 pixels per inch snaps heights to 0.75 points', () => {
    const g = geometry(
      { rows: row(1, [], 'ht="14.4" customHeight="1"') },
      { maxDigitWidth: 7, pixelsPerInch: 96, defaultRowHeight: 15, rowScale: 1 }
    );
    expect(g.rows.sizeOf(0)).toBe(14.25);
  });

  test('a row without @ht takes the height of its fonts, and a row with @ht keeps it', () => {
    const rows = [row(1, [c('A1', 1)]), row(2, [c('B2', 1)], 'ht="12" customHeight="1"')].join('');
    const g = geometry({ rows }, REFERENCE, () => 26);
    expect([g.rows.sizeOf(0), g.rows.sizeOf(1), g.rows.sizeOf(2)]).toEqual([26, 12, 15]);
    // A custom default row height turns the font height off.
    const custom = geometry(
      { rows, formatPr: 'defaultRowHeight="18" customHeight="1"' },
      REFERENCE,
      () => 26
    );
    expect(custom.rows.sizeOf(0)).toBe(18);
  });

  test('grid-structure.pdf: print scales every row by 14 / 15 and rounds down', () => {
    const rows = [
      row(1, [c('A1', 1)]),
      row(2, [], 'ht="10" customHeight="1"'),
      row(3, [], 'ht="20" customHeight="1"'),
      row(4, [], 'ht="33.75" customHeight="1"'),
    ].join('');
    const g = geometry({ rows }, { ...REFERENCE, rowScale: 14 / 15 });
    expect([0, 1, 2, 3, 4].map((i) => g.rows.sizeOf(i))).toEqual([14, 9, 18, 31, 14]);
  });
});

describe('the inverse conversions of a resize drag', () => {
  test('columnWidthForPixels gives back every whole pixel width', () => {
    for (const mdw of [5, 6, 7, 8, 9, 11, 13, 20, 40])
      for (let p = 0; p <= 1500; p += 1) {
        const w = columnWidthForPixels(p, mdw);
        if (w >= 255) break;
        expect(columnPixels(w, mdw)).toBe(p);
        if (p > 0) expect(columnPixels(w - 1 / 256, mdw)).toBeLessThan(p);
      }
    expect(columnWidthForPixels(64, 7)).toBe(9.07421875);
    expect(columnWidthForPixels(10_000, 7)).toBe(255);
  });

  test('rowHeightForPixels gives back every whole pixel height', () => {
    for (const metrics of [
      { pixelsPerInch: 96, rowScale: 1 },
      { pixelsPerInch: 72, rowScale: 14 / 15 },
      { pixelsPerInch: 96 * 1.75, rowScale: 1 },
    ]) {
      const ptPerPx = 72 / metrics.pixelsPerInch;
      for (let p = 0; p <= 400; p += 1) {
        const ht = rowHeightForPixels(p, metrics);
        if (ht >= 409) break;
        expect(Math.floor((ht / ptPerPx) * metrics.rowScale + 1e-9)).toBe(p);
      }
    }
    expect(rowHeightForPixels(27, { pixelsPerInch: 96, rowScale: 1 })).toBe(20.25);
    expect(rowHeightForPixels(5000, { pixelsPerInch: 72, rowScale: 1 })).toBe(409);
  });
});
