import { beforeAll, describe, expect, test } from 'bun:test';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx } from '../../store/package/package.ts';
import { workbookIndex } from '../../store/workbook/workbook-index.ts';
import {
  autoRowHeight,
  fontPixels,
  gridMetrics,
  pixelWidth,
  rowSize,
  PRINT_PROFILE,
  SCREEN_PROFILE,
} from '../metrics.ts';
import type { TextMeasurer } from '../text/measurer.ts';
import { substituteMeasurer } from './test-fonts.ts';

let measurer: TextMeasurer;
beforeAll(async () => {
  measurer = await substituteMeasurer();
});

const calibri = (sizePt: number) => ({ family: 'Calibri', bold: false, italic: false, sizePt });

function workbook(defaultFont?: string) {
  const read = readXlsx(
    buildXlsx({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }],
      ...(defaultFont ? { defaultFont } : {}),
    })
  );
  if (!read.ok) throw new Error('read');
  return { pkg: read.pkg, index: workbookIndex(read.pkg)! };
}

describe('font pixels (screen geometry)', () => {
  test('Calibri 11: MDW 6, line height 10 + 3', () => {
    const f = fontPixels(measurer, calibri(11), SCREEN_PROFILE);
    expect([f.maxDigitWidth, f.ascent, f.descent, f.lineHeight]).toEqual([6, 10, 3, 13]);
  });

  test('Calibri 20: MDW 10 and a 26 point row on screen', () => {
    const f = fontPixels(measurer, calibri(20), SCREEN_PROFILE);
    expect(f.maxDigitWidth).toBe(10);
    expect(autoRowHeight(f, SCREEN_PROFILE)).toBe(26);
  });

  test('a Calibri 11 row is 15 points on screen and 14 points in print', () => {
    const f = fontPixels(measurer, calibri(11), SCREEN_PROFILE);
    expect(autoRowHeight(f, SCREEN_PROFILE)).toBe(15);
    expect(rowSize(autoRowHeight(f, PRINT_PROFILE), PRINT_PROFILE)).toBe(14);
    expect(rowSize(33.75, SCREEN_PROFILE)).toBe(33);
    expect(rowSize(33.75, PRINT_PROFILE)).toBe(31);
  });

  test('each advance rounds on its own, as the glyphs of basic-values.pdf start', () => {
    const f = fontPixels(measurer, calibri(11), PRINT_PROFILE);
    // `Shared text` advances 5 6 5 4 5 6 2 4 5 5 4 in the reference PDF: 51 points.
    expect(pixelWidth(measurer.advances(f.face, 'Shared text'), f.sizePx)).toBe(51);
    expect(pixelWidth(measurer.advances(f.face, 'TRUE'), f.sizePx)).toBe(23);
    expect(pixelWidth(measurer.advances(f.face, '1234568'), f.sizePx)).toBe(42);
  });
});

describe('gridMetrics', () => {
  test('the default font of the fixtures gives the reference metrics', () => {
    const { pkg, index } = workbook();
    expect(gridMetrics(pkg, index, measurer, PRINT_PROFILE)).toEqual({
      maxDigitWidth: 6,
      pixelsPerInch: 72,
      defaultRowHeight: 15,
      rowScale: 14 / 15,
    });
    expect(gridMetrics(pkg, index, measurer, SCREEN_PROFILE).rowScale).toBe(1);
  });

  test('a Calibri 20 default font widens the digit and heightens the row', () => {
    const { pkg, index } = workbook('<sz val="20"/><name val="Calibri"/>');
    const m = gridMetrics(pkg, index, measurer, SCREEN_PROFILE);
    expect([m.maxDigitWidth, m.defaultRowHeight]).toEqual([10, 26]);
  });
});
