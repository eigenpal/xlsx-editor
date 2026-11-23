// The limit on the wrapped text that one sheet measures for its row heights.

import { beforeAll, describe, expect, test } from 'bun:test';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx, type XlsxPackage } from '../../store/package/package.ts';
import { WorkbookDocumentStore } from '../../store/store/document-store.ts';
import { MAX_CELL_LINES } from '../cell-text.ts';
import { PRINT_PROFILE } from '../metrics.ts';
import { WRAP_MEASURE_BUDGET_CHARS } from '../row-height.ts';
import { layoutSheet } from '../sheet-layout.ts';
import type { TextMeasurer } from '../text/measurer.ts';
import { substituteMeasurer } from './test-fonts.ts';

let measurer: TextMeasurer;
beforeAll(async () => {
  measurer = await substituteMeasurer();
});

const VIEW = { width: 540, height: 720, scrollX: 0, scrollY: 0 };
const WRAP =
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf>';
// Narrow letters: the length alone gives more lines than layout breaks.
const SHORT = 'il '.repeat(100);

/**
 * Rows of wrapped text: `long` rows that use the whole budget, and then one row of `last`, a text or
 * the XML of a cell at a reference.
 */
function workbook(long: number, last: string | ((ref: string) => string) = SHORT): XlsxPackage {
  const text = 'x'.repeat(Math.ceil(WRAP_MEASURE_BUDGET_CHARS / 2));
  const rows = Array.from({ length: long }, (_, r) => row(r + 1, [c(`A${r + 1}`, text, { s: 1 })]));
  const ref = `A${long + 1}`;
  rows.push(row(long + 1, [typeof last === 'string' ? c(ref, last, { s: 1 }) : last(ref)]));
  const read = readXlsx(
    buildXlsx({ sheets: [{ name: 'S', rows: rows.join('') }], styles: { cellXfs: [WRAP] } })
  );
  if (!read.ok) throw new Error('read');
  return read.pkg;
}

const heightOf = (pkg: XlsxPackage, r: number) => {
  const out = layoutSheet(pkg, 'S', VIEW, { measurer, profile: PRINT_PROFILE });
  if (!out) throw new Error('no sheet');
  return out.geometry.rows.sizeOf(r);
};

describe('the wrapped text that a sheet measures', () => {
  test('a cell inside the budget takes the lines that layout breaks', () => {
    // 14 points for each Calibri 11 line in print: the text breaks at its spaces.
    const height = heightOf(workbook(0), 0);
    expect(height % 14).toBe(0);
    expect(height).toBeGreaterThan(14);
  });

  test('a cell past the budget takes its lines from its length', () => {
    const pkg = workbook(2);
    const estimated = heightOf(pkg, 2);
    expect(estimated % 14).toBe(0);
    expect(estimated / 14).toBeLessThanOrEqual(MAX_CELL_LINES);
    expect(estimated).toBeGreaterThan(heightOf(workbook(0), 0));
  });

  test('a cell past the budget starts a line at each line break', () => {
    const breaks = 'a\nb\r\nc\rd\ne';
    const measured = heightOf(workbook(0, breaks), 0);
    expect(measured).toBe(5 * 14);
    expect(heightOf(workbook(2, breaks), 2)).toBe(measured);
  });

  test('a final line break starts no line past the budget, as in layout', () => {
    const measured = heightOf(workbook(0, 'a\n'), 0);
    expect(measured).toBe(14);
    expect(heightOf(workbook(2, 'a\n'), 2)).toBe(measured);
  });

  test('a CR at the end of a run and an LF in the next run start one line', () => {
    const runs = (ref: string) =>
      `<c r="${ref}" s="1" t="inlineStr"><is><r><t>a&#13;</t></r><r><t>&#10;b</t></r></is></c>`;
    const measured = heightOf(workbook(0, runs), 0);
    expect(measured).toBe(2 * 14);
    expect(heightOf(workbook(2, runs), 2)).toBe(measured);
  });

  test('an edit elsewhere keeps the height, also when earlier cells come from the cache', () => {
    const store = new WorkbookDocumentStore(workbook(2));
    const before = heightOf(store.package, 2);
    const edit = store.transact([
      { op: 'setCellValue', sheet: 'S', ref: 'C9', value: { type: 'number', value: 1 } },
    ]);
    expect(edit.ok).toBe(true);
    expect(heightOf(store.package, 2)).toBe(before);
  });

  test('a cell that an earlier layout measured takes its lines from its length past the budget', () => {
    const store = new WorkbookDocumentStore(workbook(2));
    // With fixed heights, the long rows spend no budget, so the short cell is measured.
    const fixed = store.transact([
      { op: 'setRowHeight', sheet: 'S', first: 1, last: 2, height: 15 },
    ]);
    expect(fixed.ok).toBe(true);
    expect(heightOf(store.package, 2)).toBe(heightOf(workbook(0), 0));
    // The undo gives the long rows back the budget, and the short cell keeps its element.
    expect(store.undo().ok).toBe(true);
    expect(heightOf(store.package, 2)).toBe(heightOf(workbook(2), 2));
  });
});
