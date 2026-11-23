// A value write keeps the look of a cell in a formatted row or column, by one rule (`buildCell`):
// a new cell takes the row format (`x:row/@s` with `customFormat`) or the column format
// (`x:col/@style`), and a cleared cell with no format of its own stays with `s="0"`, as
// `clearFormats` keeps it. ECMA-376 Part 1, 18.3.1.73 and 18.3.1.13.

import { describe, expect, test } from 'bun:test';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx, xmlPartOf, type XlsxPackage } from '../package/package.ts';
import { attr } from '../package/query.ts';
import { cellKey, sheetIndex } from '../workbook/sheet-index.ts';
import { columnFormatAt, columnFormats } from '../workbook/inherited-format.ts';
import type { Element } from '../package/tree.ts';
import { WorkbookDocumentStore } from '../store/document-store.ts';
import type { CellInput, WorkbookOp } from '../store/ops.ts';

/** Row 2 has a yellow fill (xf 1); column C shows dates (xf 2, `numFmtId` 14). */
function storeOf(): WorkbookDocumentStore {
  const read = readXlsx(
    buildXlsx({
      sheets: [
        {
          name: 'S',
          cols: '<cols><col min="3" max="3" width="10" style="2" customWidth="1"/></cols>',
          rows: row(2, [c('A2', 1)], 's="1" customFormat="1"') + row(3, [c('B3', 1)]),
        },
      ],
      styles: {
        fills: [
          '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>',
        ],
        cellXfs: [
          '<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>',
          '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
        ],
      },
    })
  );
  if (!read.ok) throw new Error(JSON.stringify(read.error));
  return new WorkbookDocumentStore(read.pkg);
}

const styleAt = (pkg: XlsxPackage, r: number, col: number) => {
  const part = xmlPartOf(pkg, '/xl/worksheets/sheet1.xml')!;
  const el = sheetIndex(part).cells.get(cellKey(r, col))?.el;
  return el ? (attr(el, 's') ?? '') : null;
};

const set = (ref: string, value: CellInput): WorkbookOp => ({
  op: 'setCellValue',
  sheet: 'S',
  ref,
  value,
});
const seven: CellInput = { type: 'number', value: 7 };

describe('a new cell takes the format of its row or column', () => {
  test('a typed value in a filled row keeps the fill; in a date column it is a date', () => {
    const store = storeOf();
    expect(store.transact([set('B2', seven)]).ok).toBe(true);
    expect(styleAt(store.package, 1, 1)).toBe('1');
    expect(store.transact([set('C5', seven)]).ok).toBe(true);
    expect(styleAt(store.package, 4, 2)).toBe('2');
    // A cell outside both keeps the default format.
    expect(store.transact([set('E5', seven)]).ok).toBe(true);
    expect(styleAt(store.package, 4, 4)).toBe('');
  });

  test('a batch of values, as a paste writes them, takes the same formats', () => {
    const store = storeOf();
    expect(store.transact([set('B2', seven), set('C6', seven), set('E6', seven)]).ok).toBe(true);
    expect(styleAt(store.package, 1, 1)).toBe('1');
    expect(styleAt(store.package, 5, 2)).toBe('2');
    expect(styleAt(store.package, 5, 4)).toBe('');
  });
});

describe('a cleared cell with no format keeps the default look', () => {
  test('alone, in a batch, and as a range, it stays with s="0"', () => {
    for (const ops of [
      [set('A2', { type: 'empty' })],
      [set('A2', { type: 'empty' }), set('E9', seven)],
      [{ op: 'clearValues', sheet: 'S', ref: 'A2' } as WorkbookOp],
    ]) {
      const store = storeOf();
      expect(store.transact(ops).ok).toBe(true);
      expect(styleAt(store.package, 1, 0)).toBe('0');
    }
  });

  test('a cleared cell outside a formatted row or column goes away', () => {
    const store = storeOf();
    expect(store.transact([set('B3', { type: 'empty' })]).ok).toBe(true);
    expect(styleAt(store.package, 2, 1)).toBeNull();
  });
});

describe('an inherited index past the cell formats', () => {
  test('counts as 0, so typing in that row or column is not refused', () => {
    const read = readXlsx(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            cols: '<cols><col min="3" max="3" width="10" style="999" customWidth="1"/></cols>',
            rows: row(5, [c('A5', 1)], 's="999" customFormat="1"'),
          },
        ],
      })
    );
    if (!read.ok) throw new Error(JSON.stringify(read.error));
    const store = new WorkbookDocumentStore(read.pkg);
    expect(store.transact([set('B5', seven)]).ok).toBe(true);
    expect(styleAt(store.package, 4, 1)).toBe('');
    expect(store.transact([set('C7', seven), set('D5', seven)]).ok).toBe(true);
    expect(styleAt(store.package, 6, 2)).toBe('');
    expect(store.transact([{ op: 'clearValues', sheet: 'S', ref: 'A5' }]).ok).toBe(true);
  });
});

describe('columnFormatAt', () => {
  /** The `x:worksheet` root of sheet S with the given `x:cols` children. */
  function rootWith(cols: string): Element {
    const read = readXlsx(
      buildXlsx({
        sheets: [{ name: 'S', cols: `<cols>${cols}</cols>`, rows: row(1, [c('A1', 1)]) }],
      })
    );
    if (!read.ok) throw new Error(JSON.stringify(read.error));
    return xmlPartOf(read.pkg, '/xl/worksheets/sheet1.xml')!.root;
  }

  test('gives the format that columnFormats gives for each column', () => {
    const root = rootWith(
      '<col min="1" max="2" style="3"/><col min="4" max="4" style="0"/>' +
        '<col min="5" max="9" style="7"/><col min="12" max="12" style="2"/>'
    );
    for (let col = 0; col < 14; col += 1)
      expect(columnFormatAt(root, col)).toBe(columnFormats(root, 0, 13).get(col));
  });

  test('overlapping ranges give one format for each column in any window', () => {
    // The first range by `min` wins: column C takes style 2 of `A:E`, not style 1 of `C`.
    const root = rootWith('<col min="3" max="3" style="1"/><col min="1" max="5" style="2"/>');
    for (let col = 0; col < 6; col += 1) {
      const want = columnFormatAt(root, col);
      expect(want).toBe(col < 5 ? 2 : undefined);
      expect(columnFormats(root, col, col).get(col)).toBe(want);
      expect(columnFormats(root, 0, 5).get(col)).toBe(want);
      expect(columnFormats(root, col, 5).get(col)).toBe(want);
    }
    // Ranges out of order, overlapping, and with style 0.
    const odd = rootWith(
      '<col min="3" max="8" style="7"/><col min="1" max="5" style="3"/><col min="4" max="4" style="0"/>'
    );
    for (let col = 0; col < 10; col += 1) {
      expect(columnFormatAt(odd, col)).toBe(col < 5 ? 3 : col < 8 ? 7 : undefined);
      expect(columnFormats(odd, col, col).get(col)).toBe(columnFormatAt(odd, col));
    }
  });

  test('16,384 lookups on 16,384 x:col elements build the spans once', () => {
    const cols = Array.from(
      { length: 16_384 },
      (_, i) => `<col min="${i + 1}" max="${i + 1}" style="${(i % 3) + 1}"/>`
    ).join('');
    const root = rootWith(cols);
    const start = performance.now();
    for (let col = 0; col < 16_384; col += 1) expect(columnFormatAt(root, col)).toBe((col % 3) + 1);
    // A walk over every x:col for each lookup takes seconds; a binary search takes milliseconds.
    expect(performance.now() - start).toBeLessThan(2000);
  });
});
