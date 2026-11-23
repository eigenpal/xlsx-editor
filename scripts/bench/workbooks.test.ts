// The bench workbooks at their default sizes open: each part stays under the XML part limits, so
// no part is refused or kept as bytes, and the timed work is the real work of each size.

import { expect, test } from 'bun:test';
import { readZip } from '../../packages/ooxml-core/src/zip.ts';
import {
  XML_DEFAULT_MAX_ELEMENTS,
  XML_HARD_MAX_BYTES,
} from '../../packages/ooxml-core/src/xml-reader.ts';
import {
  BENCH_SIZES,
  columnName,
  gridWorkbook,
  SHEET_COLS,
  SHEET_ROWS,
  sheetsWorkbook,
  stringsWorkbook,
} from './workbooks.ts';

/** The start tags of an XML part: each `<` that does not open an end tag or a declaration. */
function startTags(bytes: Uint8Array): number {
  let n = 0;
  for (let i = 0; i < bytes.length - 1; i += 1)
    if (bytes[i] === 0x3c && bytes[i + 1] !== 0x2f && bytes[i + 1] !== 0x3f) n += 1;
  return n;
}

/** The parts of `bytes`, with each XML part checked against the part limits. */
function partsUnderLimits(bytes: Uint8Array): ReadonlyMap<string, Uint8Array> {
  const read = readZip(bytes);
  if (!read.ok) throw new Error(`the zip did not read: ${read.reason}`);
  for (const [name, part] of read.entries) {
    if (!name.endsWith('.xml') && !name.endsWith('.rels')) continue;
    expect({ name, ok: part.length <= XML_HARD_MAX_BYTES }).toEqual({ name, ok: true });
    expect({ name, ok: startTags(part) <= XML_DEFAULT_MAX_ELEMENTS }).toEqual({ name, ok: true });
  }
  return read.entries;
}

const cellsOf = (part: Uint8Array | undefined): number =>
  new TextDecoder().decode(part).split('<c ').length - 1;

test('the grid workbook holds its default rows of 20 cells under the part limits', () => {
  const parts = partsUnderLimits(gridWorkbook());
  expect(cellsOf(parts.get('/xl/worksheets/sheet1.xml'))).toBe(
    BENCH_SIZES.gridRows * BENCH_SIZES.gridCols
  );
}, 60_000);

test('the sheets workbook holds 200 worksheets under the part limits', () => {
  const parts = partsUnderLimits(sheetsWorkbook());
  const sheets = [...parts.keys()].filter((n) => n.startsWith('/xl/worksheets/sheet'));
  expect(sheets.length).toBe(BENCH_SIZES.sheets);
  expect(cellsOf(parts.get('/xl/worksheets/sheet200.xml'))).toBe(SHEET_ROWS * SHEET_COLS);
}, 60_000);

test('the strings workbook holds its default shared strings under the part limits', () => {
  const parts = partsUnderLimits(stringsWorkbook());
  const sst = new TextDecoder().decode(parts.get('/xl/sharedStrings.xml'));
  expect(sst.split('<si>').length - 1).toBe(BENCH_SIZES.strings);
  expect(cellsOf(parts.get('/xl/worksheets/sheet1.xml'))).toBe(BENCH_SIZES.strings);
}, 60_000);

test('column names follow A to Z, then AA', () => {
  expect([0, 1, 25, 26, 27, 701, 702].map(columnName)).toEqual([
    'A',
    'B',
    'Z',
    'AA',
    'AB',
    'ZZ',
    'AAA',
  ]);
});
