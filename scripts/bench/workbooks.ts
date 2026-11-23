// The large workbooks of the speed benches, built at bench time, so the repository holds no large
// binary and no private data. Each build gives the same bytes.
//
// - `gridWorkbook`: one worksheet of `rows` x `cols` cells. Column A holds shared strings, the
//   other columns hold numbers. The default is 20,000 rows x 20 columns.
// - `sheetsWorkbook`: `sheets` worksheets of 100 x 10 numbers. The default is 200 sheets.
// - `stringsWorkbook`: a shared string table of `count` distinct items, and one worksheet whose
//   cells name each item once, 20 to a row. The default is 200,000 items.
//
// Each part stays under the XML part limit (`XML_HARD_MAX_BYTES`, 64 MiB) and the element limit
// (`XML_DEFAULT_MAX_ELEMENTS`), so each workbook opens and none is refused or kept as bytes.
// `workbooks.test.ts` checks this for the default sizes.
//
// Same role as PPTX Editor's `scripts/bench/text-deck.ts` at d4502552 (EigenPal), for workbooks.

import { buildXlsx, row, type FixtureSheet } from '../fixtures/builder.ts';

/**
 * The default sizes of the edit bench. At these sizes, two local runs of `bun run bench:edit`, with
 * its default 40 runs, took 3.2 to 3.3 GiB of memory at their peak and 30 to 31 s. Larger sizes
 * take memory in proportion, and `--rows` and `--strings` choose them.
 */
export const BENCH_SIZES = Object.freeze({
  gridRows: 20_000,
  gridCols: 20,
  sheets: 200,
  strings: 200_000,
});

/**
 * The sizes of the browser benches, smaller than `BENCH_SIZES` to keep a margin under the 3 GiB cap
 * of `scripts/hang/watchdog.ts`, which stops the browser. The cap is a hang limit and stays. At
 * 20,000 rows, a local run of the edit mode took 2159 MiB. At 250,000 strings, three runs of the
 * memory mode took 2294 to 2836 MiB, because each commit of a new text copies the shared string
 * table. Peaks on a CI runner, with an earlier form of these benches:
 * 50,000 rows went above the cap, 20,000 rows took up to 2048 MiB, 500,000 strings up to
 * 3066 MiB, and 250,000 strings up to 2405 MiB with 80 edits. At these sizes, two local runs of
 * the edit mode took up to 1977 MiB on the grid, 1338 MiB on the sheets, and 1829 MiB on the
 * strings workbook. Two local runs of the memory mode took up to 1468, 912, and 1275 MiB. The
 * browser benches use round sizes that keep a margin under the cap.
 */
export const BROWSER_SIZES = Object.freeze({
  gridRows: 15_000,
  sheets: BENCH_SIZES.sheets,
  strings: 125_000,
});

/** The name of the first sheet of every bench workbook. */
export const BENCH_SHEET = 'Data';

/** The column name of the zero-based column `index` (A, B, ..., Z, AA, ...). */
export function columnName(index: number): string {
  let name = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26))
    name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
}

/** The number in row `r`, column `x` (both zero-based): short, so a large sheet stays small. */
export const gridNumber = (r: number, x: number): number => (r * 31 + x * 7) % 1000;

/** The text of shared string `i`. */
export const sharedText = (i: number): string => `w${i}`;

/** The number of distinct shared strings in column A of `gridWorkbook`. */
const GRID_STRINGS = 1000;

/** `rows` of `cols` cells; column A names a shared string when `strings` is set. */
function gridRows(rows: number, cols: number, strings: boolean): string {
  const names = Array.from({ length: cols }, (_, x) => columnName(x));
  const out: string[] = [];
  for (let r = 0; r < rows; r += 1) {
    const n = r + 1;
    const cells: string[] = [];
    for (let x = 0; x < cols; x += 1) {
      if (x === 0 && strings) cells.push(`<c r="A${n}" t="s"><v>${r % GRID_STRINGS}</v></c>`);
      else cells.push(`<c r="${names[x]}${n}"><v>${gridNumber(r, x)}</v></c>`);
    }
    out.push(row(n, cells));
  }
  return out.join('');
}

/** One worksheet of `rows` x `cols` cells. */
export function gridWorkbook(
  rows: number = BENCH_SIZES.gridRows,
  cols: number = BENCH_SIZES.gridCols
): Uint8Array {
  return buildXlsx({
    sheets: [{ name: BENCH_SHEET, rows: gridRows(rows, cols, true) }],
    strings: Array.from({ length: GRID_STRINGS }, (_, i) => sharedText(i)),
  });
}

/** The rows of each worksheet of `sheetsWorkbook`. */
export const SHEET_ROWS = 100;
/** The columns of each worksheet of `sheetsWorkbook`. */
export const SHEET_COLS = 10;

/** `sheets` worksheets of `SHEET_ROWS` x `SHEET_COLS` numbers. */
export function sheetsWorkbook(sheets: number = BENCH_SIZES.sheets): Uint8Array {
  const body = gridRows(SHEET_ROWS, SHEET_COLS, false);
  const list: FixtureSheet[] = Array.from({ length: sheets }, (_, i) => ({
    name: i === 0 ? BENCH_SHEET : `Sheet ${i + 1}`,
    rows: body,
  }));
  return buildXlsx({ sheets: list });
}

/** A shared string table of `count` items, and a worksheet that names each item once. */
export function stringsWorkbook(
  count: number = BENCH_SIZES.strings,
  cols: number = BENCH_SIZES.gridCols
): Uint8Array {
  const names = Array.from({ length: cols }, (_, x) => columnName(x));
  const out: string[] = [];
  for (let r = 0; r * cols < count; r += 1) {
    const n = r + 1;
    const cells: string[] = [];
    for (let x = 0; x < cols && r * cols + x < count; x += 1)
      cells.push(`<c r="${names[x]}${n}" t="s"><v>${r * cols + x}</v></c>`);
    out.push(row(n, cells));
  }
  return buildXlsx({
    sheets: [{ name: BENCH_SHEET, rows: out.join('') }],
    rawStrings: Array.from({ length: count }, (_, i) => `<si><t>${sharedText(i)}</t></si>`),
  });
}
