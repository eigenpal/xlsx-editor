import { describe, expect, test } from 'bun:test';
import { buildXlsx, c, row, type FixtureSheet } from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx, xmlPartOf } from '../../store/package/package.ts';
import { sheetByName, workbookIndex } from '../../store/workbook/workbook-index.ts';
import { AxisIndex } from '../axis.ts';
import {
  frozenPanes,
  mergeAnchoredAt,
  mergeAt,
  mergeCoverage,
  mergedRanges,
  mergesIn,
  parseRange,
  visibleSpan,
  windowPanes,
  type CellRange,
} from '../panes.ts';

function worksheet(sheet: Omit<FixtureSheet, 'name'>) {
  const read = readXlsx(buildXlsx({ sheets: [{ name: 'S', ...sheet }] }));
  if (!read.ok) throw new Error('read');
  const index = workbookIndex(read.pkg)!;
  return xmlPartOf(read.pkg, sheetByName(index, 'S')!.partName)!.root;
}

const rows = new AxisIndex(1_048_576, 15, []);
const cols = new AxisIndex(16_384, 50, []);

describe('frozen panes', () => {
  test('a frozen pane gives its split; a split pane that is not frozen gives none', () => {
    const frozen = worksheet({
      rows: '',
      viewChildren:
        '<pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/>',
    });
    expect(frozenPanes(frozen)).toEqual({ rows: 2, cols: 1 });
    const split = worksheet({
      rows: '',
      viewChildren: '<pane xSplit="1200" ySplit="900" topLeftCell="B3" state="split"/>',
    });
    expect(frozenPanes(split)).toEqual({ rows: 0, cols: 0 });
  });

  test('the spec scenario: rows 1 and 2 stay while the rest scrolls down 50 rows', () => {
    const viewport = { width: 500, height: 300, scrollX: 0, scrollY: 50 * 15 };
    const panes = windowPanes(rows, cols, { rows: 2, cols: 0 }, viewport);
    expect(panes).toHaveLength(2);
    const [top, body] = panes;
    expect(top!.range.firstRow).toBe(0);
    expect(top!.range.lastRow).toBe(1);
    expect(top!.y).toBe(0);
    // The body starts 50 rows below row 3, right under the frozen rows.
    expect(body!.range.firstRow).toBe(52);
    expect(body!.y).toBe(30);
    expect(body!.range.lastRow).toBe(52 + Math.ceil((300 - 30) / 15) - 1);
  });

  test('four panes with frozen rows and columns, and a partly scrolled first body cell', () => {
    const panes = windowPanes(
      rows,
      cols,
      { rows: 1, cols: 1 },
      {
        width: 200,
        height: 100,
        scrollX: 25,
        scrollY: 0,
      }
    );
    expect(panes.map((p) => [p.range.firstRow, p.range.firstCol, p.x, p.y])).toEqual([
      [0, 0, 0, 0],
      [0, 1, 25, 0],
      [1, 0, 0, 15],
      [1, 1, 25, 15],
    ]);
  });
});

describe('merged ranges', () => {
  test('the spec scenario: B2:D4 is one range, found from any of its cells', () => {
    const root = worksheet({ rows: row(2, [c('B2', 'Merged')]), merges: ['B2:D4', 'F1:F1'] });
    const index = mergedRanges(root);
    expect(index.ranges).toEqual([{ firstRow: 1, firstCol: 1, lastRow: 3, lastCol: 3 }]);
    // A one-cell merge is no merge.
    expect(index.dropped).toBe(1);
    expect(mergeAt(index, 3, 3)).toEqual(index.ranges[0]);
    expect(mergeAt(index, 4, 3)).toBeUndefined();
    expect(mergesIn(index, { firstRow: 3, firstCol: 0, lastRow: 9, lastCol: 1 })).toHaveLength(1);
    expect(mergesIn(index, { firstRow: 4, firstCol: 0, lastRow: 9, lastCol: 9 })).toHaveLength(0);
  });

  test('a reference that does not parse is dropped, never expanded', () => {
    expect(parseRange('A1:XFD1048576')).toEqual({
      firstRow: 0,
      firstCol: 0,
      lastRow: 1_048_575,
      lastCol: 16_383,
    });
    expect(parseRange('A0:B2')).toBeUndefined();
    expect(parseRange('B2:A1')).toEqual({ firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 });
  });
});

describe('visibleSpan', () => {
  test('a hidden column takes no space, and the next column starts where it starts', () => {
    const hidden = new AxisIndex(16_384, 50, [{ first: 2, last: 2, size: 0 }]);
    expect(hidden.offsetOf(3)).toBe(hidden.offsetOf(2));
    expect(visibleSpan(hidden, 100, 60)).toEqual({ first: 3, last: 4 });
  });
});

describe('merge anchors', () => {
  test('a lookup of the merge that a cell anchors costs O(1), however many merges exist', () => {
    // 100,000 merges C1:D1 to C100000:D100000, and 16,000 lookups of A1 to A16000, none of
    // which anchors a merge: a scan of the merges for each lookup reads 1.6e9 ranges.
    const merges = Array.from({ length: 100_000 }, (_, k) => `C${k + 1}:D${k + 1}`);
    const index = mergedRanges(worksheet({ rows: '', merges }));
    expect(index.ranges).toHaveLength(100_000);
    const started = performance.now();
    let found = 0;
    for (let r = 0; r < 16_000; r += 1) if (mergeAnchoredAt(index, r, 0)) found += 1;
    expect(found).toBe(0);
    expect(performance.now() - started).toBeLessThan(1000);
    expect(mergeAnchoredAt(index, 4, 2)).toEqual({
      firstRow: 4,
      lastRow: 4,
      firstCol: 2,
      lastCol: 3,
    });
    expect(mergeAnchoredAt(index, 4, 3)).toBeUndefined();
  });
});

describe('merge coverage', () => {
  const inside = (ms: readonly CellRange[], r: number, c: number) =>
    ms.some((m) => r >= m.firstRow && r <= m.lastRow && c >= m.firstCol && c <= m.lastCol);

  test('staggered merges give the answer of a scan of every merge, as rows go back or skip', () => {
    let seed = 1;
    const rnd = () => (seed = (seed * 48271) % 2147483647) / 2147483647;
    const ms: CellRange[] = [];
    for (let k = 0; k < 400; k += 1) {
      const firstRow = Math.floor(rnd() * 300);
      const firstCol = k * 3;
      ms.push({
        firstRow,
        lastRow: firstRow + Math.floor(rnd() * 20),
        firstCol,
        lastCol: firstCol + Math.floor(rnd() * 3),
      });
    }
    ms.sort((a, b) => a.firstRow - b.firstRow);
    const test = mergeCoverage(ms);
    let bad = 0;
    // Every row, then rows that go back, then rows that skip ahead past whole merges.
    const passes = [
      [0, 330, 1],
      [100, 200, 1],
      [0, 330, 13],
    ] as const;
    for (const [from, to, by] of passes)
      for (let r = from; r < to; r += by)
        for (let c = 0; c < 1250; c += 7) if (test(r, c) !== inside(ms, r, c)) bad += 1;
    expect(bad).toBe(0);
  });

  test('a row that no merge starts or ends at costs no sort, however many merges reach it', () => {
    // 16,383 merges B1:B16000 to XFD1:XFD16000, and one test on each of 16,000 rows. A sort of
    // every merge on each row took about 2 s; the rows after the first sort nothing.
    const ms: CellRange[] = Array.from({ length: 16_383 }, (_, k) => ({
      firstRow: 0,
      lastRow: 15_999,
      firstCol: k + 1,
      lastCol: k + 1,
    }));
    const test = mergeCoverage(ms);
    const started = performance.now();
    let covered = 0;
    for (let r = 0; r < 16_000; r += 1) if (test(r, 0)) covered += 1;
    expect(covered).toBe(0);
    expect(performance.now() - started).toBeLessThan(500);
    expect(test(15_999, 16_383)).toBe(true);
    expect(test(16_000, 1)).toBe(false);
  });
});
