// The statistics of the browser bench.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import {
  growthOf,
  leaksOf,
  LEAK_LIMITS,
  median,
  MEMORY_WARMUP_COMMITS,
  p95,
  valuesOf,
  windowsOf,
  type Reading,
} from './measure.ts';

describe('the statistics of the browser bench', () => {
  test('the median and the p95 are samples of the list', () => {
    expect(median([5, 1, 3])).toBe(3);
    // The rank of DOCX Editor's `summarize`: ceil(0.95 n) - 1.
    expect(p95([...Array(20).keys()])).toBe(18);
    expect(p95([...Array(60).keys()])).toBe(56);
    expect(p95([...Array(100).keys()])).toBe(94);
    expect(p95([7])).toBe(7);
  });

  test('a steady rise is growth', () => {
    // Seven readings, five commits between two readings, 10 KiB for each commit.
    const readings = [0, 1, 2, 3, 4, 5, 6].map((i) => i * 50 * 1024);
    expect(growthOf(readings, 5)).toBeCloseTo(10 * 1024);
  });

  test('one high reading in the middle of the half is not growth', () => {
    const readings = [0, 0, 0, 0, 900 * 1024, 0, 0];
    expect(growthOf(readings, 5)).toBeLessThanOrEqual(0);
  });

  test('a heap that levels off grows far less than at the start', () => {
    const readings = [0, 400, 600, 700, 720, 730, 735].map((k) => k * 1024);
    // 80 KiB for each commit at the start, about 2 KiB over the second half.
    expect(growthOf(readings, 5)).toBeLessThan(3 * 1024);
  });

  test('the windows are the medians of the first and the last ten samples', () => {
    expect(windowsOf(Array.from({ length: 60 }, () => 100))).toEqual({ first: 100, last: 100 });
    // Each edit 5 ms slower than the one before.
    expect(windowsOf(Array.from({ length: 60 }, (_, i) => 100 + 5 * i))).toEqual({
      first: 125,
      last: 375,
    });
    expect(windowsOf([1, 2, 3])).toBeUndefined();
  });
});

describe('the leaks of the memory mode', () => {
  const at = (nodes: number, listeners: number, documents = 1): Reading => ({
    heapBytes: 0,
    nodes,
    listeners,
    documents,
    blobs: 0,
    fonts: 4,
    observers: 2,
    workers: 0,
    bitmaps: 0,
    intervals: 0,
    windowListeners: 9,
    documentListeners: 6,
  });

  test('nodes or listeners that keep growing are leaks, and a level count is not', () => {
    const level = [0, 1, 2, 3, 4, 5, 6].map(() => at(500, 40));
    expect(leaksOf(level, 5)).toEqual([]);
    // Six DOM nodes and two listeners for each commit.
    const growing = [0, 1, 2, 3, 4, 5, 6].map((i) => at(500 + 5 * i * 6, 40 + 5 * i * 2));
    expect(leaksOf(growing, 5)).toEqual([
      'DOM nodes +6.0 for each commit (limit 2)',
      'listeners +2.00 for each commit (limit 0.05)',
    ]);
  });

  test('one DOM node for each commit, as the memory bench measures, is not a leak', () => {
    expect(LEAK_LIMITS).toEqual({ heapKiB: 10, nodes: 2, listeners: 0.05 });
    const history = [0, 1, 2, 3, 4, 5, 6].map((i) => at(1947 + 5 * i, 242));
    expect(leaksOf(history, 5)).toEqual([]);
    // Exactly the limit passes, as in PPTX Editor's verdict; more than it fails.
    const atLimit = [0, 1, 2, 3, 4, 5, 6].map((i) => at(1947 + 10 * i, 242));
    expect(leaksOf(atLimit, 5)).toEqual([]);
    const over = [0, 1, 2, 3, 4, 5, 6].map((i) => at(1947 + 11 * i, 242));
    expect(leaksOf(over, 5)).toEqual(['DOM nodes +2.2 for each commit (limit 2)']);
    const more = [0, 1, 2, 3, 4, 5, 6].map((i) => at(1947 + 15 * i, 242));
    expect(leaksOf(more, 5)).toEqual(['DOM nodes +3.0 for each commit (limit 2)']);
  });

  test('documents that keep growing are a leak', () => {
    const growing = [0, 1, 2, 3, 4, 5, 6].map((i) => at(500, 40, 1 + 2 * i));
    expect(leaksOf(growing, 5)).toEqual(['documents 1 -> 13']);
    const once = [1, 1, 1, 2, 2, 2, 2].map((d) => at(500, 40, d));
    expect(leaksOf(once, 5)).toEqual([]);
    // Documents that grow at the start and then stay level, as PPTX Editor's verdict reads them.
    const levels = [1, 5, 9, 9, 9, 9, 9].map((d) => at(500, 40, d));
    expect(leaksOf(levels, 5)).toEqual([]);
  });

  test('a resource of the page probe that keeps growing is a leak', () => {
    // One observer for each commit that is never disconnected.
    const observers = [0, 1, 2, 3, 4, 5, 6].map((i) => ({ ...at(500, 40), observers: 2 + 5 * i }));
    expect(leaksOf(observers, 5)).toEqual(['observers 2 -> 32']);
    // One font face that the first commit adds, and no more.
    const fonts = [4, 5, 5, 5, 5, 5, 5].map((fonts) => ({ ...at(500, 40), fonts }));
    expect(leaksOf(fonts, 5)).toEqual([]);
  });

  test('a heap that keeps growing past 10 KiB for each commit is a leak', () => {
    const heap = (kib: number[]) => kib.map((k) => ({ ...at(500, 40), heapBytes: k * 1024 }));
    // 4 KiB for each commit, what a commit keeps with a full undo history.
    expect(leaksOf(heap([0, 20, 40, 60, 80, 100, 120]), 5)).toEqual([]);
    // Exactly the limit passes; 12 KiB for each commit fails.
    expect(leaksOf(heap([0, 50, 100, 150, 200, 250, 300]), 5)).toEqual([]);
    expect(leaksOf(heap([0, 60, 120, 180, 240, 300, 360]), 5)).toEqual([
      'heap +12 KiB for each commit (limit 10 KiB)',
    ]);
    // 80 KiB for each commit.
    expect(leaksOf(heap([0, 400, 800, 1200, 1600, 2000, 2400]), 5)).toEqual([
      'heap +80 KiB for each commit (limit 10 KiB)',
    ]);
  });

  test('the table gives the heap of the first reading and after the last commits', () => {
    const mib = (n: number): Reading => ({ ...at(500, 40), heapBytes: n * 1048576 });
    const row = {
      workbook: 'grid' as const,
      workbookSha256: 'a',
      cells: 1,
      samples: {},
      longestTaskMs: 0,
      peakMiB: 900,
      heapMiB: 200,
      growthKiBPerCommit: 18,
      readings: [mib(200), mib(224), mib(225)],
    };
    expect(valuesOf('memory', row).map((m) => [m.metric, m.value])).toEqual([
      ['heap MiB', 200],
      ['heap after edits MiB', 225],
      ['growth KiB per commit', 18],
      ['peak MiB', 900],
    ]);
  });

  test('the table of the edit mode gives the windows of each step', () => {
    const xs = Array.from({ length: 20 }, (_, i) => i + 1);
    const row = {
      workbook: 'grid' as const,
      workbookSha256: 'a',
      cells: 1,
      samples: { commit: xs },
      longestTaskMs: 0,
      peakMiB: 9,
    };
    expect(valuesOf('edit', row).map((m) => [m.metric, m.value])).toEqual([
      ['commit median ms', 11],
      ['commit p95 ms', 19],
      ['commit max ms', 20],
      ['commit first 10 median ms', 6],
      ['commit last 10 median ms', 16],
      ['peak MiB', 9],
    ]);
  });
});

test('the memory warmup fills the undo history of the store', () => {
  const store = readFileSync(
    join(import.meta.dir, '../../packages/core/src/store/store/document-store.ts'),
    'utf8'
  );
  expect(/const HISTORY_LIMIT = (\d+);/.exec(store)?.[1]).toBe(String(MEMORY_WARMUP_COMMITS));
});
