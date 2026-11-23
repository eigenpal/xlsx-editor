import { describe, expect, test } from 'bun:test';
import { AxisIndex } from '../axis.ts';

describe('AxisIndex', () => {
  const rows = new AxisIndex(1_048_576, 15, [
    { first: 1, last: 1, size: 30 },
    { first: 3, last: 4, size: 0 },
    { first: 1_048_575, last: 1_048_575, size: 20 },
  ]);

  test('offsets add the default size and the segment sizes', () => {
    expect(rows.offsetOf(0)).toBe(0);
    expect(rows.offsetOf(1)).toBe(15);
    expect(rows.offsetOf(2)).toBe(45);
    expect(rows.offsetOf(3)).toBe(60);
    expect(rows.offsetOf(5)).toBe(60);
    expect(rows.sizeOf(4)).toBe(0);
  });

  test('the offset of the last row is one search, not a walk', () => {
    const started = performance.now();
    expect(rows.offsetOf(1_048_575)).toBe(15 * (1_048_576 - 4) + 30);
    expect(rows.total).toBe(15 * (1_048_576 - 4) + 30 + 20);
    expect(performance.now() - started).toBeLessThan(5);
  });

  test('indexAt finds the row that holds an offset and skips hidden rows', () => {
    expect(rows.indexAt(0)).toBe(0);
    expect(rows.indexAt(14.9)).toBe(0);
    expect(rows.indexAt(15)).toBe(1);
    expect(rows.indexAt(44)).toBe(1);
    expect(rows.indexAt(45)).toBe(2);
    expect(rows.indexAt(60)).toBe(5);
    expect(rows.indexAt(rows.total + 100)).toBe(1_048_575);
    for (const i of [0, 1, 2, 5, 6, 1000, 1_048_574, 1_048_575])
      expect(rows.indexAt(rows.offsetOf(i))).toBe(i);
  });

  test('overlapping and out-of-range segments are clipped', () => {
    const cols = new AxisIndex(16_384, 48, [
      { first: 2, last: 5, size: 10 },
      { first: 4, last: 8, size: 20 },
      { first: -5, last: 0, size: 30 },
      { first: 16_000, last: 99_999_999, size: 1 },
      { first: 7, last: 3, size: 5 },
    ]);
    expect([0, 1, 2, 5, 6, 8, 9].map((i) => cols.sizeOf(i))).toEqual([30, 48, 10, 10, 20, 20, 48]);
    expect(cols.offsetOf(16_384)).toBe(cols.total);
    expect(cols.sizeOf(16_383)).toBe(1);
  });

  test('a hidden first row starts the visible axis at the next row', () => {
    const hidden = new AxisIndex(100, 15, [{ first: 0, last: 1, size: 0 }]);
    expect(hidden.indexAt(0)).toBe(2);
    expect(hidden.offsetOf(2)).toBe(0);
  });

  test('the spec scenario: scroll to row 1048576 with one explicit row height', () => {
    // A counting index: every lookup must stay a search over the one segment.
    const axis = new AxisIndex(1_048_576, 14, [{ first: 9, last: 9, size: 40 }]);
    const last = 1_048_575;
    const offset = axis.offsetOf(last);
    expect(offset).toBe(14 * (1_048_576 - 2) + 40);
    expect(axis.indexAt(offset)).toBe(last);
    expect(axis.indexAt(offset + 13.9)).toBe(last);
    const started = performance.now();
    for (let i = 0; i < 10_000; i += 1) axis.indexAt(offset - i * 14);
    expect(performance.now() - started).toBeLessThan(50);
  });
});
