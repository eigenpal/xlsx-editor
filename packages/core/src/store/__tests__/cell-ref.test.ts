import { describe, expect, test } from 'bun:test';
import { columnLetters, columnNumber, formatCellRef, parseCellRef } from '../cell-ref.ts';

describe('A1 cell references', () => {
  test('parse and format the grid corners', () => {
    expect(parseCellRef('A1')).toEqual({ row: 0, col: 0 });
    expect(parseCellRef('XFD1048576')).toEqual({ row: 1048575, col: 16383 });
    expect(formatCellRef({ row: 1048575, col: 16383 })).toBe('XFD1048576');
    expect(formatCellRef({ row: 6, col: 27 })).toBe('AB7');
  });

  test('refuse references outside the grid or malformed', () => {
    for (const bad of ['', 'A', '1', 'A0', 'A01', 'XFE1', 'A1048577', 'a1', '$A$1', 'A1B', 'AAAA1'])
      expect(parseCellRef(bad)).toBeNull();
  });

  test('a long hostile reference costs no work past the length cap', () => {
    expect(parseCellRef('A'.repeat(100_000))).toBeNull();
  });

  test('column letters round-trip', () => {
    for (const n of [1, 26, 27, 52, 702, 703, 16384])
      expect(columnNumber(columnLetters(n))).toBe(n);
    expect(() => columnLetters(16385)).toThrow(RangeError);
  });
});
