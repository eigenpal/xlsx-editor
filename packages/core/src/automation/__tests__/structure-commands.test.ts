// The structural commands lower a selection to one store op (task 7.5).

import { describe, expect, test } from 'bun:test';
import { MAX_COLS, MAX_ROWS } from '../../store/index.ts';
import { isWholeSheet, selectionBand, structureOp } from '../index.ts';

const sel = (r0: number, c0: number, r1: number, c1: number, active = { row: r0, col: c0 }) => ({
  sheetIndex: 0,
  active,
  anchor: active,
  range: { start: { row: r0, col: c0 }, end: { row: r1, col: c1 } },
});

describe('structural commands', () => {
  test('a selection is whole rows, whole columns, or neither', () => {
    expect(selectionBand(sel(2, 0, 4, MAX_COLS - 1))).toBe('rows');
    expect(selectionBand(sel(0, 1, MAX_ROWS - 1, 2))).toBe('columns');
    expect(selectionBand(sel(0, 0, MAX_ROWS - 1, MAX_COLS - 1))).toBeNull();
    expect(selectionBand(sel(1, 1, 3, 3))).toBeNull();
    expect(isWholeSheet(sel(0, 0, MAX_ROWS - 1, MAX_COLS - 1))).toBe(true);
    expect(isWholeSheet(sel(0, 0, MAX_ROWS - 1, MAX_COLS - 2))).toBe(false);
  });

  test('rows of a range or of whole rows; the active row in whole columns', () => {
    expect(structureOp('S', sel(2, 0, 4, MAX_COLS - 1), 'deleteRows')).toEqual({
      op: 'deleteRows',
      sheet: 'S',
      start: 3,
      count: 3,
    });
    expect(structureOp('S', sel(1, 1, 3, 3), 'insertRows')).toMatchObject({ start: 2, count: 3 });
    expect(
      structureOp('S', sel(0, 1, MAX_ROWS - 1, 2, { row: 6, col: 1 }), 'insertRows')
    ).toMatchObject({ start: 7, count: 1 });
    expect(
      structureOp('S', sel(2, 0, 4, MAX_COLS - 1, { row: 2, col: 5 }), 'insertColumns')
    ).toMatchObject({ op: 'insertColumns', start: 6, count: 1 });
    expect(structureOp('S', sel(0, 1, MAX_ROWS - 1, 2), 'deleteColumns')).toMatchObject({
      start: 2,
      count: 2,
    });
  });
});
