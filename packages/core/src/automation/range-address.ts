// The range addresses that the automation host reads: one cell (`A1`) or a rectangle (`A1:B2`)
// of one worksheet, in either letter case, with optional `$` marks. Every other form of an
// address (a sheet prefix, a whole row or column, a defined name) is outside the subset and is
// refused as `unsupported`, never read as something near it.

import { MAX_COLS, MAX_ROWS } from '../store/index.ts';

/** The most cells that one range operation reads or writes. */
export const MAX_RANGE_CELLS = 100_000;

/** The longest address text that the host reads. */
const MAX_ADDRESS_LENGTH = 32;

/** A rectangle of cells, 0-based and inclusive. */
export interface RangeBounds {
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}

export type RangeAddressResult =
  | { readonly ok: true; readonly bounds: RangeBounds }
  | { readonly ok: false; readonly code: 'invalid-argument' | 'unsupported' };

/** A cell reference with optional `$` marks: bounded, so it cannot backtrack. */
const CELL = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,8})$/;

/** The 0-based row and column of one cell reference, or undefined outside the grid. */
function cellOf(text: string): { row: number; col: number } | undefined | null {
  const match = CELL.exec(text);
  if (!match) return null;
  const letters = match[1]!.toUpperCase();
  let col = 0;
  for (let i = 0; i < letters.length; i += 1) col = col * 26 + (letters.charCodeAt(i) - 64);
  const row = Number(match[2]);
  if (col < 1 || col > MAX_COLS || row < 1 || row > MAX_ROWS) return undefined;
  return { row: row - 1, col: col - 1 };
}

/**
 * Read an address. A form that is not a cell or a rectangle is `unsupported`; a cell outside the
 * grid, or text that is not a string, is `invalid-argument`.
 */
export function parseRangeAddress(address: unknown): RangeAddressResult {
  if (typeof address !== 'string' || address.length === 0)
    return { ok: false, code: 'invalid-argument' };
  if (address.length > MAX_ADDRESS_LENGTH) return { ok: false, code: 'unsupported' };
  const parts = address.split(':');
  if (parts.length > 2) return { ok: false, code: 'unsupported' };
  const first = cellOf(parts[0]!);
  const last = parts.length === 2 ? cellOf(parts[1]!) : first;
  if (first === null || last === null) return { ok: false, code: 'unsupported' };
  if (first === undefined || last === undefined) return { ok: false, code: 'invalid-argument' };
  return {
    ok: true,
    bounds: {
      top: Math.min(first.row, last.row),
      left: Math.min(first.col, last.col),
      bottom: Math.max(first.row, last.row),
      right: Math.max(first.col, last.col),
    },
  };
}

/** The number of cells of a rectangle. */
export function boundsCells(bounds: RangeBounds): number {
  return (bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1);
}
