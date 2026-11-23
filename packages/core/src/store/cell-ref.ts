// A1 cell references: `x:c/@r` (`ST_CellRef`, sml.xsd) and the cells of an `ST_Ref`. The schema
// types are plain strings. The grid is bounded at `XFD1048576`, the limits of ECMA-376 Part 1,
// 18.17.5.1 (AGENTS.md, "Security"). A reference past it is invalid input at the parse boundary.
// The parser is linear and has no regex, so a hostile reference costs time in proportion to its
// length, which is capped.

/** A cell position: 0-based row and column indexes inside the grid. */
export interface CellAddress {
  readonly row: number;
  readonly col: number;
}

/** The last row of the grid, 1-based. */
export const MAX_ROWS = 1_048_576;
/** The last column of the grid, 1-based (`XFD`). */
export const MAX_COLS = 16_384;

/** The longest A1 reference: three column letters and seven row digits. */
const MAX_REF_LENGTH = 10;

/** A 1-based column number from its letters (`A` is 1), or 0 when they are not a column. */
export function columnNumber(letters: string): number {
  if (letters.length === 0 || letters.length > 3) return 0;
  let n = 0;
  for (let i = 0; i < letters.length; i += 1) {
    const c = letters.charCodeAt(i);
    if (c < 65 || c > 90) return 0;
    n = n * 26 + (c - 64);
  }
  return n <= MAX_COLS ? n : 0;
}

/** The letters of a 1-based column number. Throws outside the grid. */
export function columnLetters(column: number): string {
  if (!Number.isInteger(column) || column < 1 || column > MAX_COLS)
    throw new RangeError(`column ${column} is outside the grid`);
  let out = '';
  for (let n = column; n > 0; n = Math.floor((n - 1) / 26))
    out = String.fromCharCode(65 + ((n - 1) % 26)) + out;
  return out;
}

/**
 * Parse an uppercase A1 reference (`B7`) to a 0-based address, or `null` when it is not a cell
 * reference inside the grid. `$` markers are not accepted: cell `@r` values never hold them.
 */
export function parseCellRef(ref: string): CellAddress | null {
  if (ref.length < 2 || ref.length > MAX_REF_LENGTH) return null;
  let split = 0;
  while (split < ref.length) {
    const c = ref.charCodeAt(split);
    if (c < 65 || c > 90) break;
    split += 1;
  }
  const col = columnNumber(ref.slice(0, split));
  if (col === 0 || split === ref.length || ref.charCodeAt(split) === 48) return null;
  let row = 0;
  for (let i = split; i < ref.length; i += 1) {
    const c = ref.charCodeAt(i);
    if (c < 48 || c > 57) return null;
    row = row * 10 + (c - 48);
  }
  if (row < 1 || row > MAX_ROWS) return null;
  return { row: row - 1, col: col - 1 };
}

/** The A1 reference of a 0-based address. Throws outside the grid. */
export function formatCellRef(address: CellAddress): string {
  if (!Number.isInteger(address.row) || address.row < 0 || address.row >= MAX_ROWS)
    throw new RangeError(`row ${address.row} is outside the grid`);
  return `${columnLetters(address.col + 1)}${address.row + 1}`;
}
