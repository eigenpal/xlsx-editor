// The DOM-free part of the `contracts` lane: declaration-only public types. Nothing here names a
// DOM type, so the DOM-free lanes import this module, never `index.ts`.

/**
 * One cell of the grid, by 0-based row and column.
 *
 * @public
 */
export interface GridPoint {
  readonly row: number;
  readonly col: number;
}

/**
 * A rectangle of cells, inclusive at both ends, with `start` at the top left.
 *
 * @public
 */
export interface GridRange {
  readonly start: GridPoint;
  readonly end: GridPoint;
}

/**
 * The selection on one sheet: the active cell, and the range that the selection covers. The
 * anchor is where an extension started; the active cell stays where it was.
 *
 * @public
 */
export interface GridSelection {
  readonly sheetIndex: number;
  readonly active: GridPoint;
  readonly anchor: GridPoint;
  readonly range: GridRange;
}
