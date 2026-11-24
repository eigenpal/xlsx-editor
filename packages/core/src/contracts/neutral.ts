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
 * A selection can hold several ranges, as `x:selection/@sqref` (`ST_Sqref`, ECMA-376 Part 1,
 * 18.18.76) can be a non-contiguous set of ranges (Part 1, 18.3.1.78). Then `ranges` lists every
 * range in the order of selection, and `range` is the range that holds the active cell,
 * `ranges[activeRange]`. A selection of one range has neither field, so
 * `selection.ranges ?? [selection.range]` gives every range.
 *
 * @public
 */
export interface GridSelection {
  readonly sheetIndex: number;
  readonly active: GridPoint;
  readonly anchor: GridPoint;
  /** The range that holds the active cell: the whole selection when it is one range. */
  readonly range: GridRange;
  /**
   * Every range of a selection of two or more ranges, in the order of selection, or absent for
   * a selection of one range. A range can overlap another one.
   */
  readonly ranges?: readonly GridRange[];
  /**
   * The 0-based position in `ranges` of the range that holds the active cell, as
   * `x:selection/@activeCellId` states it. Present only with `ranges`.
   */
  readonly activeRange?: number;
}
