// The return column of a Tab run. Tab moves the active cell one column right; Enter after one or
// more Tab moves goes down one row to the column where the run started, so a user who types a
// record across the row with Tab starts the next record under its first cell. Any other move ends
// the run: an arrow key, a pointer press, Go To, a sheet change, or a move inside a multi-cell
// selection. The run lives in the editor, not in the workbook, and is never saved.
//
// A run is valid only while the selection is the one that its last Tab move made. Every other
// move replaces the selection object (`select`, `setSelection`), so no listener has to end the
// run: a stale run fails the identity check and is dropped.

import type { GridPoint, GridSelection } from '../contracts/index.ts';

/** The return column of a Tab run. @internal */
export class TabReturn {
  #run: { readonly selection: GridSelection; readonly col: number } | null = null;

  /**
   * Note a Tab move from `cell`: `before` is the selection that the move left and `after` the
   * one that it made. A move from the selection of a running Tab move keeps its column.
   */
  tabbed(before: GridSelection, cell: GridPoint, after: GridSelection): void {
    if (after === before) return;
    const col = this.#run && this.#run.selection === before ? this.#run.col : cell.col;
    this.#run = { selection: after, col };
  }

  /**
   * The column that Enter returns to from `current`, or null when no Tab run made `current`.
   * Reading it ends the run.
   */
  take(current: GridSelection): number | null {
    const run = this.#run;
    this.#run = null;
    return run && run.selection === current ? run.col : null;
  }

  /**
   * The cell that a move of `move` from `cell` starts at: for Enter (`inRange`, one row down),
   * the run's column in the row of `cell` when a Tab run made `current`, else `cell`. Enter ends
   * the run.
   */
  origin(
    current: GridSelection,
    cell: GridPoint,
    move: readonly [number, number],
    inRange: boolean
  ): GridPoint {
    if (!inRange || move[0] !== 1 || move[1] !== 0) return cell;
    const col = this.take(current);
    return col === null ? cell : { row: cell.row, col };
  }

  /** Note a move of `move` from `cell`: a Tab move (`inRange`, along the row) runs on. */
  moved(
    before: GridSelection,
    cell: GridPoint,
    after: GridSelection,
    move: readonly [number, number],
    inRange: boolean
  ): void {
    if (inRange && move[0] === 0) this.tabbed(before, cell, after);
  }

  /** End the run. */
  end(): void {
    this.#run = null;
  }
}
