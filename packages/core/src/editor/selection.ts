// Cell and range selection on one sheet: pure functions over `GridSelection`, so the keyboard,
// the pointer, and the tests share one model. The selection stays inside the grid bounds, and a
// range that meets a merged range grows to cover it, as a merged range selects as one cell.

import type { GridPoint, GridRange, GridSelection } from '../contracts/index.ts';
import { MAX_COLS, MAX_ROWS, type HistorySelection } from '../store/index.ts';
import type { CellRange } from '../layout/index.ts';

/** The merged ranges of the sheet, for selection. */
export type Merges = readonly CellRange[];

/** A point moved into the grid, with integer coordinates. */
export const clampPoint = (p: GridPoint): GridPoint => ({
  row: Math.min(MAX_ROWS - 1, Math.max(0, Math.trunc(p.row))),
  col: Math.min(MAX_COLS - 1, Math.max(0, Math.trunc(p.col))),
});

/** The range from two corners, in any order. */
export function rangeOf(a: GridPoint, b: GridPoint): GridRange {
  return {
    start: { row: Math.min(a.row, b.row), col: Math.min(a.col, b.col) },
    end: { row: Math.max(a.row, b.row), col: Math.max(a.col, b.col) },
  };
}

/** The merged range that holds a cell, if any. */
export function mergeAt(merges: Merges, p: GridPoint): CellRange | undefined {
  for (const m of merges)
    if (p.row >= m.firstRow && p.row <= m.lastRow && p.col >= m.firstCol && p.col <= m.lastCol)
      return m;
  return undefined;
}

/** A range grown until no merged range crosses its edge. At most 64 rounds. */
export function coverMerges(range: GridRange, merges: Merges): GridRange {
  let r = range;
  for (let round = 0; round < 64; round += 1) {
    let grown = false;
    for (const m of merges) {
      const meets =
        m.firstRow <= r.end.row &&
        m.lastRow >= r.start.row &&
        m.firstCol <= r.end.col &&
        m.lastCol >= r.start.col;
      if (!meets) continue;
      const next = {
        start: {
          row: Math.min(r.start.row, m.firstRow),
          col: Math.min(r.start.col, m.firstCol),
        },
        end: { row: Math.max(r.end.row, m.lastRow), col: Math.max(r.end.col, m.lastCol) },
      };
      if (
        next.start.row !== r.start.row ||
        next.start.col !== r.start.col ||
        next.end.row !== r.end.row ||
        next.end.col !== r.end.col
      ) {
        r = next;
        grown = true;
      }
    }
    if (!grown) break;
  }
  return r;
}

/** Select one cell (or the merged range that holds it). */
export function selectCell(sheetIndex: number, at: GridPoint, merges: Merges): GridSelection {
  const p = clampPoint(at);
  const m = mergeAt(merges, p);
  const active = m ? { row: m.firstRow, col: m.firstCol } : p;
  return {
    sheetIndex,
    active,
    anchor: active,
    range: m
      ? { start: active, end: { row: m.lastRow, col: m.lastCol } }
      : { start: active, end: active },
  };
}

/**
 * Extend a selection from its anchor to `to`. The active cell stays while the new range holds
 * it, as after Enter or Tab moved it inside the range; else it goes back to the anchor.
 */
export function extendTo(sel: GridSelection, to: GridPoint, merges: Merges): GridSelection {
  const range = coverMerges(rangeOf(sel.anchor, clampPoint(to)), merges);
  return { ...sel, range, active: inRange(range, sel.active) ? sel.active : sel.anchor };
}

/** Select whole columns, whole rows, or the whole sheet. */
export function selectBand(
  sheetIndex: number,
  kind: 'columns' | 'rows' | 'all',
  from: number,
  to: number
): GridSelection {
  if (kind === 'all') {
    const origin = { row: 0, col: 0 };
    return {
      sheetIndex,
      active: origin,
      anchor: origin,
      range: { start: origin, end: { row: MAX_ROWS - 1, col: MAX_COLS - 1 } },
    };
  }
  const a = Math.min(from, to);
  const b = Math.max(from, to);
  const active = kind === 'columns' ? { row: 0, col: from } : { row: from, col: 0 };
  const range =
    kind === 'columns'
      ? { start: { row: 0, col: a }, end: { row: MAX_ROWS - 1, col: b } }
      : { start: { row: a, col: 0 }, end: { row: b, col: MAX_COLS - 1 } };
  return { sheetIndex, active: clampPoint(active), anchor: clampPoint(active), range };
}

/** The next cell from `p` in a direction, stepping over a merged range and hidden lines. */
export function step(
  p: GridPoint,
  dRow: number,
  dCol: number,
  merges: Merges,
  hidden: { row(i: number): boolean; col(i: number): boolean }
): GridPoint {
  const m = mergeAt(merges, p);
  let row = p.row;
  let col = p.col;
  if (dRow > 0 && m) row = m.lastRow;
  if (dCol > 0 && m) col = m.lastCol;
  if (dRow < 0 && m) row = m.firstRow;
  if (dCol < 0 && m) col = m.firstCol;
  // A hidden row or column is skipped; the search is bounded.
  for (let i = 0; i < 4096; i += 1) {
    const next = clampPoint({ row: row + dRow, col: col + dCol });
    if (next.row === row && next.col === col) return clampPoint(p);
    row = next.row;
    col = next.col;
    if ((dRow === 0 || !hidden.row(row)) && (dCol === 0 || !hidden.col(col))) break;
  }
  return { row, col };
}

/** Whether a point lies in a range. */
export function inRange(r: GridRange, p: GridPoint): boolean {
  return p.row >= r.start.row && p.row <= r.end.row && p.col >= r.start.col && p.col <= r.end.col;
}

/**
 * The grid selection and the focus corner of a selection that a history entry keeps, on the
 * sheet at `sheetIndex`: its range, active cell, anchor, and focus, each clamped to the grid.
 * Shift+Arrow then extends the range from the anchor, at the focus, as before the edit.
 */
export function historyGridSelection(
  s: HistorySelection,
  sheetIndex: number
): { selection: GridSelection; focus: GridPoint } {
  const [top, left, bottom, right] = s.range;
  return {
    selection: {
      sheetIndex,
      anchor: clampPoint(s.anchor),
      active: clampPoint(s.active),
      range: {
        start: clampPoint({ row: top, col: left }),
        end: clampPoint({ row: bottom, col: right }),
      },
    },
    focus: clampPoint(s.focus),
  };
}

/** The words of a selection for the live region, and what they name. */
export interface SelectionWords {
  /** The selection, after its sheet when the sheet changed. */
  readonly said: string;
  /** `said` with the words of the active cell, or `said` itself for one cell. */
  readonly again: string;
  /** The sheet and the range of the selection, in no language. */
  readonly range: string;
  /** The position of the active cell, in no language. */
  readonly at: string;
  /** The words of the active cell. */
  readonly cell: string | undefined;
}

/**
 * The words of `sel`, a new selection on `sheet`, read as `selection` and its active cell as
 * `cell`. The sheet comes first when `sheetChanged`. The live region says `again` when the
 * active cell moves inside the same selection, as after Enter inside a range, or after a paste
 * over it.
 */
export function selectionWords(
  t: (key: string, params?: Record<string, string>) => string,
  sel: GridSelection,
  selection: string,
  cell: string | undefined,
  sheet: string,
  sheetChanged: boolean
): SelectionWords {
  const said = sheetChanged ? t('grid.sheetActive', { sheet, selection }) : selection;
  const again =
    cell === undefined || cell === selection
      ? said
      : t('grid.selectionCell', { selection: said, cell });
  const { start, end } = sel.range;
  const range = `${sel.sheetIndex}:${start.row},${start.col}:${end.row},${end.col}`;
  return { said, again, range, at: `${sel.active.row},${sel.active.col}`, cell };
}

/**
 * The one announcement of an undo or a redo: the step, then the sheet when the step showed
 * another sheet or changed a sheet that is not shown, then the words of the selection when
 * the step selected one. `cell` is the words of the active cell when the step changed them:
 * a selection of many cells, or a step that kept no selection, also names it, so the step says
 * the restored text of that cell.
 */
export function stepAnnouncement(
  t: (key: string, params?: Record<string, string>) => string,
  step: 'undo' | 'redo',
  selected: string | undefined,
  sheet: string | undefined,
  cell?: string
): string {
  const undo = step === 'undo';
  const selection =
    selected !== undefined && cell !== undefined && cell !== selected
      ? t('grid.selectionCell', { selection: selected, cell })
      : (selected ?? cell);
  if (selection === undefined && sheet !== undefined)
    return t(undo ? 'grid.undoneSheet' : 'grid.redoneSheet', { sheet });
  if (selection === undefined) return t(undo ? 'grid.undone' : 'grid.redone');
  if (sheet === undefined) return t(undo ? 'grid.undoneAt' : 'grid.redoneAt', { selection });
  return t(undo ? 'grid.undoneOnSheet' : 'grid.redoneOnSheet', { sheet, selection });
}
