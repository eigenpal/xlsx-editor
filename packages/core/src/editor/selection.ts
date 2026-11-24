// Cell and range selection on one sheet: pure functions over `GridSelection`, so the keyboard,
// the pointer, and the tests share one model. The selection stays inside the grid bounds, and a
// range that meets a merged range grows to cover it, as a merged range selects as one cell.
//
// A selection can hold several ranges (`x:selection/@sqref`, `ST_Sqref`, ECMA-376 Part 1,
// 18.18.76, "can be non-contiguous set of ranges", Part 1, 18.3.1.78). `addRange` adds one and
// makes it active; `extendTo` extends the active range; `selectCell` and `selectBand` start a
// selection of one range again.

import type { GridPoint, GridRange, GridSelection } from '../contracts/index.ts';
import { MAX_COLS, MAX_ROWS, type HistorySelection, type RangeBox } from '../store/index.ts';
import type { CellRange } from '../layout/index.ts';
import { selectionRanges } from '../automation/structure-commands.ts';

/**
 * The most ranges that one selection holds: a press past it adds no range, an open reads no
 * more, and `selectRanges` refuses more. It bounds the paint, the snapshot state, the commands,
 * and the announcement of a selection.
 *
 * @public
 */
export const MAX_SELECTION_RANGES = 64;

/** Whether two ranges have the same corners. */
export const sameRange = (a: GridRange, b: GridRange): boolean =>
  a.start.row === b.start.row &&
  a.start.col === b.start.col &&
  a.end.row === b.end.row &&
  a.end.col === b.end.col;

/** Whether a selection holds two or more ranges. */
export const isMultiRange = (sel: GridSelection): boolean => (sel.ranges?.length ?? 0) > 1;

/**
 * Add `range` to a selection as its last range, with `active` (a cell of it) as the active cell
 * and the anchor. Past `MAX_SELECTION_RANGES`, the selection stays as it is.
 */
export function addRange(sel: GridSelection, range: GridRange, active: GridPoint): GridSelection {
  // A range that the selection holds already moves to the end, as the active range: the list
  // keeps no two equal ranges.
  const ranges = selectionRanges(sel).filter((r) => !sameRange(r, range));
  if (ranges.length >= MAX_SELECTION_RANGES) return sel;
  if (ranges.length === 0) return { sheetIndex: sel.sheetIndex, active, anchor: active, range };
  const next = [...ranges, range];
  return {
    sheetIndex: sel.sheetIndex,
    active,
    anchor: active,
    range,
    ranges: next,
    activeRange: next.length - 1,
  };
}

/**
 * Replace the range that holds the active cell with `range`, keeping the other ranges. An
 * extension can pass through a range equal to another one, by a drag or by Shift and arrows: both
 * stay, so a further step of the same gesture never loses a range.
 */
export function withActiveRange(sel: GridSelection, range: GridRange): GridSelection {
  if (!isMultiRange(sel)) return { ...sel, range };
  const at = activeRangeIndex(sel);
  const ranges = sel.ranges!.map((r, i) => (i === at ? range : r));
  return { ...sel, range, ranges, activeRange: at };
}

/** The position of the range that holds the active cell in `selectionRanges(sel)`. */
export function activeRangeIndex(sel: GridSelection): number {
  const count = sel.ranges?.length ?? 0;
  return count > 1 ? indexIn(sel.activeRange ?? count - 1, count, 0) : 0;
}

/** `at` when it is an integer position in a list of `count` items, else `fallback`. */
const indexIn = (at: number | undefined, count: number, fallback: number): number =>
  Number.isInteger(at) && at! >= 0 && at! < count ? at! : fallback;

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
  const active = inRange(range, sel.active) ? sel.active : sel.anchor;
  return { ...withActiveRange(sel, range), active };
}

/**
 * At most this many merged range checks, for each pass of `coverMerges` (at most 64 passes),
 * cover the ranges that a sheet opens with.
 */
const OPEN_MERGE_CHECKS = 400_000;

/**
 * The selection that a sheet opens with, from its `x:selection` (`CT_Selection`, ECMA-376
 * Part 1, 18.3.1.78): the ranges of `@sqref`, each grown to cover the merged ranges it meets,
 * and the active cell `@activeCell` in the range at `@activeCellId`. When `@activeCellId` is
 * absent or out of range, the first range that holds `@activeCell` is the active range, as
 * "when this value is out of range then activeCell can be used". When no range holds the active
 * cell, or `@sqref` is absent, the selection is the active cell alone. With many merged ranges,
 * fewer ranges are read, so the merge checks stay bounded; the active range stays.
 */
export function openingSelection(
  sheetIndex: number,
  opening: {
    readonly active: GridPoint;
    readonly ranges: readonly GridRange[];
    readonly activeRange: number | undefined;
  },
  merges: Merges
): GridSelection {
  const one = selectCell(sheetIndex, opening.active, merges);
  const listed = opening.ranges.slice(0, MAX_SELECTION_RANGES);
  const holds = (r: GridRange) => inRange(r, one.active);
  let at = opening.activeRange ?? listed.findIndex(holds);
  if (at < 0 || at >= listed.length) return one;
  const room = Math.max(1, Math.floor(OPEN_MERGE_CHECKS / Math.max(1, merges.length)));
  const kept = listed.filter((_, i) => i === at || i < room - 1);
  at = Math.min(at, kept.length - 1);
  const grown = kept.map((r) =>
    coverMerges(rangeOf(clampPoint(r.start), clampPoint(r.end)), merges)
  );
  // A range equal to an earlier one, as the file lists it or as merged cells grow it, opens
  // once: the active range keeps its place among the ranges that stay.
  const ranges = grown.filter((r, i) => grown.findIndex((o) => sameRange(o, r)) === i);
  at = ranges.findIndex((r) => sameRange(r, grown[at]!));
  const range = ranges[at]!;
  const active = holds(range) ? one.active : selectCell(sheetIndex, range.start, merges).active;
  const base = { sheetIndex, active, anchor: active, range };
  return ranges.length > 1 ? { ...base, ranges, activeRange: at } : base;
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
  const gridRange = ([top, left, bottom, right]: RangeBox): GridRange => ({
    start: clampPoint({ row: top, col: left }),
    end: clampPoint({ row: bottom, col: right }),
  });
  // A history entry can come from a host transaction: at most `MAX_SELECTION_RANGES` ranges
  // reach the paint and the live region, as at open.
  const kept = s.ranges?.slice(0, MAX_SELECTION_RANGES);
  const ranges = kept && kept.length > 1 ? kept.map(gridRange) : undefined;
  // A position that is not an integer in `ranges` takes the last range.
  const at = ranges ? indexIn(s.activeRange, ranges.length, ranges.length - 1) : 0;
  // When the active range moved, the active cell moves to its first corner, so it stays inside.
  const active = clampPoint(s.active);
  const inside = (r: GridRange) =>
    active.row >= r.start.row &&
    active.row <= r.end.row &&
    active.col >= r.start.col &&
    active.col <= r.end.col;
  const corner = ranges && !inside(ranges[at]!) ? ranges[at]!.start : undefined;
  return {
    selection: {
      sheetIndex,
      anchor: corner ?? clampPoint(s.anchor),
      active: corner ?? active,
      range: ranges ? ranges[at]! : gridRange(s.range),
      ...(ranges ? { ranges, activeRange: at } : {}),
    },
    focus: corner ?? clampPoint(s.focus),
  };
}

/** The boxes of the ranges of a selection, for a history entry, when it holds several ranges. */
export function historyRanges(
  sel: GridSelection
): Pick<HistorySelection, 'ranges' | 'activeRange'> {
  if (!isMultiRange(sel)) return {};
  const box = ({ start, end }: GridRange): RangeBox => [start.row, start.col, end.row, end.col];
  return { ranges: sel.ranges!.map(box), activeRange: activeRangeIndex(sel) };
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
  const box = ({ start, end }: GridRange) => `${start.row},${start.col}:${end.row},${end.col}`;
  const range = `${sel.sheetIndex}:${selectionRanges(sel).map(box).join(' ')}`;
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
