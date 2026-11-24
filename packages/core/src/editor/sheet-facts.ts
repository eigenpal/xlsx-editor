// Facts about the cells that exist on a sheet, for scrolling and keyboard moves: the used extent
// and the data edge in each direction. Built from the sheet index (the cells that exist), never
// from a range that the file names, and cached per sheet index revision (AGENTS.md "Security").

import type { GridPoint } from '../contracts/index.ts';
import {
  MAX_COLS,
  MAX_ROWS,
  sheetIndex,
  xmlPartOf,
  type SheetIndex,
  type XlsxPackage,
} from '../store/index.ts';
import { sheetGeometry, type GridMetrics, type SheetLayout } from '../layout/index.ts';
import { frozenPanes } from '../layout/panes.ts';

export interface SheetFacts {
  /** The last row and column that hold a cell, or -1 on an empty sheet. */
  readonly lastRow: number;
  readonly lastCol: number;
  /** The rows of each column that hold a cell, ascending. */
  readonly rowsOfCol: ReadonlyMap<number, Int32Array>;
  /** The columns of each row that hold a cell, ascending. */
  readonly colsOfRow: ReadonlyMap<number, Int32Array>;
}

const cache = new WeakMap<SheetIndex, SheetFacts>();

/** The facts of a sheet. Cost is in proportion to its cells. */
export function sheetFacts(sheet: SheetIndex): SheetFacts {
  const hit = cache.get(sheet);
  if (hit) return hit;
  const byCol = new Map<number, number[]>();
  const byRow = new Map<number, number[]>();
  let lastRow = -1;
  let lastCol = -1;
  for (const entry of sheet.cells.values()) {
    if (!entry.el.children.some((n) => n.type === 'element')) continue;
    lastRow = Math.max(lastRow, entry.row);
    lastCol = Math.max(lastCol, entry.col);
    let rows = byCol.get(entry.col);
    if (!rows) byCol.set(entry.col, (rows = []));
    rows.push(entry.row);
    let cols = byRow.get(entry.row);
    if (!cols) byRow.set(entry.row, (cols = []));
    cols.push(entry.col);
  }
  const sorted = (m: Map<number, number[]>) =>
    new Map([...m].map(([k, v]) => [k, Int32Array.from(v).sort()]));
  const facts = { lastRow, lastCol, rowsOfCol: sorted(byCol), colsOfRow: sorted(byRow) };
  cache.set(sheet, facts);
  return facts;
}

/** The rows and columns that take no space: `@hidden`, or a size of 0. */
export interface HiddenLines {
  row(i: number): boolean;
  col(i: number): boolean;
  /**
   * The first row and column that can show; every line before it is hidden (the lines before a
   * frozen band). A search stops there, so its cost does not depend on the place of the band.
   */
  readonly firstRow?: number;
  readonly firstCol?: number;
}

/**
 * Any valid metrics: whether a line has a size of 0 does not depend on them (`@hidden`, a
 * `@width` or `@ht` of 0, or a custom default row height of 0), only the other sizes do.
 */
const NOMINAL: GridMetrics = {
  maxDigitWidth: 7,
  pixelsPerInch: 72,
  defaultRowHeight: 15,
  rowScale: 1,
};
const hiddenCache = new WeakMap<SheetIndex, HiddenLines>();

/**
 * The hidden rows and columns of a sheet, read from its tree, so they hold for the newest
 * revision even before the next layout. Cached per sheet index revision. The rows and columns
 * before a frozen band do not show (`layout/panes.ts`, approximation "Frozen band origin"), so
 * they count as hidden, and the keys skip them.
 */
export function hiddenLines(sheet: SheetIndex): HiddenLines {
  const hit = hiddenCache.get(sheet);
  if (hit) return hit;
  const g = sheetGeometry(sheet, NOMINAL);
  const { firstRow, firstCol } = frozenPanes(sheet.root);
  const lines: HiddenLines = {
    row: (i) => i < firstRow || g.rows.sizeOf(i) === 0,
    col: (i) => i < firstCol || g.columns.sizeOf(i) === 0,
    firstRow,
    firstCol,
  };
  hiddenCache.set(sheet, lines);
  return lines;
}

/** The first position in `list` whose value is at least `v`. */
function lowerBound(list: Int32Array, v: number): number {
  let lo = 0;
  let hi = list.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (list[mid]! < v) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

/** The bound on a search past hidden lines, as for the other keyboard moves. */
const SEARCH = 4096;
/** No hidden rows or columns. */
export const NONE_HIDDEN: HiddenLines = { row: () => false, col: () => false };

/** The hidden lines of the worksheet part `partName` of `pkg`, or none without such a part. */
export function hiddenLinesOf(
  pkg: XlsxPackage | undefined,
  partName: string | undefined
): HiddenLines {
  const part = pkg ? xmlPartOf(pkg, partName) : undefined;
  return part ? hiddenLines(sheetIndex(part)) : NONE_HIDDEN;
}

/**
 * The data edge from `at` in one direction, as a jump with the control key: the end of the run
 * of cells with values that `at` is in, or the start of the next run, or the sheet edge. A
 * hidden row or column (`hidden`) is not part of the run and takes no
 * selection: the run goes on past it, and the jump never stops on it. Each search past hidden
 * lines is bounded.
 */
export function dataEdge(
  facts: SheetFacts,
  at: GridPoint,
  dRow: number,
  dCol: number,
  lines: HiddenLines = NONE_HIDDEN
): GridPoint {
  const vertical = dRow !== 0;
  const hidden = vertical ? lines.row : lines.col;
  const min = (vertical ? lines.firstRow : lines.firstCol) ?? 0;
  const all = vertical ? facts.rowsOfCol.get(at.col) : facts.colsOfRow.get(at.row);
  const line = all ? all.filter((v) => !hidden(v)) : new Int32Array(0);
  const pos = vertical ? at.row : at.col;
  const max = (vertical ? MAX_ROWS : MAX_COLS) - 1;
  const dir = (vertical ? dRow : dCol) > 0 ? 1 : -1;
  const set = (v: number): GridPoint =>
    vertical ? { row: v, col: at.col } : { row: at.row, col: v };
  /** The first shown line from `v` on in `d`, inside the sheet, or undefined. */
  const shown = (v: number, d: 1 | -1): number | undefined => {
    // From before the first line that can show, a search forward starts at that line.
    if (d > 0) v = Math.max(v, min);
    for (let k = 0; k < SEARCH; k += 1, v += d) {
      if (v < min || v > max) return undefined;
      if (!hidden(v)) return v;
    }
    return undefined;
  };
  const has = (v: number): boolean => {
    const i = lowerBound(line, v);
    return i < line.length && line[i] === v;
  };
  const edge = () => set(shown(dir > 0 ? max : min, dir > 0 ? -1 : 1) ?? pos);
  const next = shown(pos + dir, dir as 1 | -1);
  if (next === undefined) return set(pos);
  if (line.length === 0) return edge();
  if (has(pos) && has(next)) {
    // Walk to the end of the run: values on consecutive shown lines.
    let cur = next;
    for (let k = 0; k < line.length; k += 1) {
      const after = shown(cur + dir, dir as 1 | -1);
      if (after === undefined || !has(after)) break;
      cur = after;
    }
    return set(cur);
  }
  // The next cell with a value in the direction, or the edge.
  const i = lowerBound(line, dir > 0 ? next : next + 1);
  if (dir > 0) return i < line.length ? set(line[i]!) : edge();
  return i - 1 >= 0 ? set(line[i - 1]!) : edge();
}

/** Rows and columns past the content and the window that the scroll extent adds. */
const EXTRA_ROWS = 100;
const EXTRA_COLS = 20;

/**
 * The size in points that the scroll extent covers: the content, the laid-out window, and the
 * moving corner of the selection, with room past them, clipped to the sheet. The rows and
 * columns before a frozen band do not show, so they take no room (`x:sheetView/@topLeftCell`,
 * ECMA-376 Part 1, 18.3.1.87).
 */
export function scrollExtent(
  layout: SheetLayout,
  facts: { lastRow: number; lastCol: number },
  focus: GridPoint
): { width: number; height: number } {
  const g = layout.geometry;
  let lastRow = 0;
  let lastCol = 0;
  for (const p of layout.panes) {
    for (const r of p.rows) lastRow = Math.max(lastRow, r.index);
    for (const c of p.columns) lastCol = Math.max(lastCol, c.index);
  }
  const rows = Math.min(g.rows.count, Math.max(facts.lastRow, lastRow, focus.row) + 1 + EXTRA_ROWS);
  const cols = Math.min(
    g.columns.count,
    Math.max(facts.lastCol, lastCol, focus.col) + 1 + EXTRA_COLS
  );
  const { firstRow, firstCol } = layout.frozen;
  return {
    width: Math.max(0, g.columns.offsetOf(cols) - g.columns.offsetOf(firstCol)),
    height: Math.max(0, g.rows.offsetOf(rows) - g.rows.offsetOf(firstRow)),
  };
}
