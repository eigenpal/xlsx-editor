// What a sheet opens with: its zoom, its scroll, its active cell, and its selection, as the first
// `x:sheetView` of the worksheet states them. Split from `editor-model.ts`.

import type { GridPoint, GridRange } from '../contracts/index.ts';
import {
  attr,
  child,
  children,
  parseCellRef,
  xmlPartOf,
  type XlsxPackage,
} from '../store/index.ts';
import { frozenPanes } from '../layout/index.ts';
import { readSqref } from './selection-ranges.ts';

/** What a sheet opens with: its zoom, its active cell, its selection, and its top-left cell. */
export interface SheetOpening {
  readonly zoom: number;
  readonly active: GridPoint;
  readonly topLeft: GridPoint;
  readonly headings: boolean;
  /**
   * The ranges of `x:selection/@sqref`, in file order, at most `MAX_SELECTION_RANGES`; a range
   * that does not read is left out. Empty when the attribute is absent; its schema default is
   * `A1` (`CT_Selection`, sml.xsd), which the active cell then gives, as it falls back to A1.
   */
  readonly ranges: readonly GridRange[];
  /** `x:selection/@activeCellId`, when it is an index into `ranges`. */
  readonly activeRange: number | undefined;
}

/**
 * The view of a sheet that the file states: the first `x:sheetView` (`CT_SheetView`, sml.xsd,
 * ECMA-376 Part 1, 18.3.1.87). It reads `@zoomScale` (default 100), `@topLeftCell`,
 * `@showRowColHeaders` (default true), and from one `x:selection` (`CT_Selection`, 18.3.1.78)
 * `@activeCell`, `@sqref`, and `@activeCellId`. With a frozen `x:pane` (`CT_Pane`, 18.3.1.66),
 * the scroll comes from `x:pane/@topLeftCell`, the top-left visible cell of the bottom-right
 * pane, and the selection is the one whose `@pane` is `x:pane/@activePane` (`ST_Pane`, 18.18.52;
 * both default `topLeft`).
 */
export function sheetOpening(pkg: XlsxPackage, partName: string | undefined): SheetOpening {
  const part = xmlPartOf(pkg, partName);
  const view = part ? child(child(part.root, 'x:sheetViews'), 'x:sheetView') : undefined;
  const zoomRaw = Number(attr(view, 'zoomScale') ?? '100');
  const zoom = Number.isFinite(zoomRaw) && zoomRaw >= 10 && zoomRaw <= 400 ? zoomRaw / 100 : 1;
  const pane = child(view, 'x:pane');
  // A pane that freezes no row or column, and a split pane, do not show: the editor shows the
  // view's top-left cell and the selection of the top-left pane.
  const band = part ? frozenPanes(part.root) : { rows: 0, cols: 0 };
  const frozen = band.rows > 0 || band.cols > 0;
  const shown = frozen ? (attr(pane, 'activePane') ?? 'topLeft') : 'topLeft';
  const sels = children(view, 'x:selection');
  const sel = sels.find((s) => (attr(s, 'pane') ?? 'topLeft') === shown) ?? sels[0];
  const cell = (raw: string | undefined) => (raw ? parseCellRef(raw.split(':')[0]!) : null);
  const point = (a: { row: number; col: number } | null): GridPoint =>
    a ? { row: a.row, col: a.col } : { row: 0, col: 0 };
  const paneTop = frozen ? cell(attr(pane, 'topLeftCell')) : null;
  const headings = attr(view, 'showRowColHeaders');
  const ranges = readSqref(attr(sel, 'sqref'));
  // `@activeCellId` is an `xsd:unsignedInt`. It counts the ranges of the file, and a range that
  // does not read is left out, so the index applies only when its range holds the active cell;
  // else the first range that holds it does.
  const id = attr(sel, 'activeCellId');
  const byId = id !== undefined && /^\d{1,9}$/.test(id) ? Number(id) : -1;
  // With no `@activeCell`, the active cell is the first corner of the range that
  // `@activeCellId` names (0 when absent, `CT_Selection`, sml.xsd), so the ranges still open.
  const given = cell(attr(sel, 'activeCell'));
  const named = ranges[byId >= 0 && byId < ranges.length ? byId : 0];
  const active = given ? point(given) : named ? named.start : point(null);
  const holds = (r: GridRange) =>
    active.row >= r.start.row &&
    active.row <= r.end.row &&
    active.col >= r.start.col &&
    active.col <= r.end.col;
  const index =
    byId < ranges.length && byId >= 0 && holds(ranges[byId]!) ? byId : ranges.findIndex(holds);
  return {
    zoom,
    active,
    topLeft: point(paneTop ?? cell(attr(view, 'topLeftCell'))),
    headings: headings === undefined || headings === '1' || headings === 'true',
    ranges,
    activeRange: index >= 0 && index < ranges.length ? index : undefined,
  };
}
