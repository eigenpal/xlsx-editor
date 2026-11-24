// The key table of the grid (the scroll container that holds the focus), in one place:
//
// | Key                                   | Action                                         |
// | ------------------------------------- | ---------------------------------------------- |
// | Mod+Z; Mod+Shift+Z and Mod+Y          | Undo; redo (`historyShortcutOf`)               |
// | Mod+B, Mod+I, Mod+U                   | Toggle bold, italic, underline on the selection |
// | A character, or an input method key   | Start an edit with that text                   |
// | F2                                    | Edit the active cell in place                  |
// | Backspace                             | Start an edit with empty text                  |
// | Delete                                | Clear the values of the selection, keep format |
// | Arrows, Shift+arrows, Mod+arrows      | Move, extend, jump to the data edge            |
// |                                       | (a row or column band extends as a band)       |
// | Page Up, Page Down, Home, Mod+Home/End | Move by a page, to the row start, to the ends |
// |                                       | (a hidden row or column takes no selection)    |
// | Enter, Tab (with Shift)               | Move down or right (up or left); Enter after   |
// |                                       | Tab goes down to the column where the Tab run  |
// |                                       | started (`tab-return.ts`)                      |
// |                                       | (inside a multi-cell selection: move the active |
// |                                       | cell in it, wrap, and keep the range; in a     |
// |                                       | selection of several ranges, go on to the next |
// |                                       | range, in the order of selection)              |
// | Shift+Tab in column A; Tab in the last | Leave the grid for the previous or next control |
// | Alt+Page Up, Alt+Page Down            | The previous or next sheet                     |
// | Mod+Shift+=, Mod+- on whole rows      | Insert, delete the selected rows (or columns)  |
// | Mod+A; Mod+plus, Mod+minus, Mod+0     | Select all; zoom in, out, to 100%             |
// | Shift+Space; Mod+Space                | Select the whole rows; the whole columns       |
// | Mod+C, Mod+X, Mod+V                   | Not handled here: the browser fires `copy`,    |
// |                                       | `cut`, and `paste` (`clipboard-events.ts`)     |
// | Escape                                | End a pending copy or cut; with none, arm      |
// |                                       | the exit: the next Tab or Shift+Tab leaves it  |
// |                                       | (`grid-exit.ts`, read before and after this    |
// |                                       | table in `surface-render.ts`)                  |
//
// F6 and Shift+F6 move the keyboard among the editor's regions; the adapter's viewport handles
// them, so they pass through this table.
//
// Only a key whose target is the grid itself acts: a key typed in the cell editor, or in any
// element of the host page, never reaches this table.

import type { GridPoint, GridSelection } from '../contracts/index.ts';
import { historyShortcutOf } from './history-shortcut.ts';
import {
  activeRangeIndex,
  inRange as holds,
  isMultiRange,
  mergeAt,
  selectBand,
  step,
  withActiveRange,
  type Merges,
} from './selection.ts';
import type { CellEditMode } from './cell-editor.ts';
import { selectionBand, type FormatToggle } from '../automation/index.ts';
import { selectionRanges } from '../automation/structure-commands.ts';
import { MAX_COLS, MAX_ROWS, type StructuralOpName } from '../store/index.ts';
import type { TabReturn } from './tab-return.ts';
import type { HiddenLines } from './sheet-facts.ts';

/** What the key table reads and calls on the editor. */
export interface GridKeyHost {
  readonly selection: GridSelection;
  /** The moving corner of an extended selection. */
  readonly focusCell: GridPoint;
  readonly merges: Merges;
  hidden(): HiddenLines;
  readonly pageRows: number;
  /** The last row and column that hold content, -1 when none. */
  facts(): { lastRow: number; lastCol: number };
  dataEdge(from: GridPoint, dRow: number, dCol: number): GridPoint;
  select(anchor: GridPoint, focus?: GridPoint): void;
  /**
   * Take a selection as it is, with its moving corner: a band stays a band. `reveal` scrolls
   * that cell into view.
   */
  setSelection(selection: GridSelection, focus: GridPoint, reveal?: GridPoint): void;
  /** Extend the selection from its anchor to `to` (`extendTo`), and reveal `to`. */
  extend(to: GridPoint): void;
  selectAll(): void;
  zoomStep(direction: 1 | -1): void;
  setZoom(zoom: number): void;
  cycleSheet(direction: 1 | -1): void;
  undo(): void;
  redo(): void;
  beginEdit(text: string | undefined, mode: CellEditMode): void;
  clearValues(): void;
  toggleFormat(which: FormatToggle): void;
  structure(name: StructuralOpName): void;
  /** Escape on the grid: end a pending copy or cut. False when nothing waited. */
  escape(): boolean;
  /** The return column of a Tab run (`tab-return.ts`): Enter after Tab goes back to it. */
  readonly tabReturn?: TabReturn;
}

const TOGGLES: Readonly<Record<string, FormatToggle>> = { b: 'bold', i: 'italic', u: 'underline' };

const DIRECTIONS: Readonly<Record<string, readonly [number, number]>> = {
  ArrowUp: [-1, 0],
  ArrowDown: [1, 0],
  ArrowLeft: [0, -1],
  ArrowRight: [0, 1],
};

/**
 * The nearest shown row or column to `i`: `i` itself, else the first shown one in `dir`, else
 * the first shown one back. A hidden line takes no selection, as the arrows skip it (`step`).
 * The search is bounded, and starts at `min`, the first line that can show.
 */
function shownLine(
  i: number,
  dir: 1 | -1,
  hidden: (i: number) => boolean,
  count: number,
  min = 0
): number {
  i = Math.max(i, Math.min(min, count - 1));
  if (!hidden(i)) return i;
  for (const d of [dir, -dir]) {
    for (let k = 1; k <= 4096; k += 1) {
      const at = i + d * k;
      if (at < min || at >= count) break;
      if (!hidden(at)) return at;
    }
  }
  return i;
}

/**
 * The next active cell inside `range`, in reading order for Tab (by columns for Enter), with
 * a wrap at the range edges. A hidden cell is skipped, and so is a cell that a merged range
 * covers unless it is the top-left cell of that range: only that cell holds the value
 * (quick reference, `x:mergeCell`). The search is bounded, and starts at the first row and
 * column that can show (`HiddenLines`): the lines before a frozen band take no step.
 */
export function cycleInRange(
  range: GridSelection['range'],
  at: GridPoint,
  byRow: boolean,
  dir: 1 | -1,
  hidden: HiddenLines,
  merges: Merges = []
): GridPoint {
  const bottom = Math.max(range.start.row, range.end.row);
  const right = Math.max(range.start.col, range.end.col);
  // A range that lies wholly before the band keeps its own first line.
  const firstRow = (hidden.firstRow ?? 0) <= bottom ? (hidden.firstRow ?? 0) : 0;
  const firstCol = (hidden.firstCol ?? 0) <= right ? (hidden.firstCol ?? 0) : 0;
  const top = Math.max(Math.min(range.start.row, range.end.row), firstRow);
  const left = Math.max(Math.min(range.start.col, range.end.col), firstCol);
  let { row, col } = at;
  for (let k = 0; k < 4096; k += 1) {
    if (byRow) {
      col += dir;
      if (col > right) [col, row] = [left, row + 1 > bottom ? top : row + 1];
      else if (col < left) [col, row] = [right, row - 1 < top ? bottom : row - 1];
    } else {
      row += dir;
      if (row > bottom) [row, col] = [top, col + 1 > right ? left : col + 1];
      else if (row < top) [row, col] = [bottom, col - 1 < left ? right : col - 1];
    }
    if (hidden.row(row) || hidden.col(col)) continue;
    const m = mergeAt(merges, { row, col });
    if (!m || (m.firstRow === row && m.firstCol === col)) return { row, col };
  }
  return at;
}

/**
 * Whether Enter and Tab move the active cell inside `range` and keep it: a range of more than
 * one cell, unless it is one merged cell, which is one cell (Tab steps over it, and leaves the
 * grid at its edge).
 */
function cyclesIn(range: GridSelection['range'], at: GridPoint, merges: Merges): boolean {
  if (range.start.row === range.end.row && range.start.col === range.end.col) return false;
  const m = mergeAt(merges, at);
  return !(
    m &&
    m.firstRow === range.start.row &&
    m.lastRow === range.end.row &&
    m.firstCol === range.start.col &&
    m.lastCol === range.end.col
  );
}

/** Whether `next` comes at or before `at` in the order of a step by rows or by columns. */
function wrapped(at: GridPoint, next: GridPoint, byRow: boolean, dir: 1 | -1): boolean {
  const key = (p: GridPoint) => (byRow ? p.row * 2 ** 20 + p.col : p.col * 2 ** 20 + p.row);
  return dir > 0 ? key(next) <= key(at) : key(next) >= key(at);
}

/**
 * The selection after Enter or Tab moves the active cell from `at`: the next cell of the range
 * that holds it (`cycleInRange`). In a selection of several ranges, a step past the last cell of
 * that range (the first, with `dir` -1) goes on to the first (last) cell of the next (previous)
 * range, in the order of selection, with a wrap; every range stays selected.
 */
export function cycleSelection(
  sel: GridSelection,
  at: GridPoint,
  byRow: boolean,
  dir: 1 | -1,
  hidden: { row(i: number): boolean; col(i: number): boolean },
  merges: Merges
): GridSelection {
  const ranges = selectionRanges(sel);
  const index = activeRangeIndex(sel);
  const own = ranges[index]!;
  if (cyclesIn(own, at, merges)) {
    const next = cycleInRange(own, at, byRow, dir, hidden, merges);
    if (ranges.length === 1 || !wrapped(at, next, byRow, dir)) return { ...sel, active: next };
  }
  if (ranges.length === 1) return sel;
  const to = (index + dir + ranges.length) % ranges.length;
  const range = ranges[to]!;
  // A step from the last cell wraps to the first cell, which a hidden line or a merge can move.
  const active =
    dir > 0
      ? cycleInRange(range, range.end, byRow, 1, hidden, merges)
      : cycleInRange(range, range.start, byRow, -1, hidden, merges);
  return { ...sel, active, anchor: active, range, activeRange: to };
}

/**
 * Where a committed edit moves the active cell: one step from `cell`, or, for Enter and Tab
 * (`inRange`) inside a multi-cell selection that is not one merged cell, or a selection of
 * several ranges, the next cell in it (`cycleSelection`), and the selection stays (`keep` is
 * the new selection), as the keys do without an edit.
 */
export function commitMove(
  sel: GridSelection,
  cell: GridPoint,
  move: readonly [number, number],
  inRange: boolean,
  merges: Merges,
  hidden: HiddenLines
): { at: GridPoint; keep: GridSelection | null } {
  if (inRange && (isMultiRange(sel) || cyclesIn(sel.range, cell, merges))) {
    const dir = (move[0] || move[1]) < 0 ? -1 : 1;
    const next = cycleSelection(sel, cell, move[1] !== 0, dir, hidden, merges);
    return { at: next.active, keep: next };
  }
  return { at: step(cell, move[0], move[1], merges, hidden), keep: null };
}

/** Whether a key types text: one character, with no modifier that makes it a shortcut. */
function typesText(e: KeyboardEvent): boolean {
  if (e.metaKey || (e.ctrlKey && !e.altKey)) return false;
  return [...e.key].length === 1;
}

/** Act on one keydown of the grid. */
export function handleGridKey(host: GridKeyHost, e: KeyboardEvent, grid: EventTarget): void {
  if (e.target !== grid) return;
  const mod = e.ctrlKey || e.metaKey;
  const done = () => e.preventDefault();
  const history = historyShortcutOf(e);
  if (history) {
    done();
    if (history === 'undo') host.undo();
    else host.redo();
    return;
  }
  const toggle = mod && !e.altKey && !e.shiftKey ? e.key.toLowerCase() : '';
  if (Object.hasOwn(TOGGLES, toggle)) return (done(), host.toggleFormat(TOGGLES[toggle]!));
  // On a selection of whole rows or columns, Mod+Shift+= (Mod++ where `+` needs no Shift)
  // inserts and Mod+- deletes them; on any other selection the keys zoom.
  const band = mod && !e.altKey ? selectionBand(host.selection) : null;
  if (band && (e.key === '+' || (e.shiftKey && e.key === '=')))
    return (done(), host.structure(band === 'rows' ? 'insertRows' : 'insertColumns'));
  if (band && !e.shiftKey && e.key === '-')
    return (done(), host.structure(band === 'rows' ? 'deleteRows' : 'deleteColumns'));
  if (mod && (e.key === '=' || e.key === '+')) return (done(), host.zoomStep(1));
  if (mod && e.key === '-') return (done(), host.zoomStep(-1));
  if (mod && e.key === '0') return (done(), host.setZoom(1));
  if (mod && (e.key === 'a' || e.key === 'A')) return (done(), host.selectAll());
  // Escape ends a pending copy or cut; with none, the key goes on to the page, and
  // `grid-exit.ts` arms the exit from the grid. An Escape that the page used, such as to close a
  // message, does nothing more here, and neither does a held Escape.
  if (e.key === 'Escape' && !mod && !e.altKey && !e.shiftKey) {
    if (!e.defaultPrevented && !e.repeat && host.escape()) done();
    return;
  }
  if (e.isComposing || e.key === 'Process' || e.keyCode === 229) {
    // An input method starts composing: open the editor and let the composition land in it.
    host.beginEdit('', 'enter');
    return;
  }
  if (e.key === ' ' && !e.altKey && e.shiftKey !== mod) {
    // Shift+Space selects the rows of the selection, Mod+Space its columns; neither types.
    done();
    const r = host.selection.range;
    const rows = e.shiftKey;
    const band = selectBand(
      host.selection.sheetIndex,
      rows ? 'rows' : 'columns',
      rows ? r.start.row : r.start.col,
      rows ? r.end.row : r.end.col
    );
    const keep = { active: host.selection.active, anchor: host.selection.anchor };
    // In a selection of several ranges, the band takes the place of the range of the active cell.
    return host.setSelection(
      { ...withActiveRange(host.selection, band.range), ...keep },
      host.focusCell
    );
  }
  if (typesText(e)) return (done(), host.beginEdit(e.key, 'enter'));
  if (e.key === 'F2') return (done(), host.beginEdit(undefined, 'edit'));
  if (e.key === 'Backspace') return (done(), host.beginEdit('', 'enter'));
  if (e.key === 'Delete') return (done(), host.clearValues());
  if (e.altKey && (e.key === 'PageDown' || e.key === 'PageUp'))
    return (done(), host.cycleSheet(e.key === 'PageDown' ? 1 : -1));
  const sel = host.selection;
  const from = e.shiftKey ? host.focusCell : sel.active;
  let to: GridPoint | null = null;
  const dir = Object.hasOwn(DIRECTIONS, e.key) ? DIRECTIONS[e.key]! : null;
  if (dir) {
    to = mod
      ? host.dataEdge(from, dir[0], dir[1])
      : step(from, dir[0], dir[1], host.merges, host.hidden());
  } else if (e.key === 'PageDown' || e.key === 'PageUp') {
    const dir = e.key === 'PageDown' ? 1 : -1;
    const target = Math.min(MAX_ROWS - 1, Math.max(0, from.row + host.pageRows * dir));
    const hidden = host.hidden();
    to = { row: shownLine(target, dir, hidden.row, MAX_ROWS, hidden.firstRow), col: from.col };
  } else if (e.key === 'Home') {
    // A hidden first row or column takes no selection: the first shown one does.
    const hidden = host.hidden();
    const col = shownLine(0, 1, hidden.col, MAX_COLS, hidden.firstCol);
    const row = () => shownLine(0, 1, hidden.row, MAX_ROWS, hidden.firstRow);
    to = mod ? { row: row(), col } : { row: from.row, col };
  } else if (e.key === 'End' && mod) {
    const facts = host.facts();
    const hidden = host.hidden();
    to = {
      row: shownLine(Math.max(0, facts.lastRow), -1, hidden.row, MAX_ROWS, hidden.firstRow),
      col: shownLine(Math.max(0, facts.lastCol), -1, hidden.col, MAX_COLS, hidden.firstCol),
    };
  } else if (e.key === 'Tab' || e.key === 'Enter') {
    const back = e.shiftKey;
    if (isMultiRange(sel) || cyclesIn(sel.range, sel.active, host.merges)) {
      // Inside a multi-cell selection, the keys move the active cell and keep the range; in a
      // selection of several ranges, they go on to the next range.
      done();
      const next = cycleSelection(
        sel,
        sel.active,
        e.key === 'Tab',
        back ? -1 : 1,
        host.hidden(),
        host.merges
      );
      const focus = next.range === sel.range ? host.focusCell : next.active;
      host.setSelection(next, focus, next.active);
      return;
    }
    const d: [number, number] = e.key === 'Tab' ? [0, back ? -1 : 1] : [back ? -1 : 1, 0];
    // Enter after a run of Tab moves goes down to the column where the run started.
    const run = host.tabReturn;
    const from = run?.origin(sel, sel.active, d, true) ?? sel.active;
    const next = step(from, d[0], d[1], host.merges, host.hidden());
    // Tab at the first or last column leaves the grid: the browser moves the focus to the
    // previous or next control, so the grid is no keyboard trap (WCAG 2.1.2).
    if (e.key === 'Tab' && next.row === sel.active.row && next.col === sel.active.col)
      return run?.end();
    done();
    host.select(next);
    run?.moved(sel, sel.active, host.selection, d, true);
    return;
  }
  if (!to) return;
  done();
  if (!e.shiftKey) return host.select(to);
  // A band of whole rows or columns extends as a band, as a drag over the headings does, so
  // Mod+- still deletes the extended rows or columns.
  const extended = selectionBand(sel);
  if (!extended) return host.extend(to);
  const rows = extended === 'rows';
  const next = selectBand(
    sel.sheetIndex,
    extended,
    rows ? sel.anchor.row : sel.anchor.col,
    rows ? to.row : to.col
  );
  // The active cell stays while the band holds it, as `extendTo` keeps it.
  const active = holds(next.range, sel.active) ? sel.active : sel.anchor;
  host.setSelection({ ...withActiveRange(sel, next.range), active, anchor: sel.anchor }, to, to);
}
