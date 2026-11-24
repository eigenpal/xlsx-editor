// The ranges of a selection as text and as store input: the `ST_Sqref` value that a sheet opens
// with and that the snapshot states (ECMA-376 Part 1, 18.18.76, "a sequence of cell references,
// space delimited"), the A1 reference of one range, the ops of a command that applies to every
// range, and the format limit of such a command. Split from `editor-model.ts`.

import type { GridPoint, GridRange, GridSelection } from '../contracts/index.ts';
import {
  MAX_FORMAT_NEW_CELLS,
  clearFormatsOverflow,
  formatCellRef,
  formatRangeOverflow,
  parseCellRef,
  sheetIndex,
  xmlPartOf,
  type RangeBox,
  type WorkbookOp,
  type XlsxPackage,
} from '../store/index.ts';
// Not on the public store entry: only this module counts the cells of several ranges.
import { clearFormatsCreatedIn, type RangeBands } from '../store/store/format-range.ts';
import { clearOps, formatOps, type ClearKind } from '../automation/index.ts';
import { selectionRanges } from '../automation/structure-commands.ts';
import type { FormatPatch } from '../store/index.ts';
import { refusalOf, type EditRefusal } from './cell-input.ts';
import type { SelectionRangeInput } from './editor-instance.ts';
import {
  MAX_SELECTION_RANGES,
  addRange,
  isMultiRange,
  clampPoint,
  extendTo,
  rangeOf,
  selectCell,
  type Merges,
} from './selection.ts';

/** A range in the store's range shape. */
export const boxOfRange = ({ start, end }: GridRange): RangeBox => [
  start.row,
  start.col,
  end.row,
  end.col,
];

/** The most characters of `x:selection/@sqref` that an open reads. */
const MAX_SQREF_CHARS = 16 * 1024;

/** The first result of `check` over the ranges of a selection that is not undefined. */
export function firstOf<T>(boxes: readonly RangeBox[], check: (box: RangeBox) => T | undefined) {
  for (const box of boxes) {
    const found = check(box);
    if (found !== undefined) return found;
  }
  return undefined;
}

/**
 * The ranges of an `ST_Sqref` value (Part 1, 18.18.76: references, space delimited), each as
 * `ST_Ref` (`A1` or `A1:B2`). It reads at most `MAX_SQREF_CHARS` characters and
 * `MAX_SELECTION_RANGES` ranges: the value is file text, so no file number sets a loop bound.
 */
export function readSqref(raw: string | undefined): GridRange[] {
  if (!raw) return [];
  const out: GridRange[] = [];
  const tokens = raw.slice(0, MAX_SQREF_CHARS).split(/\s+/, MAX_SELECTION_RANGES * 2);
  // A cut in the middle of a reference (`A1:B23` read as `A1:B2`) would name a range that the
  // file does not state, so the last reference of a cut value is left out.
  if (raw.length > MAX_SQREF_CHARS && !/\s/.test(raw[MAX_SQREF_CHARS]!)) tokens.pop();
  for (const token of tokens) {
    if (out.length >= MAX_SELECTION_RANGES) break;
    const [a, b, extra] = token.split(':', 3);
    if (!a || extra !== undefined) continue;
    const start = parseCellRef(a);
    const end = b === undefined ? start : parseCellRef(b);
    if (!start || !end) continue;
    out.push(rangeOf({ row: start.row, col: start.col }, { row: end.row, col: end.col }));
  }
  return out;
}

/** The A1 reference of a range: one cell, or `start:end`. */
export function rangeRef(range: GridRange): string {
  const a = formatCellRef(range.start);
  const b = formatCellRef(range.end);
  return a === b ? a : `${a}:${b}`;
}

/**
 * The A1 references of every range of a selection, space delimited, as `ST_Sqref` writes them
 * (Part 1, 18.18.76): `B2:C4 E5`. One range gives one reference.
 */
export function selectionSqref(sel: GridSelection): string {
  return selectionRanges(sel).map(rangeRef).join(' ');
}

/**
 * Why a format patch is refused for a selection, or null: `range-too-large` when the patch
 * would create more than `MAX_FORMAT_NEW_CELLS` cells, by the store's own count
 * (`formatRangeOverflow`), so the state and the command refuse the same ranges. The count reads
 * the cells inside the range only when the range is larger than the limit, and never walks the
 * range.
 */
export function formatRefusalOf(
  pkg: XlsxPackage,
  partName: string | undefined,
  selection: GridSelection
): EditRefusal | null {
  const ranges = selectionRanges(selection);
  if (ranges.length === 1) {
    const r = selection.range;
    if ((r.end.row - r.start.row + 1) * (r.end.col - r.start.col + 1) <= MAX_FORMAT_NEW_CELLS)
      return null;
    const part = xmlPartOf(pkg, partName);
    return part && formatRangeOverflow(sheetIndex(part), boxOfRange(r))
      ? { code: 'range-too-large' }
      : null;
  }
  // Several ranges format in one transaction, and a cell that two ranges hold is created once:
  // the new cells are the cells of the union of the ranges that no stored cell fills. They count
  // together, as one command creates at most `MAX_FORMAT_NEW_CELLS` cells.
  const cover = unionCover(ranges);
  if (cover.area <= MAX_FORMAT_NEW_CELLS) return null;
  const part = xmlPartOf(pkg, partName);
  if (!part) return null;
  const sheet = sheetIndex(part);
  // A stored cell fills at most one cell of the union: past the limit even if all of them do,
  // the refusal needs no pass over the stored cells.
  if (cover.area - sheet.cells.size > MAX_FORMAT_NEW_CELLS) return { code: 'range-too-large' };
  // One pass over the stored cells, each tested by two binary searches, not against each range.
  let existing = 0;
  for (const e of sheet.cells.values()) if (cover.holds(e.row, e.col)) existing += 1;
  return cover.area - existing > MAX_FORMAT_NEW_CELLS ? { code: 'range-too-large' } : null;
}

/**
 * The union of at most `MAX_SELECTION_RANGES` ranges: its area in cells, and whether it holds a
 * cell. The edges of the ranges cut the grid into at most 128 by 128 blocks, each inside or
 * outside every range, so the cost depends on the number of ranges, never on their size.
 */
export function unionCover(ranges: readonly GridRange[]): {
  readonly area: number;
  /** The union as row bands of column runs, for `clearFormatsCreatedIn`. */
  readonly bands: RangeBands;
  holds(row: number, col: number): boolean;
} {
  const edges = (lo: (r: GridRange) => number, hi: (r: GridRange) => number) =>
    [...new Set(ranges.flatMap((r) => [lo(r), hi(r) + 1]))].sort((a, b) => a - b);
  const rows = edges(
    (r) => r.start.row,
    (r) => r.end.row
  );
  const cols = edges(
    (r) => r.start.col,
    (r) => r.end.col
  );
  const inside = new Uint8Array(rows.length * cols.length);
  for (const r of ranges) {
    const [r0, r1] = [rows.indexOf(r.start.row), rows.indexOf(r.end.row + 1)];
    const [c0, c1] = [cols.indexOf(r.start.col), cols.indexOf(r.end.col + 1)];
    for (let i = r0; i < r1; i++) inside.fill(1, i * cols.length + c0, i * cols.length + c1);
  }
  let area = 0;
  for (let i = 0; i + 1 < rows.length; i++)
    for (let j = 0; j + 1 < cols.length; j++)
      if (inside[i * cols.length + j])
        area += (rows[i + 1]! - rows[i]!) * (cols[j + 1]! - cols[j]!);
  // The block that holds `v`: the last edge at or before it, or -1.
  const block = (edges: readonly number[], v: number) => {
    let [lo, hi] = [0, edges.length - 1];
    if (v < edges[0]!) return -1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (edges[mid]! <= v) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  };
  const bands: { top: number; bottom: number; runs: (readonly [number, number])[] }[] = [];
  for (let i = 0; i + 1 < rows.length; i++) {
    const runs: (readonly [number, number])[] = [];
    for (let j = 0; j + 1 < cols.length; j++) {
      if (!inside[i * cols.length + j]) continue;
      const last = runs[runs.length - 1];
      if (last && last[1] === cols[j]! - 1) runs[runs.length - 1] = [last[0], cols[j + 1]! - 1];
      else runs.push([cols[j]!, cols[j + 1]! - 1]);
    }
    if (runs.length) bands.push({ top: rows[i]!, bottom: rows[i + 1]! - 1, runs });
  }
  return {
    area,
    bands,
    holds(row, col) {
      const i = block(rows, row);
      const j = block(cols, col);
      return i >= 0 && j >= 0 && inside[i * cols.length + j] === 1;
    },
  };
}

/**
 * Why clearing the formats of a selection is refused, or null: `clear-too-large` when the clear
 * would create more than `MAX_FORMAT_NEW_CELLS` cells with `s="0"` (`clearFormatsOverflow` for
 * one range). The `s="0"` cells of every range of a selection of several ranges count together,
 * as one command creates them, and a cell that two ranges hold counts once. No count runs
 * while the union of the ranges holds no more cells than the limit.
 */
export function clearRefusalOf(
  part: NonNullable<ReturnType<typeof xmlPartOf>>,
  selection: GridSelection
): EditRefusal | null {
  const sheet = sheetIndex(part);
  const ranges = selectionRanges(selection);
  if (ranges.length === 1)
    return clearFormatsOverflow(part.root, sheet, boxOfRange(selection.range)) === null
      ? null
      : { code: 'clear-too-large' };
  // As for a format, a cell that two ranges hold is cleared once, so the union counts.
  const cover = unionCover(ranges);
  if (cover.area <= MAX_FORMAT_NEW_CELLS) return null;
  return clearFormatsCreatedIn(part.root, sheet, cover.bands) > MAX_FORMAT_NEW_CELLS
    ? { code: 'clear-too-large' }
    : null;
}

/** The ops of a command that applies to every range of `selection`, in the order of selection. */
export function rangeOps(
  selection: GridSelection,
  opsOf: (ref: string) => readonly WorkbookOp[]
): WorkbookOp[] {
  return selectionRanges(selection).flatMap((r) => opsOf(rangeRef(r)));
}

/**
 * The selection of `selectRanges`: each range from its `anchor` to its `focus`, grown to cover
 * the merged ranges that it meets, added in order (`addRange`), and the focus corner of the last
 * one. Null for an empty list, more than `MAX_SELECTION_RANGES` ranges, or a coordinate that is
 * not a finite number.
 */
export function rangesSelection(
  sheet: number,
  ranges: readonly SelectionRangeInput[],
  merges: Merges
): { readonly selection: GridSelection; readonly focus: GridPoint } | null {
  // A caller without types can pass any value: what is not a list of ranges, each with an
  // `anchor` point, is refused. `focus` can be null: it reads as absent.
  const point = (p: unknown): p is GridPoint =>
    typeof p === 'object' &&
    p !== null &&
    Number.isFinite((p as GridPoint).row) &&
    Number.isFinite((p as GridPoint).col);
  if (!Array.isArray(ranges) || ranges.length === 0 || ranges.length > MAX_SELECTION_RANGES)
    return null;
  const valid = (r: unknown) =>
    typeof r === 'object' &&
    r !== null &&
    point((r as SelectionRangeInput).anchor) &&
    ((r as SelectionRangeInput).focus == null || point((r as SelectionRangeInput).focus));
  if (!ranges.every(valid)) return null;
  let selection: GridSelection | null = null;
  for (const r of ranges) {
    let one = selectCell(sheet, r.anchor, merges);
    if (r.focus) one = extendTo(one, r.focus, merges);
    selection = selection ? addRange(selection, one.range, one.active) : one;
  }
  const last = ranges[ranges.length - 1]!;
  return { selection: selection!, focus: last.focus ? clampPoint(last.focus) : selection!.active };
}

/** The sheet that a command edits: its package, its sheet entry, and its selection. */
interface RangeTarget {
  readonly pkg: XlsxPackage;
  readonly entry: { readonly name: string; readonly partName?: string };
  readonly selection: GridSelection;
}

/**
 * The ops that clear every range of the selection, or the refusal. The store checks the
 * `clear-too-large` limit of each range alone; for several ranges the ranges count together.
 */
export function selectionClearOps(
  { pkg, entry, selection }: RangeTarget,
  kind: ClearKind
): { readonly ops: WorkbookOp[] } | { readonly refusal: EditRefusal } {
  const part = isMultiRange(selection) ? xmlPartOf(pkg, entry.partName) : undefined;
  const over = part && clearRefusalOf(part, selection);
  if (over) return { refusal: over };
  return { ops: rangeOps(selection, (ref) => clearOps(entry.name, ref, kind)) };
}

/**
 * The ops that apply `patch` to every range of the selection, in one transaction, or the
 * refusal. For several ranges, the new cells of every range count together.
 */
export function selectionFormatOps(
  { pkg, entry, selection }: RangeTarget,
  patch: FormatPatch
): { readonly ops: WorkbookOp[] } | { readonly refusal: EditRefusal } {
  if (isMultiRange(selection)) {
    const tooLarge = formatRefusalOf(pkg, entry.partName, selection);
    if (tooLarge) return { refusal: tooLarge };
  }
  let refusal: EditRefusal | undefined;
  const ops = rangeOps(selection, (ref) => {
    const r = formatOps(entry.name, ref, patch);
    if (!r.ok) refusal ??= refusalOf(r.rejection);
    return r.ok ? r.ops : [];
  });
  return refusal ? { refusal } : { ops };
}
