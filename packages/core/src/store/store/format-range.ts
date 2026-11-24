// The `applyFormat` and `clearFormats` ops: set or remove the cell format of the cells of a
// range, as one transaction with one undo step. A run of `applyFormat` ops on one sheet applies
// in one pass (`applyFormatRanges`), with the result of the ops one by one; ranges that overlap
// apply one by one.
//
// `applyFormat` gives each cell of the range the cell format of its own format with the patch
// applied (`style-records.ts`): a bold patch on a red cell keeps it red. Cells that do not exist
// are created with the new format, starting from the format of their row (`x:row/@s` with
// `customFormat`) or of their column (`x:col/@style`), as ECMA-376 Part 1, 18.3.1.73 and
// 18.3.1.13 give an empty cell (`workbook/inherited-format.ts`). Creation is bounded: a command
// that would create more than `MAX_FORMAT_NEW_CELLS` cells is refused with `range-too-large`,
// so one whole row formats and a whole column or the whole sheet is refused. This change writes
// no row or column format.
//
// A patch that sets the font size, the font name, or wrapping removes `x:row/@ht` from each
// changed row that has no `customHeight="1"`, so layout fits the row again; a row that the user
// sized keeps its height. `clearFormats` does the same for a row in which it removes a format
// with another font or with wrapping.
//
// `clearFormats` removes `@s` from the cells that exist in the range; a cell that then holds
// nothing goes away, unless its row or column has a format: it keeps `s="0"`, so it does not
// show that format. A position without a cell whose row or column has a format other than 0
// gets a new `s="0"` cell for the same reason, with the same bound and refusal as `applyFormat`.
// On a protected sheet (`x:sheetProtection/@sheet`), the clear keeps the protection of each
// cell (`x:protection`, 18.8.33): a cell whose protection differs from that of format 0 gets
// format 0 with its own protection, a record that the clear appends when none exists. So a clear
// never locks an unlocked cell, and never unlocks a locked one.
// It walks the positions of the range only where such a format is, so a sheet without row or
// column formats clears in the time of its cells.
//
// Both ops rebuild each changed row once and record one `restoreRow` inverse for it. A format
// edit that appends style records also records a `restoreStyles` inverse, so undo puts back the
// styles part as it was, and the package is equal to the one before the edit. Neither op nor its
// undo changes a value, so neither drops the calculation chain (`calc.ts`, `rowChange`). A clear
// of formats and a format edit change the result of `CELL("format")` and of `GET.CELL`, so both
// mark the cached values stale when a formula of a cell or a defined name calls `CELL` or a
// `GET.*` macro (`CELL_STATE_FUNCTIONS`), as a value edit does.

import type { Element, Node, XmlPart } from '../package/tree.ts';
import type { XlsxPackage } from '../package/package.ts';
import { attr, boolAttr, child } from '../package/query.ts';
import { formatCellRef } from '../cell-ref.ts';
import { cellKey, sheetIndex, type SheetIndex } from '../workbook/sheet-index.ts';
import { styleLists, type WorkbookIndex } from '../workbook/workbook-index.ts';
import {
  columnFormatAt,
  columnFormats,
  formatIndex,
  rowFormat,
} from '../workbook/inherited-format.ts';
import { pinCellRefs, sheetTarget, withPart } from './cells.ts';
import { rangeBounds, type RangeBox } from './clear-range.ts';
import { childIndex, IdSource, plainAttr, smlElement, withAttr, withChild } from './tree-edit.ts';
import { lowerNumberFormat, patchGroups, patchProblem, type FormatPatch } from './format-patch.ts';
import {
  EDGE_BOTTOM,
  EDGE_LEFT,
  EDGE_RIGHT,
  EDGE_TOP,
  EDGE_ALL,
  StyleLimit,
  StyleRecords,
} from './style-records.ts';
import { commitStyles, stylesOf } from './styles-part.ts';
import { refuse, unchanged, type OpResult, type HistoryOp } from './ops.ts';
import { widenDimension } from './dimension.ts';
import { lockReader, readSheetProtection } from './sheet-protection.ts';
import { CELL_STATE_FUNCTIONS, workbookCallsSheetFunction } from '../structure/sheet-functions.ts';

/** The most cells that one format command creates: one whole row. */
export const MAX_FORMAT_NEW_CELLS = 16_384;

/**
 * The cells that a format patch of a range would create beyond `MAX_FORMAT_NEW_CELLS`, or
 * null when it creates at most that many. The range is `[top, bottom]` by `[left, right]`, 0-based
 * and inclusive. The count is the area of the range less the cells that exist in it. It is
 * counted from the cells that exist, never by walking the range, and not at all when the range
 * is small enough or so large that even every cell of the sheet inside it would leave too many
 * (then `least` is true and `count` is a lower bound). The toolbar state and the command read
 * this one function, so they refuse the same ranges.
 */
export function formatRangeOverflow(
  index: SheetIndex,
  [top, left, bottom, right]: RangeBox
): { readonly count: number; readonly least: boolean } | null {
  const area = (bottom - top + 1) * (right - left + 1);
  if (area <= MAX_FORMAT_NEW_CELLS) return null;
  if (area - index.cells.size > MAX_FORMAT_NEW_CELLS)
    return { count: area - index.cells.size, least: true };
  let existing = 0;
  for (const e of index.cells.values())
    if (e.row >= top && e.row <= bottom && e.col >= left && e.col <= right) existing += 1;
  return area - existing > MAX_FORMAT_NEW_CELLS ? { count: area - existing, least: false } : null;
}

/**
 * A row without the height that an earlier fit wrote: `@ht` goes when `@customHeight` is not
 * set, so layout fits the row to the fonts and wrapped text of its cells again. A row that the
 * user sized (`customHeight="1"`) keeps its height (ECMA-376 Part 1, 18.3.1.73).
 */
export function withoutFittedHeight(row: Element): Element {
  return boolAttr(row, 'customHeight') ? row : withAttr(row, 'ht', undefined);
}

/**
 * Whether removing a cell format can change the height that its row fits to: the format names
 * a font other than the font of the default format, or wraps text. Memoized by format.
 */
export function heightFormats(
  pkg: XlsxPackage,
  index: WorkbookIndex
): (xf: number | undefined) => boolean {
  const lists = styleLists(pkg, index);
  const font0 = attr(lists.cellXfs[0], 'fontId') ?? '0';
  const memo = new Map<number, boolean>();
  return (xf) => {
    if (xf === undefined || xf === 0) return false;
    let hit = memo.get(xf);
    if (hit === undefined) {
      const el = lists.cellXfs[xf];
      hit =
        el !== undefined &&
        ((attr(el, 'fontId') ?? '0') !== font0 ||
          boolAttr(child(el, 'x:alignment'), 'wrapText') === true);
      memo.set(xf, hit);
    }
    return hit;
  };
}

/** The border edges of a cell for a patch: the range's outline, or every edge. */
function edgesOf(
  patch: FormatPatch,
  row: number,
  col: number,
  [top, left, bottom, right]: readonly [number, number, number, number]
): number {
  const b = patch.borders;
  if (!b || b.edges === 'none') return 0;
  if (b.edges === 'all') return EDGE_ALL;
  return (
    (col === left ? EDGE_LEFT : 0) |
    (col === right ? EDGE_RIGHT : 0) |
    (row === top ? EDGE_TOP : 0) |
    (row === bottom ? EDGE_BOTTOM : 0)
  );
}

/** The first index of `order` at or after `lo` whose key is at least `key`. */
function lowerBound(order: readonly number[], key: number, lo = 0): number {
  let hi = order.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (order[mid]! < key) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

/**
 * The 0-based column of each indexed cell in rows `top` to `bottom`. With no index problem, the
 * rows only go forward in `order`, so a binary search finds the cells of those rows, and the
 * cost is those cells, not the sheet. With a problem, it reads every cell.
 */
function columnsOf(index: SheetIndex, top: number, bottom: number): Map<Element, number> {
  const out = new Map<Element, number>();
  if (index.problems.length > 0) {
    for (const e of index.cells.values()) if (e.row >= top && e.row <= bottom) out.set(e.el, e.col);
    return out;
  }
  const { order } = index;
  const end = lowerBound(order, cellKey(bottom + 1, 0));
  for (let i = lowerBound(order, cellKey(top, 0)); i < end; i += 1) {
    const e = index.cells.get(order[i]!);
    if (e) out.set(e.el, e.col);
  }
  return out;
}

/** The worksheet part with changed and new rows, keeping every other child of `x:sheetData`. */
export function withRows(
  part: XmlPart,
  index: SheetIndex,
  changed: ReadonlyMap<number, Element>,
  ids: IdSource
): XmlPart {
  const sheetData = index.sheetData!;
  const rowOf = new Map<Element, number>();
  for (const [r, el] of index.rows) rowOf.set(el, r);
  const added = [...changed.keys()].filter((r) => !index.rows.has(r)).sort((a, b) => a - b);
  const kids: Node[] = [];
  let k = 0;
  for (const node of sheetData.children) {
    const r = node.type === 'element' ? rowOf.get(node) : undefined;
    if (r !== undefined) {
      while (k < added.length && added[k]! < r) kids.push(changed.get(added[k++]!)!);
      kids.push(changed.get(r) ?? node);
    } else kids.push(node);
  }
  while (k < added.length) kids.push(changed.get(added[k++]!)!);
  const root = withChild(part.root, childIndex(part.root, sheetData), {
    ...sheetData,
    children: kids,
  });
  return { partName: part.partName, root, nextId: ids.next };
}

/** The `restoreRow` inverses of the changed rows, in row order. */
export function rowInverses(
  sheet: string,
  index: SheetIndex,
  changed: ReadonlyMap<number, Element>
): HistoryOp[] {
  return [...changed.keys()]
    .sort((a, b) => a - b)
    .map((r) => ({ op: 'restoreRow', sheet, row: r + 1, node: index.rows.get(r) ?? null }));
}

export function applyFormatRange(
  pkg: XlsxPackage,
  sheet: string,
  ref: string,
  rawPatch: FormatPatch
): OpResult {
  // One range never overlaps itself, so the batch always applies it.
  return applyFormatRanges(pkg, sheet, [{ ref, patch: rawPatch }])!;
}

/**
 * Whether two of the ranges share a cell. The check sorts the ranges by their first row and
 * compares each with the ranges that start at or before its last row, so its cost is the sort
 * and a scan of the ranges that share rows. A range that is not valid shares no cell.
 */
function overlaps(boxes: readonly (RangeBox | undefined)[]): boolean {
  const valid = boxes.filter((b): b is RangeBox => b !== undefined).sort((a, b) => a[0] - b[0]);
  for (let i = 0; i < valid.length; i += 1) {
    const [, l, b, r] = valid[i]!;
    for (let k = i + 1; k < valid.length && valid[k]![0] <= b; k += 1) {
      const [, l2, , r2] = valid[k]!;
      if (l2 <= r && l <= r2) return true;
    }
  }
  return false;
}

/** One range of a run and its patch, lowered, with the cell formats that it has found. */
interface FormatRange {
  readonly box: RangeBox;
  readonly patch: FormatPatch;
  readonly groups: ReturnType<typeof patchGroups>;
  /** Whether the patch removes a fitted row height (`withoutFittedHeight`). */
  readonly refits: boolean;
  /** The new cell format for each old one and border edges. */
  readonly memo: Map<number, number>;
  /**
   * The number of cell formats before the range adds its records. A row or column format at
   * this index or past it shows format 0, as `inheritedFormat` gives it.
   */
  cellFormats: number;
}

/** The format that a new cell of a range starts from: the row or column format, or 0. */
const startFormat = (range: FormatRange, inherited: number): number =>
  inherited < range.cellFormats ? inherited : 0;

/** The new cell format of a range for an old one at a position: from its memo, or added. */
function newFormatOf(
  records: StyleRecords,
  range: FormatRange,
  base: number,
  row: number,
  col: number
): number {
  const edges = edgesOf(range.patch, row, col, range.box);
  const key = base * 16 + edges;
  let xf = range.memo.get(key);
  if (xf === undefined)
    range.memo.set(key, (xf = records.xfFor(base, range.patch, range.groups, edges)));
  return xf;
}

/**
 * Add the new style records of a range in row order, as one op adds them, so the new style records
 * come in the order of the ops one by one. The work is the area of the range, which the
 * overflow check bounds by the cells that exist and the cells that it creates. Throws
 * `StyleLimit` when a list is full.
 */
function addRecordsInOrder(
  index: SheetIndex,
  root: Element,
  records: StyleRecords,
  range: FormatRange
) {
  const [t, l, b, r] = range.box;
  for (let row = t; row <= b; row += 1) {
    const rowStyle = rowFormat(index.rows.get(row) ?? null);
    for (let col = l; col <= r; col += 1) {
      const cell = index.cells.get(cellKey(row, col));
      const base = cell
        ? (formatIndex(attr(cell.el, 's')) ?? 0)
        : startFormat(range, rowStyle ?? columnFormatAt(root, col) ?? 0);
      newFormatOf(records, range, base, row, col);
    }
  }
}

/** A part of one row that a format range covers: columns `left` to `right` of `range`. */
interface Segment {
  readonly left: number;
  readonly right: number;
  readonly range: FormatRange;
}

/**
 * Apply the patches of several ranges of one sheet in one pass: the sheet index, the column
 * formats, and the style records are built once, and each changed row is rebuilt once. Each
 * range has its own patch. The result equals one `applyFormat` op for each range, in order: the
 * same package, the same new style records in the same order, and the same refusal. Undefined
 * when two ranges overlap, so the caller applies them one by one. A run of `applyFormat` ops of
 * a paste costs the time of the sheet once, not once for each range or each format.
 */
export function applyFormatRanges(
  pkg: XlsxPackage,
  sheet: string,
  items: readonly { readonly ref: string; readonly patch: FormatPatch }[]
): OpResult | undefined {
  const boxes = items.map((item) => rangeBounds(item.ref));
  // Ranges that overlap apply one by one: a later range formats the result of an earlier one.
  if (overlaps(boxes)) return undefined;
  const target = sheetTarget(pkg, sheet);
  // Each range is checked in op order, as the ops one by one check it: its patch, its
  // reference, the sheet, its number format, the cells that it creates, the styles part, and
  // the style records that it adds.
  let defined: ReadonlyMap<number, string> | undefined;
  let index: SheetIndex | undefined;
  let styles: ReturnType<typeof stylesOf> | undefined;
  let records: StyleRecords | undefined;
  const ranges: FormatRange[] = [];
  for (const [i, box] of boxes.entries()) {
    const { ref, patch: rawPatch } = items[i]!;
    const problem = patchProblem(rawPatch);
    if (problem) return refuse('bad-format', problem);
    if (!box) return refuse('bad-reference', `${JSON.stringify(ref)} is not a range in the grid`);
    if ('code' in target) return { ok: false, rejection: target };
    // The ops of a run add custom ids only, so the built-in ids that the file redefines are the
    // same for each range.
    defined ??= styleLists(pkg, target.index).numFmts;
    const lowered = lowerNumberFormat(rawPatch, defined);
    if ('problem' in lowered) return refuse('unreadable-number-format', lowered.problem);
    index ??= sheetIndex(target.part);
    const overflow = formatRangeOverflow(index, box);
    if (overflow)
      return refuse(
        'range-too-large',
        `${ref} would create ${overflow.least ? 'at least ' : ''}${overflow.count} cells; one command creates at most ${MAX_FORMAT_NEW_CELLS}`
      );
    styles ??= stylesOf(pkg, target.index);
    if ('code' in styles) return { ok: false, rejection: styles };
    const patch = lowered.patch;
    const range: FormatRange = {
      box,
      patch,
      groups: patchGroups(patch),
      refits:
        patch.fontSize !== undefined ||
        patch.fontName !== undefined ||
        patch.wrapText !== undefined,
      memo: new Map(),
      cellFormats: 0,
    };
    ranges.push(range);
    // The ops one by one append new style records range by range, each in row order. Add the
    // records of this range now, so the records, their indexes, and a full list refusal are the
    // same.
    records ??= new StyleRecords(styles.part);
    range.cellFormats = records.cellFormats;
    try {
      addRecordsInOrder(index, target.part.root, records, range);
    } catch (e) {
      if (e instanceof StyleLimit) return refuse('style-limit', e.message);
      throw e;
    }
  }
  if (!index || !records || 'code' in target) return unchanged(pkg);
  // The segments of each row, by column. The work is the rows of the ranges, which the
  // overflow check above bounds for a range with no cells.
  const rows = new Map<number, Segment[]>();
  let top = Infinity;
  let bottom = -1;
  let left = Infinity;
  let right = -1;
  for (const range of ranges) {
    const [t, l, b, r] = range.box;
    top = Math.min(top, t);
    bottom = Math.max(bottom, b);
    left = Math.min(left, l);
    right = Math.max(right, r);
    for (let row = t; row <= b; row += 1) {
      let list = rows.get(row);
      if (!list) rows.set(row, (list = []));
      list.push({ left: l, right: r, range });
    }
  }
  for (const list of rows.values()) list.sort((a, b) => a.left - b.left);
  const all = records;
  const xfOf = (base: number, row: number, col: number, range: FormatRange): number =>
    newFormatOf(all, range, base, row, col);
  const colStyles = columnFormats(target.part.root, left, right);
  const colOf = columnsOf(index, top, bottom);
  const ids = new IdSource(target.part);
  const changed = new Map<number, Element>();
  // The bounds of the created cells, to widen `x:dimension/@ref` once.
  let made: [number, number, number, number] | undefined;
  try {
    for (const row of [...rows.keys()].sort((a, b) => a - b)) {
      const segments = rows.get(row)!;
      const rowEl = index.rows.get(row) ?? null;
      const rowStyle = rowFormat(rowEl);
      const kids: Node[] = [];
      // The segment that the walk is in, and the next column of it to create.
      let s = 0;
      let next = segments[0]!.left;
      let dirty = false;
      // Whether a range whose patch refits the row changed it.
      let refit = false;
      // Create the missing cells of the segments before column `until`.
      const create = (until: number) => {
        while (s < segments.length) {
          const seg = segments[s]!;
          const end = Math.min(seg.right + 1, until);
          for (; next < end; next += 1) {
            const inherited = rowStyle ?? colStyles.get(next) ?? 0;
            const base = startFormat(seg.range, inherited);
            const xf = xfOf(base, row, next, seg.range);
            // Compare with the format that the cell shows, not the raw index: a new record can
            // take the index of a row or column format past the cell formats.
            if (xf === base) continue;
            const attrs = [plainAttr('r', formatCellRef({ row, col: next }))];
            if (xf !== 0) attrs.push(plainAttr('s', String(xf)));
            kids.push(smlElement(ids, 'c', 'row', attrs));
            dirty = true;
            refit ||= seg.range.refits;
            if (!made) made = [row, next, row, next];
            else {
              made[0] = Math.min(made[0], row);
              made[1] = Math.min(made[1], next);
              made[2] = Math.max(made[2], row);
              made[3] = Math.max(made[3], next);
            }
          }
          if (until <= seg.right) return;
          s += 1;
          if (s < segments.length) next = Math.max(next, segments[s]!.left);
        }
      };
      for (const node of rowEl?.children ?? []) {
        const col = node.type === 'element' && node.kind === 'cell' ? colOf.get(node) : undefined;
        if (col === undefined) {
          // A trailing `x:extLst` follows the cells.
          if (node.type === 'element' && node.kind !== 'cell') create(Infinity);
          kids.push(node);
          continue;
        }
        create(col);
        const seg = segments[s];
        if (!seg || col < seg.left) {
          kids.push(node);
          continue;
        }
        next = col + 1;
        const raw = attr(node as Element, 's');
        const base = formatIndex(raw) ?? 0;
        const xf = xfOf(base, row, col, seg.range);
        // An unchanged format writes nothing; a malformed `@s` is written again.
        if (xf === base && (raw === undefined || formatIndex(raw) === base)) {
          kids.push(node);
          continue;
        }
        const updated = withAttr(node as Element, 's', xf === 0 ? undefined : String(xf));
        if (updated !== node) {
          dirty = true;
          refit ||= seg.range.refits;
        }
        kids.push(updated);
      }
      create(Infinity);
      if (!dirty) continue;
      const base = rowEl ?? smlElement(ids, 'row', 'sheetData', [plainAttr('r', String(row + 1))]);
      // `spans` is a hint that a new cell can make wrong; dropping it is always valid.
      const edited = withAttr({ ...base, children: kids }, 'spans', undefined);
      changed.set(row, refit ? withoutFittedHeight(edited) : edited);
    }
  } catch (e) {
    if (e instanceof StyleLimit) return refuse('style-limit', e.message);
    throw e;
  }
  if (changed.size === 0) return unchanged(pkg);
  let part = withRows(target.part, index, changed, ids);
  const inverse = rowInverses(sheet, index, changed);
  // New cells outside `x:dimension/@ref` widen it, in the same op and undo step.
  const first = made && widenDimension(part, made[0], made[1]);
  if (first) part = first.part;
  const last = made && widenDimension(part, made[2], made[3]);
  if (last) part = last.part;
  const oldDimension = first?.old ?? last?.old;
  if (oldDimension !== undefined) inverse.push({ op: 'setDimension', sheet, ref: oldDimension });
  let next = withPart(pkg, { kind: 'xml', partName: part.partName, xml: part });
  if (records.changed) {
    const committed = commitStyles(next, target.index, records.part());
    if ('code' in committed) return { ok: false, rejection: committed };
    next = committed.pkg;
    inverse.push(committed.inverse);
  }
  // A cell format changes the result of `CELL("format")`: a call marks the values stale.
  const value = workbookCallsSheetFunction(pkg, target.index, CELL_STATE_FUNCTIONS);
  return { ok: true, pkg: next, inverse, value, formula: false };
}

/**
 * The columns of `[left, right]` whose column format is not 0, ascending, and the count of the
 * `s="0"` cells that clearing the formats of the range creates: one for each position that holds
 * no cell and inherits a format other than 0 from its row or its column. The count reads the
 * rows and the cells of the sheet, and the `x:col` ranges, never the positions of the range.
 */
function clearCreations(
  root: Element,
  index: SheetIndex,
  [top, left, bottom, right]: RangeBox
): { readonly cols: readonly number[]; readonly colStyles: Map<number, number>; count: number } {
  const colStyles = columnFormats(root, left, right);
  const cols = [...colStyles.keys()].sort((a, b) => a - b);
  let rowsIn = 0;
  let count = 0;
  for (const [row, rowEl] of index.rows) {
    if (row < top || row > bottom) continue;
    rowsIn += 1;
    const own = rowFormat(rowEl);
    count += own === undefined ? cols.length : own !== 0 ? right - left + 1 : 0;
  }
  count += (bottom - top + 1 - rowsIn) * cols.length;
  if (count === 0) return { cols, colStyles, count };
  for (const e of index.cells.values()) {
    if (e.row < top || e.row > bottom || e.col < left || e.col > right) continue;
    if ((rowFormat(index.rows.get(e.row)) ?? colStyles.get(e.col) ?? 0) !== 0) count -= 1;
  }
  return { cols, colStyles, count };
}

/**
 * A union of ranges as disjoint row bands, ascending, each with disjoint column runs `[left,
 * right]`, ascending, all 0-based and inclusive.
 */
export type RangeBands = readonly {
  readonly top: number;
  readonly bottom: number;
  readonly runs: readonly (readonly [number, number])[];
}[];

/**
 * The `s="0"` cells that clearing the formats of every range of a union creates, as
 * `clearCreations` counts them for one range: a cell that two ranges hold counts once, as one
 * command clears it once. One pass reads the rows and one the cells of the sheet, never the
 * positions of the union, and each finds its band and run by a binary search.
 */
export function clearFormatsCreatedIn(root: Element, index: SheetIndex, bands: RangeBands): number {
  const info = bands.map((b) => {
    let width = 0;
    let styled = 0;
    for (const [left, right] of b.runs) {
      width += right - left + 1;
      styled += columnFormats(root, left, right).size;
    }
    return { width, styled, rowsIn: 0 };
  });
  const bandOf = (row: number): number => {
    let [lo, hi] = [0, bands.length - 1];
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const b = bands[mid]!;
      if (row < b.top) hi = mid - 1;
      else if (row > b.bottom) lo = mid + 1;
      else return mid;
    }
    return -1;
  };
  const inRuns = (runs: readonly (readonly [number, number])[], col: number): boolean => {
    let [lo, hi] = [0, runs.length - 1];
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const [left, right] = runs[mid]!;
      if (col < left) hi = mid - 1;
      else if (col > right) lo = mid + 1;
      else return true;
    }
    return false;
  };
  let count = 0;
  for (const [row, rowEl] of index.rows) {
    const i = bandOf(row);
    if (i < 0) continue;
    const band = info[i]!;
    band.rowsIn += 1;
    const own = rowFormat(rowEl);
    count += own === undefined ? band.styled : own !== 0 ? band.width : 0;
  }
  bands.forEach((b, i) => (count += (b.bottom - b.top + 1 - info[i]!.rowsIn) * info[i]!.styled));
  if (count === 0) return 0;
  for (const e of index.cells.values()) {
    const i = bandOf(e.row);
    if (i < 0 || !inRuns(bands[i]!.runs, e.col)) continue;
    if ((rowFormat(index.rows.get(e.row)) ?? columnFormatAt(root, e.col) ?? 0) !== 0) count -= 1;
  }
  return count;
}

/**
 * The `s="0"` cells that clearing the formats of a range creates, when they are more than
 * `MAX_FORMAT_NEW_CELLS`, or null. The range is 0-based and inclusive. A range no larger than
 * the limit returns null at once. The toolbar state and the command read this one count, so they
 * refuse the same ranges.
 */
export function clearFormatsOverflow(
  root: Element,
  index: SheetIndex,
  box: RangeBox
): number | null {
  const [top, left, bottom, right] = box;
  if ((bottom - top + 1) * (right - left + 1) <= MAX_FORMAT_NEW_CELLS) return null;
  const { count } = clearCreations(root, index, box);
  return count > MAX_FORMAT_NEW_CELLS ? count : null;
}

export function clearFormatsRange(pkg: XlsxPackage, sheet: string, ref: string): OpResult {
  const box = rangeBounds(ref);
  if (!box) return refuse('bad-reference', `${JSON.stringify(ref)} is not a range in the grid`);
  const [top, left, bottom, right] = box;
  const target = sheetTarget(pkg, sheet);
  if ('code' in target) return { ok: false, rejection: target };
  const index = sheetIndex(target.part);
  const { cols, colStyles, count } = clearCreations(target.part.root, index, box);
  if (count > MAX_FORMAT_NEW_CELLS)
    return refuse(
      'clear-too-large',
      `${ref} would create ${count} cells; one command creates at most ${MAX_FORMAT_NEW_CELLS}`
    );
  const colOf = columnsOf(index, top, bottom);
  const refits = heightFormats(pkg, target.index);
  const ids = new IdSource(target.part);
  const changed = new Map<number, Element>();
  // On a protected sheet, the clear keeps the protection of each cell: a cell whose protection
  // is not the protection of format 0 gets format 0 with its own protection (`keep`).
  let keep = (_s: number): number => 0;
  let records: StyleRecords | undefined;
  if (readSheetProtection(target.part.root)) {
    const styles = stylesOf(pkg, target.index);
    if ('code' in styles) return { ok: false, rejection: styles };
    const lock = lockReader(pkg, target.index);
    const kept = new StyleRecords(styles.part);
    const memo = new Map<string, number>();
    const zero = `${lock.locked(0)}:${lock.hidden(0)}`;
    records = kept;
    keep = (s) => {
      const key = `${lock.locked(s)}:${lock.hidden(s)}`;
      if (key === zero) return 0;
      let xf = memo.get(key);
      if (xf === undefined)
        memo.set(key, (xf = kept.defaultWithProtection(lock.locked(s), lock.hidden(s))));
      return xf;
    };
  }
  // The bounds of the created cells, to widen `x:dimension/@ref` once.
  let made: [number, number, number, number] | undefined;
  // A row is read when it exists, or when a column format can need a new cell in it.
  const rows =
    count > 0 && cols.length > 0
      ? Array.from({ length: bottom - top + 1 }, (_, i) => top + i)
      : [...index.rows.keys()].filter((r) => r >= top && r <= bottom).sort((a, b) => a - b);
  try {
    clearRows();
  } catch (e) {
    if (e instanceof StyleLimit) return refuse('style-limit', e.message);
    throw e;
  }
  function clearRows(): void {
    for (const row of rows) {
      const rowEl = index.rows.get(row) ?? null;
      let dirty = false;
      let refit = false;
      let created = false;
      const own = rowFormat(rowEl);
      const kids: Node[] = [];
      // The positions without a cell that inherit a format other than 0 get `s="0"`, so they do
      // not show it: every position of a row with a format, or the formatted columns.
      let next = left;
      let k = 0;
      const add = (col: number) => {
        const s = String(keep(own ?? colStyles.get(col) ?? 0));
        kids.push(
          smlElement(ids, 'c', 'row', [
            plainAttr('r', formatCellRef({ row, col })),
            plainAttr('s', s),
          ])
        );
        dirty = created = true;
        if (!made) made = [row, col, row, col];
        else {
          made[0] = Math.min(made[0], row);
          made[1] = Math.min(made[1], col);
          made[2] = Math.max(made[2], row);
          made[3] = Math.max(made[3], col);
        }
      };
      const create = (until: number) => {
        if (count === 0 || own === 0) return;
        if (own !== undefined) for (; next < until; next += 1) add(next);
        else {
          for (; k < cols.length && cols[k]! < until; k += 1) if (cols[k]! >= next) add(cols[k]!);
          next = Math.max(next, until);
        }
      };
      // A clear can remove a cell: each kept cell carries its `@r`, so it keeps its column.
      const pinned = rowEl ? pinCellRefs(rowEl, row).children : [];
      for (let i = 0; i < pinned.length; i += 1) {
        const node = rowEl!.children[i]!;
        const col = node.type === 'element' && node.kind === 'cell' ? colOf.get(node) : undefined;
        if (col === undefined || col < left || col > right) {
          // A trailing `x:extLst` follows the cells.
          if (col !== undefined ? col > right : node.type === 'element' && node.kind !== 'cell')
            create(right + 1);
          kids.push(pinned[i]!);
          continue;
        }
        create(col);
        next = col + 1;
        const raw = attr(node as Element, 's');
        if (raw === undefined) {
          kids.push(pinned[i]!);
          continue;
        }
        const kept = keep(formatIndex(raw) ?? 0);
        if (kept !== 0) {
          if (raw === String(kept)) kids.push(pinned[i]!);
          else {
            dirty = true;
            refit ||= refits(formatIndex(raw));
            kids.push(withAttr(pinned[i] as Element, 's', String(kept)));
          }
          continue;
        }
        dirty = true;
        refit ||= refits(formatIndex(raw));
        const cleared = withAttr(pinned[i] as Element, 's', undefined);
        // A cell with no value, no formula, and no other attribute holds nothing, unless its row
        // or column has a format: then it keeps `s="0"`, or it would show that format.
        if (
          cleared.children.length > 0 ||
          cleared.attrs.some((a) => !(a.ns === '' && a.local === 'r'))
        )
          kids.push(cleared);
        else if ((own ?? colStyles.get(col) ?? 0) !== 0)
          kids.push(raw === '0' ? pinned[i]! : withAttr(pinned[i] as Element, 's', '0'));
      }
      create(right + 1);
      if (!dirty) continue;
      const base = rowEl ?? smlElement(ids, 'row', 'sheetData', [plainAttr('r', String(row + 1))]);
      // `spans` is a hint that a new cell can make wrong; dropping it is always valid.
      const edited = { ...base, children: kids };
      const spanless = created ? withAttr(edited, 'spans', undefined) : edited;
      changed.set(row, refit ? withoutFittedHeight(spanless) : spanless);
    }
  }
  if (changed.size === 0) return unchanged(pkg);
  let part = withRows(target.part, index, changed, ids);
  const inverse = rowInverses(sheet, index, changed);
  // New cells outside `x:dimension/@ref` widen it, in the same op and undo step.
  const first = made && widenDimension(part, made[0], made[1]);
  if (first) part = first.part;
  const last = made && widenDimension(part, made[2], made[3]);
  if (last) part = last.part;
  const oldDimension = first?.old ?? last?.old;
  if (oldDimension !== undefined) inverse.push({ op: 'setDimension', sheet, ref: oldDimension });
  let next = withPart(pkg, { kind: 'xml', partName: part.partName, xml: part });
  if (records?.changed) {
    const committed = commitStyles(next, target.index, records.part());
    if ('code' in committed) return { ok: false, rejection: committed };
    next = committed.pkg;
    inverse.push(committed.inverse);
  }
  // A cell format changes the result of `CELL("format")`: a call marks the values stale.
  const value = workbookCallsSheetFunction(pkg, target.index, CELL_STATE_FUNCTIONS);
  return { ok: true, pkg: next, inverse, value, formula: false };
}
