// The resize ops of design D3 of the change `workbook-structure-editing`: set the width of a
// range of columns or of listed columns, and the height of a range of rows or of listed rows, of
// one worksheet.
//
// - `setColumnWidth` writes `x:cols/x:col` with `@width` and `@customWidth="1"` over the range
//   (ECMA-376 Part 1, 18.3.1.13 and 18.3.1.17; `CT_Col`, sml.xsd). An `x:col` that the range
//   crosses splits at the range edges, and each piece keeps its other attributes (`style`,
//   `hidden`, `outlineLevel`), except that the resized piece drops `@bestFit`, because its width is
//   no longer a best fit. Columns of the range that no `x:col` covers get a new `x:col` with no
//   other attribute. A sheet without `x:cols` gets one, in its schema place before `x:sheetData`
//   (18.3.1.99).
// - `setRowHeight` writes `x:row/@ht` and `@customHeight="1"` (18.3.1.73; `CT_Row`). A row of the
//   range with no `x:row` gets an empty one, in row order. A sheet is sparse, so the range can
//   create at most `MAX_RESIZE_NEW_ROWS` rows; a larger one is refused with `too-many-rows`. Rows
//   out of order in `x:sheetData` refuse the op with `invalid-sheet`, so it never writes a second
//   `x:row` for one row.
// - `setColumnWidths` writes the width of each listed column in one pass over `x:cols`, as
//   `setColumnWidth` writes one range. A fit of many columns costs one pass, not one for each.
// - `setFittedRowHeights` writes the height that fits the content of each listed row, in one
//   pass over `x:sheetData`: `@ht` without `@customHeight`, because a fitted height is not a
//   height set by hand (18.3.1.73), so a later font change or value edit fits the row again
//   (`format-range.ts`, `cell-batch.ts`). It sets only rows that have an `x:row`, and it creates
//   none.
//
// A drawing anchor that keeps its size or position in points (`xdr:twoCellAnchor` with
// `@editAs` `oneCell` or `absolute`), or the anchor of an embedded control or object, refuses a
// resize that would change it, with `resize-anchor` and the shape or control name as the
// subject when the file gives one (`resizeBlocker`). A fit counts the loss of a true
// `@customHeight` (`1` or `true`) as such a change, and a column that only gains `@customWidth`
// at its width as none.
//
// A column width changes the result of `CELL("width")` and of a `GET.*` macro such as
// `GET.CELL`, and a row height changes the result of `GET.CELL`. So `setColumnWidth` and
// `setColumnWidths` mark the cached values stale, as a value edit does, when a formula of a cell
// or a defined name calls `CELL` or a `GET.*` macro, and `setRowHeight` and
// `setFittedRowHeights` mark them stale when one calls a `GET.*` macro.
//
// `x:sheetFormatPr` stays as it is. Each op loops over the elements that the file holds and over
// the range that the op names, never over a range that the file states. The inverse puts back
// the worksheet part (`restoreParts`), as the inverse of a structural edit does.

import type { Element, Node } from '../package/tree.ts';
import type { XlsxPackage } from '../package/package.ts';
import { attr } from '../package/query.ts';
import { MAX_COLS, MAX_ROWS } from '../cell-ref.ts';
import { resizeBlocker } from '../structure/drawing.ts';
import {
  CELL_STATE_FUNCTIONS,
  ROW_HEIGHT_FUNCTIONS,
  workbookCallsSheetFunction,
} from '../structure/sheet-functions.ts';
import { sameNode } from './calc.ts';
import { sheetTarget, withPart, type SheetTarget } from './cells.ts';
import {
  IdSource,
  plainAttr,
  smlElement,
  withAttr,
  withChild,
  withParentDeclarations,
} from './tree-edit.ts';
import { refuse, unchanged, type OpResult, type WorkbookOp } from './ops.ts';

/**
 * The widest column, in characters: this editor's limit. The schema types `x:col/@width` as a
 * double with no bound (`sml.xsd`); `axis-size.test.ts` keeps the limit.
 *
 * @public
 */
export const MAX_COLUMN_WIDTH = 255;
/**
 * The tallest row, in points: this editor's limit. The schema types `x:row/@ht` as a double
 * with no bound (`sml.xsd`); `axis-size.test.ts` keeps the limit.
 *
 * @public
 */
export const MAX_ROW_HEIGHT = 409;
/**
 * The most `x:row` elements that one `setRowHeight` creates: as many as one format command.
 *
 * @public
 */
export const MAX_RESIZE_NEW_ROWS = 16_384;

/** The resize ops, as `ops.ts` declares them. */
export type AxisSizeOp = Extract<WorkbookOp, { readonly op: 'setColumnWidth' | 'setRowHeight' }>;

/**
 * A number as the file writes it: at most 8 decimals, with no exponent, so `1e-7` is written
 * `0.0000001`. Trailing zeros go.
 */
const numberText = (n: number): string =>
  (Math.round(n * 1e8) / 1e8).toFixed(8).replace(/\.?0+$/, '');

/** Apply a resize op. */
export function applyAxisSize(pkg: XlsxPackage, op: AxisSizeOp): OpResult {
  const columns = op.op === 'setColumnWidth';
  const limit = columns ? MAX_COLS : MAX_ROWS;
  const { first, last } = op;
  if (
    !Number.isInteger(first) ||
    !Number.isInteger(last) ||
    first < 1 ||
    last < first ||
    last > limit
  )
    return refuse(
      'bad-reference',
      `${op.op} ${first} to ${last} is outside the ${columns ? 'columns' : 'rows'} of the grid`
    );
  const size = columns ? op.width : op.height;
  const max = columns ? MAX_COLUMN_WIDTH : MAX_ROW_HEIGHT;
  if (typeof size !== 'number' || !Number.isFinite(size) || size < 0 || size > max)
    return refuse('bad-size', `${op.op}: ${size} is outside 0 to ${max}`);
  const target = sheetTarget(pkg, op.sheet);
  if ('code' in target) return { ok: false, rejection: target };
  const part = target.part;
  const ids = new IdSource(part);
  const sized = columns
    ? withColumnWidths(part.root, target.sheetData, ids, [{ first, last, width: numberText(size) }])
    : withRowHeight(part.root, target.sheetData, ids, first, last, numberText(size));
  if ('code' in sized) return { ok: false, rejection: sized };
  // `withColumnWidths` returns `{ root, changed }`, and `withRowHeight` returns the root itself.
  const root = 'root' in sized ? sized.root : sized;
  if (sameNode(part.root, root)) return unchanged(pkg);
  const anchor = resizeBlocker(pkg, part, columns ? 'columns' : 'rows', first - 1, last - 1);
  if (anchor) return { ok: false, rejection: anchor };
  return resized(pkg, target, root, ids, columns);
}

/** The result of a resize that changed the worksheet root to `root`. */
function resized(
  pkg: XlsxPackage,
  target: SheetTarget,
  root: Element,
  ids: IdSource,
  columns: boolean
): OpResult {
  const part = target.part;
  const next = { partName: part.partName, root, nextId: ids.next };
  return {
    ok: true,
    pkg: withPart(pkg, { kind: 'xml', partName: part.partName, xml: next }),
    inverse: [
      {
        op: 'restoreParts',
        parts: [{ partName: part.partName, part: pkg.parts.get(part.partName)! }],
      },
    ],
    // A size changes the result of `CELL("width")` or `GET.CELL`: a call marks the values stale.
    value: workbookCallsSheetFunction(
      pkg,
      target.index,
      columns ? CELL_STATE_FUNCTIONS : ROW_HEIGHT_FUNCTIONS
    ),
    formula: false,
  };
}

/** Apply `setColumnWidths`. */
export function applyColumnWidths(
  pkg: XlsxPackage,
  op: Extract<WorkbookOp, { readonly op: 'setColumnWidths' }>
): OpResult {
  // Adjacent columns with one width join in one run.
  const runs: { first: number; last: number; width: string }[] = [];
  let previous = 0;
  for (const entry of Array.isArray(op.columns) ? op.columns : []) {
    const { column, width } = (entry ?? {}) as { column?: unknown; width?: unknown };
    if (
      typeof column !== 'number' ||
      !Number.isInteger(column) ||
      column <= previous ||
      column > MAX_COLS
    )
      return refuse('bad-reference', `setColumnWidths: column ${column} is out of order or range`);
    if (
      typeof width !== 'number' ||
      !Number.isFinite(width) ||
      width < 0 ||
      width > MAX_COLUMN_WIDTH
    )
      return refuse('bad-size', `setColumnWidths: ${width} is outside 0 to ${MAX_COLUMN_WIDTH}`);
    const text = numberText(width);
    const run = runs[runs.length - 1];
    if (run && run.last === column - 1 && run.width === text) run.last = column;
    else runs.push({ first: column, last: column, width: text });
    previous = column;
  }
  if (runs.length === 0) return refuse('bad-reference', 'setColumnWidths: no columns');
  const target = sheetTarget(pkg, op.sheet);
  if ('code' in target) return { ok: false, rejection: target };
  const part = target.part;
  const ids = new IdSource(part);
  const sized = withColumnWidths(part.root, target.sheetData, ids, runs);
  if ('code' in sized) return { ok: false, rejection: sized };
  // A column already at its width is no change, even where its `x:col` would split.
  if (sameNode(part.root, sized.root) || sized.changed.length === 0) return unchanged(pkg);
  // An anchor refuses only a width change that meets it.
  for (const [first, last] of sized.moved) {
    const anchor = resizeBlocker(pkg, part, 'columns', first - 1, last - 1);
    if (anchor) return { ok: false, rejection: anchor };
  }
  return resized(pkg, target, sized.root, ids, true);
}

/** Apply `setFittedRowHeights`. */
export function applyFittedRowHeights(
  pkg: XlsxPackage,
  op: Extract<WorkbookOp, { readonly op: 'setFittedRowHeights' }>
): OpResult {
  const heights = new Map<number, string>();
  let previous = 0;
  for (const entry of Array.isArray(op.rows) ? op.rows : []) {
    const { row, height } = (entry ?? {}) as { row?: unknown; height?: unknown };
    if (typeof row !== 'number' || !Number.isInteger(row) || row <= previous || row > MAX_ROWS)
      return refuse('bad-reference', `setFittedRowHeights: row ${row} is out of order or range`);
    if (
      typeof height !== 'number' ||
      !Number.isFinite(height) ||
      height < 0 ||
      height > MAX_ROW_HEIGHT
    )
      return refuse('bad-size', `setFittedRowHeights: ${height} is outside 0 to ${MAX_ROW_HEIGHT}`);
    heights.set(row, numberText(height));
    previous = row;
  }
  if (heights.size === 0) return refuse('bad-reference', 'setFittedRowHeights: no rows');
  const target = sheetTarget(pkg, op.sheet);
  if ('code' in target) return { ok: false, rejection: target };
  const part = target.part;
  const sheetData = target.sheetData;
  const found = new Set<number>();
  // Whether a row changes, and the runs of rows whose height changes or whose `@customHeight`
  // goes: an anchor refuses such a change that meets it. A row without `@customHeight` takes its
  // automatic height again after a value edit (`cell-batch.ts`), so its anchor would move later.
  let changed = false;
  const runs: [number, number][] = [];
  let lastRow = 0;
  const out: Node[] = [];
  for (const n of sheetData.children) {
    if (n.type !== 'element' || n.kind !== 'row') {
      out.push(n);
      continue;
    }
    const raw = attr(n, 'r');
    const at = raw === undefined ? lastRow + 1 : Number(raw);
    // As in `withRowHeight`: rows out of order, or two rows with one number, are refused.
    if (!(at > lastRow))
      return refuse('invalid-sheet', `setFittedRowHeights: x:row ${raw ?? at} is out of row order`);
    lastRow = at;
    const height = heights.get(lastRow);
    const el =
      height === undefined ? n : withAttr(withAttr(n, 'ht', height), 'customHeight', undefined);
    if (height !== undefined) found.add(lastRow);
    out.push(sameNode(n, el) ? n : el);
    if (sameNode(n, el)) continue;
    changed = true;
    // `@customHeight` counts only when it is true (`1` or `true`), as layout reads it.
    const custom = attr(n, 'customHeight');
    if (Number(attr(n, 'ht')) !== Number(height) || custom === '1' || custom === 'true')
      addRange(runs, lastRow, lastRow);
  }
  for (const row of heights.keys())
    if (!found.has(row))
      return refuse('bad-reference', `setFittedRowHeights: row ${row} has no x:row`);
  if (!changed) return unchanged(pkg);
  for (const [first, last] of runs) {
    const anchor = resizeBlocker(pkg, part, 'rows', first - 1, last - 1);
    if (anchor) return { ok: false, rejection: anchor };
  }
  const root = withChild(part.root, part.root.children.indexOf(sheetData), {
    ...sheetData,
    children: out,
  });
  // The op makes no new element, so the id source only carries the part's next id.
  return resized(pkg, target, root, new IdSource(part), false);
}

type Refusal = { readonly code: 'invalid-sheet'; readonly detail: string };

/** A whole-number attribute from 1 to `max`, or undefined. */
function index(el: Element, name: string, max: number): number | undefined {
  const raw = attr(el, name);
  if (raw === undefined || !/^[0-9]{1,8}$/.test(raw)) return undefined;
  const n = Number(raw);
  return n >= 1 && n <= max ? n : undefined;
}

/** A run of columns, 1-based, with one width as the file writes it. */
interface WidthRun {
  readonly first: number;
  readonly last: number;
  readonly width: string;
}

/** Add the 1-based range `a` to `b` to ascending `ranges`, joined with the last when adjacent. */
function addRange(ranges: [number, number][], a: number, b: number): void {
  const last = ranges[ranges.length - 1];
  if (last && last[1] === a - 1) last[1] = b;
  else ranges.push([a, b]);
}

/**
 * The worksheet root with each run of columns at its width, in one pass over `x:cols`, the
 * ranges whose `x:col` changes (`changed`), and the ranges whose width changes (`moved`): only
 * these move an anchor. `runs` are in ascending order and do not overlap.
 */
function withColumnWidths(
  root: Element,
  sheetData: Element,
  ids: IdSource,
  runs: readonly WidthRun[]
):
  | {
      readonly root: Element;
      readonly changed: readonly [number, number][];
      readonly moved: readonly [number, number][];
    }
  | Refusal {
  // A worksheet may hold more than one `x:cols` (Part 1, 18.3.1.17): the columns of all of them
  // go to the first, in order, which is the same set of column records. A column that moves takes
  // the namespace declarations of the list it leaves (`withParentDeclarations`). When the first
  // list declares a prefix of the root again with another URI, the column also declares the
  // binding of the root again. So its prefixes, and the prefixes that its `mc:Ignorable` names,
  // keep their meaning (ECMA-376 Part 3, 7.2).
  const lists = root.children.filter(
    (n): n is Element => n.type === 'element' && n.kind === 'cols'
  );
  const cols = lists[0];
  const shadowed = root.nsDecls.filter((d) =>
    cols?.nsDecls.some((c) => c.prefix === d.prefix && c.uri !== d.uri)
  );
  const moveOut = (node: Element, list: Element): Element => {
    const n = withParentDeclarations(node, list);
    const own = new Set(n.nsDecls.map((d) => d.prefix));
    const added = shadowed.filter((d) => !own.has(d.prefix));
    return added.length === 0 ? n : { ...n, nsDecls: [...added, ...n.nsDecls] };
  };
  // The merge removes each later list. The attributes of a later list, such as `mc:Ignorable`,
  // have no place to go, so the edit is refused rather than dropping them.
  const marked = lists.find((l) => l !== cols && l.attrs.length > 0)?.attrs[0];
  if (marked) {
    const name = marked.prefix ? `${marked.prefix}:${marked.local}` : marked.local;
    return {
      code: 'invalid-sheet',
      detail: `a column width edit cannot keep the attribute ${name} of an x:cols after the first`,
    };
  }
  const pieces: { min: number; el: Element }[] = [];
  const changed: [number, number][] = [];
  const moved: [number, number][] = [];
  const set = (el: Element, min: number, max: number, width?: string): Element => {
    let out = withAttr(withAttr(el, 'min', String(min)), 'max', String(max));
    // A width set by hand is not a best fit: the sized piece drops `@bestFit`.
    if (width !== undefined)
      out = withAttr(
        withAttr(withAttr(out, 'width', width), 'customWidth', '1'),
        'bestFit',
        undefined
      );
    return out;
  };
  /** The index of the first run that ends at or after `col`. */
  const runAt = (col: number) => {
    let lo = 0;
    let hi = runs.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (runs[mid]!.last < col) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  };
  const covered: [number, number][] = [];
  for (const list of lists)
    for (const node of list.children) {
      if (node.type !== 'element' || node.kind !== 'col')
        return { code: 'invalid-sheet', detail: 'x:cols holds a child that is not an x:col' };
      const n = list === cols ? node : moveOut(node, list);
      const min = index(n, 'min', MAX_COLS);
      const max = index(n, 'max', MAX_COLS);
      if (min === undefined || max === undefined || max < min)
        return { code: 'invalid-sheet', detail: 'an x:col has no valid min and max' };
      let r = runAt(min);
      if (r >= runs.length || runs[r]!.first > max) {
        pieces.push({ min, el: n });
        continue;
      }
      // The first piece keeps the element id; each other piece is a new element.
      let reuse = true;
      const piece = (a: number, b: number, width?: string) => {
        const el = set(n, a, b, width);
        pieces.push({ min: a, el: reuse ? el : { ...el, id: ids.id() } });
        reuse = false;
      };
      const alreadyAt = (width: string) =>
        attr(n, 'width') === width &&
        attr(n, 'customWidth') === '1' &&
        attr(n, 'bestFit') === undefined;
      let at = min;
      for (; r < runs.length && runs[r]!.first <= max; r += 1) {
        const run = runs[r]!;
        const a = Math.max(min, run.first);
        const b = Math.min(max, run.last);
        if (a > at) piece(at, a - 1);
        piece(a, b, run.width);
        covered.push([a, b]);
        if (!alreadyAt(run.width)) addRange(changed, a, b);
        if (Number(attr(n, 'width')) !== Number(run.width)) addRange(moved, a, b);
        at = b + 1;
      }
      if (at <= max) piece(at, max);
    }
  // The columns of each run that no `x:col` covers get a new one each gap.
  covered.sort((x, y) => x[0] - y[0]);
  const gap = (a: number, b: number, width: string) => {
    const attrs = [
      plainAttr('min', String(a)),
      plainAttr('max', String(b)),
      plainAttr('width', width),
      plainAttr('customWidth', '1'),
    ];
    pieces.push({ min: a, el: smlElement(ids, 'col', 'cols', attrs) });
    addRange(changed, a, b);
    addRange(moved, a, b);
  };
  let c = 0;
  for (const run of runs) {
    let next = run.first;
    while (c < covered.length && covered[c]![1] < run.first) c += 1;
    for (let i = c; i < covered.length && covered[i]![0] <= run.last; i += 1) {
      const [a, b] = covered[i]!;
      if (a > next) gap(next, a - 1, run.width);
      next = Math.max(next, b + 1);
    }
    if (next <= run.last) gap(next, run.last, run.width);
  }
  changed.sort((x, y) => x[0] - y[0]);
  moved.sort((x, y) => x[0] - y[0]);
  pieces.sort((x, y) => x.min - y.min);
  const children = pieces.map((p) => p.el);
  if (!cols) {
    const at = root.children.indexOf(sheetData);
    return {
      root: withChild(root, at, smlElement(ids, 'cols', root.kind, [], children), true),
      changed,
      moved,
    };
  }
  const kept = root.children.filter((n) => n === cols || !lists.includes(n as Element));
  return {
    root: { ...root, children: kept.map((n) => (n === cols ? { ...cols, children } : n)) },
    changed,
    moved,
  };
}

/** The worksheet root with the rows `first` to `last` at `height` points. */
function withRowHeight(
  root: Element,
  sheetData: Element,
  ids: IdSource,
  first: number,
  last: number,
  height: string
): Element | Refusal | { readonly code: 'too-many-rows'; readonly detail: string } {
  const sized = (el: Element) => withAttr(withAttr(el, 'ht', height), 'customHeight', '1');
  const existing = new Set<number>();
  let lastRow = 0;
  for (const n of sheetData.children) {
    if (n.type !== 'element' || n.kind !== 'row') continue;
    const raw = attr(n, 'r');
    const at = raw === undefined ? lastRow + 1 : Number(raw);
    // Rows in order are the only order in which the new rows have one place each.
    if (!(at > lastRow))
      return { code: 'invalid-sheet', detail: `x:row ${raw ?? at} is out of row order` };
    lastRow = at;
    if (lastRow >= first && lastRow <= last) existing.add(lastRow);
  }
  const missing = last - first + 1 - existing.size;
  if (missing > MAX_RESIZE_NEW_ROWS)
    return {
      code: 'too-many-rows',
      detail: `setRowHeight would create ${missing} rows, more than ${MAX_RESIZE_NEW_ROWS}`,
    };
  const out: Node[] = [];
  let next = first;
  const fill = (upTo: number) => {
    for (; next <= Math.min(upTo, last); next += 1)
      out.push(
        smlElement(ids, 'row', sheetData.kind, [
          plainAttr('r', String(next)),
          plainAttr('ht', height),
          plainAttr('customHeight', '1'),
        ])
      );
  };
  lastRow = 0;
  for (const n of sheetData.children) {
    if (n.type !== 'element' || n.kind !== 'row') {
      out.push(n);
      continue;
    }
    const raw = attr(n, 'r');
    lastRow = raw === undefined ? lastRow + 1 : Number(raw);
    fill(lastRow - 1);
    if (lastRow >= first && lastRow <= last) {
      out.push(sized(n));
      next = Math.max(next, lastRow + 1);
    } else out.push(n);
  }
  fill(last);
  const at = root.children.indexOf(sheetData);
  return withChild(root, at, { ...sheetData, children: out });
}
