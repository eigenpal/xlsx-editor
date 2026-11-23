// The merge ops: merge a range of one worksheet into one cell, and unmerge the merged ranges in a
// range (ECMA-376 Part 1, 18.3.1.54 `x:mergeCell` and 18.3.1.55 `x:mergeCells`). In sml.xsd,
// `CT_MergeCell/@ref` is a required `ST_Ref`, and `CT_MergeCells/@count` is an optional
// `xsd:unsignedInt`.
//
// - `mergeCells` adds one `x:mergeCell/@ref`. A sheet with no `x:mergeCells` gets one in its
//   schema place in `CT_Worksheet`, and `@count` follows the number of children. A merge keeps
//   the value of the top-left cell only, so the op refuses a range where another cell holds a
//   value or a formula (`merge-drops-values`), with the first such cell in reading order as its
//   subject: the caller clears them first, in the same transaction, when the user accepts the
//   loss. A range that meets a merged range is refused (`merge-overlap`) with that range as its
//   subject, and a range of one cell is refused (`merge-one-cell`). The new `@ref` is written
//   top-left first. The op never reads the cells that a range names: it reads the cells that
//   exist (AGENTS.md, "Security").
// - `unmergeCells` removes each `x:mergeCell` whose range meets the op range, and the empty
//   `x:mergeCells`, because `CT_MergeCells` needs at least one child. So an unmerge of one cell
//   of a merged range unmerges the whole range. A range that meets no merged range is a no-op.
//
// The inverse puts back the worksheet part (`restoreParts`), as the inverse of a resize does.

import type { Element, Node, XmlPart } from '../package/tree.ts';
import type { XlsxPackage } from '../package/package.ts';
import { attr } from '../package/query.ts';
import { formatCellRef } from '../cell-ref.ts';
import { cellKey, keyAddress, sheetIndex, type SheetIndex } from '../workbook/sheet-index.ts';
import { holdsWrappedValue, sheetTarget, withPart } from './cells.ts';
import { clearsValue, rangeBounds, type RangeBox } from './clear-range.ts';
import { IdSource, plainAttr, smlElement } from './tree-edit.ts';
import { refuse, unchanged, type OpResult, type WorkbookOp } from './ops.ts';

/** The merge ops, as `ops.ts` declares them. */
export type MergeOp = Extract<WorkbookOp, { readonly op: 'mergeCells' | 'unmergeCells' }>;

/**
 * The children of `x:worksheet` that come after `x:mergeCells` in `CT_Worksheet` (ECMA-376
 * Part 1, 18.3.1.99). A new `x:mergeCells` goes before the first of them.
 */
const AFTER_MERGE_CELLS = new Set([
  'phoneticPr',
  'conditionalFormatting',
  'dataValidations',
  'hyperlinks',
  'printOptions',
  'pageMargins',
  'pageSetup',
  'headerFooter',
  'rowBreaks',
  'colBreaks',
  'customProperties',
  'cellWatches',
  'ignoredErrors',
  'smartTags',
  'drawing',
  'legacyDrawing',
  'legacyDrawingHF',
  'drawingHF',
  'picture',
  'oleObjects',
  'controls',
  'webPublishItems',
  'tableParts',
  'extLst',
]);

const meets = (a: RangeBox, b: RangeBox): boolean =>
  a[0] <= b[2] && b[0] <= a[2] && a[1] <= b[3] && b[1] <= a[3];

/** The longest cell range reference of the grid, `XFD1048576:XFD1048576`. */
const MAX_REF_LENGTH = 21;

/** The parsed `x:mergeCell/@ref` of each worksheet root, so a selection move parses none. */
const boxesOf = new WeakMap<Element, readonly (readonly [string, RangeBox])[]>();

function mergeBoxes(root: Element): readonly (readonly [string, RangeBox])[] {
  const hit = boxesOf.get(root);
  if (hit) return hit;
  const boxes: (readonly [string, RangeBox])[] = [];
  for (const m of mergeList(root)?.children ?? []) {
    if (m.type !== 'element' || m.kind !== 'mergeCell') continue;
    const ref = attr(m, 'ref') ?? '';
    // A file-supplied ref longer than any range of the grid is no range: never split it.
    const b = ref.length <= MAX_REF_LENGTH ? rangeBounds(ref) : undefined;
    if (b) boxes.push([ref, b]);
  }
  boxesOf.set(root, boxes);
  return boxes;
}

/**
 * The first `x:mergeCell/@ref` of a worksheet root that meets `box`, or undefined. The merge op
 * and the editor's `selectionMerged` both read it, so a state and a command agree on every ref
 * that parses, also a one-cell or an overlapping ref of the file. The parsed refs are kept for
 * each root.
 */
export function mergedRefMeeting(root: Element, box: RangeBox): string | undefined {
  for (const [ref, b] of mergeBoxes(root)) if (meets(box, b)) return ref;
  return undefined;
}

/** How many `x:mergeCell/@ref` boxes of a worksheet root parse. */
export function mergedRangeCount(root: Element): number {
  return mergeBoxes(root).length;
}

/** Each parsed `x:mergeCell/@ref` box of a worksheet root that meets `box`, in file order. */
export function mergedBoxesMeeting(root: Element, box: RangeBox): readonly RangeBox[] {
  const out: RangeBox[] = [];
  for (const [, b] of mergeBoxes(root)) if (meets(box, b)) out.push(b);
  return out;
}

/** A cell whose content is about to change: whether it then holds a value or a formula. */
export interface CellOverride {
  readonly row: number;
  readonly col: number;
  readonly holds: boolean;
}

/**
 * The first cell of `box`, in reading order, that a merge of `box` would empty: a cell other than
 * the top-left cell that holds a value or a formula. Undefined when none does. The merge op refuses
 * with it, and the editor snapshot reads it so that a host can ask the user before the command.
 * It reads the cells that exist, never the cells that `box` names. `edit` is an open cell edit,
 * which a merge commits first: its cell then counts as it would be after the commit.
 */
export function firstValueLost(
  part: XmlPart,
  box: RangeBox,
  edit?: CellOverride
): string | undefined {
  const [top, left, bottom, right] = box;
  const keys = contentKeys(sheetIndex(part));
  const skip = (row: number, col: number): boolean =>
    (row === top && col === left) || (edit !== undefined && edit.row === row && edit.col === col);
  // Jump from row to row in key order: each step is one binary search, and the steps are at
  // most twice the rows of the box that hold content.
  let found: number | undefined;
  let at = lowerBound(keys, cellKey(top, left));
  while (at < keys.length) {
    const { row, col } = keyAddress(keys[at]!);
    if (row > bottom) break;
    if (col < left) at = lowerBound(keys, cellKey(row, left));
    else if (col > right) at = lowerBound(keys, cellKey(row + 1, left));
    else if (skip(row, col)) at++;
    else {
      found = keys[at];
      break;
    }
  }
  if (edit?.holds && !(edit.row === top && edit.col === left)) {
    const { row, col } = edit;
    const key = cellKey(row, col);
    const inside = row >= top && row <= bottom && col >= left && col <= right;
    if (inside && (found === undefined || key < found)) found = key;
  }
  return found === undefined ? undefined : formatCellRef(keyAddress(found));
}

/**
 * The keys of the cells that hold a value in any form, in key order, for each sheet index:
 * the cells whose value a clear removes (`clearsValue`: direct value content, a formula inside
 * `mc:AlternateContent`, or a value in `mc:AlternateContent`), and the cells with value content
 * in a child in another namespace (`holdsWrappedValue`). A clear refuses such a cell with
 * `cell-alternate-content`, so a merge that the user confirms is refused there and never hides
 * the content.
 */
const contentKeysOf = new WeakMap<SheetIndex, readonly number[]>();

function contentKeys(index: SheetIndex): readonly number[] {
  const hit = contentKeysOf.get(index);
  if (hit) return hit;
  const keys: number[] = [];
  for (const [key, e] of index.cells)
    if (clearsValue(e.el) || holdsWrappedValue(e.el)) keys.push(key);
  // The index keeps cells in document order, which a file can break: sort when it does.
  if (keys.some((k, i) => i > 0 && k < keys[i - 1]!)) keys.sort((x, y) => x - y);
  contentKeysOf.set(index, keys);
  return keys;
}

/** The first position in `keys` whose key is `key` or greater. */
function lowerBound(keys: readonly number[], key: number): number {
  let lo = 0;
  let hi = keys.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (keys[mid]! < key) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

/** The `x:mergeCells` child of a worksheet root. */
const mergeList = (root: Element): Element | undefined =>
  root.children.find((n): n is Element => n.type === 'element' && n.kind === 'mergeCells');

/** Apply a merge op. */
export function applyMerge(pkg: XlsxPackage, op: MergeOp): OpResult {
  const box = rangeBounds(op.ref);
  if (!box) return refuse('bad-reference', `${JSON.stringify(op.ref)} is not a range in the grid`);
  const target = sheetTarget(pkg, op.sheet);
  if ('code' in target) return { ok: false, rejection: target };
  const part = target.part;
  const list = mergeList(part.root);
  const merges = (list?.children ?? []).filter(
    (n): n is Element => n.type === 'element' && n.kind === 'mergeCell'
  );
  let root: Element;
  if (op.op === 'unmergeCells') {
    // The refs that `mergedRefMeeting` reads, so the state and the unmerge agree.
    const met = new Set(mergeBoxes(part.root).flatMap(([ref, b]) => (meets(box, b) ? [ref] : [])));
    const kept = merges.filter((m) => !met.has(attr(m, 'ref') ?? ''));
    if (!list || kept.length === merges.length) return unchanged(pkg);
    root = withMergeList(part.root, list, kept, null);
  } else {
    const [top, left, bottom, right] = box;
    if (top === bottom && left === right)
      return refuse('merge-one-cell', `${op.ref} is one cell, so there is nothing to merge`);
    const met = mergedRefMeeting(part.root, box);
    if (met !== undefined)
      return {
        ok: false,
        rejection: {
          code: 'merge-overlap',
          detail: `${op.ref} meets the merged ${met}`,
          subject: met,
        },
      };
    const lost = firstValueLost(part, box);
    if (lost) {
      const subject = lost;
      return {
        ok: false,
        rejection: {
          code: 'merge-drops-values',
          detail: `${op.ref}: ${subject} holds a value outside the top-left cell`,
          subject,
        },
      };
    }
    const ids = new IdSource(part);
    const from = formatCellRef({ row: top, col: left });
    const ref = `${from}:${formatCellRef({ row: bottom, col: right })}`;
    const cell = smlElement(ids, 'mergeCell', 'mergeCells', [plainAttr('ref', ref)]);
    root = withMergeList(part.root, list, [...merges, cell], ids);
    const next = { partName: part.partName, root, nextId: ids.next };
    return result(pkg, part.partName, next);
  }
  return result(pkg, part.partName, { ...part, root });
}

/** The changed package and the inverse that puts the old worksheet part back. */
function result(
  pkg: XlsxPackage,
  partName: string,
  xml: { readonly partName: string; readonly root: Element; readonly nextId: number }
): OpResult {
  return {
    ok: true,
    pkg: withPart(pkg, { kind: 'xml', partName, xml }),
    inverse: [{ op: 'restoreParts', parts: [{ partName, part: pkg.parts.get(partName)! }] }],
    value: false,
    formula: false,
  };
}

/**
 * The worksheet root with `merges` as the children of `x:mergeCells`, and `@count` set. No merge
 * removes the list; a sheet with no list gets one in its schema place (`ids` makes it).
 */
function withMergeList(
  root: Element,
  list: Element | undefined,
  merges: readonly Element[],
  ids: IdSource | null
): Element {
  const others = (list?.children ?? []).filter(
    (n) => !(n.type === 'element' && n.kind === 'mergeCell')
  );
  // No merge left removes the list with any other child: `CT_MergeCells` needs one `x:mergeCell`,
  // so a list with none is not valid, and other content in it can only be ignorable.
  if (list && merges.length === 0)
    return { ...root, children: root.children.filter((n) => n !== list) };
  if (list) {
    const count = list.attrs.filter((a) => !(a.ns === '' && a.local === 'count'));
    const next: Element = {
      ...list,
      attrs: [...count, plainAttr('count', String(merges.length))],
      children: [...merges, ...others],
    };
    return { ...root, children: root.children.map((n) => (n === list ? next : n)) };
  }
  const made = smlElement(
    ids!,
    'mergeCells',
    root.kind,
    [plainAttr('count', String(merges.length))],
    merges
  );
  const at = root.children.findIndex(
    (n: Node) => n.type === 'element' && AFTER_MERGE_CELLS.has(n.local)
  );
  const children = root.children.slice();
  children.splice(at < 0 ? children.length : at, 0, made);
  return { ...root, children };
}
