// Cell and row addressing for the store ops: the worksheet a sheet name names, the slot of a row
// and of a cell, the cell element an input builds, and the array formula ranges that refuse an
// edit (design D4). Shared by the ops of `ops.ts` and `clear-range.ts`. A value edit keeps the
// children of a cell that it does not replace; `cellChildBlocker` records how its rule for
// `mc:AlternateContent` differs from PPTX Editor's `alternate-copies.ts`.

import type { Element, Node, XmlPart } from '../package/tree.ts';
import { xmlPartOf, type PackagePart, type XlsxPackage } from '../package/package.ts';
import { attr, child } from '../package/query.ts';
import { formatCellRef, parseCellRef } from '../cell-ref.ts';
import { sheetByName, workbookIndex, type WorkbookIndex } from '../workbook/workbook-index.ts';
import { sheetIndex } from '../workbook/sheet-index.ts';
import { wrappedArrayRefs } from '../workbook/cell-formula.ts';
import {
  childIndex,
  IdSource,
  plainAttr,
  smlElement,
  textNode,
  withAttr,
  withChild,
  withNamespacesOf,
} from './tree-edit.ts';
import type { CellInput, OpRejection } from './ops.ts';
import { NS } from '../package/namespaces.ts';
import { keepsSubtreeMemo } from '../package/subtree-memo-policy.ts';

export function withPart(pkg: XlsxPackage, part: PackagePart): XlsxPackage {
  const parts = new Map(pkg.parts);
  parts.set(part.partName, part);
  return { parts };
}

export interface SheetTarget {
  readonly index: WorkbookIndex;
  readonly part: XmlPart;
  readonly sheetData: Element;
}

export function sheetTarget(pkg: XlsxPackage, name: string): SheetTarget | OpRejection {
  const index = workbookIndex(pkg);
  const entry = index ? sheetByName(index, name) : undefined;
  const part = xmlPartOf(pkg, entry?.partName);
  if (!index || !entry || !part || entry.kind !== 'worksheet')
    return { code: 'unknown-sheet', detail: `no worksheet named ${JSON.stringify(name)}` };
  const cells = sheetIndex(part);
  if (cells.problems.length > 0)
    return {
      code: 'invalid-sheet',
      detail: `${entry.name}: ${cells.problems[0]!.code} at ${cells.problems[0]!.at ?? '?'}`,
    };
  if (!cells.sheetData)
    return { code: 'invalid-sheet', detail: `${entry.name}: the sheet has no x:sheetData` };
  return { index, part, sheetData: cells.sheetData };
}

/** Rebuild the worksheet part with a new `x:sheetData`. */
export function withSheetData(part: XmlPart, sheetData: Element, ids: IdSource): XmlPart {
  const old = child(part.root, 'x:sheetData')!;
  const root = withChild(part.root, childIndex(part.root, old), sheetData);
  return { partName: part.partName, root, nextId: ids.next };
}

/** Where row `row` (0-based) is in `x:sheetData`, or where it would be inserted. */
export function rowSlot(
  sheetData: Element,
  row: number
): { index: number; el: Element | null; prevRow: number } {
  const kids = sheetData.children;
  let lastRow = -1;
  for (let i = 0; i < kids.length; i += 1) {
    const node = kids[i]!;
    if (node.type !== 'element' || node.kind !== 'row') continue;
    const raw = attr(node, 'r');
    const r = raw === undefined ? lastRow + 1 : Number(raw) - 1;
    if (r === row) return { index: i, el: node, prevRow: lastRow };
    if (r > row) return { index: i, el: null, prevRow: lastRow };
    lastRow = r;
  }
  return { index: kids.length, el: null, prevRow: lastRow };
}

/**
 * Whether a row element is bare: no child element, and no attribute but `@r` and `@spans`. A
 * bare row says nothing, so a write that empties it drops it; a row with a height, a format,
 * or any other attribute stays.
 */
export function isBareRow(el: Element): boolean {
  return (
    el.children.every((n) => n.type !== 'element') &&
    el.attrs.every((a) => !a.ns && (a.local === 'r' || a.local === 'spans'))
  );
}

/**
 * The row with an explicit `@r` on each cell that has none, from its position: a cell without
 * `@r` follows the one before it. An edit that removes a cell pins the cells first, so a cell
 * after the removed one keeps its column. A row whose cells all have `@r` comes back as it is.
 */
export function pinCellRefs(rowEl: Element, row: number): Element {
  let lastCol = -1;
  let changed = false;
  const children = rowEl.children.map((n) => {
    if (n.type !== 'element' || n.kind !== 'cell') return n;
    const ref = attr(n, 'r');
    const col = ref === undefined ? lastCol + 1 : (parseCellRef(ref)?.col ?? lastCol + 1);
    lastCol = col;
    if (ref !== undefined) return n;
    changed = true;
    return withAttr(n, 'r', formatCellRef({ row, col }));
  });
  return changed ? { ...rowEl, children } : rowEl;
}

/**
 * `x:sheetData` with an explicit `@r` on each row that has none, from its position: a row
 * without `@r` is the row after the one before it. An edit that removes a row pins the rows
 * first, so a row after the removed one keeps its number.
 */
export function pinRowRefs(sheetData: Element): Element {
  let lastRow = 0;
  let changed = false;
  const children = sheetData.children.map((n) => {
    if (n.type !== 'element' || n.kind !== 'row') return n;
    const raw = attr(n, 'r');
    const row = raw === undefined ? lastRow + 1 : Number(raw);
    lastRow = row;
    if (raw !== undefined) return n;
    changed = true;
    return withAttr(n, 'r', String(row));
  });
  return changed ? { ...sheetData, children } : sheetData;
}

/** Where column `col` (0-based) is in a row, or where its cell would be inserted. */
export function cellSlot(rowEl: Element, col: number): { index: number; el: Element | null } {
  let lastCol = -1;
  for (let i = 0; i < rowEl.children.length; i += 1) {
    const node = rowEl.children[i]!;
    if (node.type !== 'element' || node.kind !== 'cell') continue;
    const ref = attr(node, 'r');
    const c = ref === undefined ? lastCol + 1 : (parseCellRef(ref)?.col ?? lastCol + 1);
    lastCol = c;
    if (c === col) return { index: i, el: node };
    if (c > col) return { index: i, el: null };
  }
  // A cell goes before any trailing non-cell child (`x:extLst`).
  let index = rowEl.children.length;
  while (index > 0) {
    const prev = rowEl.children[index - 1]!;
    if (prev.type === 'element' && prev.kind === 'cell') break;
    index -= 1;
  }
  return { index, el: null };
}

/** Whether a node is an `x:f`, `x:v`, or `x:is`: the value content of a cell (`CT_Cell`). */
const isValueChild = (n: Node): boolean =>
  n.type === 'element' && n.ns === NS.x && (n.local === 'f' || n.local === 'v' || n.local === 'is');

/**
 * The deepest nesting of `mc:AlternateContent`, or of the subtree of an element in another
 * namespace, that the cell child checks read, counted from 0 at the child of the cell: nine
 * levels are read, and the tenth is not. Markup compatibility content can nest without a
 * limit (ECMA-376 Part 3, 7.5), and the file controls it, so deeper content is not read: a
 * wrapper past it is not value content, and a subtree past it counts as value content. Both
 * branches refuse the edit, so the limit never drops content, and it can be deeper than the
 * wrapper limit of a drawing (`MAX_WRAP_DEPTH` in `structure/drawing.ts`), which types anchors.
 */
const MAX_DEPTH = 8;

/**
 * Whether a node is an `x:v` or an `x:is`: a stored value, with no formula. A formula in
 * `mc:AlternateContent` is not replaced, because the calc chain, stale value, array, and shared
 * formula checks read only a direct `x:f`: the edit refuses that cell (`cellChildBlocker`).
 */
const isStoredValue = (n: Node): boolean =>
  n.type === 'element' && n.ns === NS.x && (n.local === 'v' || n.local === 'is');

/**
 * Whether a node is an `mc:AlternateContent` whose branches each hold a stored value and nothing
 * else but white space: `x:v`, `x:is`, or such a wrapper again. It gives another form of the
 * value of its cell, so a value edit replaces it with the rest of the value. A branch with no
 * value, or with text, is other content: the edit refuses its cell (`cellChildBlocker`).
 */
function isValueAlternate(n: Node, depth = 0): boolean {
  if (n.type !== 'element' || n.ns !== NS.mc || n.local !== 'AlternateContent') return false;
  if (depth > MAX_DEPTH || n.children.length === 0) return false;
  return n.children.every(
    (branch) =>
      branch.type === 'element' &&
      branch.ns === NS.mc &&
      (branch.local === 'Choice' || branch.local === 'Fallback') &&
      branch.children.some((k) => k.type === 'element') &&
      branch.children.every((k) =>
        k.type === 'element'
          ? isStoredValue(k) || isValueAlternate(k, depth + 1)
          : k.value.trim() === ''
      )
  );
}

/** Whether an element holds nothing but white space: an empty wrapper holds no value. */
const isBlank = (n: Element): boolean =>
  n.children.every((k) => k.type === 'text' && k.value.trim() === '');

/** Whether a value edit replaces a child of a cell: its value content, in any form. */
const isReplaced = (n: Node): boolean => isValueChild(n) || isValueAlternate(n);

// Work counters for tests, as `sheetIndexTestRecorder` (`workbook/sheet-index.ts`): a count is
// exact on every machine, a wall-clock limit is not.
let subtreeWalks = 0;
let subtreeNodes = 0;

/** @internal Counts the subtree walks of the cell child checks and the elements that they read. */
export function cellChildTestRecorder(): {
  readonly walks: number;
  readonly nodes: number;
  reset(): void;
} {
  return {
    get walks() {
      return subtreeWalks;
    },
    get nodes() {
      return subtreeNodes;
    },
    reset() {
      subtreeWalks = 0;
      subtreeNodes = 0;
    },
  };
}

/**
 * The answers of the cell child checks for each element, after DOCX Editor's subtree memos
 * (`storyCarriesCommentAnchor` in `packages/core/src/binding/story-text-reads.ts` at 6c11927b).
 * A tree node never changes, and the sender controls how wide a subtree is, so a batch that
 * reads or writes one cell many times walks each wide or deep subtree once. Only the nodes that
 * `keepsSubtreeMemo` admits get an entry: a node of leaves answers again from those leaves.
 */
const wrapsValueOf = new WeakMap<Element, boolean>();
const wrappedCells = new WeakMap<Element, boolean>();
const blockingChildOf = new WeakMap<Element, Element | null>();

/**
 * Whether an element in another namespace holds value content (`x:f`, `x:v`, or `x:is`) in its
 * subtree, which a consumer that processes its content would read as the value of the cell. One
 * pass over the subtree, once per admitted element; past `MAX_DEPTH` it counts as value content.
 */
function wrapsValue(n: Element): boolean {
  const known = wrapsValueOf.get(n);
  if (known !== undefined) return known;
  const found = walkWrapsValue(n);
  if (keepsSubtreeMemo(n)) wrapsValueOf.set(n, found);
  return found;
}

function walkWrapsValue(n: Element): boolean {
  subtreeWalks += 1;
  const stack: [Element, number][] = [[n, 0]];
  for (let top = stack.pop(); top; top = stack.pop()) {
    const [el, depth] = top;
    for (const k of el.children) {
      if (k.type !== 'element') continue;
      subtreeNodes += 1;
      if (isValueChild(k) || depth >= MAX_DEPTH) return true;
      stack.push([k, depth + 1]);
    }
  }
  return false;
}

/**
 * Whether a cell holds its value in an `mc:AlternateContent` child that an edit replaces. The
 * arrow keeps `some` from passing the child index as the nesting depth.
 */
export const holdsAlternateValue = (cell: Element): boolean =>
  cell.children.some((n) => isValueAlternate(n));

/**
 * Whether a cell holds value content (`x:f`, `x:v`, or `x:is`) in a child in another namespace,
 * `mc:AlternateContent` included, in any form. Read once per cell that the memo policy admits.
 */
export function holdsWrappedValue(cell: Element): boolean {
  const known = wrappedCells.get(cell);
  if (known !== undefined) return known;
  const found = cell.children.some((n) => n.type === 'element' && n.ns !== NS.x && wrapsValue(n));
  if (keepsSubtreeMemo(cell)) wrappedCells.set(cell, found);
  return found;
}

/** The first child of `cell` that a value edit can neither keep nor replace. */
function blockingChild(cell: Element): Element | null {
  const known = blockingChildOf.get(cell);
  if (known !== undefined) return known;
  // An empty wrapper holds no value, so the cell keeps it.
  const kept = cell.children.find(
    (n): n is Element =>
      n.type === 'element' &&
      !isReplaced(n) &&
      ((n.ns === NS.mc && !isBlank(n)) || (n.ns !== NS.x && wrapsValue(n)))
  );
  if (keepsSubtreeMemo(cell)) blockingChildOf.set(cell, kept ?? null);
  return kept ?? null;
}

// A line comment, not TSDoc, so that the shipped declarations carry no repository path.
// PPTX Editor removes the copies of a shape in the other branches of `mc:AlternateContent` that
// an edit leaves out of date (`core/src/automation/commands/alternate-copies.ts` of that
// repository at 9d7d60b7). A cell follows that rule for a wrapper that holds only a stored value:
// the edit replaces it. It refuses other content instead, because a formula in a branch feeds the
// calc chain, array, and shared formula rules, which read only a direct `x:f`, and an unknown
// branch can hold the value in a form that the editor cannot read; removing it would drop data.
/**
 * Why a value edit or a clear cannot keep a child of the cell `old`, or undefined. Two kinds of
 * child can hold another form of the value that the edit would leave out of date: a non-empty
 * child in the markup compatibility namespace that is not value content, and an element in
 * another namespace with value content in it. The edit refuses that cell with
 * `cell-alternate-content`, with the sheet and the cell as the subject, such as `Data!B5`.
 * `buildCell` keeps every other child.
 */
export function cellChildBlocker(
  sheet: string,
  ref: string,
  old: Element
): OpRejection | undefined {
  const kept = blockingChild(old);
  if (!kept) return undefined;
  return {
    code: 'cell-alternate-content',
    subject: `${sheet}!${ref}`,
    detail: `${ref} has the child ${kept.prefix ? `${kept.prefix}:` : ''}${kept.local}, which the edit can neither keep nor replace`,
  };
}

/**
 * The new cell element, or `null` when the cell goes away (an empty value on a cell with no
 * format and no other child). The cell keeps its format (`@s`), its id, its prefix, its namespace
 * declarations, and every child that the edit does not replace, in its place: `x:extLst`, an
 * element in another namespace, and content that the schema does not allow. Type, value,
 * formula, inline string, value metadata, and an `mc:AlternateContent` that holds only value
 * content are replaced. The new value content goes
 * where the first replaced child was or before the first `x:extLst`, whichever comes first, or
 * else last. The caller refuses a cell that `cellChildBlocker` names first. The new content
 * keeps the order of the `CT_Cell` sequence (sml.xsd): `x:f`, `x:v`, `x:is`, then `x:extLst`.
 *
 * One rule keeps the look of a cell in a formatted row or column (`inherited`, the format of
 * `inheritedFormat`; ECMA-376 Part 1, 18.3.1.73 and 18.3.1.13, draw an absent cell with it): a
 * new cell takes that format, and a cleared cell with no format of its own stays with `s="0"`,
 * so it does not take the row or column format, as `clearFormats` does.
 */
export function buildCell(
  ids: IdSource,
  ref: string,
  old: Element | null,
  value: CellInput,
  sstIndex: number | undefined,
  formulaText: string | undefined,
  inherited = 0
): Element | null {
  const own = attr(old ?? undefined, 's');
  const empty = value.type === 'empty';
  const style =
    own !== undefined || inherited === 0
      ? own
      : old
        ? empty
          ? '0'
          : undefined
        : empty
          ? undefined
          : String(inherited);
  const keep = (old?.attrs ?? []).filter(
    (a) =>
      !(a.ns === '' && (a.local === 'r' || a.local === 't' || a.local === 'vm' || a.local === 'cm'))
  );
  // White space between the children is not content: a cell with nothing else goes away.
  const others = (old?.children ?? []).filter(
    (n) => !isReplaced(n) && !(n.type === 'text' && n.value.trim() === '')
  );
  if (value.type === 'empty' && style === undefined && others.length === 0) return null;
  const attrs = [plainAttr('r', ref), ...keep];
  if (own === undefined && style !== undefined) attrs.push(plainAttr('s', style));
  const children: Node[] = [];
  const v = (text: string) => smlElement(ids, 'v', 'cell', [], [textNode(ids, text)]);
  // `@t` is an `ST_CellType` (sml.xsd): absent for a number (the default `n`), `b`, or `s`.
  // A date writes `d`, which the Transitional schema does not list (`CellInput`).
  switch (value.type) {
    case 'number':
      children.push(v(String(value.value)));
      break;
    case 'boolean':
      attrs.push(plainAttr('t', 'b'));
      children.push(v(value.value ? '1' : '0'));
      break;
    case 'string':
      attrs.push(plainAttr('t', 's'));
      children.push(v(String(sstIndex)));
      break;
    case 'date':
      attrs.push(plainAttr('t', 'd'));
      children.push(v(value.iso));
      break;
    case 'formula':
      // No cached value: the cell shows nothing and is stale until a consumer recalculates.
      children.push(smlElement(ids, 'f', 'cell', [], [textNode(ids, formulaText ?? value.text)]));
      break;
    case 'empty':
      break;
  }
  // The new value content takes the place of the first replaced child or of the first `x:extLst`,
  // whichever comes first (`CT_Cell` puts `x:extLst` last), or else goes last.
  const kids = old?.children ?? [];
  const firstOf = (test: (n: Node) => boolean) => {
    const i = kids.findIndex(test);
    return i < 0 ? kids.length : i;
  };
  const at = Math.min(
    firstOf(isReplaced),
    firstOf((n) => n.type === 'element' && n.ns === NS.x && n.local === 'extLst')
  );
  const out: Node[] = [];
  for (let i = 0; i <= kids.length; i += 1) {
    if (i === at) for (const k of children) out.push(k);
    const n = kids[i];
    if (n && !isReplaced(n)) out.push(n);
  }
  const el = smlElement(ids, 'c', 'row', attrs, out);
  return old ? { ...withNamespacesOf(old, el), id: old.id } : el;
}

const isDigit = (c: number): boolean => c >= 48 && c <= 57;

/** Whether `s` has `n` digits at `i`, as a number from `min` to `max`. */
function digitsAt(s: string, i: number, n: number, min: number, max: number): boolean {
  let value = 0;
  for (let k = i; k < i + n; k += 1) {
    if (!isDigit(s.charCodeAt(k))) return false;
    value = value * 10 + s.charCodeAt(k) - 48;
  }
  return value >= min && value <= max;
}

/**
 * Whether `s` is an ISO 8601 calendar date (`YYYY-MM-DD`) with an optional time
 * (`THH:MM`, `THH:MM:SS`, `THH:MM:SS.fff`) and no time zone: the form of `t="d"` values that
 * the cascade reads. One pass, no regular expression.
 */
export function isIsoDate(s: string): boolean {
  if (s.length < 10 || s.length > 29) return false;
  if (!digitsAt(s, 0, 4, 1, 9999) || s[4] !== '-' || s[7] !== '-') return false;
  if (!digitsAt(s, 5, 2, 1, 12) || !digitsAt(s, 8, 2, 1, 31)) return false;
  const year = Number(s.slice(0, 4));
  const month = Number(s.slice(5, 7));
  const leap = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
  const days = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month - 1]!;
  if (Number(s.slice(8, 10)) > days) return false;
  if (s.length === 10) return true;
  if (s[10] !== 'T' || s.length < 16) return false;
  if (!digitsAt(s, 11, 2, 0, 23) || s[13] !== ':' || !digitsAt(s, 14, 2, 0, 59)) return false;
  if (s.length === 16) return true;
  if (s[16] !== ':' || s.length < 19 || !digitsAt(s, 17, 2, 0, 59)) return false;
  if (s.length === 19) return true;
  return s[19] === '.' && s.length > 20 && digitsAt(s, 20, s.length - 20, 0, 999_999_999);
}

type ArrayRanges = readonly (readonly [number, number, number, number, string])[] & {
  /** The masters of the worksheet. */
  readonly masters: number;
  /** The masters whose `@ref` names no range of the grid. */
  readonly unreadable: number;
};

const arrayRanges = new WeakMap<XmlPart, ArrayRanges>();

/**
 * The array and data table ranges (`x:f/@t` of `array` or `dataTable`, with `@ref`) of a
 * worksheet, as 0-based bounds and the reference. A master in a direct `x:f` or inside an
 * `mc:AlternateContent` child of its cell counts. Read once per part from its cells that exist.
 */
function arrayRangesOf(part: XmlPart): ArrayRanges {
  let ranges = arrayRanges.get(part);
  if (!ranges) {
    const found: (readonly [number, number, number, number, string])[] = [];
    let masters = 0;
    let unreadable = 0;
    const add = (ref: string): void => {
      masters += 1;
      const parts = ref.split(':');
      const [a, b = a] = parts;
      const from = parseCellRef(a!);
      const to = parseCellRef(b!);
      // A reference of more than two parts is not a range (ST_Ref): the first two still count,
      // as an edit inside them is refused, and the master counts as unreadable too.
      if (!from || !to || parts.length > 2) unreadable += 1;
      if (!from || !to) return;
      // The range as the reader understood it, not the raw `@ref`: a refusal names it.
      const one = from.row === to.row && from.col === to.col;
      const text = one ? formatCellRef(from) : `${formatCellRef(from)}:${formatCellRef(to)}`;
      found.push([from.row, from.col, to.row, to.col, text]);
    };
    for (const entry of sheetIndex(part).cells.values()) {
      const f = child(entry.el, 'x:f');
      const ref = attr(f, 'ref');
      const t = attr(f, 't');
      if ((t === 'array' || t === 'dataTable') && ref) add(ref);
      for (const wrapped of wrappedArrayRefs(entry.el)) add(wrapped);
    }
    arrayRanges.set(part, (ranges = Object.assign(found, { masters, unreadable })));
  }
  return ranges;
}

/**
 * How many array and data table masters a worksheet has, and how many of them have an `@ref`
 * that names no range of the grid, which could cover any cell. Read once per part.
 */
export function arrayRangeCount(part: XmlPart): {
  readonly all: number;
  readonly unreadable: number;
} {
  const ranges = arrayRangesOf(part);
  return { all: ranges.masters, unreadable: ranges.unreadable };
}

/**
 * The first array or data table range that meets the 0-based bounds, or undefined. An edit
 * inside such a range is refused.
 */
export function arrayFormulaIn(
  part: XmlPart,
  top: number,
  left: number,
  bottom: number,
  right: number
): string | undefined {
  for (const [r0, c0, r1, c1, ref] of arrayRangesOf(part))
    if (top <= r1 && bottom >= r0 && left <= c1 && right >= c0) return ref;
  return undefined;
}

/** The array or data table range that holds a cell, or undefined. */
export function arrayFormulaAt(part: XmlPart, row: number, col: number): string | undefined {
  return arrayFormulaIn(part, row, col, row, col);
}
