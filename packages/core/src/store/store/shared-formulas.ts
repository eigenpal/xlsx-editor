// Unshare a shared formula before an edit overwrites its master cell (design D4).
//
// A shared formula keeps its text in one master cell (`x:f/@t="shared"` with `@ref` and `@si`).
// In `CT_CellFormula` (sml.xsd), `@t` is an `ST_CellFormulaType`, `@ref` an `ST_Ref`, and `@si`
// an `xsd:unsignedInt`.
// Each child cell names the master by `@si` and holds no text: its formula is the master's text
// moved by the child's offset. An edit that overwrites or clears the master would leave the
// children without a definition, so the store first writes each child as its own formula, with
// the tokenizer, and keeps each child's cached value. When the tokenizer cannot read the master
// text, or a moved reference leaves the grid, the edit is refused: a save never writes a child
// whose master is gone.
//
// Several masters are unshared in one pass over the cells: each child finds its master by `@si`.

import type { Element, Node, XmlPart } from '../package/tree.ts';
import { attr, child, textOf } from '../package/query.ts';
import { formatCellRef, parseCellRef } from '../cell-ref.ts';
import { sheetIndex, type CellEntry } from '../workbook/sheet-index.ts';
import { tokenizeFormula, type FormulaToken } from '../formula/tokenizer.ts';
import { formatTokens, moveFormula } from '../formula/format.ts';
import { wrappedSharedChildren } from '../workbook/cell-formula.ts';
import {
  childIndex,
  IdSource,
  smlElement,
  textNode,
  withChild,
  withNamespacesOf,
} from './tree-edit.ts';
import type { HistoryOp, OpRejection } from './ops.ts';

/** Whether a cell is the master of a shared formula. */
export function isSharedMaster(cell: Element): boolean {
  const f = child(cell, 'x:f');
  return attr(f, 't') === 'shared' && attr(f, 'ref') !== undefined;
}

const wrappedChildMaps = new WeakMap<XmlPart, ReadonlyMap<string, readonly CellEntry[]>>();

/**
 * The first child of a shared formula (`@si`) whose `x:f` is inside `mc:AlternateContent` in its
 * cell, other than the master's own cell, or none. The unshare reads only a direct `x:f`, so it
 * cannot rewrite that child: an edit or a clear of the master refuses. A clear that also covers
 * that child is refused for its cell first (`clearBlocker`). It reads the cells of the sheet
 * index: a sheet with wrapped rows refuses every edit before it comes here.
 */
function wrappedChildOf(part: XmlPart, si: string, master: CellEntry): CellEntry | undefined {
  let map = wrappedChildMaps.get(part);
  if (!map) {
    const found = new Map<string, CellEntry[]>();
    for (const entry of sheetIndex(part).cells.values())
      for (const s of wrappedSharedChildren(entry.el)) {
        const list = found.get(s);
        if (list) list.push(entry);
        else found.set(s, [entry]);
      }
    wrappedChildMaps.set(part, (map = found));
  }
  return map.get(si)?.find((e) => e.el !== master.el);
}

/**
 * The subject of a refusal for a cell of a formula range: the range `ref` (a master's `@ref`) as
 * the reader reads it, when it holds the cell, so the message names the whole range. Else the
 * cell. A shared child that blocks its master gets the master's range, which holds both cells.
 */
export function formulaRangeSubject(
  sheet: string,
  ref: string | undefined,
  cell: { readonly row: number; readonly col: number }
): string {
  const [a, b = a] = (ref ?? '').split(':');
  const from = parseCellRef(a!);
  const to = parseCellRef(b!);
  const holds =
    from &&
    to &&
    cell.row >= from.row &&
    cell.row <= to.row &&
    cell.col >= from.col &&
    cell.col <= to.col;
  if (!holds) return `${sheet}!${formatCellRef(cell)}`;
  const one = from.row === to.row && from.col === to.col;
  return one
    ? `${sheet}!${formatCellRef(from)}`
    : `${sheet}!${formatCellRef(from)}:${formatCellRef(to)}`;
}

const masterRange = (master: CellEntry): string | undefined => attr(child(master.el, 'x:f'), 'ref');

const masterMaps = new WeakMap<XmlPart, ReadonlyMap<string, CellEntry>>();

/** The master cell of each shared formula of a worksheet, by `@si`. Read once per part. */
function sharedMasters(part: XmlPart): ReadonlyMap<string, CellEntry> {
  let map = masterMaps.get(part);
  if (!map) {
    const found = new Map<string, CellEntry>();
    for (const entry of sheetIndex(part).cells.values()) {
      if (!isSharedMaster(entry.el)) continue;
      const si = attr(child(entry.el, 'x:f'), 'si');
      if (si !== undefined && !found.has(si)) found.set(si, entry);
    }
    masterMaps.set(part, (map = found));
  }
  return map;
}

/**
 * The formula text of a cell, as a formula bar shows it without `=`: the text of `x:f`, or for
 * a shared formula child, the master's text moved by the child's offset. Undefined when the
 * cell has no formula, or the child's master text cannot be read. `f` is the cell's direct
 * `x:f` by default; the formula bar passes the `x:f` of a wrapper branch (`wrappedFormula`), so
 * a wrapped shared child shows the formula of its direct master too.
 */
export function sharedFormulaText(
  part: XmlPart,
  entry: CellEntry,
  f: Element | undefined = child(entry.el, 'x:f')
): string | undefined {
  if (!f) return undefined;
  const own = textOf(f);
  if (own.length > 0 || attr(f, 't') !== 'shared') return own.length > 0 ? own : undefined;
  const si = attr(f, 'si');
  const master = si === undefined ? undefined : sharedMasters(part).get(si);
  if (!master) return undefined;
  const moved = moveFormula(textOf(child(master.el, 'x:f')), {
    rows: entry.row - master.row,
    cols: entry.col - master.col,
  });
  return moved.ok ? moved.text : undefined;
}

/** The tokenizer result of each master `x:f`. The tree is immutable, so an element keeps it. */
const masterTokens = new WeakMap<Element, ReturnType<typeof tokenizeFormula>>();

function tokensOf(f: Element): ReturnType<typeof tokenizeFormula> {
  let read = masterTokens.get(f);
  if (!read) masterTokens.set(f, (read = tokenizeFormula(textOf(f))));
  return read;
}

/** The formula tokens of each shared master, by `@si`. */
type SharedTokens = Map<string, { master: CellEntry; tokens: readonly FormulaToken[] }>;

/**
 * The formula tokens of shared masters by `@si`, or the refusal of their unshare: a master with
 * a child inside `mc:AlternateContent`, or a master whose formula text cannot be read. The
 * unshare and the clear state of the toolbar (`clearRangeBlocker`) read this one check.
 */
function readMasters(
  part: XmlPart,
  sheet: string,
  masters: readonly CellEntry[]
): SharedTokens | OpRejection {
  const bySi: SharedTokens = new Map();
  const checked = new Set<string>();
  for (const master of masters) {
    const f = child(master.el, 'x:f')!;
    const si = attr(f, 'si');
    // The rewrite reads only a direct `x:f`, so it cannot rewrite a wrapped child. Each `@si`
    // is checked once, so many masters with one `@si` cost one pass over its children.
    const wrapped =
      si !== undefined && !checked.has(si) ? wrappedChildOf(part, si, master) : undefined;
    if (si !== undefined) checked.add(si);
    if (wrapped)
      return {
        code: 'shared-formula',
        detail: `${formatCellRef(master)} is a shared formula master with the child ${formatCellRef(wrapped)} inside mc:AlternateContent`,
        // The range that holds the child, so the message names the cells that the clear meets.
        subject: formulaRangeSubject(sheet, masterRange(master), wrapped),
        causeCell: `${sheet}!${formatCellRef(wrapped)}`,
        masterCell: `${sheet}!${formatCellRef(master)}`,
      };
    // The toolbar state reads this check on each publish, so each master is read once.
    const read = tokensOf(f);
    if (!read.ok)
      return {
        code: 'shared-formula',
        detail: `${formatCellRef(master)} holds a shared formula whose text cannot be read (${read.problem} at ${read.at})`,
        subject: formulaRangeSubject(sheet, masterRange(master), master),
        causeCell: `${sheet}!${formatCellRef(master)}`,
        causeUnreadable: true,
        masterCell: `${sheet}!${formatCellRef(master)}`,
      };
    if (si !== undefined && !bySi.has(si)) bySi.set(si, { master, tokens: read.tokens });
  }
  return bySi;
}

/** Why the shared masters `masters` cannot unshare (`readMasters`), or undefined. */
export function unshareBlocker(
  part: XmlPart,
  sheet: string,
  masters: readonly CellEntry[]
): OpRejection | undefined {
  const read = readMasters(part, sheet, masters);
  return 'code' in read ? read : undefined;
}

/**
 * The worksheet part with every child of the shared formulas whose masters are `masters`
 * written as its own formula, and the ops that restore the rows it changed. The master cells do
 * not change: the edit that follows replaces them. One pass over the cells.
 */
export function unshareChildren(
  part: XmlPart,
  sheet: string,
  masters: readonly CellEntry[]
): { readonly part: XmlPart; readonly inverse: readonly HistoryOp[] } | OpRejection {
  if (masters.length === 0) return { part, inverse: [] };
  const read = readMasters(part, sheet, masters);
  if ('code' in read) return read;
  const bySi = read;
  const ids = new IdSource(part);
  // New cells by their row element, in document order.
  const rows = new Map<Element, { row: number; cells: Map<Element, Element> }>();
  const index = sheetIndex(part);
  for (const key of index.order) {
    const entry = index.cells.get(key)!;
    const cf = child(entry.el, 'x:f');
    if (!cf || attr(cf, 't') !== 'shared' || attr(cf, 'ref') !== undefined) continue;
    const si = attr(cf, 'si');
    const shared = si === undefined ? undefined : bySi.get(si);
    if (!shared || shared.master === entry) continue;
    const { master, tokens } = shared;
    const moved = formatTokens(tokens, {
      rows: entry.row - master.row,
      cols: entry.col - master.col,
    });
    if (!moved.ok)
      return {
        code: 'shared-formula',
        detail: `${formatCellRef(entry)}: the shared formula of ${formatCellRef(master)} moves out of the grid`,
        subject: formulaRangeSubject(sheet, masterRange(master), entry),
        causeCell: `${sheet}!${formatCellRef(entry)}`,
        masterCell: `${sheet}!${formatCellRef(master)}`,
      };
    const keep = cf.attrs.filter(
      (a) => !(a.ns === '' && (a.local === 't' || a.local === 'si' || a.local === 'ref'))
    );
    const own = smlElement(ids, 'f', 'cell', keep, [textNode(ids, moved.text)]);
    const cell = withChild(entry.el, childIndex(entry.el, cf), {
      ...withNamespacesOf(cf, own),
      id: cf.id,
    });
    let group = rows.get(entry.rowEl);
    if (!group) rows.set(entry.rowEl, (group = { row: entry.row, cells: new Map() }));
    group.cells.set(entry.el, cell);
  }
  const sheetData = child(part.root, 'x:sheetData')!;
  const children: Node[] = sheetData.children.slice();
  const inverse: HistoryOp[] = [];
  for (const [rowEl, group] of rows) {
    const at = childIndex(sheetData, rowEl);
    children[at] = {
      ...rowEl,
      children: rowEl.children.map((n) => (n.type === 'element' ? (group.cells.get(n) ?? n) : n)),
    };
    inverse.push({ op: 'restoreRow', sheet, row: group.row + 1, node: rowEl });
  }
  const root = withChild(part.root, childIndex(part.root, sheetData), { ...sheetData, children });
  return { part: { partName: part.partName, root, nextId: ids.next }, inverse };
}
