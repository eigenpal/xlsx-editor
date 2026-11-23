// The structural rules of the typed model, checked on the parts a transaction touched.
//
// Same role as pptx-editor/packages/core/src/store/store/part-checks.ts at 1943d2c9 (EigenPal):
// a transaction that leaves a part broken is refused like a refused op. The checks cost time in
// proportion to the touched content: for a worksheet, the cells of the rows whose element
// changed are read, and the row order is read once, as the edit itself copied the row list. A
// row or a cell without `@r` takes the number after the one before it. A rule that failed
// before the transaction in an untouched row does not refuse it, and neither does a bad format
// or shared string index that the part already held (`heldBadIndexes`).
//
// Every text and attribute value of a changed row, of an appended shared string, and of a new style
// record must be text that XML 1.0 can hold (`xmlInvalidAt`), so no committed transaction can make
// a save throw. The ops write such text escaped (`encodeXstring`) or refuse it; this check is the
// guard behind them. For the same reason, an attribute value of those nodes, and of a changed
// child of the worksheet root, may name only prefixes that are declared in scope
// (`prefixValueError`), which the serializer requires. The ops keep the declarations of each
// element that they build again (`withNamespacesOf`, `withParentDeclarations`). Like
// `holdsInvalidXml`, the prefix check reads 64 levels below each node; no op builds deeper content.
//
// A transaction that touches the workbook relationship part or the content type part (both of
// ECMA-376 Part 2) must not leave a new workbook relationship or override that names a missing
// part: the saved file would need a repair. An entry that named a missing part before the
// transaction does not refuse it. Nor may it add a second content type override for one part
// name.

import type { Element, Node } from '../package/tree.ts';

import {
  CONTENT_TYPES_PART,
  equivalentPartName,
  relationshipsOf,
  relsPartName,
  xmlPartOf,
  type XlsxPackage,
} from '../package/package.ts';
import { asciiFold, normalizePartName } from '@eigenpal/ooxml-core';
import { xmlInvalidAt } from '../workbook/xstring.ts';
import { attr, child, children } from '../package/query.ts';
import { NS } from '../package/namespaces.ts';
import { prefixValueError } from '../package/xml-names.ts';
import { parseCellRef } from '../cell-ref.ts';
import { sharedStringItems, styleLists, workbookIndex } from '../workbook/workbook-index.ts';

/**
 * Rows of `after` whose element is not in `before`, by identity, or a message when a new child
 * of `x:sheetData` is not an `x:row` element.
 */
function changedRows(before: Element | undefined, after: Element): Element[] | string {
  const old = new Set(before?.children ?? []);
  const out: Element[] = [];
  for (const node of after.children) {
    if (old.has(node)) continue;
    if (node.type !== 'element' || node.kind !== 'row' || node.ns !== NS.x || node.local !== 'row')
      return 'a new child of sheetData is not a row element';
    out.push(node);
  }
  return out;
}

/** Whether a node holds a text or an attribute value that XML 1.0 cannot hold. */
function holdsInvalidXml(node: Node, depth = 0): boolean {
  if (node.type === 'text') return xmlInvalidAt(node.value) >= 0;
  if (node.type !== 'element' || depth > 64) return false;
  for (const a of node.attrs) if (xmlInvalidAt(a.value) >= 0) return true;
  for (const n of node.children) if (holdsInvalidXml(n, depth + 1)) return true;
  return false;
}

/** The declarations in scope inside `el`, from the declarations in scope at its parent. */
function scopeIn(parent: ReadonlyMap<string, string>, el: Element): ReadonlyMap<string, string> {
  if (el.nsDecls.length === 0) return parent;
  const out = new Map(parent);
  for (const d of el.nsDecls) out.set(d.prefix, d.uri);
  return out;
}

/** The name of an element or an attribute as the file spells it, as the reader reports it. */
const qualified = (n: { readonly prefix?: string; readonly local: string }): string =>
  n.prefix ? `${n.prefix}:${n.local}` : n.local;

/** The declarations in scope inside the root of a part, as the reader starts them. */
const rootScope = (root: Element) => scopeIn(new Map([['xml', NS.xml]]), root);

/**
 * The first attribute value in `node` that names a prefix that is not declared in scope
 * (`prefixValueError`: `mc:Ignorable` and the other prefix and name lists), as text, or
 * undefined. `scope` holds the declarations in scope at the parent of `node`.
 */
function undeclaredPrefix(
  node: Node,
  scope: ReadonlyMap<string, string>,
  depth = 0
): string | undefined {
  if (node.type !== 'element' || depth > 64) return undefined;
  const inner = scopeIn(scope, node);
  for (const a of node.attrs) {
    const error = prefixValueError(a, node.ns, node.local, inner);
    if (error) return `${qualified(node)}/@${qualified(a)}: ${error.detail}`;
  }
  for (const n of node.children) {
    const found = undeclaredPrefix(n, inner, depth + 1);
    if (found) return found;
  }
  return undefined;
}

/**
 * The first new node of a styles part, or a changed list's own attributes, that names an
 * undeclared prefix, as text, or undefined. The same nodes as `stylesHoldInvalidXml`.
 */
function stylesUndeclaredPrefix(before: Element | undefined, after: Element): string | undefined {
  const old = new Set<Node>();
  for (const list of before?.children ?? []) {
    old.add(list);
    if (list.type === 'element') for (const n of list.children) old.add(n);
  }
  const scope = rootScope(after);
  for (const list of after.children) {
    if (old.has(list) || list.type !== 'element') continue;
    const inner = scopeIn(scope, list);
    for (const a of list.attrs) {
      const error = prefixValueError(a, list.ns, list.local, inner);
      if (error) return `${qualified(list)}/@${qualified(a)}: ${error.detail}`;
    }
    for (const n of list.children) {
      const found = old.has(n) ? undefined : undeclaredPrefix(n, inner);
      if (found) return found;
    }
  }
  return undefined;
}

/**
 * Whether a styles part holds a record that `before` does not hold, with a value that XML 1.0
 * cannot hold. Format edits append records and keep the old nodes, so only the new nodes of the
 * part's lists are read.
 */
function stylesHoldInvalidXml(before: Element | undefined, after: Element): boolean {
  const old = new Set<Node>();
  for (const list of before?.children ?? []) {
    old.add(list);
    if (list.type === 'element') for (const n of list.children) old.add(n);
  }
  for (const list of after.children) {
    if (old.has(list)) continue;
    if (list.type !== 'element') {
      if (holdsInvalidXml(list)) return true;
      continue;
    }
    for (const a of list.attrs) if (xmlInvalidAt(a.value) >= 0) return true;
    for (const n of list.children) if (!old.has(n) && holdsInvalidXml(n)) return true;
  }
  return false;
}

/**
 * The format and shared string indexes that the cells of `sheetData` hold and that are past
 * their lists, as `s:<index>` and `t:<index>` keys. A transaction may keep such an index that
 * the file already held (unknown content never locks editing), but never add one. One pass
 * over the cells, read only when a changed row holds a bad index.
 */
function heldBadIndexes(
  sheetData: Element | undefined,
  cellXfs: number,
  strings: number
): Set<string> {
  const out = new Set<string>();
  for (const row of sheetData?.children ?? []) {
    if (row.type !== 'element' || row.kind !== 'row') continue;
    for (const cell of row.children) {
      if (cell.type !== 'element' || cell.kind !== 'cell') continue;
      const s = badFormat(cell, cellXfs);
      if (s) out.add(s);
      const t = badString(cell, strings);
      if (t) out.add(t);
    }
  }
  return out;
}

/** `s:<index>` when the cell's format index is past the cell formats, else undefined. */
function badFormat(cell: Element, cellXfs: number): string | undefined {
  const s = attr(cell, 's');
  return s !== undefined && !(Number(s) >= 0 && Number(s) < Math.max(cellXfs, 1))
    ? `s:${s}`
    : undefined;
}

/** `t:<index>` when the cell's shared string index is past the table, else undefined. */
function badString(cell: Element, strings: number): string | undefined {
  if (attr(cell, 't') !== 's') return undefined;
  const text = child(cell, 'x:v')?.children[0];
  const raw = text?.type === 'text' ? text.value : '';
  const index = Number(raw);
  return Number.isInteger(index) && index >= 0 && index < strings ? undefined : `t:${raw}`;
}

/**
 * The first broken rule of one row, or undefined. A bad index that the part held before the
 * transaction (`held`) is not the transaction's own, so it does not refuse it.
 */
function rowViolation(
  row: Element,
  rowNumber: number,
  cellXfs: number,
  strings: number,
  held: () => { has(key: string): boolean }
): string | undefined {
  const r = String(rowNumber);
  if (holdsInvalidXml(row)) return `row ${r}: a value holds a character that XML cannot hold`;
  let lastCol = -1;
  for (const node of row.children) {
    if (node.type !== 'element' || node.kind !== 'cell') continue;
    const ref = attr(node, 'r');
    // A cell without `@r` follows the one before it (quick reference, "Cells and values"); an
    // undo can put such a cell back from an earlier revision.
    const address = ref === undefined ? { row: -1, col: lastCol + 1 } : parseCellRef(ref);
    if (!address) return `row ${r}: a cell has no reference in the grid`;
    if (ref !== undefined && rowNumber > 0 && address.row + 1 !== rowNumber)
      return `row ${r}: cell ${ref} is in the wrong row`;
    if (address.col <= lastCol) return `row ${r}: cell ${ref} is out of order`;
    lastCol = address.col;
    const s = badFormat(node, cellXfs);
    if (s && !held().has(s))
      return `cell ${ref}: format index ${s.slice(2)} is past the cell formats`;
    const t = badString(node, strings);
    if (t && !held().has(t)) return `cell ${ref}: shared string index is past the table`;
  }
  return undefined;
}

/**
 * The missing parts that an internal workbook relationship or a content type override names,
 * as folded part names.
 */
function danglingTargets(pkg: XlsxPackage, mainPart: string): Map<string, string> {
  const out = new Map<string, string>();
  for (const r of relationshipsOf(pkg, mainPart))
    if (!r.external && r.target && !equivalentPartName(pkg, r.target))
      out.set(asciiFold(r.target), r.target);
  for (const n of xmlPartOf(pkg, CONTENT_TYPES_PART)?.root.children ?? []) {
    if (n.type !== 'element' || n.ns !== NS.ct || n.local !== 'Override') continue;
    const norm = normalizePartName(attr(n, 'PartName') ?? '');
    if (norm.ok && !equivalentPartName(pkg, norm.partName))
      out.set(asciiFold(norm.partName), norm.partName);
  }
  return out;
}

/** The number of content type overrides of each part, by folded part name. */
function overrideCounts(pkg: XlsxPackage): Map<string, number> {
  const out = new Map<string, number>();
  for (const n of xmlPartOf(pkg, CONTENT_TYPES_PART)?.root.children ?? []) {
    if (n.type !== 'element' || n.ns !== NS.ct || n.local !== 'Override') continue;
    const norm = normalizePartName(attr(n, 'PartName') ?? '');
    if (!norm.ok) continue;
    const key = asciiFold(norm.partName);
    out.set(key, (out.get(key) ?? 0) + 1);
  }
  return out;
}

/** Every bad index counts as held. */
const ALL_HELD = { has: (_key: string) => true };

/**
 * The first rule a transaction broke in the parts it touched, or undefined. An undo or a redo
 * (`history`) puts back nodes of an earlier revision: a bad index that such a node holds was
 * held in its own revision, so it does not refuse the step.
 */
export function touchedPartViolation(
  before: XlsxPackage,
  after: XlsxPackage,
  parts: readonly string[],
  history = false
): string | undefined {
  const index = workbookIndex(after);
  if (!index) return 'the package has no workbook part';
  const worksheets = new Set(index.sheets.map((s) => s.partName));
  const relsKey = asciiFold(relsPartName(index.mainPart));
  const ctKey = asciiFold(CONTENT_TYPES_PART);
  if (parts.some((p) => asciiFold(p) === relsKey || asciiFold(p) === ctKey)) {
    const old = workbookIndex(before);
    const was = old ? danglingTargets(before, old.mainPart) : new Map<string, string>();
    for (const [key, t] of danglingTargets(after, index.mainPart))
      if (!was.has(key))
        return `a workbook relationship or a content type override names ${t}, which is missing`;
    const had = overrideCounts(before);
    for (const [key, n] of overrideCounts(after))
      if (n > 1 && n > (had.get(key) ?? 0))
        return `two content type overrides name the part ${key}`;
  }
  let limits: { cellXfs: number; strings: number } | undefined;
  for (const name of parts) {
    if (name === index.sharedStringsPart) {
      // Items are only ever appended: check the new ones. Only `x:si` counts, so an item
      // added before `x:extLst`, or after text between items, is still checked.
      const old = children(xmlPartOf(before, name)?.root, 'x:si').length;
      const items = children(xmlPartOf(after, name)?.root, 'x:si');
      const root = xmlPartOf(after, name)?.root;
      for (let i = old; i < items.length; i += 1) {
        if (holdsInvalidXml(items[i]!))
          return `${name}: a string holds a character XML cannot hold`;
        const prefix = root && undeclaredPrefix(items[i]!, rootScope(root));
        if (prefix) return `${name}: ${prefix}`;
      }
      continue;
    }
    if (name === index.stylesPart) {
      const stylesRoot = xmlPartOf(after, name)?.root;
      const stylesBefore = xmlPartOf(before, name)?.root;
      if (stylesRoot && stylesHoldInvalidXml(stylesBefore, stylesRoot))
        return `${name}: a style record holds a character XML cannot hold`;
      const prefix = stylesRoot && stylesUndeclaredPrefix(stylesBefore, stylesRoot);
      if (prefix) return `${name}: ${prefix}`;
      continue;
    }
    if (!worksheets.has(name)) continue;
    const part = xmlPartOf(after, name);
    const sheetData = part ? child(part.root, 'x:sheetData') : undefined;
    if (!sheetData) continue;
    const old = child(xmlPartOf(before, name)?.root, 'x:sheetData');
    limits ??= {
      cellXfs: styleLists(after, index).cellXfs.length,
      strings: sharedStringItems(after, index).length,
    };
    const changed = changedRows(old, sheetData);
    if (typeof changed === 'string') return `${name}: ${changed}`;
    const changedSet = new Set(changed);
    // A row without `@r` is the row after the one before it, so it counts in the order too.
    const numberOf = new Map<Element, number>();
    let lastRow = 0;
    for (const node of sheetData.children) {
      if (node.type !== 'element' || node.kind !== 'row') continue;
      const raw = attr(node, 'r');
      const r = raw === undefined ? lastRow + 1 : Number(raw);
      if (!Number.isInteger(r)) continue;
      if (r <= lastRow && changedSet.has(node)) return `${name}: row ${r} is out of order`;
      if (changedSet.has(node)) numberOf.set(node, r);
      lastRow = r;
    }
    let held: Set<string> | undefined;
    const lim = limits;
    const heldOf = history
      ? () => ALL_HELD
      : () => (held ??= heldBadIndexes(old, lim.cellXfs, lim.strings));
    for (const row of changed) {
      const v = rowViolation(row, numberOf.get(row) ?? -1, limits.cellXfs, limits.strings, heldOf);
      if (v) return `${name}: ${v}`;
    }
    // The changed rows, and each other changed child of the worksheet root (such as the
    // `x:cols` that a width edit merges), name only prefixes that are declared in scope.
    const sheetScope = rootScope(part!.root);
    const dataScope = scopeIn(sheetScope, sheetData);
    for (const row of changed) {
      const prefix = undeclaredPrefix(row, dataScope);
      if (prefix) return `${name}: ${prefix}`;
    }
    const oldChildren = new Set(xmlPartOf(before, name)?.root.children ?? []);
    for (const node of part!.root.children) {
      if (node === sheetData || oldChildren.has(node)) continue;
      const prefix = undeclaredPrefix(node, sheetScope);
      if (prefix) return `${name}: ${prefix}`;
    }
  }
  return undefined;
}
