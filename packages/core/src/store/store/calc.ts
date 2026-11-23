// The calculation side effects of an edit: the stale flag of design D3 and the calculation
// chain of design D5. Both are part of the transaction that makes them necessary, and neither
// enters the undo history:
//
// - D3: a transaction that changes a value in a workbook that holds a formula cell after the
//   transaction sets `x:calcPr/@fullCalcOnLoad="1"` (`CT_CalcPr`, sml.xsd; default false), so a
//   consumer recalculates. Undo does not clear it, because the engine cannot prove that the cached
//   values are current again.
// - D5: a transaction that adds, removes, or changes a formula drops `/xl/calcChain.xml`, its
//   workbook relationship, and its content type override, so no consumer reads a stale chain.
//   The chain is optional (ECMA-376 Part 1, 18.6). Because the drop is outside the undo
//   history, an undo can put back a workbook relationship part or a content type part that
//   still names the dropped chain. Every transaction therefore also drops a chain
//   relationship or override whose part is gone (`withoutDanglingCalcChain`).

import type { Element, Node, XmlPart } from '../package/tree.ts';
import {
  CONTENT_TYPES_PART,
  contentTypeOf,
  partOf,
  relationshipsOf,
  relsPartName,
  xmlPartOf,
  type XlsxPackage,
} from '../package/package.ts';
import { asciiFold, normalizePartName } from '@eigenpal/ooxml-core';
import { NS, REL } from '../package/namespaces.ts';
import { attr, boolAttr, child } from '../package/query.ts';
import { workbookIndex } from '../workbook/workbook-index.ts';
import { cellHasFormula } from '../workbook/cell-formula.ts';
import { IdSource, plainAttr, smlElement, withAttr, withChild } from './tree-edit.ts';

/**
 * Whether a typed cell of a row holds a formula, also in a cell nested in it
 * ({@link cellOrNestedHasFormula}). It reads only the typed cells of the row: an edit cannot
 * change a row-level `mc:AlternateContent` or a generic `x:c`.
 */
export function rowHasFormula(row: Element | null | undefined): boolean {
  if (!row) return false;
  for (const n of row.children)
    if (n.type === 'element' && n.kind === 'cell' && cellOrNestedHasFormula(n)) return true;
  return false;
}

/** Whether two nodes have the same name, attributes in any order, and the same content. */
export function sameNode(a: Node, b: Node): boolean {
  if (a === b) return true;
  if (a.type !== b.type) return false;
  if (a.type === 'text' || b.type === 'text')
    return a.type === 'text' && b.type === 'text' && a.value === b.value;
  if (a.type !== 'element' || b.type !== 'element') return false;
  if (a.ns !== b.ns || a.local !== b.local || a.attrs.length !== b.attrs.length) return false;
  for (const x of a.attrs)
    if (!b.attrs.some((y) => y.ns === x.ns && y.local === x.local && y.value === x.value))
      return false;
  if (a.children.length !== b.children.length) return false;
  for (let i = 0; i < a.children.length; i += 1)
    if (!sameNode(a.children[i]!, b.children[i]!)) return false;
  return true;
}

/**
 * The cells of a row that hold content other than a cell format, in order: a cell with only
 * `@r` and `@s` holds no value and no formula, so it is left out.
 */
function contentCells(row: Element | null | undefined): Element[] {
  const out: Element[] = [];
  for (const n of row?.children ?? []) {
    if (n.type !== 'element' || n.kind !== 'cell') continue;
    if (
      n.children.length > 0 ||
      n.attrs.some((a) => a.ns !== '' || (a.local !== 'r' && a.local !== 's'))
    )
      out.push(n);
  }
  return out;
}

/** Whether two cells are equal when their `@s` is ignored. */
function sameContent(a: Element, b: Element): boolean {
  const strip = (e: Element): Element => ({
    ...e,
    attrs: e.attrs.filter((x) => !(x.ns === '' && x.local === 's')),
  });
  return sameNode(strip(a), strip(b));
}

/**
 * What replacing row `before` with row `after` changes for design D3 and D5: `value` when a
 * cell value or formula differs, and `formula` when that change can touch a formula. A change
 * of cell formats, row properties, or empty formatted cells changes neither, so the undo of a
 * format edit neither marks the workbook stale nor drops the calculation chain.
 */
export function rowChange(
  before: Element | null | undefined,
  after: Element | null | undefined
): { readonly value: boolean; readonly formula: boolean } {
  const a = contentCells(before);
  const b = contentCells(after);
  let same = a.length === b.length;
  for (let i = 0; same && i < a.length; i += 1) same = sameContent(a[i]!, b[i]!);
  if (same) return { value: false, formula: false };
  return { value: true, formula: rowHasFormula(before) || rowHasFormula(after) };
}

const formulaParts = new WeakMap<XmlPart, boolean>();

/**
 * Whether a parsed sheet part holds a formula cell: an `x:f` inside an `x:c` (`CT_Cell`, Part 1,
 * 18.3.1.4) in its tree, whatever its root element. `sheetMayHoldFormula` counts a root in
 * another namespace before it calls this scan. The stale flag reads it for each
 * parsed sheet part, and counts a sheet part that the reader keeps as bytes as one with formulas.
 *
 * The scan covers the cells that the sheet index skips: rows and cells inside
 * `mc:AlternateContent` (every branch), an `x:f` inside `mc:AlternateContent` in a cell, and an
 * `x:c` that the typed tree keeps generic. A cell counts by the rule of {@link cellHasFormula}: an
 * `x:f` child, or an `x:f` child of a branch of an `mc:AlternateContent` child. A cell out of
 * place inside another cell counts too. Rule and validation formulas (`x:formula`, `x:formula1`,
 * `xm:f`) are not in a cell and have no cached value. It reads the nodes that exist, a fixed
 * number of times each, with an explicit stack, so the parser's element count cap bounds it.
 * Cached per part revision.
 */
export function partHasFormula(part: XmlPart): boolean {
  const hit = formulaParts.get(part);
  if (hit !== undefined) return hit;
  const found = treeHasFormula(part.root, false);
  formulaParts.set(part, found);
  return found;
}

/**
 * Whether a cell holds a formula, or a cell out of place inside it holds one: the rule of
 * {@link partHasFormula} for one cell. The calculation chain drop reads it to err to the safe
 * side: an edit or a clear keeps a cell out of place as other content, so the chain can drop for
 * a formula that stays.
 */
export function cellOrNestedHasFormula(cell: Element): boolean {
  // A cell with no child that can hold another cell needs no stack: the cell rule answers.
  if (!cell.children.some(holdsCell)) return cellHasFormula(cell);
  return treeHasFormula(cell, true);
}

/** Whether a child of a cell can hold another cell: a nested `x:c`, or an `mc` element. */
const holdsCell = (n: Node): boolean =>
  n.type === 'element' && (n.ns === NS.mc || (n.ns === NS.x && n.local === 'c'));

/**
 * The one pass of the scan. Each `x:c` counts by {@link cellHasFormula}. Inside a cell, also in
 * its `mc:AlternateContent` children, it goes down only into a nested `x:c` and an `mc` element,
 * which can hold an `x:c`. It pushes no element that has no child to go down into, so a value
 * cell costs no stack entry. It uses an explicit stack.
 */
function treeHasFormula(root: Element, inCell: boolean): boolean {
  if (inCell && cellHasFormula(root)) return true;
  // Two stacks: elements outside cells, and elements inside a cell, which go down only into a
  // nested `x:c` and `mc` elements. No entry is allocated for each element.
  const outside: Element[] = inCell ? [] : [root];
  const inside: Element[] = inCell ? [root] : [];
  while (outside.length > 0 || inside.length > 0) {
    const cell = inside.length > 0;
    const el = cell ? inside.pop()! : outside.pop()!;
    for (const n of el.children) {
      if (n.type !== 'element') continue;
      if (n.ns === NS.x && n.local === 'c') {
        const read = readCell(n);
        if (read === FORMULA) return true;
        if (read === DEEPER) inside.push(n);
      } else if (!cell) {
        if (n.children.length > 0) outside.push(n);
      } else if (n.ns === NS.mc && n.children.some(holdsCell)) inside.push(n);
    }
  }
  return false;
}

const NONE = 0;
const FORMULA = 1;
const DEEPER = 2;

/**
 * One pass over the children of a cell: {@link FORMULA} for a direct `x:f`, or for a wrapped
 * one by {@link cellHasFormula}; {@link DEEPER} when a child can hold another cell; else
 * {@link NONE}. The wrapper rule runs only for a cell with an `mc:AlternateContent` child, so a
 * value cell costs one read of its children. Its formula answer must stay equal to
 * {@link cellHasFormula}, which a test table checks.
 */
function readCell(c: Element): typeof NONE | typeof FORMULA | typeof DEEPER {
  let wrapper = false;
  let deeper = false;
  for (const k of c.children) {
    if (k.type !== 'element') continue;
    if (k.ns === NS.x) {
      if (k.local === 'f') return FORMULA;
      if (k.local === 'c') deeper = true;
    } else if (k.ns === NS.mc) {
      deeper = true;
      if (k.local === 'AlternateContent') wrapper = true;
    }
  }
  if (wrapper && cellHasFormula(c)) return FORMULA;
  return deeper ? DEEPER : NONE;
}

/**
 * The content types of the sheet parts whose root holds no cells: a chartsheet (Chartsheet Part,
 * Part 1, 12.3.2) and a dialog sheet (Dialogsheet Part, Part 1, 12.3.7).
 */
const CELL_FREE_SHEETS = new Set(
  [
    'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml',
  ].map(asciiFold)
);

/**
 * Whether a sheet part can hold a formula cell. A parsed part is scanned. The reader keeps a
 * sheet part that the engine does not model (a macro sheet) as unread bytes, so the engine
 * cannot see its cells: such a part counts as one with a formula, unless its content type holds
 * no cells or is not XML. A chartsheet or a dialog sheet never counts, parsed or not. A stale
 * flag that a consumer did not need costs one recalculation; a missing one keeps old cached
 * values.
 */
function sheetPartMayHoldFormula(pkg: XlsxPackage, partName: string | undefined): boolean {
  const part = partOf(pkg, partName);
  if (!part || !partName) return false;
  const type = contentTypeOf(pkg, partName);
  // A content type can carry parameters after `;`: the media type alone decides.
  const folded = type === undefined ? undefined : asciiFold(type.split(';')[0]!.trim());
  // A part that is not XML (an image that a hostile `x:sheet` names) holds no cells.
  if (folded !== undefined && (CELL_FREE_SHEETS.has(folded) || !folded.endsWith('xml')))
    return false;
  return part.kind !== 'xml' || sheetMayHoldFormula(part.xml);
}

/**
 * Whether a parsed sheet part can hold a formula cell: {@link partHasFormula}, or a root in
 * another namespace than the Transitional SpreadsheetML namespace (a Strict root), whose cells
 * the scan does not read. The stale flag and the stale notice both read it.
 */
export function sheetMayHoldFormula(part: XmlPart): boolean {
  return part.root.ns !== NS.x || partHasFormula(part);
}

/**
 * Whether any sheet part that `x:sheets` names holds a formula cell (design D3). Every sheet
 * counts, not only worksheets: a macro sheet holds formula cells too.
 */
export function workbookHasFormula(pkg: XlsxPackage): boolean {
  const index = workbookIndex(pkg);
  if (!index) return false;
  for (const sheet of index.sheets) if (sheetPartMayHoldFormula(pkg, sheet.partName)) return true;
  return false;
}

/** The children of `x:workbook` that come before `x:calcPr` (`CT_Workbook`, Part 1, 18.2.27). */
const BEFORE_CALC_PR = new Set([
  'fileVersion',
  'fileSharing',
  'workbookPr',
  'workbookProtection',
  'bookViews',
  'sheets',
  'functionGroups',
  'externalReferences',
  'definedNames',
]);

/**
 * The package with `x:calcPr/@fullCalcOnLoad="1"`, or undefined when it is already set or the
 * package has no workbook part. A new `x:calcPr` goes in its schema position.
 */
export function withFullCalcOnLoad(pkg: XlsxPackage): XlsxPackage | undefined {
  const index = workbookIndex(pkg);
  const part = xmlPartOf(pkg, index?.mainPart);
  if (!index || !part) return undefined;
  const calcPr = child(part.root, 'x:calcPr');
  if (boolAttr(calcPr, 'fullCalcOnLoad') === true) return undefined;
  const ids = new IdSource(part);
  let root: Element;
  if (calcPr) {
    const at = part.root.children.indexOf(calcPr);
    root = withChild(part.root, at, withAttr(calcPr, 'fullCalcOnLoad', '1'));
  } else {
    let at = 0;
    part.root.children.forEach((n, i) => {
      if (n.type === 'element' && n.ns === NS.x && BEFORE_CALC_PR.has(n.local)) at = i + 1;
    });
    const el = smlElement(ids, 'calcPr', 'workbook', [plainAttr('fullCalcOnLoad', '1')]);
    root = withChild(part.root, at, el, true);
  }
  const parts = new Map(pkg.parts);
  const xml: XmlPart = { partName: part.partName, root, nextId: ids.next };
  parts.set(part.partName, { kind: 'xml', partName: part.partName, xml });
  return { parts };
}

/**
 * The package without the calculation chain: the part, the workbook relationship to it, and its
 * content type override. Undefined when the package has none of them.
 */
export function withoutCalcChain(pkg: XlsxPackage): XlsxPackage | undefined {
  const index = workbookIndex(pkg);
  if (!index) return undefined;
  const rels = relationshipsOf(pkg, index.mainPart).filter(
    (r) => r.type === REL.calcChain && !r.external
  );
  const targets = new Set(rels.flatMap((r) => (r.target ? [asciiFold(r.target)] : [])));
  targets.add(asciiFold('/xl/calcChain.xml'));
  const parts = new Map(pkg.parts);
  let changed = false;
  for (const name of pkg.parts.keys())
    if (targets.has(asciiFold(name))) {
      parts.delete(name);
      changed = true;
    }
  const relsName = relsPartName(index.mainPart);
  const relsPart = xmlPartOf(pkg, relsName);
  if (relsPart && rels.length > 0) {
    const ids = new Set(rels.map((r) => r.id));
    const root = {
      ...relsPart.root,
      children: relsPart.root.children.filter(
        (n) =>
          !(
            n.type === 'element' &&
            n.ns === NS.rel &&
            n.local === 'Relationship' &&
            ids.has(attr(n, 'Id') ?? '') &&
            attr(n, 'Type') === REL.calcChain
          )
      ),
    };
    parts.set(relsPart.partName, {
      kind: 'xml',
      partName: relsPart.partName,
      xml: { ...relsPart, root },
    });
    changed = true;
  }
  const ct = xmlPartOf(pkg, CONTENT_TYPES_PART);
  if (ct) {
    const kept = ct.root.children.filter((n) => {
      if (n.type !== 'element' || n.ns !== NS.ct || n.local !== 'Override') return true;
      const norm = normalizePartName(attr(n, 'PartName') ?? '');
      return !(norm.ok && targets.has(asciiFold(norm.partName)));
    });
    if (kept.length !== ct.root.children.length) {
      parts.set(ct.partName, {
        kind: 'xml',
        partName: ct.partName,
        xml: { ...ct, root: { ...ct.root, children: kept } },
      });
      changed = true;
    }
  }
  return changed ? { parts } : undefined;
}

/**
 * The package without a calculation chain relationship or override whose part is gone, or
 * undefined when the chain part is present or nothing names a missing chain. An undo of an edit
 * that dropped the chain can put such an entry back, because the drop is not in the history.
 */
export function withoutDanglingCalcChain(pkg: XlsxPackage): XlsxPackage | undefined {
  const index = workbookIndex(pkg);
  if (!index) return undefined;
  if (partOf(pkg, '/xl/calcChain.xml')) return undefined;
  for (const r of relationshipsOf(pkg, index.mainPart))
    if (r.type === REL.calcChain && !r.external && r.target && partOf(pkg, r.target))
      return undefined;
  return withoutCalcChain(pkg);
}
