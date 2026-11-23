// What an internal paste depends on (design D2 of the change `workbook-structure-editing`): the
// snapshot of a copy is used only while everything it depends on is unchanged. A copy records
// it, and a paste checks it:
//
// - every sheet that a copied formula prefix names, read by the tokenizer, and the source
//   sheet of a cut, or of a copied formula with a name that the source sheet scopes: the
//   `x:sheet` element that the name finds, by identity (none when no sheet has the name). The
//   tree is immutable, so a rename, a delete, or a new sheet with that name changes it, also a
//   delete and a new sheet with the same name, which can reuse the part name and `sheetId`. A
//   move or an edit of another sheet keeps the element, and an undo restores the same element.
//   A sheet that no copied formula names does not count. A prefix with a workbook (`[1]Jan!`)
//   names a sheet of another workbook, so it does not count; `[0]Jan!` names this workbook,
//   so it counts. A copy (not a cut) depends on its source sheet for a scoped name only when
//   it pastes on another sheet: on the source sheet itself, the name binds as it did. The
//   source sheet is the one whose `x:sheet` element has the node id that it had at the copy,
//   so a rename keeps it, and a delete and a new sheet with the same name do not.
// - the `x:xf` element of every copied cell format index, with the font, fill, and border
//   elements and the number format code that it uses. The tree is immutable, so an element
//   that did not change is the same object; an element that a later edit rebuilt with the same
//   content compares equal by structure.
//
// A copy with formulas also ends at an edit that can change what their text refers to: an insert
// or a delete of rows or columns on any sheet, a rename, a move, or a delete of a sheet that a formula names,
// and an undo or a redo of such an edit (`WorkbookDocumentStore.referenceEdits`). A formula with a
// reference across a span of sheets (`Jan:Mar!A1`) also ends at an add, a rename, a move, or a
// delete of any sheet, because the span can then hold other sheets that the formula does not name.
//
// The store identity, that rule, and the revision rule of a cut stay in `clipboard.ts`.

import {
  child,
  children,
  intAttr,
  sheetByName,
  sheetNameKey,
  styleLists,
  tokenizeFormula,
  xmlPartOf,
  type Element,
  type Node,
  type WorkbookIndex,
  type XlsxPackage,
} from '../store/index.ts';

/** A sheet that the copy depends on: its name, and the `x:sheet` element that the name found. */
interface SheetDep {
  readonly name: string;
  readonly element: Element | null;
}

/** A cell format index that the copy depends on, and what it reached at the copy. */
interface FormatDep {
  readonly index: number;
  readonly elements: readonly (Element | undefined)[];
  readonly numFmt: string | undefined;
}

/** What a copy depends on. */
export interface CopyDeps {
  readonly sheets: readonly SheetDep[];
  readonly formats: readonly FormatDep[];
  /** Whether a copied formula has a reference across a span of sheets of this workbook. */
  readonly span: boolean;
  /** The source sheet of a copy with a scoped name: a dependency for another sheet only. */
  readonly scoped?: SheetDep;
}

/** The `x:sheet` element that a sheet name finds, or null. */
export function sheetElementOf(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  name: string
): Element | null {
  const entry = sheetByName(index, name);
  if (!entry) return null;
  const sheets = child(xmlPartOf(pkg, index.mainPart)?.root, 'x:sheets');
  return children(sheets, 'x:sheet')[entry.index] ?? null;
}

/** The elements and number format code that a cell format index reaches. */
function formatState(pkg: XlsxPackage, index: WorkbookIndex, xf: number): Omit<FormatDep, 'index'> {
  const lists = styleLists(pkg, index);
  const el = lists.cellXfs[xf];
  const at = (list: readonly Element[], key: string) => {
    const i = intAttr(el, key);
    return i === undefined ? undefined : list[i];
  };
  const numFmtId = intAttr(el, 'numFmtId');
  return {
    elements: [
      el,
      at(lists.fonts, 'fontId'),
      at(lists.fills, 'fillId'),
      at(lists.borders, 'borderId'),
    ],
    numFmt: numFmtId === undefined ? undefined : lists.numFmts.get(numFmtId),
  };
}

/**
 * What a copy depends on: the sheets that `formulas` (formula texts without `=`) name, the
 * source sheet `sourceSheet` of a cut, the source sheet `scopedSheet` of a copy whose formulas
 * use a name that the source sheet scopes, and the cell format indexes `formats`.
 */
export function copyDeps(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  formulas: readonly string[],
  formats: ReadonlySet<number>,
  sourceSheet?: string,
  scopedSheet?: string
): CopyDeps {
  const names = new Map<string, string>();
  const add = (name: string) => {
    const key = sheetNameKey(name);
    if (!names.has(key)) names.set(key, name);
  };
  if (sourceSheet !== undefined) add(sourceSheet);
  let span = false;
  for (const text of formulas) {
    const read = tokenizeFormula(text);
    if (!read.ok) continue;
    for (const token of read.tokens) {
      const kind = token.kind;
      const named = kind === 'reference' || kind === 'name' || kind === 'error';
      const prefix = named ? token.sheet : undefined;
      // `[0]` names this workbook, so its sheets count as a local prefix's do.
      const local = prefix && (prefix.book === undefined || prefix.book === '0');
      if (local) for (const name of prefix.sheets) add(name);
      if (local && prefix.sheets.length > 1) span = true;
    }
  }
  return {
    sheets: [...names.values()].map((name) => ({
      name,
      element: sheetElementOf(pkg, index, name),
    })),
    formats: [...formats].map((xf) => ({ index: xf, ...formatState(pkg, index, xf) })),
    span,
    ...(scopedSheet !== undefined
      ? { scoped: { name: scopedSheet, element: sheetElementOf(pkg, index, scopedSheet) } }
      : {}),
  };
}

/** Whether two nodes have the same content. */
export function sameNode(a: Node | undefined, b: Node | undefined): boolean {
  if (a === b) return true;
  if (!a || !b || a.type !== b.type) return false;
  if (a.type === 'text') return b.type === 'text' && a.value === b.value;
  if (b.type !== 'element') return false;
  if (a.ns !== b.ns || a.local !== b.local) return false;
  if (a.attrs.length !== b.attrs.length || a.children.length !== b.children.length) return false;
  for (let i = 0; i < a.attrs.length; i += 1) {
    const x = a.attrs[i]!;
    const y = b.attrs[i]!;
    if (x.ns !== y.ns || x.local !== y.local || x.value !== y.value) return false;
  }
  return a.children.every((child, i) => sameNode(child, b.children[i]));
}

/**
 * Whether everything that a copy depends on is unchanged in the workbook, for a paste on the
 * sheet `target`: the source sheet of a scoped name counts only when `target` is another sheet.
 */
export function depsHold(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  deps: CopyDeps,
  target: string
): boolean {
  for (const dep of deps.sheets)
    if (sheetElementOf(pkg, index, dep.name) !== dep.element) return false;
  const scoped = deps.scoped;
  const sameSheet = scoped?.element && sheetElementOf(pkg, index, target)?.id === scoped.element.id;
  if (scoped && !sameSheet && sheetElementOf(pkg, index, scoped.name) !== scoped.element)
    return false;
  for (const dep of deps.formats) {
    const now = formatState(pkg, index, dep.index);
    if (now.numFmt !== dep.numFmt) return false;
    if (!now.elements.every((el, i) => sameNode(el, dep.elements[i]))) return false;
  }
  return true;
}

/**
 * Whether a formula of `formulas` uses a name with no sheet prefix that `locals` holds (the
 * names that the source sheet scopes, as `sheetNameKey` folds them). A function name counts,
 * because a defined name can hold a function (`myFn(1)`).
 */
export function usesLocalName(formulas: readonly string[], locals: ReadonlySet<string>): boolean {
  if (locals.size === 0) return false;
  for (const text of formulas) {
    const read = tokenizeFormula(text);
    if (!read.ok) continue;
    for (const token of read.tokens)
      if (
        ((token.kind === 'name' && !token.sheet) || token.kind === 'function') &&
        locals.has(sheetNameKey(token.text))
      )
        return true;
  }
  return false;
}
