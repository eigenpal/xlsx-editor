// The functions whose result depends on the names or the positions of the sheets, not on a
// reference with a sheet prefix: `INDIRECT`, `CELL`, `INFO`, `SHEET`, `SHEETS`, and
// `FORMULATEXT` (stored as `_xlfn.SHEET`, `_xlfn.SHEETS`, and `_xlfn.FORMULATEXT`), and every
// macro function `GET.*`, such as `GET.WORKBOOK`, `GET.DOCUMENT`, and `GET.CELL`, which a
// defined name can call. An add, a rename, or a move of a sheet changes no reference, but it can
// change the result of a call to one of them: a rename rewrites the formula text that
// `FORMULATEXT` returns. So the sheet ops mark the cached values stale (design D1 and D3) when a
// formula of a cell or a defined name calls one. A column width changes the result of
// `CELL("width")` and of `GET.CELL`, so `setColumnWidth` marks them stale when a formula calls
// one (`CELL_STATE_FUNCTIONS`). The format ops use the same set. A row height changes the result
// of `GET.CELL`, so `setRowHeight` marks them stale when a formula calls a `GET.*` macro
// (`ROW_HEIGHT_FUNCTIONS`).
//
// A name that ends with a full stop, as `get.`, is a prefix: it matches every function whose
// name starts with it. The scan reads each formula text once. A text that does not hold the
// name of one of the functions is not tokenized. A text that the tokenizer cannot read counts
// as a call, so the scan never keeps a stale value as current.

import type { Element } from '../package/tree.ts';
import { xmlPartOf, type XlsxPackage } from '../package/package.ts';
import { NS } from '../package/namespaces.ts';
import { child, textOf } from '../package/query.ts';
import { asciiFold } from '@eigenpal/ooxml-core';
import type { WorkbookIndex } from '../workbook/workbook-index.ts';
import { tokenizeFormula } from '../formula/tokenizer.ts';
import { isAlternateContent, wrappedDefinedNames } from './wrapped-formulas.ts';

/** The prefix of the macro functions `GET.*`, folded. A name that ends with `.` is a prefix. */
const GET_MACROS = 'get.';

/** The functions that read sheet names or positions, folded, without the `_xlfn.` prefix. */
export const SHEET_FUNCTIONS: ReadonlySet<string> = new Set([
  'indirect',
  'cell',
  'info',
  'sheet',
  'sheets',
  'formulatext',
  GET_MACROS,
]);

/**
 * The functions whose result depends on the state of a cell, such as its cell format or its
 * column width: `CELL` (`CELL("format")`, `CELL("width")`) and the macro functions `GET.*`.
 */
export const CELL_STATE_FUNCTIONS: ReadonlySet<string> = new Set(['cell', GET_MACROS]);

/** The functions whose result depends on a row height: the macro functions `GET.*`. */
export const ROW_HEIGHT_FUNCTIONS: ReadonlySet<string> = new Set([GET_MACROS]);

/** Whether the folded function name `name` is in `names`, or starts with a prefix of it. */
function inNames(name: string, names: ReadonlySet<string>): boolean {
  if (names.has(name)) return true;
  for (const n of names) if (n.endsWith('.') && name.startsWith(n)) return true;
  return false;
}

/** The depth of elements that the scan reads, as the mention scan of `blockers.ts` does. */
const MAX_SCAN_DEPTH = 64;

/** Whether formula text calls a function of `names`, or cannot be read. */
export function callsSheetFunction(
  text: string,
  names: ReadonlySet<string> = SHEET_FUNCTIONS
): boolean {
  // A fast test before the tokenizer: the folded text holds a name of one of the functions.
  const folded = asciiFold(text);
  let maybe = false;
  for (const name of names) if (folded.includes(name)) maybe = true;
  if (!maybe) return false;
  const read = tokenizeFormula(text);
  if (!read.ok) return true;
  for (const t of read.tokens) {
    if (t.kind !== 'function') continue;
    const name = asciiFold(t.text).replace(/^_xlfn\./, '');
    if (inNames(name, names)) return true;
  }
  return false;
}

/** Whether an `x:f` under `el` calls one of `names`. Deeper content counts as a call. */
function formulasCall(el: Element, names: ReadonlySet<string>, depth = 0): boolean {
  if (depth > MAX_SCAN_DEPTH) return true;
  for (const c of el.children) {
    if (c.type !== 'element') continue;
    if (c.ns === NS.x && c.local === 'f') {
      if (callsSheetFunction(textOf(c), names)) return true;
    } else if (formulasCall(c, names, depth + 1)) return true;
  }
  return false;
}

/**
 * The scan result of each package and set of names. A package does not change, so a format op
 * or a resize does not scan every formula again until a commit gives a new package.
 */
const scanCache = new WeakMap<XlsxPackage, Map<ReadonlySet<string>, boolean>>();

/**
 * Whether a formula of a worksheet cell or a defined name of the workbook calls one of `names`:
 * by default, a function whose result depends on the names or the positions of the sheets. The
 * result is kept for each package and set of names. `index` is the index of `pkg`.
 */
export function workbookCallsSheetFunction(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  names: ReadonlySet<string> = SHEET_FUNCTIONS
): boolean {
  let byNames = scanCache.get(pkg);
  if (!byNames) scanCache.set(pkg, (byNames = new Map()));
  const known = byNames.get(names);
  if (known !== undefined) return known;
  const calls = scanCalls(pkg, index, names);
  byNames.set(names, calls);
  return calls;
}

function scanCalls(pkg: XlsxPackage, index: WorkbookIndex, names: ReadonlySet<string>): boolean {
  for (const sheet of index.sheets) {
    const data = child(xmlPartOf(pkg, sheet.partName)?.root, 'x:sheetData');
    if (data && formulasCall(data, names)) return true;
  }
  const root = xmlPartOf(pkg, index.mainPart)?.root;
  const defined = child(root, 'x:definedNames');
  for (const n of defined?.children ?? []) {
    if (n.type !== 'element') continue;
    const list = isAlternateContent(n) ? wrappedDefinedNames(n) : [n];
    if (list.some((d) => callsSheetFunction(textOf(d), names))) return true;
  }
  // A name list in a wrapper of `x:workbook` (`wrapped-formulas.ts`): its names read the sheets too.
  for (const n of root?.children ?? [])
    if (
      isAlternateContent(n) &&
      wrappedDefinedNames(n).some((d) => callsSheetFunction(textOf(d), names))
    )
      return true;
  return false;
}
