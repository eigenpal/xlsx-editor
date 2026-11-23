// Cell values as the file stores them (design D3): one result for each `ST_CellType`
// (ECMA-376 Part 1, 18.18.11) of `x:c/@t` (`CT_Cell`, default `n`), inline strings, and `t="d"`
// ISO 8601 dates (`d` is not in the transitional `sml.xsd` enumeration). The engine never
// evaluates a formula: a formula cell shows its cached `x:v`.
//
// Stale state (design D3): `x:calcPr/@fullCalcOnLoad` (`CT_CalcPr`) makes every formula result of
// the workbook stale, and `x:sheetCalcPr/@fullCalcOnLoad` (`CT_SheetCalcPr`) makes the formula
// results of one sheet stale. A formula cell with no cached value is stale too. A cell whose `x:f`
// is in a branch of an `mc:AlternateContent` child counts as a formula cell by the shared rule
// (`cellHasFormula`). Its cached value is the cell's `x:v`, else the branch value of
// `wrappedValue`.

import type { Element } from '../store/package/tree.ts';
import type { XlsxPackage } from '../store/package/package.ts';
import { attr, child, textOf } from '../store/package/query.ts';
import { cellHasFormula, wrappedValue } from '../store/workbook/cell-formula.ts';
import { sharedStringItems, type WorkbookIndex } from '../store/workbook/workbook-index.ts';
import { decodeXstring } from '../store/workbook/xstring.ts';
import { NS } from '../store/package/namespaces.ts';
import type { CascadeProblem } from './problems.ts';

/** One run of a rich string, as stored. `richTextRuns` (cell-format.ts) resolves its font. */
export interface TextRun {
  /** The text of the run (`x:r/x:t`, `CT_RElt`, `ST_Xstring`), with `_xHHHH_` escapes decoded. */
  readonly text: string;
}

/** The `x:rPr` (`CT_RPrElt`) of each stored run, kept out of the public run type. */
const runProps = new WeakMap<TextRun, Element>();

/** The `x:rPr` of a run that `stringItem` read, or undefined. */
export function runPropsOf(run: TextRun): Element | undefined {
  return runProps.get(run);
}

export type CellValue =
  | { readonly kind: 'empty' }
  | { readonly kind: 'number'; readonly value: number }
  | { readonly kind: 'string'; readonly text: string; readonly runs?: readonly TextRun[] }
  | { readonly kind: 'boolean'; readonly value: boolean }
  | { readonly kind: 'error'; readonly code: string }
  /** `t="d"`: the ISO 8601 text as stored. The number format converts it to a serial. */
  | { readonly kind: 'date'; readonly iso: string };

export interface ResolvedValue {
  readonly value: CellValue;
  /** The cell holds `x:f`. */
  readonly formula: boolean;
  /** The value is a cached formula result that can be out of date. */
  readonly stale: boolean;
  /** Why a stored value cannot be shown, for the caller to record with the cell reference. */
  readonly problem?: CascadeProblem;
}

const EMPTY: CellValue = Object.freeze({ kind: 'empty' });

/**
 * The error values of `ST_CellType` `e` (Part 1, 18.17.3 lists the error literals), and the
 * error values that workbooks store for dynamic arrays, data types, and external data.
 */
const ERRORS = new Set([
  '#NULL!',
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
  '#GETTING_DATA',
  '#SPILL!',
  '#CALC!',
  '#FIELD!',
  '#BLOCKED!',
  '#CONNECT!',
  '#BUSY!',
  '#UNKNOWN!',
]);

/** Whether `code` is an error value that a cell of type `e` can store. */
export function isErrorValue(code: string): boolean {
  return ERRORS.has(code);
}

/** The text of a string item, and its runs when it has `x:r` children. */
export interface StringItem {
  /** The text of the item: its own `x:t` text, then the text of each run in order, decoded. */
  readonly text: string;
  /** The runs of a rich string (`x:r`), in order. Absent when the item has no runs. */
  readonly runs?: readonly TextRun[];
}

/**
 * The string items that were read. A tree node never changes, and the sender controls how many
 * children an item has, so a layout or a paint must not walk the same item again.
 */
const readItems = new WeakMap<Element, StringItem>();

/**
 * The text of an `x:si` or `x:is` string item (`CT_Rst`), and its runs when it has `x:r`
 * children.
 */
export function stringItem(item: Element | undefined): StringItem {
  if (!item) return { text: '' };
  const known = readItems.get(item);
  if (known) return known;
  const read = readStringItem(item);
  readItems.set(item, read);
  return read;
}

function readStringItem(item: Element): StringItem {
  const t = child(item, 'x:t');
  const runs: TextRun[] = [];
  for (const node of item.children) {
    if (node.type !== 'element' || node.local !== 'r') continue;
    const run: TextRun = { text: decodeXstring(textOf(child(node, 'x:t'))) };
    const props = child(node, 'x:rPr');
    if (props) runProps.set(run, props);
    runs.push(run);
  }
  if (runs.length === 0) return { text: t ? decodeXstring(textOf(t)) : '' };
  let text = t ? decodeXstring(textOf(t)) : '';
  for (const run of runs) text += run.text;
  return { text, runs };
}

/** Whether the formula results of a sheet are stale from the file. */
export function sheetStale(index: WorkbookIndex, worksheetRoot: Element): boolean {
  if (index.fullCalcOnLoad) return true;
  const v = attr(child(worksheetRoot, 'x:sheetCalcPr'), 'fullCalcOnLoad');
  return v === '1' || v === 'true';
}

/**
 * The stored value of a cell. `stale` is the sheet's stale state from {@link sheetStale}; it
 * applies to formula cells only.
 */
export function cellValue(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  cell: Element,
  stale: boolean
): ResolvedValue {
  const formula = cellHasFormula(cell);
  const raw = storedText(cell, formula);
  const t = attr(cell, 't') ?? 'n';
  const done = (value: CellValue, problem?: CascadeProblem): ResolvedValue => ({
    value,
    formula,
    // A formula with no cached value shows nothing and is stale.
    stale: formula && (stale || raw === undefined),
    ...(problem ? { problem } : {}),
  });
  if (t === 'inlineStr') {
    const is = child(cell, 'x:is');
    if (!is) return done(raw === undefined ? EMPTY : { kind: 'string', text: decodeXstring(raw) });
    const s = stringItem(is);
    return done({ kind: 'string', text: s.text, ...(s.runs ? { runs: s.runs } : {}) });
  }
  if (raw === undefined) return done(EMPTY);
  switch (t) {
    case 'n': {
      const n = parseDouble(raw);
      return n !== undefined
        ? done({ kind: 'number', value: n })
        : done(EMPTY, { code: 'bad-number' });
    }
    case 's': {
      const items = sharedStringItems(pkg, index);
      // The same index as `cellSharedStringIndex(cell)`, from the text that this function read.
      const i = decimalIndex(raw);
      if (i === undefined || i >= items.length)
        return done(EMPTY, { code: 'shared-string-out-of-range' });
      const s = stringItem(items[i]);
      return done({ kind: 'string', text: s.text, ...(s.runs ? { runs: s.runs } : {}) });
    }
    case 'str':
      return done({ kind: 'string', text: decodeXstring(raw) });
    case 'b':
      if (raw === '1' || raw === 'true') return done({ kind: 'boolean', value: true });
      if (raw === '0' || raw === 'false') return done({ kind: 'boolean', value: false });
      return done(EMPTY, { code: 'bad-boolean' });
    case 'e':
      return done({ kind: 'error', code: ERRORS.has(raw) ? raw : raw.slice(0, 16) });
    case 'd':
      return done({ kind: 'date', iso: raw.slice(0, 64) });
    default:
      return done(EMPTY, { code: 'unknown-cell-type' });
  }
}

/** The string items that were checked by {@link stringItemHidesText}. */
const checkedItems = new WeakMap<Element, boolean>();

/** Whether an `x:t` holds an element, whose text `stringItem` would join with its own. */
const textHoldsElement = (t: Element): boolean => t.children.some((k) => k.type === 'element');

/**
 * Whether a string item (`x:si` or `x:is`) holds text in a form that {@link stringItem} does not
 * read: an element in another namespace (`mc:AlternateContent` included) as a child of the item
 * or of one of its runs, or an element inside an `x:t`. Such an item reads as `""` or as part of
 * its text. One pass over the children of the item and of each run, with no deeper walk. Read
 * once per item.
 */
export function stringItemHidesText(item: Element | undefined): boolean {
  if (!item) return false;
  const known = checkedItems.get(item);
  if (known !== undefined) return known;
  const foreign = (k: Element): boolean =>
    k.ns !== NS.x || (k.local === 't' && textHoldsElement(k));
  let hides = false;
  for (const node of item.children) {
    if (node.type !== 'element') continue;
    if (foreign(node)) hides = true;
    else if (node.local === 'r')
      hides = node.children.some((k) => k.type === 'element' && foreign(k));
    if (hides) break;
  }
  checkedItems.set(item, hides);
  return hides;
}

/**
 * The string item of a cell that stores one: the `x:is` of an inline string, or the `x:si` that
 * a shared string cell names, as {@link cellValue} reads it. Undefined for any other cell.
 */
export function cellStringItem(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  cell: Element
): Element | undefined {
  const t = attr(cell, 't');
  if (t === 'inlineStr') return child(cell, 'x:is');
  if (t !== 's') return undefined;
  const raw = storedText(cell, cellHasFormula(cell));
  const i = raw === undefined ? undefined : decimalIndex(raw);
  return i === undefined ? undefined : sharedStringItems(pkg, index)[i];
}

const isDigit = (c: number): boolean => c >= 48 && c <= 57;

/**
 * A finite `xsd:double` (XML Schema), as in `x:c/x:v` of a number cell (`CT_Cell`; the schema
 * types `x:v` as `ST_Xstring`): an optional sign, digits with an optional point, and an optional
 * exponent. Refuses hex, binary, and octal literals, surrounding spaces,
 * `INF`, and `NaN`. One pass, no regular expression.
 */
export function parseDouble(raw: string): number | undefined {
  const n = raw.length;
  if (n === 0 || n > 64) return undefined;
  let i = 0;
  if (raw[i] === '+' || raw[i] === '-') i++;
  let digits = 0;
  while (i < n && isDigit(raw.charCodeAt(i))) {
    i++;
    digits++;
  }
  if (raw[i] === '.') {
    i++;
    while (i < n && isDigit(raw.charCodeAt(i))) {
      i++;
      digits++;
    }
  }
  if (digits === 0) return undefined;
  if (raw[i] === 'e' || raw[i] === 'E') {
    i++;
    if (raw[i] === '+' || raw[i] === '-') i++;
    let expDigits = 0;
    while (i < n && isDigit(raw.charCodeAt(i))) {
      i++;
      expDigits++;
    }
    if (expDigits === 0) return undefined;
  }
  if (i !== n) return undefined;
  const value = Number(raw);
  return Number.isFinite(value) ? value : undefined;
}

/**
 * The shared string index that a cell names, as {@link cellValue} reads it, or undefined when the
 * cell stores no valid index. Layout keeps it to see whether the item of a cell is the same.
 */
export function cellSharedStringIndex(cell: Element): number | undefined {
  const raw = storedText(cell, cellHasFormula(cell));
  return raw === undefined ? undefined : decimalIndex(raw);
}

/** The text of the stored value of a cell, or undefined when it stores none. */
function storedText(cell: Element, formula: boolean): string | undefined {
  // A wrapped formula keeps its cached value in a wrapper branch when the cell has none.
  const v = child(cell, 'x:v') ?? (formula && !child(cell, 'x:f') ? wrappedValue(cell) : undefined);
  return v ? textOf(v) : undefined;
}

/** A shared string index: decimal digits only, at most 9, no sign and no spaces. */
function decimalIndex(raw: string): number | undefined {
  if (raw.length === 0 || raw.length > 9) return undefined;
  let value = 0;
  for (let k = 0; k < raw.length; k++) {
    const c = raw.charCodeAt(k);
    if (!isDigit(c)) return undefined;
    value = value * 10 + (c - 48);
  }
  return value;
}
