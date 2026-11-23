// Conditional formatting (Part 1, 18.3.1.18 `x:conditionalFormatting`, 18.3.1.10 `x:cfRule`):
// the rules of a worksheet, their evaluation for one cell, and the differential format overlay.
// Schema types (transitional `sml.xsd`): `x:conditionalFormatting` is `CT_ConditionalFormatting`
// with `@sqref` (`ST_Sqref`), and `x:cfRule` is `CT_CfRule` with `@type` (`ST_CfType`),
// `@dxfId` (`ST_DxfId`), `@priority` (`xsd:int`, required), and `@stopIfTrue` (default false).
//
// The engine evaluates `cellIs` rules whose operands are literals: a number, a quoted string,
// `TRUE`, or `FALSE`. Every other rule is not shown: it never paints, and it never stops a
// lower rule. A rule reads the stored value of its cell; it never evaluates a formula.

import type { Element } from '../store/package/tree.ts';
import type { XlsxPackage } from '../store/package/package.ts';
import { attr } from '../store/package/query.ts';
import { NS } from '../store/package/namespaces.ts';
import { parseCellRef } from '../store/cell-ref.ts';
import { styleLists, type WorkbookIndex } from '../store/workbook/workbook-index.ts';
import { overlayDxf, type ResolvedCellFormat } from './cell-format.ts';
import { workbookTheme } from './theme.ts';
import { parseDouble, type CellValue } from './values.ts';

/** At most this many `x:cfRule` elements are read from one sheet. */
export const MAX_RULES = 4_096;
/** At most this many ranges are read from one `@sqref`. */
export const MAX_RANGES = 512;
/** At most this many rule and range pairs are read from one sheet. */
export const MAX_PAIRS = 65_536;
/** At most this many rules are walked for one cell. */
export const MAX_RULES_PER_CELL = 64;
/** At most this many rule ranges are kept for one row, and checked for each of its cells. */
export const MAX_ENTRIES_PER_ROW = 1_024;
/** At most this many rows keep their rule ranges in the cache of a sheet. */
const ROW_CACHE_CAP = 4_096;
/** At most this many rule ranges are kept in the cache of a sheet, over all its rows. */
const ENTRY_CACHE_CAP = 262_144;
/** The longest `@sqref` that is read. */
const MAX_SQREF_LENGTH = 16_384;
/** The longest `x:formula` operand that is read. */
const MAX_OPERAND_LENGTH = 8_192;

/** `ST_ConditionalFormattingOperator`, Part 1, 18.18.15, for `cellIs`. */
export type CellIsOperator =
  | 'lessThan'
  | 'lessThanOrEqual'
  | 'equal'
  | 'notEqual'
  | 'greaterThanOrEqual'
  | 'greaterThan'
  | 'between'
  | 'notBetween';

const OPERATORS: ReadonlySet<string> = new Set<CellIsOperator>([
  'lessThan', 'lessThanOrEqual', 'equal', 'notEqual', 'greaterThanOrEqual', 'greaterThan',
  'between', 'notBetween',
]); // prettier-ignore

/** A literal operand of a `cellIs` rule. */
export type Operand =
  | { readonly kind: 'number'; readonly value: number }
  | {
      readonly kind: 'string';
      readonly text: string;
      /** The text in lowercase, made once when the rule is read. */
      readonly lower: string;
    }
  | { readonly kind: 'boolean'; readonly value: boolean };

/** A range of a rule, 0-based and inclusive. */
export interface RuleRange {
  readonly firstRow: number;
  readonly firstCol: number;
  readonly lastRow: number;
  readonly lastCol: number;
}

/** Why the engine does not show a rule. */
export type NotShownReason = 'needs-formula' | 'unsupported' | 'malformed';

/** One `x:cfRule` (`CT_CfRule`) with the ranges of its `x:conditionalFormatting`. */
export interface ConditionalRule {
  /** 0-based ordinal of the `x:cfRule` in document order. */
  readonly ordinal: number;
  /** `@priority` (`xsd:int`). */
  readonly priority: number;
  readonly ranges: readonly RuleRange[];
  /** The `x:dxf` index (`@dxfId`, `ST_DxfId`), when the rule names one. */
  readonly dxfId: number | undefined;
  /** `@stopIfTrue` (`xsd:boolean`, default false). */
  readonly stopIfTrue: boolean;
  /** The test of a shown rule. */
  readonly test?: { readonly operator: CellIsOperator; readonly operands: readonly Operand[] };
  /** Why the rule is not shown. */
  readonly notShown?: NotShownReason;
}

/**
 * The rule ranges that the cells of one layout may check, over all cells. Past it, the rest of
 * the cells of the layout keep their cell formats.
 */
export const LAYOUT_CHECK_BUDGET = 4_000_000;

/** A count of rule-range checks that callers share across the cells of one layout. */
export interface CheckBudget {
  left: number;
  /**
   * Charge each row once as a new row, also a row in the cache, so the result depends only on
   * the cells and not on the cache history. A copy sets it and walks its rows in order; a layout
   * does not.
   */
  readonly everyRow?: boolean;
  /** With `everyRow`, the last row that was charged: the other cells of that row pay no more. */
  charged?: number;
}

/** The rules of a worksheet in walk order: ascending `@priority`, then document order. */
export interface SheetRules {
  readonly rules: readonly ConditionalRule[];
  /** `x:cfRule` elements past {@link MAX_RULES} or {@link MAX_PAIRS}, which are not read. */
  readonly skipped: number;
  /** The rule and range pairs that were read. */
  readonly pairs: number;
  /** The pairs of the shown rules: the cost of building the ranges of one row. */
  readonly shownPairs: number;
}

const EMPTY: SheetRules = Object.freeze({ rules: [], skipped: 0, pairs: 0, shownPairs: 0 });
const cache = new WeakMap<Element, SheetRules>();

/** The digits of a number without its leading zeros (one zero for zero), or undefined. */
function digitsOf(raw: string): string | undefined {
  if (raw.length === 0 || raw.length > 64) return undefined;
  for (let i = 0; i < raw.length; i++) {
    const c = raw.charCodeAt(i);
    if (c < 48 || c > 57) return undefined;
  }
  let i = 0;
  while (i < raw.length - 1 && raw.charCodeAt(i) === 48) i++;
  return raw.slice(i);
}

/**
 * An `xsd:unsignedInt` (`ST_DxfId`): an optional `+`, leading zeros, and at most 4,294,967,295,
 * or undefined.
 */
function integer(raw: string | undefined): number | undefined {
  if (raw === undefined) return undefined;
  const digits = digitsOf(raw.startsWith('+') ? raw.slice(1) : raw);
  if (digits === undefined || digits.length > 10) return undefined;
  const n = Number(digits);
  return n <= 4_294_967_295 ? n : undefined;
}

/** An `xsd:int`: an optional sign, leading zeros, and the 32-bit range, or undefined. */
function xsdInt(raw: string | undefined): number | undefined {
  if (raw === undefined) return undefined;
  const sign = raw.startsWith('-') ? -1 : 1;
  const digits = digitsOf(raw.startsWith('-') || raw.startsWith('+') ? raw.slice(1) : raw);
  if (digits === undefined || digits.length > 10) return undefined;
  const n = sign * Number(digits);
  return n >= -2_147_483_648 && n <= 2_147_483_647 ? n : undefined;
}

/**
 * The ranges of `@sqref` (`ST_Sqref`), or undefined when a range cannot be read or there are too
 * many.
 */
export function sqrefRanges(raw: string | undefined): RuleRange[] | undefined {
  if (raw === undefined || raw.length === 0 || raw.length > MAX_SQREF_LENGTH) return undefined;
  const ranges: RuleRange[] = [];
  // `ST_Sqref` is a list: any XML white space separates its ranges.
  for (const token of raw.split(/[ \t\r\n]/)) {
    if (token.length === 0) continue;
    if (ranges.length >= MAX_RANGES) return undefined;
    const colon = token.indexOf(':');
    const a = parseCellRef(colon < 0 ? token : token.slice(0, colon));
    const b = colon < 0 ? a : parseCellRef(token.slice(colon + 1));
    if (!a || !b) return undefined;
    ranges.push({
      firstRow: Math.min(a.row, b.row),
      firstCol: Math.min(a.col, b.col),
      lastRow: Math.max(a.row, b.row),
      lastCol: Math.max(a.col, b.col),
    });
  }
  return ranges.length > 0 ? ranges : undefined;
}

/** A literal operand, or undefined when the formula text is anything else. */
export function literalOperand(raw: string): Operand | undefined {
  if (raw.length > MAX_OPERAND_LENGTH) return undefined;
  // XML white space only: a no-break space is part of the operand.
  let start = 0;
  let end = raw.length;
  while (start < end && ' \t\r\n'.includes(raw[start]!)) start++;
  while (end > start && ' \t\r\n'.includes(raw[end - 1]!)) end--;
  const text = raw.slice(start, end);
  if (text.length >= 2 && text.startsWith('"') && text.endsWith('"')) {
    const inner = text.slice(1, -1);
    // A quote inside a string literal is doubled; a single quote ends the literal early.
    if (inner.replaceAll('""', '').includes('"')) return undefined;
    const unquoted = inner.replaceAll('""', '"');
    return { kind: 'string', text: unquoted, lower: unquoted.toLowerCase() };
  }
  const upper = text.toUpperCase();
  if (upper === 'TRUE' || upper === 'FALSE') return { kind: 'boolean', value: upper === 'TRUE' };
  const n = parseDouble(text);
  return n === undefined ? undefined : { kind: 'number', value: n };
}

function readRule(el: Element, ordinal: number, ranges: RuleRange[]): ConditionalRule {
  // `xsd:int`, `ST_DxfId`, and `xsd:boolean` collapse white space.
  const priority = xsdInt(attr(el, 'priority')?.trim());
  const dxfRaw = attr(el, 'dxfId')?.trim();
  const dxfId = integer(dxfRaw);
  const stop = attr(el, 'stopIfTrue')?.trim();
  const base = {
    ordinal,
    priority: priority ?? Number.MAX_SAFE_INTEGER,
    ranges,
    dxfId,
    stopIfTrue: stop === '1' || stop === 'true',
  };
  if (priority === undefined || (dxfRaw !== undefined && dxfId === undefined))
    return { ...base, notShown: 'malformed' };
  if (attr(el, 'type')?.trim() !== 'cellIs') return { ...base, notShown: 'unsupported' };
  const operator = attr(el, 'operator')?.trim();
  if (operator === undefined) return { ...base, notShown: 'malformed' };
  // The text operators of `ST_ConditionalFormattingOperator` belong to other rule types.
  if (!OPERATORS.has(operator)) return { ...base, notShown: 'unsupported' };
  const wanted = operator === 'between' || operator === 'notBetween' ? 2 : 1;
  const operands: Operand[] = [];
  for (const node of el.children) {
    if (node.type !== 'element' || node.ns !== NS.x || node.local !== 'formula') continue;
    if (operands.length === wanted) return { ...base, notShown: 'malformed' };
    const text = node.children.map((c) => (c.type === 'text' ? c.value : '')).join('');
    const operand = literalOperand(text);
    if (!operand) return { ...base, notShown: 'needs-formula' };
    operands.push(operand);
  }
  if (operands.length !== wanted) return { ...base, notShown: 'malformed' };
  return { ...base, test: { operator: operator as CellIsOperator, operands } };
}

/** The conditional formatting rules of a worksheet root. Cached per root. */
export function sheetRules(worksheet: Element): SheetRules {
  const hit = cache.get(worksheet);
  if (hit) return hit;
  const rules: ConditionalRule[] = [];
  let skipped = 0;
  let pairs = 0;
  let ordinal = 0;
  for (const block of worksheet.children) {
    if (block.type !== 'element' || block.ns !== NS.x || block.local !== 'conditionalFormatting')
      continue;
    const ranges = sqrefRanges(attr(block, 'sqref'));
    for (const el of block.children) {
      if (el.type !== 'element' || el.ns !== NS.x || el.local !== 'cfRule') continue;
      const at = ordinal++;
      const count = ranges?.length ?? 0;
      if (rules.length >= MAX_RULES || pairs + count > MAX_PAIRS) {
        skipped += 1;
        continue;
      }
      pairs += count;
      const rule = readRule(el, at, ranges ?? []);
      rules.push(ranges ? rule : { ...rule, test: undefined, notShown: 'malformed' });
    }
  }
  if (rules.length === 0 && skipped === 0) {
    cache.set(worksheet, EMPTY);
    return EMPTY;
  }
  rules.sort((a, b) => a.priority - b.priority || a.ordinal - b.ordinal);
  const shownPairs = rules.reduce((n, r) => (r.test ? n + r.ranges.length : n), 0);
  const result = { rules, skipped, pairs, shownPairs };
  cache.set(worksheet, result);
  return result;
}

const notShownCache = new WeakMap<Element, boolean>();

/**
 * Whether a worksheet has conditional formatting that the grid can leave out: a rule that is
 * not shown, rules past the read limits, more shown rules or ranges than the caps of one cell
 * or one row take, or rules in an extension list (`x:extLst`), which the engine keeps and never
 * reads. The chrome tells the user, so no rule is dropped silently.
 */
export function sheetHasHiddenRules(worksheet: Element): boolean {
  const hit = notShownCache.get(worksheet);
  if (hit !== undefined) return hit;
  const sheet = sheetRules(worksheet);
  let hidden = sheet.skipped > 0 || sheet.rules.some((r) => r.notShown !== undefined);
  // A cell can lose shown rules to the caps of one cell and one row: say so when it can occur.
  const shown = sheet.rules.filter((r) => r.test);
  if (shown.length > MAX_RULES_PER_CELL) hidden = true;
  if (shown.reduce((n, r) => n + r.ranges.length, 0) > MAX_ENTRIES_PER_ROW) hidden = true;
  for (const list of worksheet.children) {
    if (hidden) break;
    if (list.type !== 'element' || list.ns !== NS.x || list.local !== 'extLst') continue;
    for (const ext of list.children) {
      if (ext.type !== 'element' || ext.local !== 'ext') continue;
      if (ext.children.some((c) => c.type === 'element' && c.local === 'conditionalFormattings'))
        hidden = true;
    }
  }
  notShownCache.set(worksheet, hidden);
  return hidden;
}

/** One range of a rule that meets a row: the rule and the columns of the range. */
interface RowEntry {
  readonly rule: ConditionalRule;
  readonly firstCol: number;
  readonly lastCol: number;
}

/** The ranges of the shown rules that meet one row, and the columns that they span. */
interface RowRanges {
  readonly entries: readonly RowEntry[];
  readonly firstCol: number;
  readonly lastCol: number;
}

interface RowCache {
  readonly rows: Map<number, RowRanges>;
  entries: number;
}

const rowCaches = new WeakMap<SheetRules, RowCache>();

/**
 * The ranges of the shown rules that meet `row`, in walk order. Built once for each row in one
 * pass over the rule and range pairs, and cached for each sheet.
 */
function rowEntries(sheet: SheetRules, row: number): RowRanges {
  let cache = rowCaches.get(sheet);
  if (!cache) {
    cache = { rows: new Map(), entries: 0 };
    rowCaches.set(sheet, cache);
  }
  const hit = cache.rows.get(row);
  if (hit) return hit;
  const entries: RowEntry[] = [];
  let firstCol = Infinity;
  let lastCol = -1;
  // A cell checks only the first ranges of its row, so the list stops there.
  for (const rule of sheet.rules) {
    if (entries.length === MAX_ENTRIES_PER_ROW) break;
    if (!rule.test) continue;
    for (const r of rule.ranges) {
      if (entries.length === MAX_ENTRIES_PER_ROW) break;
      if (row < r.firstRow || row > r.lastRow) continue;
      entries.push({ rule, firstCol: r.firstCol, lastCol: r.lastCol });
      firstCol = Math.min(firstCol, r.firstCol);
      lastCol = Math.max(lastCol, r.lastCol);
    }
  }
  const ranges = { entries, firstCol, lastCol };
  if (cache.rows.size >= ROW_CACHE_CAP || cache.entries + entries.length > ENTRY_CACHE_CAP) {
    cache.rows.clear();
    cache.entries = 0;
  }
  cache.rows.set(row, ranges);
  cache.entries += entries.length;
  return ranges;
}

/** The ranges of a row, with the cost of a row that is not in the cache taken from `budget`. */
function chargedRow(sheet: SheetRules, row: number, budget?: CheckBudget): RowRanges {
  // A row that is not in the cache costs one pass over the pairs of the shown rules.
  if (budget?.everyRow) {
    if (budget.charged !== row) budget.left -= sheet.shownPairs;
    budget.charged = row;
  } else if (budget && !rowCaches.get(sheet)?.rows.has(row)) budget.left -= sheet.shownPairs;
  return rowEntries(sheet, row);
}

/**
 * Whether a shown rule meets a row. Layout asks it before it checks the blank cells of the row,
 * so a row that no rule meets costs no check for each cell. Past the budget, it is false: the
 * cells keep their formats, as {@link conditionalFormat} keeps them.
 */
export function rowHasShownRules(worksheet: Element, row: number, budget?: CheckBudget): boolean {
  if (budget && budget.left <= 0) return false;
  const sheet = sheetRules(worksheet);
  if (sheet.rules.length === 0) return false;
  return chargedRow(sheet, row, budget).entries.length > 0;
}

/** The rows and rule ranges that the cache of a sheet keeps, for tests of the memory bound. */
export const rowCacheSizeForTest = (sheet: SheetRules): { rows: number; entries: number } => {
  const cache = rowCaches.get(sheet);
  return { rows: cache?.rows.size ?? 0, entries: cache?.entries ?? 0 };
};

/** The rule ranges that the cache keeps for `row`, for tests of the memory bound. */
export const rowEntriesForTest = (sheet: SheetRules, row: number): number =>
  rowEntries(sheet, row).entries.length;

/**
 * A cell value as the rules compare it: a number, a boolean, or a string in lowercase. Made once
 * for each cell, so a cell lowercases its text once whatever the number of rules. Undefined for
 * an error or a date value, which no rule compares.
 */
type Key =
  | { readonly kind: 'number'; readonly value: number }
  | { readonly kind: 'boolean'; readonly value: boolean }
  | { readonly kind: 'string'; readonly lower: string }
  | { readonly kind: 'empty' };

function keyOf(value: CellValue): Key | undefined {
  switch (value.kind) {
    case 'number':
    case 'boolean':
    case 'empty':
      return value;
    case 'string':
      return { kind: 'string', lower: value.text.toLowerCase() };
    default:
      return undefined;
  }
}

const rank = (kind: string) => (kind === 'number' ? 0 : kind === 'string' ? 1 : 2);

/**
 * The order of two values in a comparison: numbers before strings before booleans. Strings
 * compare without case. An empty value is 0 against a number and the empty string against a
 * string (`conditional-rules.pdf`, B16 to B18). Returns undefined when the cell value cannot be
 * compared. Two strings compare at a cost of their shorter length, which `budget` pays in units
 * of 16 code units.
 */
function compare(key: Key | undefined, operand: Operand, budget?: CheckBudget): number | undefined {
  if (!key) return undefined;
  if (key.kind === 'empty') {
    if (operand.kind === 'number') return compare({ kind: 'number', value: 0 }, operand);
    if (operand.kind === 'string') return compare({ kind: 'string', lower: '' }, operand);
    return undefined;
  }
  const byKind = rank(key.kind) - rank(operand.kind);
  if (byKind !== 0) return Math.sign(byKind);
  if (key.kind === 'string' && operand.kind === 'string') {
    if (budget) budget.left -= Math.min(key.lower.length, operand.lower.length) >> 4;
    return key.lower < operand.lower ? -1 : key.lower > operand.lower ? 1 : 0;
  }
  const left = (key as { readonly value: number | boolean }).value;
  const right = (operand as { readonly value: number | boolean }).value;
  return left < right ? -1 : left > right ? 1 : 0;
}

/**
 * The result of a rule for a cell value: undefined for a rule that is not shown, and false for
 * an error or date value, and for an empty value under a boolean operand.
 */
export function ruleIsTrue(rule: ConditionalRule, value: CellValue): boolean | undefined {
  return keyIsTrue(rule, keyOf(value));
}

function keyIsTrue(
  rule: ConditionalRule,
  key: Key | undefined,
  budget?: CheckBudget
): boolean | undefined {
  const test = rule.test;
  if (!test) return undefined;
  const [a, b] = test.operands;
  const x = compare(key, a!, budget);
  if (x === undefined) return false;
  switch (test.operator) {
    case 'lessThan':
      return x < 0;
    case 'lessThanOrEqual':
      return x <= 0;
    case 'equal':
      return x === 0;
    case 'notEqual':
      return x !== 0;
    case 'greaterThanOrEqual':
      return x >= 0;
    case 'greaterThan':
      return x > 0;
    case 'between':
    case 'notBetween': {
      const y = compare(key, b!, budget);
      if (y === undefined) return false;
      // The two bounds may come in either order.
      const inside = (x >= 0 && y <= 0) || (x <= 0 && y >= 0);
      return test.operator === 'between' ? inside : !inside;
    }
  }
}

/**
 * `format` with the formats of the true rules over the cell at `row` and `col` laid over it.
 * The walk takes rules by ascending priority; each property takes the first value that a true
 * rule sets, and a true rule with `stopIfTrue` ends the walk. A rule that is not shown is
 * skipped. At most {@link MAX_ENTRIES_PER_ROW} rule ranges of the row are checked, and at
 * most {@link MAX_RULES_PER_CELL} rules over the cell are walked.
 */
export function conditionalFormat(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  worksheet: Element,
  row: number,
  col: number,
  value: CellValue,
  format: ResolvedCellFormat,
  budget?: CheckBudget
): ResolvedCellFormat {
  if (budget && budget.left <= 0) return format;
  const sheet = sheetRules(worksheet);
  if (sheet.rules.length === 0) return format;
  const ranges = chargedRow(sheet, row, budget);
  // A cell outside the columns of every range of its row checks no range.
  if (col < ranges.firstCol || col > ranges.lastCol) return format;
  const { entries } = ranges;
  const lists = styleLists(pkg, index);
  const applied: Element[] = [];
  let walked = 0;
  let last: ConditionalRule | undefined;
  // The cell value as the rules compare it, made at the first rule that meets the cell.
  let key: Key | undefined | null = null;
  if (budget) budget.left -= entries.length;
  for (const entry of entries) {
    // A rule with two ranges over the cell is walked once.
    if (col < entry.firstCol || col > entry.lastCol || entry.rule === last) continue;
    last = entry.rule;
    if (++walked > MAX_RULES_PER_CELL) break;
    if (key === null) key = keyOf(value);
    if (keyIsTrue(entry.rule, key, budget) !== true) continue;
    const dxfId = entry.rule.dxfId;
    const dxf = dxfId === undefined ? undefined : lists.dxfs[dxfId];
    if (dxf) applied.push(dxf);
    if (entry.rule.stopIfTrue) break;
  }
  if (applied.length === 0) return format;
  const theme = workbookTheme(pkg, index);
  // The first true rule wins each property, so lay the formats from the last to the first.
  let out = format;
  for (let i = applied.length - 1; i >= 0; i--)
    out = overlayDxf(out, applied[i]!, theme, lists.indexedColors);
  return out;
}
