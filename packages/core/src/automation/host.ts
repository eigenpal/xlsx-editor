// The automation host over a store, after PPTX Editor's `packages/core/src/automation/host.ts`
// and `server-host.ts` at d4502552, which follow DOCX Editor's at adf0e178. Differences: the
// operations are the workbook subset of `protocol.ts`; there are no handles, bindings,
// selection, or events; a worksheet id is its part name; and there is no browser host.
//
// THE BATCH RULE. Every operation of a batch reads the workbook as it was before the batch, and
// the value writes commit in one transaction at the end, through `WorkbookDocumentStore.transact`,
// the only write path. A range read of values or formats after a write in the same batch is
// refused (`read-after-write`), so a batch never reads a workbook that it has half changed, and
// no read costs a pass over a changed worksheet. If an operation is refused, nothing is
// committed, and the response names the failing operation and marks every other one `skipped`.
//
// VALUES AS STORED. A read gives the value that the file stores. A cell with a formula, and a cell
// of an array formula or data table range (`x:f/@t` `array` or `dataTable` with `@ref`,
// `CT_CellFormula` and `ST_CellFormulaType`, sml.xsd), holds a result that the engine does not
// calculate, so a value read or a write that meets one is refused. A worksheet whose array
// ranges cannot all be read (an `@ref` that names no range), or that has more than
// `MAX_ARRAY_RANGES` of them, refuses every value read and write. A worksheet whose cells the
// sheet index cannot place (`SheetIndex.problems`) refuses each read, as the store refuses each
// write to it. A number format read gives the stored code or the ECMA-376 code of a built-in id,
// and refuses a format with no such code, so a caller never sees `General` for another format.
// A cell that holds value content in a child in another namespace, `mc:AlternateContent`
// included (`holdsWrappedValue`), refuses its value read: the read cannot give that value, and an
// empty read written back would remove it. So does a cell whose string item, its `x:is` or the
// `x:si` that it names (both `CT_Rst`), holds text in such a child (`stringItemHidesText`).
//
// REFUSED BEFORE THE STORE. The planner refuses each cell write that the store would refuse on
// its own: a text longer than `MAX_CELL_TEXT`, a cell whose children the edit can neither keep
// nor replace (`cellChildBlocker`), a locked cell of a protected worksheet (`cellLocked`, the
// rule of the store's `lockedCellIn` and `protectionRefusal`), and a new text that the shared string table cannot
// take. A run of value writes that the store cannot apply in one pass falls back to one op at a
// time, and a store refusal costs a search over prefixes of the batch, so these checks keep a
// refused batch to one planning pass. The editor checks the protection before its edit too. A
// value write changes no cell format and no protection, so the workbook before the batch gives
// the answer for each write.
//
// ONE PASS FOR EACH WORKSHEET. The store applies a run of value writes to one worksheet in one
// pass, and starts a new pass at each change of worksheet. The writes of a batch go to the store
// grouped by worksheet part (`bySheet`), in their order within each part, so a batch that moves
// between worksheets costs one pass for each worksheet, not one for each move. Writes to two
// worksheets do not depend on each other, so the grouped order gives the same workbook.
//
// LIMITS. A batch carries at most `MAX_BATCH_OPERATIONS` operations; one range reaches at most
// `MAX_RANGE_CELLS` cells; the reads and writes of a batch reach at most `MAX_BATCH_CELLS` cells,
// its writes at most `MAX_BATCH_WRITE_CELLS`, and its worksheet lists at most
// `MAX_BATCH_LISTED_WORKSHEETS` worksheets. A cell with more than `MAX_CELL_CHILDREN` child nodes
// refuses its value read and write. A worksheet with more than `MAX_MERGES` merged ranges, the
// layout limit, refuses each value write, and the writes of a batch compare their bounds with at
// most `MAX_BATCH_MERGE_CHECKS` merged ranges in all. Each is a count, never a time, as DOCX
// Editor counts the nodes of a fragment (`MAX_FRAGMENT_NODES`, `packages/core/src/store/store/
// tree-op-fragment.ts` at ad686574).

import {
  EN_US,
  builtinFormatCode,
  cellFormat,
  cellValue,
  dateToSerial,
  formatCodeOf,
  type CellValue,
} from '../cascade/index.ts';
import { cellStringItem, isErrorValue, stringItemHidesText } from '../cascade/values.ts';
import {
  arrayFormulaIn,
  arrayRangeCount,
  attr,
  cellChildBlocker,
  cellHasFormula,
  cellKey,
  columnFormatAt,
  formatCellRef,
  readSheetProtection,
  readXlsx,
  rowFormat,
  saveXlsx,
  sheetIndex,
  workbookIndex,
  xmlPartOf,
  type RangeBox,
  type SheetEntry,
  type SheetIndex,
  type WorkbookIndex,
  type WorkbookOp,
  type XlsxPackage,
  type XmlPart,
  WorkbookDocumentStore,
} from '../store/index.ts';
import { parseCellRef } from '../store/cell-ref.ts';
import type { Element } from '../store/package/tree.ts';
import { holdsWrappedValue } from '../store/store/cells.ts';
import { lockReader, protectionChecks } from '../store/store/sheet-protection.ts';
import { hostWork, memoized, readMemo, type ReadMemo } from './host-memo.ts';
import { formatIndex } from '../store/workbook/inherited-format.ts';
import { sheetsByKey } from '../store/workbook/workbook-index.ts';
import { MAX_MERGES } from '../layout/panes.ts';
import { mergedBoxesMeeting, mergedRangeCount } from '../store/store/merge-cells.ts';
import { MAX_CELL_TEXT, applyOps } from '../store/store/ops.ts';
import { openDiagnosticOf } from './open-diagnostics.ts';
import { sharedStringIndices } from '../store/store/shared-strings.ts';
import type {
  AutomationBatchRequest,
  AutomationBatchResponse,
  AutomationCellValue,
  AutomationCellWrite,
  AutomationError,
  AutomationErrorCode,
  AutomationHost,
  AutomationOpenResult,
  AutomationOperation,
  AutomationOperationResult,
  AutomationRangeRead,
  AutomationValue,
  AutomationWorksheetRead,
} from './protocol.ts';
import {
  MAX_RANGE_CELLS,
  boundsCells,
  parseRangeAddress,
  type RangeBounds,
} from './range-address.ts';

/** The most operations one batch may carry. */
export const MAX_BATCH_OPERATIONS = 10_000;

/** The most cells that the range reads and writes of one batch may reach in all. */
export const MAX_BATCH_CELLS = 1_000_000;

/** The most cells that the value writes of one batch may reach in all. */
export const MAX_BATCH_WRITE_CELLS = 10_000;

/**
 * The most merged ranges that the value writes of one batch may compare their bounds with, in
 * all. Each write compares its bounds with each merged range of its worksheet.
 */
export const MAX_BATCH_MERGE_CHECKS = 50_000_000;

/** The most worksheets that the lists of one batch may give in all. */
export const MAX_BATCH_LISTED_WORKSHEETS = 100_000;

/**
 * The most array formula and data table ranges of a worksheet whose values a batch reads or
 * writes. Each range operation compares its bounds with each of them.
 */
export const MAX_ARRAY_RANGES = 10_000;

/**
 * The most child nodes of a cell whose value a batch reads or writes. Each operation on a cell
 * reads its children, so a cell with more refuses, and a batch of operations on one cell stays
 * bounded.
 */
export const MAX_CELL_CHILDREN = 1_000;

/** The refusal of one operation, thrown inside the planner and caught by the batch loop. */
class Refusal {
  constructor(readonly error: AutomationError) {}
}

function automationError(
  code: AutomationErrorCode,
  message: string,
  detail?: string,
  cell?: string
): AutomationError {
  return Object.freeze({
    code,
    message,
    ...(detail === undefined ? {} : { detail }),
    ...(cell === undefined ? {} : { cell }),
  });
}

/** Refuse the operation. `cell` is the A1 reference of the cell that the refusal is about. */
function refuseWith(
  code: AutomationErrorCode,
  message: string,
  detail?: string,
  cell?: string
): never {
  throw new Refusal(automationError(code, message, detail, cell));
}

const SKIPPED: AutomationOperationResult = Object.freeze({ status: 'skipped' as const });

/** A batch that answered nothing: the failure at its index, `skipped` everywhere else. */
function refused(
  count: number,
  index: number,
  error: AutomationError,
  revision: number
): AutomationBatchResponse {
  const results: AutomationOperationResult[] = [];
  for (let at = 0; at < count; at += 1)
    results.push(at === index ? { status: 'error', error } : SKIPPED);
  return { ok: false, results, revision, changed: false };
}

/**
 * The sheet a workbook opens on: `x:workbookView/@activeTab` (`CT_BookView`, default 0) when it
 * is a visible worksheet, else the first visible worksheet, else -1. The same rule as the
 * editor's opening sheet.
 */
function openingWorksheet(index: WorkbookIndex): number {
  const ok = (at: number) => {
    const s = index.sheets[at];
    return !!s && s.state === 'visible' && s.kind === 'worksheet';
  };
  if (ok(index.activeTab)) return index.activeTab;
  return index.sheets.findIndex((_, at) => ok(at));
}

/**
 * The worksheets of a batch, indexed once, so that each lookup is one map read. The directory
 * lives for one batch, and it keeps the worksheet that each selector name found, so a name is
 * looked up once for each batch however many operations select it. It folds no name: a name
 * that one worksheet holds exactly needs no fold, and the store keeps the folded names of each
 * workbook index (`sheetsByKey`), so a long name is folded once for each index, not each batch.
 */
interface WorksheetDirectory {
  readonly index: WorkbookIndex;
  readonly reads: readonly AutomationWorksheetRead[];
  readonly byId: ReadonlyMap<string, SheetEntry>;
  /** The worksheets of each exact name. */
  readonly byName: ReadonlyMap<string, readonly SheetEntry[]>;
  readonly active: SheetEntry | undefined;
  /** The worksheet that each selector name found, or undefined when it found none or two. */
  readonly found: Map<string, SheetEntry | undefined>;
}

function push<K, V>(map: Map<K, V[]>, key: K, value: V): void {
  const list = map.get(key);
  if (list) list.push(value);
  else map.set(key, [value]);
}

function worksheetDirectory(index: WorkbookIndex): WorksheetDirectory {
  const reads: AutomationWorksheetRead[] = [];
  const byId = new Map<string, SheetEntry>();
  const byName = new Map<string, SheetEntry[]>();
  for (const entry of index.sheets) {
    if (entry.kind !== 'worksheet' || entry.partName === undefined) continue;
    reads.push(Object.freeze({ id: entry.partName, name: entry.name }));
    if (!byId.has(entry.partName)) byId.set(entry.partName, entry);
    push(byName, entry.name, entry);
  }
  const active = index.sheets[openingWorksheet(index)];
  return {
    index,
    reads: Object.freeze(reads),
    byId,
    byName,
    active: active?.kind === 'worksheet' && active.partName !== undefined ? active : undefined,
    found: new Map(),
  };
}

const isWorksheet = (entry: SheetEntry): boolean =>
  entry.kind === 'worksheet' && entry.partName !== undefined;

/** The worksheet that `selector` names, or a refusal. */
function worksheetOf(directory: WorksheetDirectory, selector: unknown): SheetEntry {
  const s = (typeof selector === 'object' && selector !== null ? selector : {}) as {
    id?: unknown;
    name?: unknown;
    active?: unknown;
  };
  if (Object.keys(s).length === 1) {
    if (typeof s.id === 'string')
      return (
        directory.byId.get(s.id) ??
        refuseWith('item-not-found', 'the workbook no longer holds the worksheet', 'gone')
      );
    if (typeof s.name === 'string') {
      const name = s.name;
      const found = memoized(directory.found, name, () => {
        hostWork.names += 1;
        // As the store's `sheetByName`: the one exact match, else the one match without case.
        // The one exact match is also the store's answer when other names match without case.
        const exact = directory.byName.get(name) ?? [];
        if (exact.length === 1) return exact[0];
        hostWork.folds += 1;
        const matches = sheetsByKey(directory.index, name).filter(isWorksheet);
        const list = matches.length === 1 ? matches : exact;
        return list.length === 1 ? list[0] : undefined;
      });
      return (
        found ?? refuseWith('item-not-found', 'the workbook has no worksheet of that name', 'name')
      );
    }
    if (s.active === true)
      return (
        directory.active ??
        refuseWith('item-not-found', 'the workbook has no active worksheet', 'no-active')
      );
  }
  return refuseWith('invalid-argument', 'a selector has one of id, name, or active', 'sheet');
}

function boundsOf(address: unknown): RangeBounds {
  const parsed = parseRangeAddress(address);
  if (!parsed.ok) refuseWith(parsed.code, 'the address is not in the subset', 'address');
  if (boundsCells(parsed.bounds) > MAX_RANGE_CELLS)
    refuseWith('resource-limit', 'the range has too many cells', 'range-cells');
  return parsed.bounds;
}

/**
 * The sheet name of an address, with its `!`: quoted only when formula text needs the quotes. A
 * name that reads as a cell reference, an R1C1 reference, or a boolean is quoted. Each pattern
 * is linear.
 */
function addressSheetText(name: string): string {
  const plain =
    /^[\p{L}_][\p{L}\p{N}_]*$/u.test(name) &&
    !/^[A-Za-z]{1,3}[0-9]+$/.test(name) &&
    !/^[Rr][0-9]*(?:[Cc][0-9]*)?$/.test(name) &&
    !/^[Cc][0-9]*$/.test(name) &&
    !/^(?:true|false)$/i.test(name);
  return plain ? `${name}!` : `'${name.replace(/'/g, "''")}'!`;
}

/** The address of `bounds` on `entry`, with the sheet name text made once for each batch. */
function addressText(entry: SheetEntry, bounds: RangeBounds, memo: ReadMemo): string {
  const first = formatCellRef({ row: bounds.top, col: bounds.left });
  const last = formatCellRef({ row: bounds.bottom, col: bounds.right });
  const sheet = memoized(memo.sheets, entry, () => {
    hostWork.sheets += 1;
    return addressSheetText(entry.name);
  });
  return `${sheet}${first === last ? first : `${first}:${last}`}`;
}

/** The worksheet part of `entry` with its sheet index, or a refusal. */
function worksheetPart(pkg: XlsxPackage, entry: SheetEntry) {
  const part = xmlPartOf(pkg, entry.partName);
  const sheet = part ? sheetIndex(part) : undefined;
  // A cell that the index cannot place would read as an empty cell, and so would each cell of a
  // worksheet whose `x:sheetData` is not in place, such as one inside a wrapper. The store
  // refuses each write to both.
  if (!part || !sheet || sheet.problems.length > 0 || !sheet.sheetData)
    refuseWith('unsupported', 'the worksheet cannot be read', 'unreadable-sheet');
  return { part, sheet };
}

/** Refuse values that meet a formula range, or a worksheet whose formula ranges are unknown. */
function checkArrayRanges(part: XmlPart, bounds: RangeBounds): void {
  const count = arrayRangeCount(part);
  if (count.unreadable > 0 || count.all > MAX_ARRAY_RANGES)
    refuseWith('unsupported', 'the formula ranges of the worksheet are unknown', 'formula-ranges');
  const range = arrayFormulaIn(part, bounds.top, bounds.left, bounds.bottom, bounds.right);
  if (range !== undefined) {
    const ref = firstCell(range, bounds);
    refuseWith('unsupported', 'a cell holds a formula result', 'formula-cell', ref);
  }
}

/**
 * The first cell of `bounds` in an array or data table range that `arrayFormulaIn` gives, such
 * as `B2`. The range starts at its first cell, which the reader made with `formatCellRef`.
 */
function firstCell(range: string, bounds: RangeBounds): string | undefined {
  const at = parseCellRef(range.split(':')[0]!);
  if (!at) return undefined;
  return formatCellRef({ row: Math.max(at.row, bounds.top), col: Math.max(at.col, bounds.left) });
}

/** A stored value as the object model gives it. */
function plainValue(value: CellValue, index: WorkbookIndex, cell: string): AutomationCellValue {
  switch (value.kind) {
    case 'empty':
      return '';
    case 'number':
    case 'boolean':
      return value.value;
    case 'string':
      return value.text;
    case 'error':
      // A stored error code that is no error value reads as shortened text: refused instead.
      return isErrorValue(value.code)
        ? value.code
        : refuseWith('unsupported', 'a stored value cannot be read', 'unreadable-value', cell);
    case 'date':
      return (
        dateToSerial(value.iso, index.date1904) ??
        refuseWith('unsupported', 'a stored date cannot be read', 'unreadable-value', cell)
      );
  }
}

/** The cascade problems that leave the number format of a cell unknown. */
const FORMAT_PROBLEMS: ReadonlySet<string> = new Set([
  'cell-format-out-of-range',
  'number-format-id-malformed',
]);

/**
 * The number format code of the cell format index text `s`: the stored code, or the ECMA-376 code
 * of a built-in id (Part 1, 18.8.30). A format with no such code is refused, with the cell at
 * `row`, `col`: a cell format index out of range, a malformed or undefined id, or a built-in id
 * whose code depends on the locale.
 */
function formatAt(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  memo: ReadMemo,
  s: string | undefined,
  row: number,
  col: number
): string {
  const code = memoized(memo.codes, s, () => {
    hostWork.codes += 1;
    const { format, problems } = cellFormat(pkg, index, s);
    return problems.some((problem) => FORMAT_PROBLEMS.has(problem.code))
      ? undefined
      : format.formatCode === undefined && format.numFmtId !== 0
        ? builtinFormatCode(format.numFmtId, EN_US)
        : formatCodeOf(format, EN_US);
  });
  if (code === undefined) {
    const ref = formatCellRef({ row, col });
    refuseWith('unsupported', 'a number format cannot be read', 'unreadable-format', ref);
  }
  return code;
}

/** The value of the existing cell `el` at `ref`, or a refusal when the read cannot give it. */
function valueOf(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  el: Element,
  ref: string
): AutomationCellValue {
  hostWork.values += 1;
  if (el.children.length > MAX_CELL_CHILDREN)
    refuseWith('resource-limit', 'a cell holds too many elements', 'cell-children', ref);
  if (holdsWrappedValue(el))
    refuseWith('unsupported', 'a stored value cannot be read', 'unreadable-value', ref);
  const resolved = cellValue(pkg, index, el, false);
  if (resolved.formula)
    refuseWith('unsupported', 'a cell holds a formula result', 'formula-cell', ref);
  if (resolved.problem)
    refuseWith('unsupported', 'a stored value cannot be read', 'unreadable-value', ref);
  if (stringItemHidesText(cellStringItem(pkg, index, el)))
    refuseWith('unsupported', 'a stored value cannot be read', 'unreadable-value', ref);
  return plainValue(resolved.value, index, ref);
}

function readRange(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  entry: SheetEntry,
  bounds: RangeBounds,
  want: { readonly values: boolean; readonly numberFormat: boolean },
  memo: ReadMemo
): AutomationRangeRead {
  const address = addressText(entry, bounds, memo);
  if (!want.values && !want.numberFormat) return Object.freeze({ address });
  const { part, sheet } = worksheetPart(pkg, entry);
  if (want.values) checkArrayRanges(part, bounds);
  const values: AutomationCellValue[][] = [];
  const formats: string[][] = [];
  for (let row = bounds.top; row <= bounds.bottom; row += 1) {
    const valueRow: AutomationCellValue[] = [];
    const formatRow: string[] = [];
    // The row format of the empty cells, read once for each row element in the batch.
    const rowEl = want.numberFormat ? sheet.rows.get(row) : undefined;
    const rowStyle = rowEl
      ? memoized(memo.rows, rowEl, () => {
          hostWork.rows += 1;
          return rowFormat(rowEl);
        })
      : undefined;
    for (let col = bounds.left; col <= bounds.right; col += 1) {
      const el = sheet.cells.get(cellKey(row, col))?.el;
      if (want.values)
        // Only a cell that exists can refuse a read, so an empty cell makes no reference.
        valueRow.push(
          el
            ? memoized(memo.values, el, () => valueOf(pkg, index, el, formatCellRef({ row, col })))
            : ''
        );
      if (want.numberFormat) {
        const s = el
          ? memoized(memo.styles, el, () => {
              hostWork.styles += 1;
              return attr(el, 's');
            })
          : String(rowStyle ?? columnFormatAt(sheet.root, col) ?? 0);
        formatRow.push(formatAt(pkg, index, memo, s, row, col));
      }
    }
    if (want.values) values.push(valueRow);
    if (want.numberFormat) formats.push(formatRow);
  }
  return Object.freeze({
    address,
    ...(want.values ? { values } : {}),
    ...(want.numberFormat ? { numberFormat: formats } : {}),
  });
}

/** One cell write of the protocol, checked. */
function cellWrite(value: unknown): AutomationCellWrite {
  if (value === null) return null;
  const v = (typeof value === 'object' ? value : {}) as { type?: unknown; value?: unknown };
  if (v.type === 'number' && typeof v.value === 'number' && Number.isFinite(v.value))
    return { type: 'number', value: v.value };
  if (v.type === 'string' && typeof v.value === 'string') return { type: 'string', value: v.value };
  if (v.type === 'boolean' && typeof v.value === 'boolean')
    return { type: 'boolean', value: v.value };
  if (v.type === 'empty') return { type: 'empty' };
  return refuseWith('invalid-argument', 'a cell write this protocol does not define', 'values');
}

/** Whether a cell write stores a value: not `null` (leave the cell) and not an empty cell. */
function storesValue(value: unknown): boolean {
  const write = cellWrite(value);
  return write !== null && write.type !== 'empty';
}

/**
 * Refuse a write that stores a value in a cell that a merged range covers and does not start:
 * the merge hides that cell. The paste of the editor refuses a range that meets a merge
 * (`merged-range`), and the merge op refuses to hide a value (`merge-drops-values`). One pass
 * over the grid makes a table of counts, so each merge costs one lookup, whatever its size.
 */
function checkMergedCells(
  merges: readonly RangeBox[],
  bounds: RangeBounds,
  grid: readonly (readonly unknown[])[]
): void {
  const rows = bounds.bottom - bounds.top + 1;
  const cols = bounds.right - bounds.left + 1;
  const width = cols + 1;
  // `sum[r * width + c]`: the cells that store a value in the rows before r and the columns
  // before c of the grid.
  const sum = new Uint32Array((rows + 1) * width);
  for (let r = 0; r < rows; r += 1)
    for (let c = 0; c < cols; c += 1)
      sum[(r + 1) * width + c + 1] =
        (storesValue(grid[r]![c]) ? 1 : 0) +
        sum[r * width + c + 1]! +
        sum[(r + 1) * width + c]! -
        sum[r * width + c]!;
  const at = (r: number, c: number) => sum[r * width + c]!;
  for (const [mergeTop, mergeLeft, mergeBottom, mergeRight] of merges) {
    const top = Math.max(mergeTop, bounds.top) - bounds.top;
    const left = Math.max(mergeLeft, bounds.left) - bounds.left;
    const bottom = Math.min(mergeBottom, bounds.bottom) - bounds.top;
    const right = Math.min(mergeRight, bounds.right) - bounds.left;
    const anchor = mergeTop === bounds.top + top && mergeLeft === bounds.left + left;
    const stored =
      at(bottom + 1, right + 1) - at(top, right + 1) - at(bottom + 1, left) + at(top, left);
    if (stored - (anchor && storesValue(grid[top]![left]) ? 1 : 0) === 0) continue;
    for (let r = top; r <= bottom; r += 1)
      for (let c = left; c <= right; c += 1)
        if (!(anchor && r === top && c === left) && storesValue(grid[r]![c])) {
          const ref = formatCellRef({ row: bounds.top + r, col: bounds.left + c });
          refuseWith('unsupported', 'a merged range hides the cell', 'merged-cell', ref);
        }
  }
}

function writeOps(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  entry: SheetEntry,
  bounds: RangeBounds,
  grid: unknown,
  budget: Budget
): WorkbookOp[] {
  const rows = bounds.bottom - bounds.top + 1;
  const cols = bounds.right - bounds.left + 1;
  if (
    !Array.isArray(grid) ||
    grid.length !== rows ||
    grid.some((r) => !Array.isArray(r) || r.length !== cols)
  )
    refuseWith('invalid-argument', 'values need the rows and columns of the range', 'values');
  const { part, sheet } = worksheetPart(pkg, entry);
  checkArrayRanges(part, bounds);
  // Each write compares its bounds with each merged range, so the count of them is capped, for
  // one write and for the batch.
  const mergeCount = mergedRangeCount(part.root);
  if (mergeCount > MAX_MERGES)
    refuseWith('resource-limit', 'the worksheet has too many merged ranges', 'merged-ranges');
  budget.mergeChecks += mergeCount;
  if (budget.mergeChecks > MAX_BATCH_MERGE_CHECKS)
    refuseWith(
      'resource-limit',
      'the writes of the batch meet too many merged ranges',
      'merge-checks'
    );
  const box: RangeBox = [bounds.top, bounds.left, bounds.bottom, bounds.right];
  const merges = mergedBoxesMeeting(part.root, box);
  if (merges.length > 0) checkMergedCells(merges, bounds, grid as unknown[][]);
  const ops: WorkbookOp[] = [];
  const guarded = memoized(budget.reads.guarded, part.root, () => {
    hostWork.guarded += 1;
    return readSheetProtection(part.root) !== null;
  });
  for (let r = 0; r < rows; r += 1)
    for (let c = 0; c < cols; c += 1) {
      const write = cellWrite((grid as unknown[][])[r]![c]);
      if (write === null) continue;
      const at = { row: bounds.top + r, col: bounds.left + c };
      const ref = formatCellRef(at);
      const cell = sheet.cells.get(cellKey(at.row, at.col));
      if (cell && cell.el.children.length > MAX_CELL_CHILDREN)
        refuseWith('resource-limit', 'a cell holds too many elements', 'cell-children', ref);
      if (cell && cellHasFormula(cell.el))
        refuseWith('unsupported', 'a cell holds a formula', 'formula-cell', ref);
      // The store applies a run of writes in one pass, and a write that the pass cannot take
      // sends the run to one pass for each write. These checks refuse such a write here.
      if (write.type === 'string' && write.value.length > MAX_CELL_TEXT)
        refuseWith('transaction-refused', 'the text is too long for a cell', 'text-too-long', ref);
      const blocker = cell && cellChildBlocker(entry.name, ref, cell.el);
      if (blocker)
        refuseWith('transaction-refused', 'the store refused the edit', blocker.code, ref);
      if (guarded && cellLocked(pkg, index, sheet, at.row, at.col, cell?.el, budget.reads))
        refuseWith('transaction-refused', 'the cell is locked', 'locked', ref);
      ops.push({ op: 'setCellValue', sheet: entry.name, ref, value: write });
    }
  return ops;
}

/**
 * Whether a write of one value to the cell at 0-based `row`, `col` changes a locked cell: the
 * rule of the store's `lockedCellIn` for one cell. The cell that exists gives its format (`@s`);
 * else the format that a new cell takes, from its row or column. Each format text, row format,
 * and lock is read once for each batch, so a write costs map reads, never a new scan of the
 * attributes of a cell, a row, or an `x:protection`.
 */
function cellLocked(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  sheet: SheetIndex,
  row: number,
  col: number,
  el: Element | undefined,
  memo: ReadMemo
): boolean {
  let s: number;
  if (el) {
    const text = memoized(memo.styles, el, () => {
      hostWork.styles += 1;
      return attr(el, 's');
    });
    s = formatIndex(text) ?? 0;
  } else {
    const rowEl = sheet.rows.get(row);
    const rowStyle = rowEl
      ? memoized(memo.rows, rowEl, () => {
          hostWork.rows += 1;
          return rowFormat(rowEl);
        })
      : undefined;
    // As the store's `inheritedFormat`. The lock reader reads an index past the cell formats as
    // format 0, as `inheritedFormat` does, so no bound is needed here.
    s = rowStyle ?? columnFormatAt(sheet.root, col) ?? 0;
  }
  return memoized(memo.locks, s, () => {
    hostWork.locks += 1;
    return (memo.locker ??= lockReader(pkg, index)).locked(s);
  });
}

/** The text that a write op stores, as a list of none or one. */
function textOf(op: WorkbookOp): string[] {
  return op.op === 'setCellValue' && op.value.type === 'string' ? [op.value.value] : [];
}

/** The counters of one batch, and what its reads resolved. */
interface Budget {
  cells: number;
  writeCells: number;
  mergeChecks: number;
  listed: number;
  readonly reads: ReadMemo;
}

function spend(budget: Budget, bounds: RangeBounds, write: boolean): void {
  budget.cells += boundsCells(bounds);
  if (budget.cells > MAX_BATCH_CELLS)
    refuseWith('resource-limit', 'the batch reaches too many cells', 'batch-cells');
  if (!write) return;
  budget.writeCells += boundsCells(bounds);
  if (budget.writeCells > MAX_BATCH_WRITE_CELLS)
    refuseWith('resource-limit', 'the writes of the batch reach too many cells', 'write-cells');
}

/** What one operation gives: a value, or the ops of a write. */
type Planned = { readonly value: AutomationValue; readonly ops?: readonly WorkbookOp[] };

function plan(
  operation: AutomationOperation,
  pkg: XlsxPackage,
  index: WorkbookIndex,
  directory: WorksheetDirectory,
  budget: Budget,
  written: boolean
): Planned {
  switch ((operation as { op?: unknown } | null)?.op) {
    case 'listWorksheets':
      budget.listed += directory.reads.length;
      if (budget.listed > MAX_BATCH_LISTED_WORKSHEETS)
        refuseWith('resource-limit', 'the lists give too many worksheets', 'batch-worksheets');
      return { value: directory.reads };
    case 'getWorksheet': {
      const o = operation as Extract<AutomationOperation, { op: 'getWorksheet' }>;
      try {
        const entry = worksheetOf(directory, o.sheet);
        return { value: Object.freeze({ id: entry.partName!, name: entry.name }) };
      } catch (error) {
        if (o.orNull === true && error instanceof Refusal && error.error.code === 'item-not-found')
          return { value: null };
        throw error;
      }
    }
    case 'readRange': {
      const o = operation as Extract<AutomationOperation, { op: 'readRange' }>;
      const entry = worksheetOf(directory, o.sheet);
      const bounds = boundsOf(o.address);
      const want = { values: o.values === true, numberFormat: o.numberFormat === true };
      if (want.values || want.numberFormat) {
        if (written)
          refuseWith('unsupported', 'a read follows a write in the batch', 'read-after-write');
        spend(budget, bounds, false);
      }
      return { value: readRange(pkg, index, entry, bounds, want, budget.reads) };
    }
    case 'writeValues': {
      const o = operation as Extract<AutomationOperation, { op: 'writeValues' }>;
      // Two `x:sheet` entries can name one worksheet part. The writes to the part go to the store
      // under the name of its first entry, so they keep their order and form one run.
      const named = worksheetOf(directory, o.sheet);
      const entry = directory.byId.get(named.partName ?? '') ?? named;
      const bounds = boundsOf(o.address);
      spend(budget, bounds, true);
      return { value: null, ops: writeOps(pkg, index, entry, bounds, o.values, budget) };
    }
    default:
      return refuseWith('unknown-operation', 'this protocol does not define the operation');
  }
}

/**
 * The ops of `writes`, grouped by worksheet in the order that each worksheet first comes, and in
 * their own order within a worksheet. The planner names each worksheet part by one sheet name,
 * so a group is one part.
 */
function bySheet(writes: readonly { readonly ops: readonly WorkbookOp[] }[]): WorkbookOp[] {
  const groups = new Map<string, WorkbookOp[]>();
  for (const write of writes)
    for (const op of write.ops) push(groups, 'sheet' in op ? op.sheet : '', op);
  return [...groups.values()].flat();
}

/**
 * The index of the first write of `writes` that the store refuses. A prefix of the writes that
 * the store refuses stays refused with more writes after it, and a write to another worksheet
 * does not change whether a write is refused, so a binary search over the grouped writes finds
 * the first one in about log2(n) applies, each one pass over each changed worksheet.
 */
function firstRefusedWrite(
  pkg: XlsxPackage,
  writes: readonly { readonly ops: readonly WorkbookOp[] }[]
): number {
  let low = 0;
  let high = writes.length - 1;
  while (low < high) {
    const middle = (low + high) >> 1;
    hostWork.searches += 1;
    if (applyOps(pkg, bySheet(writes.slice(0, middle + 1))).ok) low = middle + 1;
    else high = middle;
  }
  return low;
}

/** The first cell, as `A1`, that the protection of its worksheet refuses a value write of. */
function lockedCell(pkg: XlsxPackage, ops: readonly WorkbookOp[]): string | undefined {
  const refusal = protectionChecks();
  for (const op of ops)
    if (op.op === 'setCellValue' && refusal(pkg, op)?.lock === 'cell') return op.ref;
  return undefined;
}

/**
 * Open XLSX bytes as a headless host. Opening is the bounded package reader and one store, as
 * the editor opens a workbook. Malformed bytes answer `invalid-argument`, and bytes over a reader
 * limit `resource-limit` with the limit. Each refusal gives its public open diagnostic. The
 * answer carries no file content.
 *
 * @public
 */
export function openAutomationHost(bytes: Uint8Array | ArrayBuffer): AutomationOpenResult {
  const read = readXlsx(bytes);
  if (read.ok)
    return { ok: true, host: createStoreAutomationHost(new WorkbookDocumentStore(read.pkg)) };
  const { code, limit } = openDiagnosticOf(read.error);
  return limit === undefined
    ? { ok: false, code: 'invalid-argument', reason: code }
    : { ok: false, code: 'resource-limit', reason: code, limit };
}

/** A headless host over a store. */
export function createStoreAutomationHost(store: WorkbookDocumentStore): AutomationHost {
  const capabilities = Object.freeze({ document: true, save: true });
  let disposed = false;

  const execute = (request: AutomationBatchRequest): AutomationBatchResponse => {
    const revision = store.revision;
    const operations: unknown = request?.operations;
    const count = Array.isArray(operations) ? operations.length : 0;
    const fail = (
      at: number,
      code: AutomationErrorCode,
      message: string,
      detail?: string,
      cell?: string
    ) => refused(Math.max(count, 1), at, automationError(code, message, detail, cell), revision);
    if (disposed) return fail(0, 'disposed', 'this host was disposed');
    if (!Array.isArray(operations))
      return fail(0, 'invalid-argument', 'a batch has an array of operations', 'operations-list');
    if (count > MAX_BATCH_OPERATIONS)
      return fail(0, 'resource-limit', 'too many operations', 'operations');
    if (count === 0) return { ok: true, results: [], revision, changed: false };
    const expected = request.expectedRevision;
    if (expected !== undefined && expected !== revision)
      return fail(0, 'stale-revision', 'the workbook changed after the caller read it');
    const pkg = store.package;
    const index = workbookIndex(pkg);
    if (!index)
      return fail(0, 'unsupported', 'the package has no workbook part', 'unreadable-sheet');
    const directory = worksheetDirectory(index);
    const budget: Budget = {
      cells: 0,
      writeCells: 0,
      mergeChecks: 0,
      listed: 0,
      reads: readMemo(),
    };
    const writes: { readonly at: number; readonly ops: readonly WorkbookOp[] }[] = [];
    const results: AutomationOperationResult[] = [];
    for (let at = 0; at < count; at += 1) {
      try {
        const planned = plan(operations[at], pkg, index, directory, budget, writes.length > 0);
        if (planned.ops && planned.ops.length > 0) writes.push({ at, ops: planned.ops });
        results.push(Object.freeze({ status: 'ok' as const, value: planned.value }));
      } catch (error) {
        if (error instanceof Refusal) return refused(count, at, error.error, revision);
        throw error;
      }
    }
    if (writes.length === 0) return { ok: true, results, revision, changed: false };
    // A shared string table that cannot take text refuses the first write of text here, for the
    // reason above.
    const texts = writes.flatMap((write) => write.ops.flatMap(textOf));
    const table = texts.length > 0 ? sharedStringIndices(pkg, index, texts) : undefined;
    if (table && 'code' in table) {
      const at = writes.find((write) => write.ops.some((op) => textOf(op).length > 0))!.at;
      return fail(at, 'transaction-refused', 'the store refused the edit', table.code);
    }
    const before = store.revision;
    const committed = store.transact(bySheet(writes), {
      label: 'automation',
      baseRevision: revision,
    });
    if (!committed.ok) {
      if (committed.rejection.code === 'stale-revision')
        return fail(writes[0]!.at, 'stale-revision', 'another writer changed the workbook first');
      const write = writes[firstRefusedWrite(pkg, writes)]!;
      const code = committed.rejection.code;
      return fail(
        write.at,
        'transaction-refused',
        'the store refused the edit',
        code,
        code === 'locked' ? lockedCell(pkg, write.ops) : undefined
      );
    }
    return { ok: true, results, revision: store.revision, changed: store.revision !== before };
  };

  return {
    capabilities,
    revision: () => store.revision,
    execute,
    save() {
      if (disposed)
        return { ok: false, error: automationError('disposed', 'this host was disposed') };
      return { ok: true, bytes: saveXlsx(store.package) };
    },
    dispose() {
      disposed = true;
    },
  };
}
