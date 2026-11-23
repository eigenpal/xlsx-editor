// The decimal steps of a selection, DOM-free. After
// docx-editor/packages/core/src/editor/surface-structure.ts at f48510f2 (EigenPal), `stepIndent`
// and `canAdjustIndent`: each item of the selection steps from its own value, an item that
// cannot step stays as it is, and the command is enabled when one item can step. Here an item
// is a cell, and its value is its number format. One step writes each new code only to the
// cells that it comes from, in one transaction, so one undo restores every cell.
//
// `General` numbers are the exception, on a selection of more cells: `General` shows each number
// with its own places, so a step of each from its own places would never line up a column. In each
// column, every `General` number takes the step of the first `General` number of that column, when
// that step succeeds. The first is the active cell in its own column, and else the first in row
// order. When that step fails, each `General` number steps from its own places. A `General` number
// in exponent form steps alone: it neither sets nor takes the step of its column. So the active
// cell and each cell with an explicit code move exactly one place, and another `General` number
// takes the code of its column: a decrease can show 7 as 7.0. One cell, or one merged range, steps
// from its own places.
//
// Intended differences: when no cell can step, the disabled reason and the refusal name the
// `DecimalSelectionProblem` of the selection (`decimal-reasons.ts`), because a cell has many ways
// not to step. And the scan has a limit: the selection holds at most `MAX_DECIMAL_STEP_CELLS` cells
// that exist, or the step is refused.
//
// The scan reads only the cells that exist inside the selection (`SheetIndex.order`), and never
// walks the grid that the selection names. In a sheet with no index problem, `order` ascends by
// `cellKey`, so binary searches jump over the cells left and right of the selection, and a whole
// column never visits an empty row. The scan stops at the first cell past the limit, so the
// state and the command apply the same limit at the same cost.
//
// A cell that holds text, a boolean, or an error gives `noNumbers`, on one cell as on more cells,
// so one such cell and a range of them disable the steps with the same reason. An empty cell has no
// value to show. On a selection of one empty cell (or of exactly one merged range), an increase on
// `General` gives `0.0`, and the step creates the cell when it has no `x:c` element. A decrease
// gives `noPlaces`, as it does for one `General` number that shows no places. An empty cell with an
// explicit code steps from its code there. In a selection of more cells, every cell with no number
// is skipped, whatever its code, the active cell too, and the step creates no cell. An empty
// position in such a selection is never written, also when its row or column format could step,
// which is intended. Among more cells, only a cell that holds a number names a reason. The reason
// is the problem of the active cell when it holds a number, else of the first selected cell that
// holds a number, or `noNumbers`. The editor uses a reason about all selected cells, `noDecimals`
// or `maxDecimals`, only when each selected cell that holds a number gives it. Else the cell that
// names the reason (the active cell, or the first cell with a number) gives `cellNoDecimals` or
// `cellMaxDecimals`, which speak of one selected cell (`Problems.reason`).

import type { GridPoint, GridSelection } from '../contracts/index.ts';
import { mergedRanges, zoomProfile, type TextMeasurer } from '../layout/index.ts';
import { shownGeneralReader } from '../layout/sheet-layout.ts';
import { mergeAnchoredAt, mergeCoverage, mergesIn } from '../layout/panes.ts';
import {
  attr,
  cellKey,
  inheritedFormat,
  keyAddress,
  formatCellRef,
  sheetIndex,
  xmlPartOf,
  type Element,
  type SheetIndex,
  type WorkbookIndex,
  type WorkbookOp,
  type XlsxPackage,
} from '../store/index.ts';
import {
  builtinFormatCode,
  cellFormat,
  cellValue,
  dateToSerial,
  formatCodeOf,
  type CellValue,
  type NumberLocale,
  type ResolvedCellFormat,
} from '../cascade/index.ts';
import { formatOps } from '../automation/index.ts';
import {
  decimalStepOf,
  generalDecimals,
  generalExponent,
  stepGeneralText,
  isGeneralCode,
  type DecimalStepProblem,
  type DecimalStepResult,
} from './number-format-commands.ts';
import type { EditRefusal } from './cell-input.ts';
import {
  decimalRefusalOf,
  MAX_DECIMAL_STEP_CELLS,
  type DecimalSelectionProblem,
} from './decimal-reasons.ts';

/** The number format of a cell format: its id, its code, and whether the code is known. */
export interface CellNumberFormat {
  readonly id: number;
  readonly code: string;
  readonly hasCode: boolean;
}

/**
 * The number format of a resolved cell format in the number locale. `hasCode` is false when
 * neither the workbook nor the locale gives the id a code: the cell then shows `General`.
 */
export function numberFormatOf(format: ResolvedCellFormat, locale: NumberLocale): CellNumberFormat {
  return {
    id: format.numFmtId,
    code: formatCodeOf(format, locale),
    hasCode:
      format.formatCode !== undefined || builtinFormatCode(format.numFmtId, locale) !== undefined,
  };
}

export { MAX_DECIMAL_STEP_CELLS, type DecimalSelectionProblem };

/**
 * The state of the two decimal steps on a selection: null when at least one cell of the
 * selection can step, or else the refusal that the step gives. The refusal is
 * `decimals-unchanged` or `decimals-too-large`, and its `reason` is the disabled reason of the
 * step's control. `toolbarCommandState` reads it: a host reads the enabled state there.
 *
 * @public
 */
export interface DecimalSteps {
  /** Null when an increase changes at least one selected cell, or the refusal of the increase. */
  readonly increase: EditRefusal | null;
  /** Null when a decrease changes at least one selected cell, or the refusal of the decrease. */
  readonly decrease: EditRefusal | null;
}

/**
 * What the scan found at the active cell, beside `DecimalSteps`. The tests read it to check the
 * step of the active cell (`__tests__/decimal-step-preview.ts`) against the ops of the command.
 * It is kept off the state object, so the snapshot holds only the public fields.
 *
 * @internal
 */
export interface DecimalStepsDetail {
  /**
   * The text that the grid shows for the `General` number of the active cell, fitted to its
   * column, as the step reads it. Null when the active cell shows no `General` number, or with
   * no text measurer: the step then reads the full text.
   */
  readonly shownText: string | null;
  /**
   * The step can write the active cell: it holds a number, or it is the only selected cell and
   * holds no text, boolean, or error.
   * Among more selected cells, the step never writes a cell with no number (an empty position,
   * text, a boolean, or an error, whatever its code), or a cell that a merged range covers.
   */
  readonly activeWritten: boolean;
  /** The number of the active cell, stored or cached from a formula, or null. */
  readonly activeNumber: number | null;
}

const NO_DETAIL: DecimalStepsDetail = Object.freeze({
  shownText: null,
  activeWritten: true,
  activeNumber: null,
});

const details = new WeakMap<DecimalSteps, DecimalStepsDetail>();

/** What the scan of `selectionDecimalSteps` found at the active cell. @internal */
export function decimalStepsDetail(steps: DecimalSteps): DecimalStepsDetail {
  return details.get(steps) ?? NO_DETAIL;
}

/** Both steps with no sheet shown, or with nothing to read. */
export const NO_DECIMAL_STEPS: DecimalSteps = Object.freeze({ increase: null, decrease: null });

/** The first index at or after `lo` whose key in `order` is at least `key`. */
function lowerBound(order: readonly number[], key: number, lo = 0): number {
  let hi = order.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (order[mid]! < key) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

/**
 * The keys of the cells that exist inside the selection, in document order, one at a time.
 * With no index problem, `order` ascends by `cellKey`: a cell left of the box jumps to the box
 * in its row, and a cell right of it jumps to the next row, so the walk costs the cells inside
 * the box and two searches for each row that holds a cell. With a problem, the order is not
 * known, and the walk reads every key of the sheet.
 */
function* boxKeys(sheet: SheetIndex, selection: GridSelection): Generator<number> {
  const { start, end } = selection.range;
  const [top, left, bottom, right] = [start.row, start.col, end.row, end.col];
  const { order } = sheet;
  if (sheet.problems.length > 0) {
    for (const key of order) {
      const { row, col } = keyAddress(key);
      if (row >= top && row <= bottom && col >= left && col <= right) yield key;
    }
    return;
  }
  const last = cellKey(bottom, right);
  let i = lowerBound(order, cellKey(top, left));
  while (i < order.length && order[i]! <= last) {
    const { row, col } = keyAddress(order[i]!);
    if (col < left) i = lowerBound(order, cellKey(row, left), i);
    else if (col > right) i = lowerBound(order, cellKey(row + 1, left), i);
    else {
      yield order[i]!;
      i += 1;
    }
  }
}

/**
 * A walk of the keys of a selection that counts them. It stops after `MAX_DECIMAL_STEP_CELLS` + 1
 * keys, so `tooLarge` is exact and the walk never reads more keys than that.
 */
class BoxWalk {
  count = 0;
  readonly #keys: Iterator<number>;

  /**
   * A walk of the keys of `selection`. On a `single` selection, the step reads the active cell
   * only (the anchor of one merged range), so the walk reads and counts no key: the cells that a
   * merged range covers never count toward the limit.
   */
  constructor(sheet: SheetIndex, selection: GridSelection, single: boolean) {
    this.#keys = single ? [].values() : boxKeys(sheet, selection);
  }

  get tooLarge(): boolean {
    return this.count > MAX_DECIMAL_STEP_CELLS;
  }

  /** The next key, or undefined at the end or past the limit. */
  next(): number | undefined {
    if (this.tooLarge) return undefined;
    const r = this.#keys.next();
    if (r.done) return undefined;
    this.count += 1;
    return this.tooLarge ? undefined : r.value;
  }

  /** Count the keys that are left, up to the limit, with no cell read. */
  drain(): void {
    while (this.next() !== undefined);
  }
}

/** The problem of one cell, as a selection names it. */
type CellProblem = DecimalStepProblem | 'noPlaces' | 'noNumbers';

/** The problems of the cells of a step that no cell can take. */
class Problems {
  /** The problem of the active cell. */
  active: CellProblem | undefined;
  /** The problem of the first other selected cell that holds a number. */
  first: DecimalStepProblem | undefined;
  /** The problem of each selected cell that holds a number, the active cell too. */
  readonly all = new Set<DecimalStepProblem>();

  /** Record the problem of a cell. Among more cells, only a cell with a number names one. */
  add(steps: CellSteps, cell: StepCell, problem: DecimalStepProblem, single: boolean): void {
    if (cell.value !== null) this.all.add(problem);
    if (cell.active && (single || cell.value !== null))
      this.active = single && cell.other ? 'noNumbers' : steps.problemOf(cell, problem);
    else if (cell.value !== null) this.first ??= problem;
  }

  /**
   * The reason of a step that no cell can take: the problem of the active cell (`noNumber` when it
   * holds no number; on one cell, `noNumbers` when it holds text, a boolean, or an error), else the
   * problem of the first selected cell that holds a number. Among more cells, `noDecimals` and
   * `maxDecimals` speak of all selected cells: they apply only when every selected cell that holds
   * a number gives them, and else give `cellNoDecimals` or `cellMaxDecimals`, which speak of that
   * one cell.
   */
  reason(single: boolean): DecimalSelectionProblem {
    const p = this.active ?? 'noNumber';
    if (single) return p === 'noNumber' ? 'noPlaces' : p;
    // A cell with a number never gives `noNumber`; the check keeps the type exact.
    const base =
      p !== 'noNumber'
        ? p
        : this.first === undefined || this.first === 'noNumber'
          ? 'noNumbers'
          : this.first;
    if (base !== 'noDecimals' && base !== 'maxDecimals') return base;
    if (this.all.size === 1) return base;
    return base === 'noDecimals' ? 'cellNoDecimals' : 'cellMaxDecimals';
  }
}

/** One cell that a step reads: its place, its cell format index, and its number. */
interface StepCell {
  readonly row: number;
  readonly col: number;
  readonly s: string | undefined;
  readonly value: number | null;
  readonly active: boolean;
  /** The cell has an `x:c` element. */
  readonly exists: boolean;
  /** The cell holds a value that is not a number: text, a boolean, or an error. */
  readonly other: boolean;
}

/** The cell of a position, as the scan reads it. */
function stepCell(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  sheet: SheetIndex,
  row: number,
  col: number,
  active: boolean
): StepCell {
  const entry = sheet.cells.get(cellKey(row, col));
  if (!entry) {
    const s = String(inheritedFormat(sheet, row, col));
    return { row, col, s, value: null, active, exists: false, other: false };
  }
  const value = cellValue(pkg, index, entry.el, false).value;
  const number = numberOf(value, index);
  return {
    row,
    col,
    s: attr(entry.el, 's'),
    value: number,
    active,
    exists: true,
    other: number === null && value.kind !== 'empty',
  };
}

/**
 * Whether a cell lies inside a merged range of the selection but is not its top-left cell: the
 * grid shows only the anchor of a merge, so a step reads and writes the anchor only. The test
 * reads the merges that meet the selection, one row at a time (`mergeCoverage`, which starts
 * again by itself when the rows go back).
 */
function coveredBy(root: Element, { range: { start, end } }: GridSelection) {
  const merges = mergedRanges(root);
  const area = { firstRow: start.row, lastRow: end.row, firstCol: start.col, lastCol: end.col };
  const inside = mergesIn(merges, area);
  if (inside.length === 0) return () => false;
  const test = mergeCoverage(inside);
  return (row: number, col: number): boolean =>
    test(row, col) && !mergeAnchoredAt(merges, row, col);
}

/**
 * The cells that a step reads: the active cell first, then each other cell of the selection
 * that exists, from `walk`, in document order, with no cell that a merged range covers. Among
 * more cells, that holds for the active cell too: the step then never writes it, and
 * `DecimalSteps.activeWritten` is false. Each cell is read when the caller asks for it.
 */
function* stepCells(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  sheet: SheetIndex,
  selection: GridSelection,
  walk: BoxWalk,
  single: boolean
): Generator<StepCell> {
  const { active } = selection;
  const covered = coveredBy(sheet.root, selection);
  // A band selection can put the active cell under a merged range whose anchor is outside the
  // band: the grid does not show that cell, so among more cells the step skips it.
  if (single || !covered(active.row, active.col))
    yield stepCell(pkg, index, sheet, active.row, active.col, true);
  for (let key = walk.next(); key !== undefined; key = walk.next()) {
    const { row, col } = keyAddress(key);
    if ((row === active.row && col === active.col) || covered(row, col)) continue;
    yield stepCell(pkg, index, sheet, row, col, false);
  }
}

/**
 * The text that the grid shows for a `General` number at a cell, fitted to its column
 * (`shownGeneralReader` of the layout), or undefined. With no reader, as with no text
 * measurer, `General` steps from the text that it shows at full width.
 */
export type ShownGeneralText = (row: number, col: number) => string | undefined;

/** The number of a cell: a stored number or the cached number of a formula, or null. */
export function numberAt(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  partName: string | undefined,
  at: GridPoint
): number | null {
  const part = xmlPartOf(pkg, partName);
  const entry = part ? sheetIndex(part).cells.get(cellKey(at.row, at.col)) : undefined;
  return entry ? numberOf(cellValue(pkg, index, entry.el, false).value, index) : null;
}

/**
 * The number that the grid formats for a value: a number, or the serial of a date cell
 * (`t="d"`), as the grid shows it. Else null.
 */
function numberOf(value: CellValue, index: WorkbookIndex): number | null {
  if (value.kind === 'number') return value.value;
  if (value.kind === 'date') return dateToSerial(value.iso, index.date1904) ?? null;
  return null;
}

/**
 * The reader of the text that the grid shows for a `General` number on the active sheet, at
 * the zoom that the user sees, or undefined with no workbook or no text measurer. The decimal
 * steps read it, so that they step from the shown places.
 */
export function shownGeneralOf(
  opened: { readonly pkg: XlsxPackage; readonly index: WorkbookIndex } | null,
  active: number,
  measurer: TextMeasurer | null,
  zoom: number | undefined,
  locale: NumberLocale
): ShownGeneralText | undefined {
  const name = opened?.index.sheets[active]?.name;
  if (!opened || !measurer || name === undefined || zoom === undefined) return undefined;
  return shownGeneralReader(opened.pkg, name, { measurer, profile: zoomProfile(zoom), locale });
}

/**
 * The step of each cell, with one parse for each cell format and, for `General`, for each
 * shown text. `General` steps from the text that the grid shows for the number of the cell
 * (`shown`), or else from the full text. Among more cells, a cell with no number (text, a
 * boolean, an error, or nothing) gives `noNumber`, whatever its code. On a selection of one
 * cell, a cell with text, a boolean, or an error gives `noNumber` too, and an increase on an empty
 * `General` cell gives `0.0`.
 */
class CellSteps {
  readonly #formats = new Map<string, CellNumberFormat & { readonly general: boolean }>();
  readonly #steps = new Map<string, DecimalStepResult>();
  /**
   * The step of the first `General` number in each column of a selection of more cells, by
   * column and direction (`2 * col`, plus 1 for an increase). The selection holds at most
   * `MAX_DECIMAL_STEP_CELLS` cells, so the map holds at most two entries for each cell.
   */
  readonly #shared = new Map<number, DecimalStepResult>();

  constructor(
    private readonly pkg: XlsxPackage,
    private readonly index: WorkbookIndex,
    private readonly locale: NumberLocale,
    private readonly single: boolean,
    private readonly shown: ShownGeneralText | undefined
  ) {}

  /** The shown text of a `General` cell at a place, as `of` reads it, or null. */
  shownAt(at: { readonly row: number; readonly col: number }): string | null {
    return this.shown?.(at.row, at.col) ?? null;
  }

  of(cell: StepCell, step: 1 | -1): DecimalStepResult {
    const s = cell.s ?? '';
    let format = this.#formats.get(s);
    if (!format) {
      const read = numberFormatOf(cellFormat(this.pkg, this.index, cell.s).format, this.locale);
      format = { ...read, general: read.hasCode && isGeneralCode(read.code) };
      this.#formats.set(s, format);
    }
    // Among more cells, a cell with no number keeps its format, whatever its code. On one cell,
    // a cell with text, a boolean, or an error keeps it too: only an empty cell takes a code.
    if (cell.value === null && (!this.single || cell.other))
      return { ok: false, problem: 'noNumber' };
    if (format.general && cell.value !== null) {
      // On more cells, every `General` number of a column takes the step of the first `General`
      // number of that column (the active cell first, then row order), so the column lines up.
      // When that step fails, each one steps from its own places. A number in exponent form
      // steps alone: it neither sets nor takes the step of its column.
      const own = this.#general(cell, cell.value, format, s, step);
      if (this.single || own.scientific) return own.result;
      const key = 2 * cell.col + (step > 0 ? 1 : 0);
      const shared = this.#shared.get(key);
      if (shared === undefined) this.#shared.set(key, own.result);
      return shared?.ok ? shared : own.result;
    }
    const key = `${step}|${s}`;
    let result = this.#steps.get(key);
    if (!result) this.#steps.set(key, (result = decimalStepOf(format, step, cell.value)));
    return result;
  }

  /**
   * The step of one `General` number, from the text that the grid shows for it, and whether
   * that text is in exponent form.
   */
  #general(
    cell: StepCell,
    value: number,
    format: CellNumberFormat,
    s: string,
    step: 1 | -1
  ): { readonly result: DecimalStepResult; readonly scientific: boolean } {
    // The text that the grid shows decides the places and the exponent form. A text that shows
    // no number, such as `#`, and a host with no reader step from the full text.
    const text = this.shown?.(cell.row, cell.col);
    const fitted =
      text === undefined ? undefined : stepGeneralText(step, text, this.locale.decimal);
    if (fitted) return { result: fitted, scientific: text!.includes('E') };
    const exponent = generalExponent(value);
    const key = `${step}|${s}|${generalDecimals(value)}|${exponent}`;
    let result = this.#steps.get(key);
    if (!result) this.#steps.set(key, (result = decimalStepOf(format, step, value)));
    return { result, scientific: exponent };
  }

  /**
   * The problem of the only selected cell: `noPlaces` for a `General` number that shows no
   * places, so that the reason speaks of the cell and not of a number format.
   */
  problemOf(cell: StepCell, problem: DecimalStepProblem): CellProblem {
    const general = this.#formats.get(cell.s ?? '')?.general === true;
    return this.single && general && cell.value !== null && problem === 'noDecimals'
      ? 'noPlaces'
      : problem;
  }
}

/**
 * Whether a selection is one cell, or exactly one merged range, which shows as one cell: the
 * only case in which a step creates a cell.
 */
/** Whether the selection is one cell, or exactly one merged range of the sheet `root`. */
export function isSingleSelection(
  root: Element | undefined,
  { range: { start, end } }: GridSelection
): boolean {
  if (start.row === end.row && start.col === end.col) return true;
  if (root === undefined) return false;
  return mergedRanges(root).ranges.some(
    (m) =>
      m.firstRow === start.row &&
      m.firstCol === start.col &&
      m.lastRow === end.row &&
      m.lastCol === end.col
  );
}

/** What the scan of one direction found. */
interface Direction {
  readonly step: 1 | -1;
  /** A cell can step. */
  ok: boolean;
  /** The problems of the cells that cannot step. */
  readonly problems: Problems;
}

/**
 * The state of the two decimal steps on a selection. Each step is enabled when one cell of the
 * selection can step; when none can, it gives the reason of `DecimalSelectionProblem`. The cell
 * reads stop at the first cell that can step in each direction, so the common case reads the
 * active cell only. The key walk then counts the keys that are left, with no cell read, up to
 * `MAX_DECIMAL_STEP_CELLS` + 1, so the limit is exact.
 */
export function selectionDecimalSteps(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  partName: string | undefined,
  selection: GridSelection,
  locale: NumberLocale,
  shown?: ShownGeneralText
): DecimalSteps {
  const part = xmlPartOf(pkg, partName);
  if (!part) return NO_DECIMAL_STEPS;
  const sheet = sheetIndex(part);
  const single = isSingleSelection(part.root, selection);
  const walk = new BoxWalk(sheet, selection, single);
  const steps = new CellSteps(pkg, index, locale, single, shown);
  const direction = (step: 1 | -1): Direction => ({ step, ok: false, problems: new Problems() });
  const state = { increase: direction(1), decrease: direction(-1) };
  let activeWritten = single;
  for (const cell of stepCells(pkg, index, sheet, selection, walk, single)) {
    if (cell.active) activeWritten = (single && !cell.other) || cell.value !== null;
    let open = false;
    for (const d of [state.increase, state.decrease]) {
      if (d.ok) continue;
      const r = steps.of(cell, d.step);
      // In a selection of more cells, an empty position is never written.
      if (r.ok && (cell.exists || single)) d.ok = true;
      else if (!r.ok) d.problems.add(steps, cell, r.problem, single);
      open ||= !d.ok;
    }
    if (!open) break;
  }
  walk.drain();
  const shownText = activeWritten ? steps.shownAt(selection.active) : null;
  const activeNumber = numberAt(pkg, index, partName, selection.active);
  const refusal = (d: Direction): EditRefusal | null =>
    d.ok ? null : decimalRefusalOf(d.problems.reason(single), single, locale);
  const tooLarge = walk.tooLarge ? decimalRefusalOf('tooLarge', single, locale) : null;
  const result: DecimalSteps = {
    increase: tooLarge ?? refusal(state.increase),
    decrease: tooLarge ?? refusal(state.decrease),
  };
  details.set(result, { shownText, activeWritten, activeNumber });
  return result;
}

/** The ops of a decimal step on a selection, or the refusal. */
export type DecimalStepOps =
  | { readonly ok: true; readonly ops: readonly WorkbookOp[] }
  | { readonly ok: false; readonly refusal: EditRefusal };

/** One run of cells in a row that get one new code: columns `left` to `right`. */
interface Run {
  readonly left: number;
  readonly right: number;
  readonly code: string;
}

/**
 * The `applyFormat` ops of a decimal step on a selection: one for each block of rows with the
 * same runs of new codes. The store applies a run of them as one pass. Refused with
 * `decimals-too-large` past `MAX_DECIMAL_STEP_CELLS` cells that exist inside the selection, and
 * with `decimals-unchanged` when no cell can step. The refusal is the one that
 * `selectionDecimalSteps` gives for the direction, with its `reason`.
 */
export function decimalStepOps(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  entry: { readonly name: string; readonly partName: string | undefined },
  selection: GridSelection,
  locale: NumberLocale,
  step: 1 | -1,
  shown?: ShownGeneralText
): DecimalStepOps {
  const part = xmlPartOf(pkg, entry.partName);
  if (!part) return { ok: false, refusal: { code: 'not-editable' } };
  const sheet = sheetIndex(part);
  const single = isSingleSelection(part.root, selection);
  const walk = new BoxWalk(sheet, selection, single);
  const steps = new CellSteps(pkg, index, locale, single, shown);
  const rows = new Map<number, { col: number; code: string }[]>();
  const problems = new Problems();
  for (const cell of stepCells(pkg, index, sheet, selection, walk, single)) {
    const r = steps.of(cell, step);
    if (!r.ok) {
      // The rule of `selectionDecimalSteps`, so the refusal names the reason that the state gives.
      problems.add(steps, cell, r.problem, single);
      continue;
    }
    // In a selection of more cells, an empty position is never written.
    if (!cell.exists && !single) continue;
    let row = rows.get(cell.row);
    if (!row) rows.set(cell.row, (row = []));
    row.push({ col: cell.col, code: r.code });
  }
  if (walk.tooLarge) return { ok: false, refusal: decimalRefusalOf('tooLarge', single, locale) };
  if (rows.size === 0)
    return { ok: false, refusal: decimalRefusalOf(problems.reason(single), single, locale) };
  const ops: WorkbookOp[] = [];
  for (const b of runBlocks(rows)) {
    const a = formatCellRef({ row: b.top, col: b.run.left });
    const z = formatCellRef({ row: b.bottom, col: b.run.right });
    const r = formatOps(entry.name, a === z ? a : `${a}:${z}`, {
      numberFormat: { code: b.run.code },
    });
    if (!r.ok) return { ok: false, refusal: { code: 'unreadable-number-format' } };
    ops.push(...r.ops);
  }
  return { ok: true, ops };
}

/** A run of cells over rows `top` to `bottom`. */
interface Block {
  top: number;
  bottom: number;
  readonly run: Run;
}

/**
 * The runs of each row, joined over the next rows that have the same run: a column of cells
 * with one new code is one block. At most one block for each cell.
 */
function runBlocks(rows: ReadonlyMap<number, readonly { col: number; code: string }[]>): Block[] {
  const blocks: Block[] = [];
  // The blocks that end at the row before, by run.
  let open = new Map<string, Block>();
  for (const row of [...rows.keys()].sort((a, b) => a - b)) {
    const cells = [...rows.get(row)!].sort((a, b) => a.col - b.col);
    const next = new Map<string, Block>();
    let k = 0;
    while (k < cells.length) {
      const first = cells[k]!;
      let last = first;
      while (
        k + 1 < cells.length &&
        cells[k + 1]!.col === last.col + 1 &&
        cells[k + 1]!.code === first.code
      )
        last = cells[++k]!;
      k += 1;
      const run: Run = { left: first.col, right: last.col, code: first.code };
      const key = `${run.left}:${run.right}:${run.code}`;
      const before = open.get(key);
      if (before && before.bottom === row - 1) {
        before.bottom = row;
        next.set(key, before);
      } else {
        const block: Block = { top: row, bottom: row, run };
        blocks.push(block);
        next.set(key, block);
      }
    }
    open = next;
  }
  return blocks;
}
