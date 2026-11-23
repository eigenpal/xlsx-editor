// Why a decimal step changes no cell: the problems of a selection, the disabled reason key of
// each, and the refusal that carries the reason. The selection state (`DecimalSteps`) and the
// command give the same refusal, so the disabled control and the refusal text say the same
// words. After docx-editor/packages/core/src/editor/selected-chart.ts at 708f28b8 (EigenPal),
// whose view state names the disabled reason of its controls: a host shows the reason and
// never maps a problem to a key.

import { EN_US, type NumberLocale } from '../cascade/index.ts';
import { MAX_FORMAT_CODE_LENGTH, MAX_FORMAT_NEW_CELLS } from '../store/index.ts';
import type { EditRefusal } from './cell-input.ts';
import { MAX_STEP_DECIMALS } from './number-format-commands.ts';
import type { ChromeDisabledReason, DisabledReasonKey } from './toolbar-commands.ts';

/**
 * Why a decimal step cannot change any cell of a selection:
 *
 * - The problem of the active cell's code, from `generalSection` to `missingCode`, as the
 *   decimal step names them. In a selection of more than one
 *   cell, that holds only when the active cell holds a number. Otherwise it is the problem of
 *   the first selected cell with a number, or `noNumbers`. A cell with no number cannot step
 *   there, whatever its code, the active cell too, and the step creates no cell.
 * - `noNumbers`: no selected cell holds a number. One cell (or one merged range) that holds text,
 *   a boolean, or an error gives it too, whatever its code.
 * - `noPlaces`: a decrease on a selection of one empty cell (or one merged range) that is
 *   `General`, or on one `General` number that shows no places. An increase on one empty `General`
 *   cell gives `0.0`.
 * - Among more cells, when the cells give different problems, the reason names the problem of
 *   the active cell when it holds a number, and else of the first selected cell with a number.
 *   `noDecimals` and `maxDecimals` speak of all selected cells, so they apply only when each
 *   selected cell with a number gives them. Else that cell's problem is `cellNoDecimals` or
 *   `cellMaxDecimals`, which speak of one selected cell.
 * - `tooLarge`: the selection holds more than 16,384 cells that exist. One
 *   merged range never gives it: the step reads only its anchor.
 *
 * The code problems:
 *
 * - `generalSection`: a `General` section with a color, a condition, a literal, or other
 *   sections, as in `[Red]General` or `General;-0`.
 * - `date`: a date or time section.
 * - `text`: the code shows text in place of a number: the value as text (`@`), or only literals.
 * - `fraction`: a fraction section, as in `# ?/?`.
 * - `noDecimals`: a decrease on a code with no decimal place to remove.
 * - `maxDecimals`: an increase on a section with 30 places, the most that a step writes.
 * - `tooLong`: the new code would be longer than 255 characters.
 * - `unreadable`: the step cannot change the places of the code. The parser cannot read it, the
 *   step would make a comma scale the number or add or remove a group separator, or the code
 *   holds a control character.
 * - `noCode`: a built-in id that has no code in the number locale.
 * - `missingCode`: a custom id that the workbook gives no code.
 *
 * @internal
 */
export type DecimalSelectionProblem =
  | 'generalSection'
  | 'date'
  | 'text'
  | 'fraction'
  | 'noDecimals'
  | 'maxDecimals'
  | 'tooLong'
  | 'unreadable'
  | 'noCode'
  | 'missingCode'
  | 'noNumbers'
  | 'noPlaces'
  | 'tooLarge'
  | 'cellNoDecimals'
  | 'cellMaxDecimals';

/** The most cells that exist inside a selection that a decimal step reads. */
export const MAX_DECIMAL_STEP_CELLS = MAX_FORMAT_NEW_CELLS;

/**
 * The disabled reason key of each problem of a decimal step on one cell, or on exactly one
 * merged range. For more cells, use `DECIMAL_SELECTION_PROBLEM_KEYS`. A test keeps each one in the
 * catalog. `tooLong` takes `limit` (`MAX_FORMAT_CODE_LENGTH`), `maxDecimals` and
 * `cellMaxDecimals` take `MAX_STEP_DECIMALS`, and `tooLarge` takes `MAX_DECIMAL_STEP_CELLS`, each
 * grouped in the number locale. The map is the lookup table of `decimalReasonOf`, which adds the
 * parameters. `noNumbers` on one cell is a cell that holds text, a boolean, or an error, and its
 * key speaks of that cell. The map also names the problems that only a selection gives
 * (`tooLarge`, `cellNoDecimals`, and `cellMaxDecimals`), so that it is total over
 * `DecimalSelectionProblem`.
 *
 * @internal
 */
export const DECIMAL_STEP_PROBLEM_KEYS = {
  noNumbers: 'disabledReason.decimalsNoNumber',
  noPlaces: 'disabledReason.decimalsNoShownPlaces',
  generalSection: 'disabledReason.decimalsGeneralSection',
  date: 'disabledReason.decimalsDate',
  text: 'disabledReason.decimalsText',
  fraction: 'disabledReason.decimalsFraction',
  noDecimals: 'disabledReason.decimalsNoDecimals',
  maxDecimals: 'disabledReason.decimalsMaxDecimals',
  tooLong: 'disabledReason.decimalsTooLong',
  unreadable: 'disabledReason.decimalsUnreadable',
  noCode: 'disabledReason.decimalsNoCode',
  missingCode: 'disabledReason.decimalsMissingCode',
  tooLarge: 'disabledReason.decimalsTooLarge',
  cellNoDecimals: 'disabledReason.decimalsNoDecimals',
  cellMaxDecimals: 'disabledReason.decimalsMaxDecimals',
} as const satisfies Readonly<Record<DecimalSelectionProblem, DisabledReasonKey>>;

/**
 * The disabled reason key of each problem of a decimal step on a selection of more than one
 * cell. These keys speak of "a selected cell", because the
 * problem can come from a cell that is not the active cell. For one cell, and for the internal
 * `decimalStepOf`, use `DECIMAL_STEP_PROBLEM_KEYS`. The parameters are the same. The map is the
 * lookup table of `decimalReasonOf`. `noPlaces` keeps its one-cell words, because only one cell
 * gives it.
 *
 * @internal
 */
export const DECIMAL_SELECTION_PROBLEM_KEYS = {
  noNumbers: 'disabledReason.decimalsNoNumberSelection',
  noPlaces: 'disabledReason.decimalsNoShownPlaces',
  generalSection: 'disabledReason.decimalsGeneralSectionSelection',
  date: 'disabledReason.decimalsDateSelection',
  text: 'disabledReason.decimalsTextSelection',
  fraction: 'disabledReason.decimalsFractionSelection',
  noDecimals: 'disabledReason.decimalsNoDecimalsSelection',
  maxDecimals: 'disabledReason.decimalsMaxDecimalsSelection',
  tooLong: 'disabledReason.decimalsTooLongSelection',
  unreadable: 'disabledReason.decimalsUnreadableSelection',
  noCode: 'disabledReason.decimalsNoCodeSelection',
  missingCode: 'disabledReason.decimalsMissingCodeSelection',
  tooLarge: 'disabledReason.decimalsTooLarge',
  cellNoDecimals: 'disabledReason.decimalsNoDecimalsCellSelection',
  cellMaxDecimals: 'disabledReason.decimalsMaxDecimalsCellSelection',
} as const satisfies Readonly<Record<DecimalSelectionProblem, DisabledReasonKey>>;

/** A whole number with the group separator of `locale`, such as `16,384`. */
export function groupDigits(n: number, locale: NumberLocale): string {
  // A replacer function: a group separator such as `$` is never read as a pattern.
  return String(Math.trunc(n)).replace(/\B(?=(\d{3})+$)/g, () => locale.group);
}

/** The disabled reason keys of the decimal steps, for a check of a refusal's `reason`. */
export const DECIMAL_REASON_KEYS: ReadonlySet<DisabledReasonKey> = new Set([
  ...Object.values(DECIMAL_STEP_PROBLEM_KEYS),
  ...Object.values(DECIMAL_SELECTION_PROBLEM_KEYS),
]);

/**
 * The reason of a decimal step problem, ready for a catalog: a key and its `limit`, grouped in
 * the number locale. `single` picks the key that speaks of one cell (`DECIMAL_STEP_PROBLEM_KEYS`)
 * over the key that speaks of a selected cell.
 */
export function decimalReasonOf(
  problem: DecimalSelectionProblem,
  single: boolean,
  locale: NumberLocale = EN_US
): ChromeDisabledReason {
  const key = (single ? DECIMAL_STEP_PROBLEM_KEYS : DECIMAL_SELECTION_PROBLEM_KEYS)[problem];
  const limit =
    problem === 'tooLong'
      ? MAX_FORMAT_CODE_LENGTH
      : problem === 'maxDecimals' || problem === 'cellMaxDecimals'
        ? MAX_STEP_DECIMALS
        : problem === 'tooLarge'
          ? MAX_DECIMAL_STEP_CELLS
          : undefined;
  return limit === undefined ? { key } : { key, params: { limit: groupDigits(limit, locale) } };
}

/**
 * The refusal of a decimal step that changes no cell: `decimals-too-large` for `tooLarge`, and
 * else `decimals-unchanged`. `reason` holds the disabled reason of the step's control, and
 * `detail` the problem, for logs.
 */
export function decimalRefusalOf(
  problem: DecimalSelectionProblem,
  single: boolean,
  locale: NumberLocale = EN_US
): EditRefusal {
  return {
    code: problem === 'tooLarge' ? 'decimals-too-large' : 'decimals-unchanged',
    detail: problem,
    reason: decimalReasonOf(problem, single, locale),
  };
}
