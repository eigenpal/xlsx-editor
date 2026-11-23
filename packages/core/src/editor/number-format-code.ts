// The check of a typed number format code by the store's own rules (ECMA-376 Part 1, 18.8.31):
// the checks of the format patch `{ numberFormat: { code } }` that the apply runs. A leaf
// module, so the toolbar commands and the number format entry both import it with no cycle.

import { lowerNumberFormat, MAX_FORMAT_CODE_LENGTH, patchProblem } from '../store/index.ts';

/**
 * Every catalog key that `numberFormatCodeProblem` can return. A test keeps each one in the
 * catalog.
 *
 * @public
 */
export const NUMBER_FORMAT_CODE_PROBLEM_KEYS = [
  'numberFormat.problem.empty',
  'numberFormat.problem.tooLong',
  'numberFormat.problem.invalidCharacter',
  'numberFormat.problem.unreadable',
] as const;

/**
 * Why a typed number format code cannot be applied: a key of the shared catalog and its
 * parameters. The adapters translate it.
 *
 * @public
 */
export interface NumberFormatCodeProblem {
  readonly key: (typeof NUMBER_FORMAT_CODE_PROBLEM_KEYS)[number];
  readonly params?: Readonly<Record<string, string>>;
}

/**
 * Why the store refuses a number format code, or null when it applies the code: the checks of
 * the format patch `{ numberFormat: { code } }`, which the apply runs. A code of white space
 * only is refused as empty, because it shows nothing.
 *
 * @public
 */
export function numberFormatCodeProblem(code: string): NumberFormatCodeProblem | null {
  // A code of white space only shows nothing that the user can see: it is refused as empty.
  if (code.trim().length === 0) return { key: 'numberFormat.problem.empty' };
  if (code.length > MAX_FORMAT_CODE_LENGTH)
    return {
      key: 'numberFormat.problem.tooLong',
      params: { limit: String(MAX_FORMAT_CODE_LENGTH) },
    };
  const patch = { numberFormat: { code } };
  if (patchProblem(patch) !== undefined) return { key: 'numberFormat.problem.invalidCharacter' };
  return 'problem' in lowerNumberFormat(patch) ? { key: 'numberFormat.problem.unreadable' } : null;
}
