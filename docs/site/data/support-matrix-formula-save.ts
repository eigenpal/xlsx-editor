/**
 * The rows of the support matrix for the calculation chain and the recalculation flag after an
 * edit. `support-matrix.ts` lists them in its order; the rules of that file hold here.
 */

import type { SupportFeature } from './support-matrix.ts';

const CORE = 'packages/core/src';

export const FORMULA_SAVE_FEATURES: readonly SupportFeature[] = [
  {
    id: 'save.calc-chain',
    name: 'Calculation chain after a formula edit',
    category: 'save',
    status: 'supported',
    notes:
      'The calculation chain (the `x:calcChain` part, ECMA-376 Part 1, 18.6) lists the order in which formulas calculate. An edit or an undo that writes or removes a formula drops it, and the next consumer builds it again. The formula can be in an `x:c` nested in the edited cell. A value edit or a clear of a cell with a formula inside `mc:AlternateContent` is refused with `cell-alternate-content`, so the chain stays.',
    evidence: {
      saveVerified: ['formulas'],
      tests: [
        `${CORE}/store/__tests__/formula-edits.test.ts`,
        `${CORE}/store/__tests__/formula-scan.test.ts`,
      ],
    },
  },
  {
    id: 'save.full-calc-on-load',
    name: 'Recalculation flag after an edit',
    category: 'save',
    status: 'supported',
    notes:
      'An edit that can change a result sets `x:calcPr/@fullCalcOnLoad` when any sheet holds a formula cell, so the next consumer calculates again. Such an edit is an edit of a value or a formula, a row or column insert or delete, or a sheet delete. A sheet add, rename, or move is such an edit only when a formula or a defined name can read the changed name or position. The check scans every parsed sheet whose root element is in the Transitional SpreadsheetML namespace. It also finds a formula inside `mc:AlternateContent`, in an `x:c` outside a row, and in an `x:c` inside another cell. A sheet whose root element is in another namespace counts as a sheet with formulas. A sheet that the editor keeps as unread bytes, such as a macro sheet, counts too, when its content type is XML or it has none. A chartsheet and a dialog sheet hold no cells, so they never count. The save check opens the saved and edited copies of `wrapped-only-formula`. The only formula of that workbook is in `mc:AlternateContent`, so only the scan can set the flag. Unit tests cover the other forms after save and reopen. Each form writes the same `fullCalcOnLoad` attribute.',
    evidence: {
      saveVerified: ['basic-values', 'formulas', 'wrapped-formula', 'wrapped-only-formula'],
      tests: [
        `${CORE}/store/__tests__/formula-edits.test.ts`,
        `${CORE}/store/__tests__/formula-scan.test.ts`,
        `${CORE}/store/__tests__/structure-wrapped.test.ts`,
        `${CORE}/cascade/__tests__/values.test.ts`,
      ],
    },
  },
];
