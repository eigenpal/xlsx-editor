/**
 * The rows of the support matrix for the merge commands. `support-matrix.ts` lists them in its
 * order; the rules of that file hold here.
 */

import type { SupportFeature } from './support-matrix.ts';

const CORE = 'packages/core/src';

export const MERGE_FEATURES: readonly SupportFeature[] = [
  {
    id: 'editing.merge-cells',
    name: 'Merge and unmerge cells',
    category: 'editing',
    status: 'supported',
    notes:
      'Merge, merge and center, and unmerge act on the selection. Each command is one undo step. A merged range (`x:mergeCell`, ECMA-376 Part 1, 18.3.1.54) keeps the value of its top-left cell only. A merge over other values is refused with `merge-drops-values`, and the refusal names the first of those cells. A value or a formula inside `mc:AlternateContent` or in an element in another namespace counts too. The host can then pass `dropValues`. Then the other values and formulas are cleared in the same undo step. A cell with such content refuses that clear with `cell-alternate-content`, so the merge is not applied. A merge of one cell is refused with `merge-one-cell`. A range that meets a merged range is refused with `merge-overlap`. Unmerge removes each merged range that the selection meets. A protected sheet refuses both commands. `mergeCommandState` gives the state of each command for host controls, and the first cell that a merge would empty, so a host can ask first. The save check opens a copy of `grid-structure` after a merge and an unmerge.',
    evidence: {
      saveVerified: ['grid-structure'],
      tests: [
        `${CORE}/store/__tests__/structure-wrapped.test.ts`,
        `${CORE}/store/__tests__/merge-cells.test.ts`,
        `${CORE}/editor/__tests__/merge-editing.test.ts`,
        `${CORE}/editor/__tests__/cell-children.test.ts`,
        `${CORE}/editor/__tests__/view-mode.test.ts`,
      ],
    },
  },
];
