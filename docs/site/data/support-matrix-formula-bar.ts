/**
 * The rows of the support matrix for the formula bar: the content field, which edits the
 * active cell, and the name box, which selects a typed reference. `support-matrix.ts` lists
 * them in its order; the rules of that file hold here.
 */

import type { SupportFeature } from './support-matrix.ts';

const CORE = 'packages/core/src';

export const FORMULA_BAR_FEATURES: readonly SupportFeature[] = [
  {
    id: 'editing.formula-bar',
    name: 'Formula bar input',
    category: 'editing',
    status: 'partial',
    notes:
      'Typing in the content field of the formula bar edits the active cell. The grid shows the same text as you type. Enter, Tab, Shift+Enter, and Shift+Tab commit and move, and Escape cancels. Alt+Enter adds a line break. A refused commit keeps the text in the field and shows the reason. An input method composes with no early commit. A text with more than one line shows a line count. A press on it, or on the gap beside the field, puts the caret at the end of the shown line, in view. With Shift, while the field has the keyboard, the press extends the text selection. The field is read-only in view mode. For a formula cell, the field shows a direct `x:f` first. Else it reads the `mc:AlternateContent` wrappers in document order: in each, `mc:Fallback` first, then each `mc:Choice`. A shared formula child in a branch shows the formula of its master, moved to the child, when the master holds its `x:f` directly. A cell with an empty `x:f` has no formula text, so the field shows the value. No reference measures which wrapped formula a consumer reads (approximation "Wrapped formula values"). The field shows a cell text of more than 65,536 UTF-16 code units shortened and read-only, with a note. F2 in the grid edits the cell. An edit applies only a text of at most 32,767 UTF-16 code units, the limit of a cell value.',
    evidence: {
      approximations: ['wrapped-formula-values'],
      saveVerified: ['basic-values'],
      tests: [
        `${CORE}/editor/__tests__/formula-bar-edit.test.ts`,
        `${CORE}/editor/__tests__/several-ranges.test.ts`,
        `${CORE}/editor/__tests__/stale-notice.test.ts`,
        'packages/react/src/__tests__/formula-bar.test.tsx',
      ],
    },
  },
  {
    id: 'editing.formula-point-mode',
    name: 'Point to cells while you type a formula',
    category: 'editing',
    status: 'unsupported',
    notes:
      'A press on a cell while you type a formula does not insert a reference. It commits the edit, or shows why the edit cannot commit. Type the reference instead.',
    evidence: {
      tests: [
        `${CORE}/editor/__tests__/formula-bar-edit.test.ts`,
        `${CORE}/editor/__tests__/cell-editing.test.ts`,
      ],
    },
  },
  {
    id: 'viewer.name-box',
    name: 'Name box: go to a cell, a range, or a defined name',
    category: 'viewer',
    status: 'supported',
    notes:
      'Enter in the name box selects a typed cell, range, whole columns, or whole rows, with an optional sheet name, in any letter case. A defined name goes to its range when its text is one reference with a sheet name; another defined name is refused with `go-to-name`. A sheet that is not a visible worksheet is refused with `go-to-sheet`, and other text with `go-to-reference`. Several references, space delimited (`B2:C4 E5` or `A:A C:C`), select those ranges on the shown sheet. A sheet name with a space cannot be one of them. More than `MAX_SELECTION_RANGES` references are refused with `too-many-ranges`. With several ranges selected, the name box shows the active cell, and Enter with no typed text keeps every range.',
    evidence: {
      tests: [
        `${CORE}/editor/__tests__/formula-bar-edit.test.ts`,
        'packages/react/src/__tests__/formula-bar.test.tsx',
      ],
    },
  },
];
