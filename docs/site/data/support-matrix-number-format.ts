/**
 * The rows of the support matrix for the number format commands: the custom number format
 * entry (the number format dialog and its apply), the entries of the number format list, and
 * the decimal place steps. `support-matrix.ts` lists them in its order; the rules of that file
 * hold here.
 */

import type { SupportFeature } from './support-matrix.ts';

const CORE = 'packages/core/src';

export const NUMBER_FORMAT_FEATURES: readonly SupportFeature[] = [
  {
    id: 'editing.custom-number-format',
    name: 'Custom number format entry',
    category: 'editing',
    status: 'supported',
    notes:
      'The number format dialog takes a typed format code and lists the custom codes of the workbook. Its preview of the active cell uses the number format engine of the grid. The dialog refuses a code that the engine cannot read, with the reason, before the apply. It also refuses a code over 255 characters or with a control character. A new code gets a custom number format from id 164. An equal code that the workbook holds keeps its id, and a built-in code is written as its built-in id. A built-in code stays a custom code when the workbook gives that id another code. A code equal to a redefined built-in id also gets a custom id. The apply is one undo step. The grid shows a built-in id that has no code in the number locale as `General`. The dialog opens with `General` for it, and an apply of `General` writes the `General` format. A code keeps its locale, calendar, or numeral system selector in the file. The save check covers a new custom code in one `colors-fonts` edit.',
    evidence: {
      saveVerified: ['colors-fonts'],
      tests: [
        `${CORE}/editor/__tests__/number-format-entry.test.ts`,
        `packages/react/src/__tests__/number-format-dialog.test.tsx`,
        `packages/react/src/__tests__/number-format-host.test.tsx`,
        `${CORE}/store/__tests__/format-ops.test.ts`,
      ],
    },
  },
  {
    id: 'editing.number-format-list',
    name: 'Number format list',
    category: 'editing',
    status: 'supported',
    notes:
      'The list offers General, Number, Currency, Accounting, Percentage, Scientific, Fraction, Short date, Long date, Time, and Text. An entry writes its built-in id only when the workbook gives that id no other code. Otherwise it writes the code of the id in the number locale. Currency and Accounting take the currency symbol of the number locale, and the symbol goes before the number in every number locale. Long date writes `dddd, mmmm d, yyyy` in every number locale. The list leaves both out when that symbol is not valid. A pick of the entry that a single cell shows writes nothing and adds no undo step. A code that differs from an entry only in its decimal places shows the name of that entry. The save check covers the codes that the list writes for Scientific, Currency, Accounting, and Long date.',
    evidence: {
      saveVerified: ['number-formats-more'],
      tests: [
        `${CORE}/editor/__tests__/number-format-commands.test.ts`,
        `${CORE}/editor/__tests__/number-format-list.test.ts`,
        `${CORE}/store/__tests__/number-format-lowering.test.ts`,
        'packages/react/src/__tests__/number-format-commands.test.tsx',
      ],
    },
  },
  {
    id: 'editing.decimal-places',
    name: 'Increase and decrease decimal places',
    category: 'editing',
    status: 'partial',
    notes:
      'Each selected cell steps from its own number format code, one decimal place more or less in each number section, in one undo step. On more cells, the General numbers of a column take the step of its first General number. So a decrease can add a place to a whole number, and an increase can show fewer places of another number. The `@xlsx-editor.dev/react` README gives the step and the rules of each code under Decimal places. In the zero section, `?` padding beside a literal other than a dash, or with no fill, can be lost (approximation `decimal-padding`). Decimal places that a literal or a second point splits, as in `0.00" "0`, do not always step back (approximation `decimal-split-places`). A cell keeps its format when the step would scale the number or change its group separator. A code with a control character, as a tab in the file, keeps its format, and the other selected cells step. A cell that holds text, a boolean, or an error does not step. If no selected cell holds a number, both controls are disabled. On a selection that is too large to format, the steps stay enabled with a hint, because they change only the cells that hold numbers. One empty cell steps from its own code, and an increase on one empty General cell writes `0.0`. A date cell (`t="d"`) steps from the serial number that the grid shows. The step reads at most 16,384 cells that exist inside the selection. The save check covers one increase on a range in which three cells step and three keep a date or time.',
    evidence: {
      approximations: ['decimal-padding', 'decimal-split-places'],
      saveVerified: ['number-formats-more'],
      tests: [
        `${CORE}/editor/__tests__/number-format-commands.test.ts`,
        `${CORE}/editor/__tests__/decimal-step-selection.test.ts`,
        `${CORE}/editor/__tests__/decimal-padding.test.ts`,
        'packages/react/src/__tests__/number-format-commands.test.tsx',
      ],
    },
  },
];
