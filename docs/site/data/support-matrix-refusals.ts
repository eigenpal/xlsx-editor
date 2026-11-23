/**
 * The rows of the support matrix for refused edits. `support-matrix.ts` lists them in its order;
 * the rules of that file hold here.
 */

import type { SupportFeature } from './support-matrix.ts';

const CORE = 'packages/core/src';

export const REFUSAL_FEATURES: readonly SupportFeature[] = [
  {
    id: 'refusals.structural',
    name: 'Structural edits over content that holds references',
    category: 'refusals',
    status: 'supported',
    notes:
      'Insert and delete are refused when preserved content can refer to the moved cells. That content is a chart, a pivot table, or a sheet that is not a worksheet, anywhere in the workbook. On the sheet, it is any element or related part outside a fixed allowlist. Examples are conditional formatting, data validation, a table, a filter, and comments. It is also a table or query table on another sheet that names the sheet. It is also a workbook part outside a fixed allowlist, such as data connections. A formula or a defined name inside `mc:AlternateContent` that the edit would change refuses it too. The refusal names the content in words, such as "conditional formatting on sheet Data", and keeps the part path in its detail for logs. Other editing stays available.',
    evidence: {
      tests: [
        `${CORE}/store/__tests__/structure-refusals.test.ts`,
        `${CORE}/store/__tests__/structure-cases.test.ts`,
        `${CORE}/store/__tests__/structure-wrapped.test.ts`,
        `${CORE}/editor/__tests__/structure-editing.test.ts`,
        `${CORE}/editor/__tests__/toolbar-commands.test.ts`,
      ],
    },
  },
  {
    id: 'refusals.whole-sheet',
    name: 'Structural edits on a whole-sheet selection',
    category: 'refusals',
    status: 'supported',
    notes:
      'Insert and delete are refused when the whole sheet is selected, because they would act on every row or column. The toolbar disables them with the reason.',
    evidence: {
      tests: [
        `${CORE}/editor/__tests__/structure-editing.test.ts`,
        `${CORE}/editor/__tests__/toolbar-commands.test.ts`,
      ],
    },
  },
  {
    id: 'refusals.cells',
    name: 'Cell edits that would break the file',
    category: 'refusals',
    status: 'supported',
    notes:
      'An edit inside an array formula or a data table is refused and writes nothing, also when its master is inside `mc:AlternateContent`. While a clear of the selection would change a formula range, Clear contents and Clear all are disabled. Their reason names the range, or the shared formula master to leave out. When the formula of a shared child would move out of the grid, those controls stay enabled. The clear is then refused when it runs. Formula text that cannot be read is refused, and so is R1C1 formula text. Formula text with a character that XML cannot hold is also refused. Typed formula text that calls a function outside the predefined functions of ECMA-376 Part 1, 18.17.7, is refused. A column width or a row height outside its limits is refused. A value edit, a paste, a clear, or a cut of a cell that holds another form of its value is refused with `cell-alternate-content`. A cut of a wrapped formula is the exception. The message names the cell. That form is markup compatibility content other than the value, or an element in another namespace with value content in it. A cut of a cell whose formula is inside `mc:AlternateContent` gives `copy-wrapped-formula`, as its copy does. A cut of a cell whose only value is in markup compatibility content is refused with `cut-wrapped-value`. A value edit, a clear, or a paste keeps each other child of a cell in its place. A column width edit is refused when an `x:cols` after the first has an attribute that the merge of the column lists would drop. An edit that would leave a namespace prefix undeclared is refused with `invalid-part`.',
    evidence: {
      tests: [
        `${CORE}/store/__tests__/cell-children.test.ts`,
        `${CORE}/store/__tests__/namespace-declarations.test.ts`,
        `${CORE}/store/__tests__/store-ops.test.ts`,
        `${CORE}/store/__tests__/edit-safety.test.ts`,
        `${CORE}/store/__tests__/formula-edits.test.ts`,
        `${CORE}/store/__tests__/formula-scan.test.ts`,
        `${CORE}/store/__tests__/axis-size.test.ts`,
        `${CORE}/editor/__tests__/cell-editing.test.ts`,
        `${CORE}/editor/__tests__/clipboard.test.ts`,
        `${CORE}/editor/__tests__/resize.test.ts`,
      ],
    },
  },
  {
    id: 'refusals.codes',
    name: 'Refusal codes and messages',
    category: 'refusals',
    status: 'supported',
    notes:
      'A refused edit writes nothing and returns one code, with a message in the shared strings. Cell codes: `unknown-sheet`, `bad-reference`, `bad-value`, `text-too-long`, `invalid-sheet`, `cell-alternate-content`, `shared-formula`. That refusal protects an array or data table formula range, an unreadable shared master, or a shared child that would move out of the grid. The array or data table master can be inside `mc:AlternateContent`. It also protects a shared master whose child has its `x:f` there. Its message names the cell or the formula range that blocks the edit, and the cell whose formula the editor cannot rewrite for a shared master. Formula codes: `unreadable-formula`, `formula-syntax`, `r1c1-formula`, `unknown-formula-reference`, `invalid-character`, `unsupported-function`. Format codes: `bad-format`, `unreadable-number-format`, `range-too-large`, `clear-too-large`, `decimals-too-large`, `decimals-unchanged`, `style-limit`. Workbook codes: `not-editable`, `unsupported-package`, `invalid-part`, `stale-revision`, `save-failed`. Structure codes: `structure-blocked`, `grid-full`, `whole-sheet`. Sheet codes: `sheet-range-blocked`, `sheet-name-too-long`, `sheet-name-invalid`, `sheet-name-taken`, `sheet-name-ambiguous`, `last-sheet`, `bad-position`. Clipboard codes: `merged-range`, `paste-html`, `copy-too-large`, `copy-conditional-limit`, `copy-array-formula`, `copy-unreadable-formula`, `copy-wrapped-formula`, `paste-unreadable-formula`, `paste-name-scope`, `paste-reference-outside`, `cut-wrapped-value`, `cut-expired`, `copy-expired`. Resize codes: `bad-size`, `too-many-rows`, `resize-anchor`. Merge codes: `merge-one-cell`, `merge-overlap`, `merge-drops-values`. History codes: `nothing-to-undo`, `nothing-to-redo`, `edit-pending`, which refuses a redo or a decimal step. Name box codes: `go-to-reference`, `go-to-name`, `go-to-sheet`. Input method code: `composing`. Load code: `loading`, which refuses every command, a save included, while a workbook loads. Dropped edit code: `edit-dropped`, which no command returns: the snapshot keeps it when the editor drops an open edit with changed text. A refused save keeps its code in the snapshot, marked as a save, and its message says that the workbook was not saved. Each refusal has one whole-sentence message in the catalog, and the status strip shows it in full. Mode code: `view-mode`. Protection code: `locked`. Its lock field names the rule that refused it: cell, format, structure, size, merge, or workbook. The workbook rule refuses a sheet add, rename, delete, or move. A refused commit of an open cell edit keeps the cell editor open, and its message tells the user to cancel the cell edit.',
    evidence: {
      tests: [
        `${CORE}/editor/__tests__/cell-editing.test.ts`,
        `${CORE}/editor/__tests__/refusal-text.test.ts`,
        `${CORE}/editor/__tests__/toolbar-commands.test.ts`,
        `${CORE}/store/__tests__/edit-safety.test.ts`,
        `${CORE}/store/__tests__/sheet-ops.test.ts`,
        `${CORE}/store/__tests__/formula-scan.test.ts`,
        `${CORE}/editor/__tests__/clipboard.test.ts`,
        `${CORE}/editor/__tests__/formula-bar-edit.test.ts`,
        `scripts/check-support-matrix.test.ts`,
      ],
    },
  },
];
