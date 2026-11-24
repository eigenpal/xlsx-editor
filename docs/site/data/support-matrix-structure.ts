/**
 * The rows of the support matrix for the structure edits of the change
 * `workbook-structure-editing`: the clipboard, the resize, and the sheet commands.
 * `support-matrix.ts` lists them in its order; the rules of that file hold here.
 */

import type { SupportFeature } from './support-matrix.ts';

const CORE = 'packages/core/src';

export const STRUCTURE_FEATURES: readonly SupportFeature[] = [
  {
    id: 'editing.clipboard',
    name: 'Copy, cut, and paste',
    category: 'editing',
    status: 'partial',
    notes:
      'A paste in the workbook writes values, formulas, and cell formats as one undo step. A copied formula moves by the offset. A copy of a cell whose formula is inside `mc:AlternateContent` is refused with `copy-wrapped-formula`. A cut of it is refused with the same code. So its cached value never pastes as a constant. A cut moves the values, formulas, and cell formats of its cells, but references to the cut cells do not move. Hyperlinks, comments, data validation, and conditional formatting stay with the source cells. A cut formula keeps the cells that it refers to, also on another sheet. On another sheet, a copied or cut name that the source sheet scopes gets that sheet name. For a paste on another sheet, the copy then depends on that sheet. A function name that the source sheet scopes refuses a paste on another sheet. A formula can hold a name with no sheet prefix that the target sheet scopes and the source sheet does not. A paste of such a copied or cut formula on another sheet is refused. A paste of a cut after another edit is refused, and the user cuts again. A second paste of a cut is a copy. A paste is refused when a sheet that a copied formula names, or a copied cell format, changed after the copy. A delete and an add of that sheet is a change. A paste of copied formulas is also refused after an insert or a delete of rows or columns. A rename, a move, or a delete of a sheet that a copied formula names also refuses it. For a reference across a span of sheets, an add, a rename, a move, or a delete of any sheet refuses it. An undo or a redo of such an edit does not bring the copy back. An expired copy stays refused until the user copies again. A moved reference that leaves the sheet is refused. Rich text pastes as plain text, and an error value pastes as text. A tab or a line break inside a copied cell becomes a space in the copied plain text. A date whose shown text does not read back to the same day, such as `2-Jan`, copies its ISO 8601 text. Pasted plain text reads as typed input: a tab starts a cell and a line break starts a row. This holds inside a quoted value too, and its quotes stay. A typed form on a General cell also sets its number format (`editing.cell-values`). A sheet can refuse formatting cells, or the list of cell formats or number formats can be full. Then the paste stores the number and keeps the cell format. A paste of HTML only and a range that meets a merged range are refused. On a protected sheet, a paste is refused when a target cell is locked. It is also refused when it gives a cell another cell format and the protection does not allow formatting cells. A new cell format that changes whether a cell is locked or hidden is refused with the rule `cell`, whatever the protection allows. A cut whose source cells the editor cannot clear is refused at the cut. A refused cut says that the cells were not cut. The paste clears only the source cells that it does not write, before it writes. So a paste over part of the cut source can still be refused, when that clear alone is refused. On a protected sheet, a clear keeps whether each cell is locked or hidden. A paste that changes a cell format marks the cached results for recalculation when a formula calls `CELL` or a macro function `GET.*`. A copy of more than 16,384 cells, an array formula, or a data table is refused. A dashed outline in the editor overlay marks the source cells of a pending copy or cut. It is not part of the workbook. A paste that gets only plain text never moves cut cells. It ends a pending cut, and the live region says that the cells stay. Such a paste writes the shown text, so a formula becomes its shown value, and a number takes its rounded shown text.',
    evidence: {
      fixtures: ['clipboard-edit'],
      saveVerified: ['clipboard-edit'],
      approximations: [
        'cut-references',
        'cut-cell-attachments',
        'pasted-values',
        'pasted-quotes',
        'copied-line-breaks',
        'shadow-root-clipboard',
      ],
      tests: [
        `${CORE}/editor/__tests__/typed-entry.test.ts`,
        `${CORE}/editor/__tests__/typed-entry-copy.test.ts`,
        `scripts/fixtures/editing.test.ts`,
        `${CORE}/store/__tests__/cell-format-index.test.ts`,
        `${CORE}/editor/__tests__/clipboard.test.ts`,
        `${CORE}/editor/__tests__/clipboard-cases.test.ts`,
        `${CORE}/editor/__tests__/clipboard-expiry.test.ts`,
        `${CORE}/editor/__tests__/clipboard-mark.test.ts`,
        `${CORE}/editor/__tests__/clipboard-events.test.ts`,
        `${CORE}/store/__tests__/sheet-protection.test.ts`,
        `${CORE}/store/__tests__/sheet-ops-edges.test.ts`,
        `packages/react/src/__tests__/clipboard.test.tsx`,
      ],
    },
  },
  {
    id: 'editing.resize',
    name: 'Resize columns and rows',
    category: 'editing',
    status: 'partial',
    notes:
      'A drag of the right edge of a column heading sets a custom width, up to 255 characters. A drag of the bottom edge of a row heading sets a custom height, up to 409 points. Each drag is one undo step. A selection of whole columns or rows that holds the dragged one resizes all of them. A height for more than 16,384 rows that have no row record is refused. Escape cancels the drag. A double-click on the edge fits the column to its widest shown text, or the row to its tallest font and wrapped lines. Each fit is one undo step. A selection of whole columns or rows that holds the edge fits each of them. A fitted row is not a custom height, so a later font or value edit fits it again. On a sheet whose default row height is set by hand, such an edit gives the row that default height instead. Layout gives that height to each row without a height. A fitted row holds its wrapped lines at each zoom step from 50% and at the current zoom. At a lower zoom, its last wrapped line can be cut. A fit has no save check: the save check of `resize-edit` covers a drag, and a fitted row writes `@ht` without `@customHeight`. For a column width, text that wraps, rotated text, and justified or distributed text do not count. Text centered across cells, text of a merged range over several columns, and text in a hidden row do not count either. A column with more conditional formats than layout checks keeps its width. A row with rotated text keeps its height, and hidden columns and rows keep their size. A column or row of size 0 can be dragged wider from the right or bottom side of its edge. A drawing or control that keeps its size refuses a resize of its columns or rows with `resize-anchor`. One that keeps its position also refuses a resize before it. A drawing part that the editor cannot read refuses every resize of its sheet with `resize-anchor`. The message says that the editor cannot read the drawing. A resize keeps the cells of legacy drawing anchors, such as comment boxes. On a protected sheet, a width is refused with `locked` unless the protection allows formatting columns. A height is refused unless it allows formatting rows. A width marks the cached results for recalculation when a formula or a defined name calls `CELL` or a macro function `GET.*`. Both can read the width. A height marks them when a formula or a defined name calls a macro function `GET.*`, such as `GET.CELL`, which can read the height.',
    evidence: {
      fixtures: ['resize-edit'],
      saveVerified: ['resize-edit'],
      approximations: ['legacy-drawing-anchors', 'row-auto-height', 'wrap-measure-budget'],
      tests: [
        `scripts/fixtures/editing.test.ts`,
        `${CORE}/store/__tests__/axis-size.test.ts`,
        `${CORE}/store/__tests__/drawing-kept-as-bytes.test.ts`,
        `${CORE}/store/__tests__/sheet-ops-calc.test.ts`,
        `${CORE}/store/__tests__/sheet-protection.test.ts`,
        `${CORE}/layout/__tests__/grid-geometry.test.ts`,
        `${CORE}/editor/__tests__/resize.test.ts`,
        `${CORE}/editor/__tests__/fit.test.ts`,
        `${CORE}/layout/__tests__/fit.test.ts`,
      ],
    },
  },
  {
    id: 'editing.sheets',
    name: 'Add, rename, delete, and move sheets',
    category: 'editing',
    status: 'partial',
    notes:
      'A rename rewrites formulas, defined names, hyperlink locations, and hyperlink targets such as `#Data!A1`. When the new name holds a space or a character outside ASCII, such a target becomes a hyperlink location. A delete turns each reference to the sheet into `#REF!` and removes its sheet-scoped names. It removes the printer settings, the drawing, and the images that only the sheet uses. A hyperlink whose target names the deleted sheet gets the location `#REF!`. A chart, a pivot table, or a sheet that is not a worksheet anywhere in the workbook refuses a rename or a delete. A table, conditional formatting, data validation, drawing, or control that names the sheet also refuses them. So does a relationship of the workbook or the package to a part outside a fixed allowlist, such as data connections or revisions. A drawing of the deleted sheet that links a shape to the sheet goes with it. A delete refuses a sheet with a related part other than hyperlinks, printer settings, images, and a drawing. It also refuses a sheet whose drawing relates to a part other than an image or a hyperlink, such as a diagram. A custom workbook view names its sheet by `@activeSheetId`, so it refuses only a delete of that sheet. A sheet range refuses a move that changes the sheets between its ends or their order. It refuses a delete of its start or end sheet. A sheet range of another workbook (`[1]Jan:Mar!A1`) keeps its text. A local sheet range whose start or end reads as a cell refuses a rename. A sheet range in content that the editor does not rewrite refuses any move. An `x:row` or an `x:c` inside `mc:AlternateContent` that names the sheet refuses a rename or a delete. A formula or a defined name inside `mc:AlternateContent` keeps its text in each branch. A rename, a delete, or a move that would change that text is refused. An add, a delete, or a move that changes the position of the local scope sheet of such a name is refused too. The refusal names the cell or the name, or the wrapper when a name list in `x:workbook` names the sheet. A new name has at most 31 UTF-16 code units and no `_xHHHH_` escape sequence. Names compare without case, and `Maße` and `MASSE` are two names. A new name that equals another after canonical composition (NFC) is refused. When a file gives two sheets the same name without case, a delete or a row or column edit of either is refused with `sheet-name-ambiguous`. So is a rename of either while a reference names one of them. A formula or a defined name can name a sheet before an add or rename gives a sheet that name. Then its cached results are marked for recalculation when the file opens. Some functions read sheet names or positions: `INDIRECT`, `CELL`, `INFO`, `SHEET`, `SHEETS`, `FORMULATEXT`, and each macro function `GET.*` in a formula or a defined name. An add, a rename, or a move marks the results of their callers in the same way. The sheet titles in `docProps/app.xml` keep their old text. A file can store a sheet name of more than 31 characters. Then a move or a delete of any sheet, and an insert or delete of rows or columns, is refused with `sheet-name-too-long`, which names that sheet. A rename can shorten the name. When the workbook holds macros, an add, a rename, a delete, or a move is refused, and the sheet tabs disable them. Macro code can name any sheet or read it by its position. When `x:workbookProtection/@lockStructure` is true, an add, a rename, a delete, or a move is refused with `locked`, and the sheet tabs disable them. The tab of a sheet that is not a worksheet is disabled, and the tabs do not rename, delete, or move it. Its tooltip says why. The `moveSheet` command refuses to move such a sheet with `structure-blocked`, which names the sheet, as a delete does. Its message says that the editor can move or delete only a worksheet. A worksheet moves past it. When the workbook content refuses a rename or a delete of any sheet, the sheet menu disables both with that reason. That content is a chart, a pivot table, a sheet that is not a worksheet, or a relationship outside the allowlist. Each worksheet tab then shows the reason in its tooltip and its description.',
    evidence: {
      fixtures: ['sheet-ops', 'sheet-add'],
      saveVerified: ['sheet-ops', 'sheet-add'],
      approximations: ['sheet-titles'],
      tests: [
        `${CORE}/editor/__tests__/macro-enabled.test.ts`,
        `scripts/fixtures/editing.test.ts`,
        `${CORE}/store/__tests__/sheet-ops.test.ts`,
        `${CORE}/store/__tests__/sheet-ops-edges.test.ts`,
        `${CORE}/store/__tests__/structure-wrapped.test.ts`,
        `${CORE}/store/__tests__/sheet-name-fold.test.ts`,
        `${CORE}/store/__tests__/sheet-ops-cases.test.ts`,
        `${CORE}/store/__tests__/sheet-ops-links.test.ts`,
        `${CORE}/store/__tests__/sheet-ops-drawing.test.ts`,
        `${CORE}/store/__tests__/sheet-range-ops.test.ts`,
        `${CORE}/editor/__tests__/sheet-commands.test.ts`,
        `${CORE}/editor/__tests__/sheet-command-state.test.ts`,
        `${CORE}/store/__tests__/sheet-ops-calc.test.ts`,
        `${CORE}/store/__tests__/sheet-protection.test.ts`,
        `packages/react/src/__tests__/sheet-tabs.test.tsx`,
        `packages/react/src/__tests__/sheet-tabs-rename.test.tsx`,
        `packages/react/src/__tests__/sheet-tabs-focus.test.tsx`,
      ],
    },
  },
];
