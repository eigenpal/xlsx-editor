## 1. Sheets

- [x] 1.1 Add the ops `addSheet`, `renameSheet`, `deleteSheet`, and `moveSheet` with their inverses and refusals, by design D1.
- [x] 1.2 Add the editor commands and the sheet tab chrome (add button, context menu, inline rename), with i18n strings, by design D4.

## 2. Clipboard

- [x] 2.1 Add the op `setCellFormatIndex` and the clipboard module: copy, cut, internal paste, plain-text paste, and the refusals, by design D2.
- [x] 2.2 Wire the copy, cut, and paste events of the grid surface in the editor (`packages/core/src/editor/clipboard-events.ts`).
- [x] 2.3 Add the outline of a pending copy or cut, and end a copy at an edit that can change what its formulas refer to, by design D2.

## 3. Resize

- [x] 3.1 Add the ops `setColumnWidth` and `setRowHeight`, by design D3.
- [x] 3.2 Add the header resize drag with its preview, cancel, and commit.
- [x] 3.3 Add the fit to content on a double-click on a heading edge, with the ops `setColumnWidths` and `setFittedRowHeights`, by design D3.

## 4. Evidence

- [x] 4.1 Add the fixtures `sheet-ops`, `sheet-add`, `clipboard-edit`, and `resize-edit` with save-check records.
- [x] 4.2 Add the support matrix rows, the approximations of D1, D2, and D3, and the docs.
