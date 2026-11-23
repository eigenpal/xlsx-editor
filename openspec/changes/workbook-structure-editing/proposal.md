## Why

The editor opens, shows, and edits cells, cell formats, rows, and columns. A usable base editor also needs the sheet operations of the sheet tabs, the clipboard, and the resize of columns and rows. Without them, a user cannot organize a workbook, move data between cells, or make a column wide enough for its text.

## What Changes

- Add, rename, delete, and move sheets from the sheet tabs. Formula references and defined names stay valid, or the edit is refused with a typed reason.
- Copy, cut, and paste a cell range inside the workbook, with values, formulas, and cell formats. Paste plain text from outside the editor. Refuse a paste that holds only HTML.
- Drag the edge of a column heading or a row heading to set the column width or the row height.
- Each edit is one transaction of typed ops through `WorkbookDocumentStore.transact`, and one undo step.
- Add i18n strings, refusal codes, fixtures with save-check records, tests, and support matrix rows for the three features.

## Capabilities

### New Capabilities

- `workbook-structure-editing`: sheet operations, the clipboard of a cell range, and the resize of columns and rows.

### Modified Capabilities

- `workbook-cell-editing`: the requirement "Ops carry intent" of the first change `typed-ooxml-workbook-editor` gets one exception. The op `setCellFormatIndex` carries a cell format index, because a paste gives each cell the cell format of the copy snapshot. A format edit never changes the record at an index, so a paste applies the format from the time of the copy. A paste is refused with `copy-expired` when the record at a copied index no longer holds the copied format.
- `workbook-cell-editing`: the requirement "Sheet protection" of the first change gets the rules of this change. The lock rules add `size` and `workbook`, and the commands add `cut` and `paste`. A new cell format that changes the protection of a cell is refused whatever `@formatCells` states. A clear of formats on a protected sheet keeps the protection of each cell.
- `workbook-cell-values`: the requirement "Stale formula results are visible" of the first change gets the edits that change no cell value but change a function result. A format edit, a clear of formats, and a paste of a cell format mark the cached results stale when a formula or a defined name calls `CELL` or a macro function `GET.*`. A column width, a row height, and a sheet op mark them stale by the rules of designs D1, D2, and D3.

This change archives after the first change `typed-ooxml-workbook-editor`, because its MODIFIED requirements need the `workbook-cell-editing` and `workbook-cell-values` specs that the first change adds.

The first change keeps its other requirements, and this change adds ops and editor commands next to them. The row and column edits of that change (`workbook-cell-editing`) take three rules of this change, which design D1 records:

- A new refusal `sheet-name-ambiguous` when a file gives the name of the edited sheet to another sheet too.
- A new refusal `sheet-name-too-long` when a stored sheet name has more than 31 characters.
- A hyperlink relationship target that is a fragment, such as `#Data!A5`, moves with the cells that it names.

The sheet protection of the first change (`x:sheetProtection`) takes six rules of this change, which designs D1, D2, and D3 record:

- `SheetProtection` gets `allowFormatColumns` and `allowFormatRows`, and a locked column width or row height is refused with `locked` and the new lock rule `size`.
- A paste that gives a cell a new cell format is a format edit, and a paste in the format that the cell already shows is not.
- The workbook protection (`x:workbookProtection/@lockStructure`) refuses a sheet op with `locked` and the new lock rule `workbook`.
- A new cell format that changes the `@locked` or `@hidden` protection of a cell is refused with `locked` and the lock rule `cell`, whatever `@formatCells` states.
- On a protected sheet, a clear of formats keeps the `@locked` and `@hidden` protection of each cell.
- A cut whose paste could not clear its source is refused at the cut with `locked` and the command `cut`.

The format ops of the first change take one rule of this change, which design D2 records:

- A format edit (`applyFormat`) and a clear of formats mark the cached results stale when a formula or a defined name calls `CELL` or a macro function `GET.*`, because a cell format changes their results. The saved workbook then asks for a full calculation on load.

## Non-goals

Merged cell edits, hidden rows, columns, and sheets, fill handles, outlines, insert with shift (insert of copied cells), paste of HTML or of rich content from other applications, paste special modes, the move of references on cut and paste, sheet copy, tab colors, auto fit of a column or a row, and a keyboard command for resize. These belong to later changes.

## Impact

The store gets the ops `addSheet`, `renameSheet`, `deleteSheet`, `moveSheet`, `setCellFormatIndex`, `setColumnWidth`, and `setRowHeight` (`packages/core/src/store/`). The editor gets sheet commands, a clipboard module, and a header resize module (`packages/core/src/editor/`). The editor also wires the clipboard events and the resize drag of the grid surface. The React adapter gets the tab actions (`packages/react`). `packages/i18n` gets the strings. The fixture catalog gets sheet, clipboard, and resize edits with save-check records. Collaboration stays out of scope, as design D10 of the first change states.
