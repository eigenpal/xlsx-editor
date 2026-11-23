## ADDED Requirements

### Requirement: Sheet names

The store SHALL refuse a sheet name that is empty, longer than 31 UTF-16 code units, holds one of `[ ] : * ? / \`, a control character, another character that XML 1.0 cannot hold, or an `ST_Xstring` escape sequence (`_xHHHH_`), starts or ends with an apostrophe, or matches another sheet name without case, also after both names take their canonical composition (NFC).

#### Scenario: A name that is taken

- **WHEN** the user renames a sheet to the name of another sheet in other case
- **THEN** the store refuses the edit with `sheet-name-taken`

#### Scenario: A name with a colon

- **WHEN** the user renames a sheet to `A:B`
- **THEN** the store refuses the edit with `sheet-name-invalid`

#### Scenario: A name with an escape sequence

- **WHEN** the user renames a sheet to `_x0041_`
- **THEN** the store refuses the edit with `sheet-name-invalid`

#### Scenario: Canonically equal names

- **WHEN** the workbook has the sheet `Café`, with U+00E9, and the user adds a sheet `Café`, with `e` and U+0301
- **THEN** the store refuses the edit with `sheet-name-taken`

#### Scenario: Names that differ only by a character with a longer uppercase

- **WHEN** the workbook has the sheets `Maße` and `MASSE`, and the user renames `Maße`
- **THEN** only `Maße` and the references to it change

#### Scenario: Two sheets of a file with the same name without case

- **WHEN** a file has the sheets `Data` and `DATA`, a formula holds `=DATA!A1`, and the user renames or deletes one of them
- **THEN** the store refuses the edit with `sheet-name-ambiguous`, with the other sheet as the subject

#### Scenario: A rename of a sheet name that no reference names

- **WHEN** a file has the sheets `Data` and `DATA`, no reference names either of them, and the user renames `Data` to `X`
- **THEN** the sheet gets the new name, and a later edit of `DATA` is not refused for the name

#### Scenario: A sheet name of another workbook

- **WHEN** `Sum!A1` holds `=SUM([1]Jan:Mar!A1)`, and the user renames the local sheet `Jan`
- **THEN** the sheet gets the new name, and `Sum!A1` keeps its text

### Requirement: Add a sheet

The store SHALL add a worksheet with a new part, relationship, content type override, and `sheetId`, at the requested position.

#### Scenario: Add and save

- **WHEN** the user adds a sheet and saves
- **THEN** the saved workbook has the new sheet at that position, and it opens with no repair prompt

#### Scenario: A formula that already names the new name

- **WHEN** `Sheet1!A1` holds `=Data!A1` with a cached error, and the user adds a sheet named `Data`
- **THEN** the saved workbook asks for a full calculation on load, so the cached result is not shown as current

#### Scenario: A function that reads sheet names or positions

- **WHEN** a cell formula or a defined name calls `INDIRECT`, `CELL`, `INFO`, `_xlfn.SHEET`, `_xlfn.SHEETS`, or `_xlfn.FORMULATEXT`, or a defined name calls a macro function `GET.*`, such as `GET.WORKBOOK` or `GET.DOCUMENT`, and the user adds, renames, or moves a sheet
- **THEN** the cached results are marked stale, and the saved workbook asks for a full calculation on load

### Requirement: Rename a sheet

The store SHALL rewrite every reference to a renamed sheet in formulas, defined names, internal hyperlink locations, and hyperlink relationship targets that are a fragment (`#Data!A1`). It SHALL refuse the rename when the workbook holds a chart, a pivot table, a sheet that is not a worksheet, or a relationship outside the allowlist of design D4 of the first change. It SHALL also refuse the rename when other content that it does not rewrite names the sheet.

#### Scenario: A formula on another sheet

- **WHEN** `Sheet2!A1` holds `=Sheet1!B2` and the user renames `Sheet1` to `Data 1`
- **THEN** `Sheet2!A1` holds `='Data 1'!B2`

#### Scenario: A formula that already names the new name of a rename

- **WHEN** `Sheet1!A1` holds `=Data!A1` with a cached error, and the user renames `Sheet2` to `Data`
- **THEN** the saved workbook asks for a full calculation on load, so the cached result is not shown as current

#### Scenario: A chart in the workbook

- **WHEN** the workbook holds a chart that does not name the sheet, and the user renames the sheet
- **THEN** the store refuses the edit with `structure-blocked`, with the chart part as the subject

#### Scenario: The sheet menu in a workbook with a chart

- **WHEN** the workbook holds a chart, a pivot table, a sheet that is not a worksheet, or a workbook or package relationship outside the allowlist, and the user opens the menu of a sheet tab
- **THEN** Rename and Delete are disabled with the reason `sheetEditBlocked`, which names the part. F2 or a double-click starts no rename. Each worksheet tab shows that reason in its title and its accessible description

#### Scenario: The tab of a chartsheet

- **WHEN** the workbook holds a chartsheet
- **THEN** its tab is disabled, opens no menu, and shows the reason `notWorksheet` in its title and its accessible description: "This sheet is not a worksheet, or its part is missing. The editor cannot show, rename, delete, or move it."

#### Scenario: A drawing that names the sheet

- **WHEN** a shape on another sheet has a `textlink` to `Sheet1!$A$1`, and the user renames `Sheet1`
- **THEN** the store refuses the edit with `structure-blocked`, with the drawing part as the subject

### Requirement: Delete a sheet

The store SHALL remove a deleted sheet with its part, relationships, and override, turn each reference to it, with its address, into `#REF!` (a hyperlink whose relationship target names it takes `#REF!` as its location), remove its sheet-scoped defined names, and SHALL refuse to delete the last visible sheet.

#### Scenario: A reference to the deleted sheet

- **WHEN** `Sheet2!A1` holds `=Sheet1!B2` and the user deletes `Sheet1`
- **THEN** `Sheet2!A1` holds `=#REF!`

#### Scenario: A range operand

- **WHEN** `Sheet2!A1` holds `=Sheet1!A1:Sheet1!B2` and the user deletes `Sheet1`
- **THEN** `Sheet2!A1` holds `=#REF!`

#### Scenario: A stored sheet name longer than 31 characters

- **WHEN** a file stores a sheet name of 40 characters, and the user deletes or moves another sheet, or inserts a row
- **THEN** the store refuses the edit with `sheet-name-too-long`, with the full stored name as the subject, and a rename of that sheet to a shorter name is applied

#### Scenario: A custom workbook view of the sheet

- **WHEN** a custom workbook view has the `sheetId` of the sheet as `@activeSheetId`, and the user deletes the sheet
- **THEN** the store refuses the edit with `structure-blocked`

#### Scenario: Undo a delete in a workbook with a calculation chain

- **WHEN** the workbook has formulas and a calculation chain, and the user deletes a sheet, undoes, and saves
- **THEN** the saved workbook has no relationship or content type override to the dropped chain

#### Scenario: The last visible sheet

- **WHEN** the user deletes the only visible sheet
- **THEN** the store refuses the edit with `last-sheet`

### Requirement: Move a sheet

The store SHALL move a sheet to a new position and keep `@localSheetId` and the active tab on their sheets. It SHALL refuse with `sheet-range-blocked` a move that changes the sheets between the ends of a 3-D reference or which end comes first in tab order, and a delete of the start or end sheet of one. The editor SHALL refuse a move of a sheet that is not a worksheet with `structure-blocked`, with the sheet name as the subject.

#### Scenario: A sheet-scoped name

- **WHEN** a defined name is scoped to the second sheet and the user moves that sheet to the front
- **THEN** the name has `@localSheetId` 0

#### Scenario: A workbook with no view

- **WHEN** the workbook has no `x:bookViews`, so its active tab is the first sheet, and the user moves the first sheet to the end
- **THEN** the workbook has one `x:workbookView` with `@activeTab` 2 and no `@firstSheet`

#### Scenario: A move outside a 3-D reference

- **WHEN** `Sum!A1` holds `=SUM(Jan:Mar!A1)` and the user moves the sheet `Notes`, which is not between `Jan` and `Mar`, to the front
- **THEN** the sheet moves, and `Sum!A1` keeps its text

#### Scenario: A move into a 3-D reference

- **WHEN** `Sum!A1` holds `=SUM(Jan:Mar!A1)` and the user moves `Notes` between `Jan` and `Feb`
- **THEN** the store refuses the edit with `sheet-range-blocked`, with `Sum!A1` as the subject

#### Scenario: A swap of the ends of a 3-D reference

- **WHEN** `Sum!A1` holds `=SUM(Jan:Feb!A1)` and the user moves `Feb` before `Jan`
- **THEN** the store refuses the edit with `sheet-range-blocked`, with `Sum!A1` as the subject

#### Scenario: A move of a chartsheet

- **WHEN** the sheets are the worksheets `A` and `B` and the chartsheet `Chart1`, and a host calls `moveSheet` for `Chart1`
- **THEN** the editor refuses the edit with `structure-blocked`, with `Chart1` as the subject, and a move of `A` past `Chart1` applies

#### Scenario: A delete inside a 3-D reference

- **WHEN** `Sum!A1` holds `=SUM(Jan:Mar!A1)` and the user deletes `Feb`
- **THEN** the sheet goes, and `Sum!A1` keeps its text

### Requirement: Copy and paste in the workbook

The editor SHALL paste a copied range at the active cell with its values, formulas moved by the offset, and cell formats, as one transaction and one undo step. The editor SHALL refuse the paste with `paste-reference-outside` when a moved reference would leave the sheet.

#### Scenario: A relative formula

- **WHEN** the user copies `B1` that holds `=A1` and pastes it at `B3`
- **THEN** `B3` holds `=A3` with the cell format of `B1`

#### Scenario: A reference that leaves the sheet

- **WHEN** the user copies `B2` that holds `=A1` and pastes it at `A2`
- **THEN** the editor refuses the paste with `paste-reference-outside`, and `A2` does not change

#### Scenario: A sheet range that reads two ways

- **WHEN** the user copies `C1` that holds `=SUM(A1:S!B2)` and pastes it at `D4`
- **THEN** the editor refuses the paste with `paste-unreadable-formula`, and `D4` does not change

#### Scenario: A same-sheet cut of a sheet range that reads two ways

- **WHEN** the user cuts `C1` that holds `=SUM(A1:S!B2)` and pastes it at `D4` of the same sheet
- **THEN** the editor refuses the paste with `paste-unreadable-formula`, because the check of typed input refuses the text, and `C1` keeps its formula

#### Scenario: A copied formula that fails the check of typed input

- **WHEN** `S!A1` holds the formula `1+`, and the user copies it and pastes it
- **THEN** the editor refuses the paste with `paste-unreadable-formula`, with `S!A1` as the subject

#### Scenario: Undo a paste

- **WHEN** the user pastes a range and then undoes once
- **THEN** every target cell has its value and cell format from before the paste

### Requirement: Copy expiry at a change of a sheet or a cell format

The editor SHALL refuse the paste with `copy-expired` when a copied cell format changed after the copy, or when a sheet that a copied formula names was renamed, moved, or deleted after the copy. A copy with a name that the source sheet scopes SHALL expire at a rename or a delete of the source sheet only for a paste on another sheet. The editor SHALL refuse every later paste of an expired copy with the same code.

#### Scenario: A sheet rename after the copy

- **WHEN** `T!A1` holds `=S!A1`, the user copies it, renames `S`, and pastes
- **THEN** the editor refuses the paste with `copy-expired`

#### Scenario: A delete and an add of the named sheet after the copy

- **WHEN** `S!A1` holds `=Data!A1`, the user copies it, deletes `Data`, adds a sheet `Data`, and pastes
- **THEN** the editor refuses the paste with `copy-expired`

#### Scenario: A change of a copied cell format after the copy

- **WHEN** the user makes `A1` bold, copies it, undoes the bold edit, makes `B1` italic, and pastes
- **THEN** the editor refuses the paste with `copy-expired`

#### Scenario: A rename of the source sheet after the copy of a scoped name

- **WHEN** `S!A1` holds `=Rate`, `S` scopes the name `Rate`, the user copies the cell, renames `S`, and pastes on the sheet `T`
- **THEN** the editor refuses the paste with `copy-expired`

#### Scenario: A paste on the source sheet after the copy of a scoped name

- **WHEN** `S!A1` holds `=Rate`, `S` scopes the name `Rate`, the user copies the cell, renames `S`, and pastes on the renamed sheet
- **THEN** the target cell holds `=Rate`

### Requirement: Copy expiry at a row or column edit

A copy with formulas SHALL expire with `copy-expired` at an insert or a delete of rows or columns on any sheet, and at an undo or a redo of such an edit. A copy with no formula SHALL NOT expire at these edits.

#### Scenario: A row insert after the copy

- **WHEN** `B5` holds `=$A$5`, the user copies it, inserts a row above it, and pastes at `D20`
- **THEN** the editor refuses the paste with `copy-expired`, and `D20` does not change

#### Scenario: A column delete after the copy

- **WHEN** `B5` holds `=$A$5`, the user copies it, deletes column `A`, and pastes at `D20`
- **THEN** the editor refuses the paste with `copy-expired`, and `D20` does not change

#### Scenario: A row insert on another sheet after the copy

- **WHEN** `S!B5` holds `=$A$5`, the user copies it, inserts a row on the sheet `T`, and pastes on `S`
- **THEN** the editor refuses the paste with `copy-expired`, although the formula does not name `T`

#### Scenario: An undo of a row insert after the copy

- **WHEN** the user copies a formula, inserts a row, undoes the insert, and pastes
- **THEN** the editor refuses the paste with `copy-expired`

#### Scenario: A redo of a row insert after the copy

- **WHEN** the user inserts a row, undoes the insert, copies a formula, redoes the insert, and pastes
- **THEN** the editor refuses the paste with `copy-expired`

#### Scenario: A row insert after a copy of values

- **WHEN** `A5` holds the number `7`, the user copies it, inserts a row, and pastes at `D20`
- **THEN** `D20` holds `7`

### Requirement: Copy expiry at a sheet edit with a span

A copy with a formula that holds a reference across a span of sheets SHALL expire with `copy-expired` at an add, a rename, a move, or a delete of any sheet, and at an undo or a redo of such an edit.

#### Scenario: A move of a sheet outside a span after the copy

- **WHEN** `S!A1` holds `=SUM(Jan:Mar!A1)`, the user copies it, moves the sheet `Other` that the span does not hold, and pastes
- **THEN** the editor refuses the paste with `copy-expired`, although the formula does not name `Other`

#### Scenario: An add of a sheet after the copy of a span

- **WHEN** `S!A1` holds `=SUM(Jan:Mar!A1)`, the user copies it, adds a sheet, and pastes
- **THEN** the editor refuses the paste with `copy-expired`

#### Scenario: A rename of a sheet inside a span after the copy

- **WHEN** `S!A1` holds `=SUM(Jan:Mar!A1)`, the user copies it, renames the sheet `Feb`, and pastes
- **THEN** the editor refuses the paste with `copy-expired`

#### Scenario: A delete of a sheet inside a span after the copy

- **WHEN** `S!A1` holds `=SUM(Jan:Mar!A1)`, the user copies it, deletes the sheet `Feb`, and pastes
- **THEN** the editor refuses the paste with `copy-expired`

#### Scenario: An undo of a sheet delete after the copy of a span

- **WHEN** the user deletes the sheet `Feb`, copies `S!A1` that holds `=SUM(Jan:Mar!A1)`, undoes the delete, and pastes
- **THEN** the editor refuses the paste with `copy-expired`

#### Scenario: A redo of a sheet delete after the copy of a span

- **WHEN** the user deletes the sheet `Feb`, undoes the delete, copies `S!A1` that holds `=SUM(Jan:Mar!A1)`, redoes the delete, and pastes
- **THEN** the editor refuses the paste with `copy-expired`

### Requirement: Edits that keep a copy

A cell edit SHALL NOT end a copy, also on a sheet that a copied formula names. An edit that the editor refuses, or that changes nothing, SHALL NOT end a copy.

#### Scenario: A cell edit on a named sheet after the copy

- **WHEN** `T!A1` holds `=S!A1`, the user copies it, types a value in `S!B1`, and pastes at `T!D20`
- **THEN** `T!D20` holds `=S!D20`

#### Scenario: A rename of a sheet that no copied formula names

- **WHEN** the user copies `B1` that holds `=A1+1`, renames another sheet, and pastes
- **THEN** the paste writes `=A2+1` at `B2`

#### Scenario: A move of a sheet out of a span after the copy

- **WHEN** `S!A1` holds `=SUM(Jan:Mar!A1)`, the user copies it, tries to move the sheet `Feb` out of the span, and pastes at `D20`
- **THEN** the editor refuses the move with `sheet-range-blocked`, the sheet order does not change, and `D20` holds `=SUM(Jan:Mar!D20)`

#### Scenario: An edit that changes nothing after the copy

- **WHEN** `B5` holds `=$A$5`, the user copies it, deletes the empty column `F`, and pastes at `D20`
- **THEN** `D20` holds `=$A$5`

### Requirement: Outline of a pending copy or cut

The editor SHALL mark the source range of a pending copy or cut with a dashed outline in the editor overlay. The outline SHALL NOT change the workbook, the canvas paint, or a print. The first Escape on the grid SHALL end a pending copy or cut on any sheet, and the live region SHALL say so. The next Escape SHALL go on to its next use.

#### Scenario: A copy shows the outline

- **WHEN** the user copies `A1:B2`
- **THEN** a dashed outline marks `A1:B2` on the sheet of the copy, and no other sheet shows it

#### Scenario: A paste that leaves the source as it was

- **WHEN** the user copies `A1` and pastes it at `C1`
- **THEN** the outline stays, and a next paste writes the copied cell again

#### Scenario: A paste that changes the source

- **WHEN** the user copies a rich text cell and pastes it at its own source
- **THEN** the outline ends, because the cell holds plain text after the paste

#### Scenario: An edit ends the outline

- **WHEN** the user copies `A1` and edits another cell
- **THEN** the outline ends

#### Scenario: Escape on a copy

- **WHEN** the user copies a formula, presses Escape on the grid, and pastes
- **THEN** the first Escape ends only the outline, and the paste writes the copied formula

#### Scenario: Escape on a cut

- **WHEN** the user cuts `A1`, presses Escape on the grid, and pastes at `C1`
- **THEN** the cut ends, `A1` does not change, and the paste writes only the copied text

#### Scenario: The next Escape goes on

- **WHEN** the user copies `A1` and presses Escape on the grid two times
- **THEN** the second Escape goes on to its next use, and Tab then leaves the grid

#### Scenario: Escape on another sheet after a copy

- **WHEN** the user copies `S!A1`, shows the sheet `T`, presses Escape on the grid, and pastes
- **THEN** the outline ends, the live region says so, and the paste writes the copied cell

#### Scenario: Escape on another sheet after a cut

- **WHEN** the user cuts `S!A1`, shows the sheet `T`, and presses Escape on the grid
- **THEN** the cut ends, the live region says so, and `S!A1` does not change

#### Scenario: A print

- **WHEN** a copy is pending and the user prints the sheet
- **THEN** the print shows no outline

### Requirement: Names in a copy to another sheet

On another sheet, the editor SHALL give a copied name that the source sheet scopes the source sheet prefix, and the copy SHALL then depend on the source sheet for a paste on another sheet. It SHALL refuse with `paste-name-scope` a paste on another sheet of a function name that the source sheet scopes. It SHALL refuse the paste with `paste-name-scope` when a copied or cut formula pasted on another sheet holds a name with no sheet prefix that the target sheet scopes and the source sheet does not.

#### Scenario: A copy of a name that the source sheet scopes

- **WHEN** `S1!A1` holds `=Rate*2+B1`, `S1` scopes `Rate`, and the user copies it and pastes it at `S2!A3`
- **THEN** the pasted cell holds `='S1'!Rate*2+B3`

#### Scenario: A copy of a name that only the target sheet scopes

- **WHEN** `S1!A1` holds `=nm+1`, `S2` scopes `nm`, and the user copies it and pastes it on `S2`
- **THEN** the editor refuses the paste with `paste-name-scope`

### Requirement: Cut

The editor SHALL clear the source of a cut range in the paste transaction, and SHALL refuse the paste of a cut with `cut-expired` when the workbook changed after the cut.

#### Scenario: Cut and paste

- **WHEN** the user cuts `A1` and pastes it at `C1`
- **THEN** `C1` holds the value and cell format of `A1`, and `A1` is empty with the default cell format

#### Scenario: Cut to another sheet

- **WHEN** `S1!B1` holds `=A1`, and the user cuts it and pastes it on `S2`
- **THEN** the pasted cell holds `='S1'!A1`, and a name that `S1` scopes gets the prefix `'S1'!`, because a sheet name with a digit is quoted

#### Scenario: A name that the target sheet scopes

- **WHEN** the cut formula holds a name with no sheet prefix that `S2` scopes and `S1` does not
- **THEN** the editor refuses the paste with `paste-name-scope`, because the name would bind to the name of `S2`

#### Scenario: An edit after the cut

- **WHEN** the user cuts `A1`, edits another cell, and pastes the cut at `C1`
- **THEN** the editor refuses the paste with `cut-expired`, and `A1` and `C1` do not change, and a second paste of the same cut is refused with `cut-expired`

#### Scenario: Another workbook opens

- **WHEN** the user cuts a range, opens another workbook, and pastes
- **THEN** no cell of the other workbook is cleared, and the paste takes the plain-text lane

### Requirement: Paste from outside

The editor SHALL read pasted plain text as tab-separated rows of typed input, and SHALL refuse a paste that holds only HTML.

#### Scenario: Two rows of text

- **WHEN** the clipboard holds the plain text `1<TAB>a` and `2<TAB>b` on two lines and the user pastes at `A1`
- **THEN** `A1` is the number 1, `B1` is the text `a`, `A2` is 2, and `B2` is `b`

#### Scenario: HTML only

- **WHEN** the clipboard holds only `text/html`
- **THEN** the editor refuses the paste with `paste-html`

### Requirement: Clipboard text export

The editor SHALL write a copied range as tab-separated shown text, with an apostrophe before the shown text of any value that starts with `=`, `+`, `-`, `@`, tab, or CR, except the shown text of a number or a date that is a plain signed number, such as `-1,234.50` or `-$1,234.50`. A tab or a line break inside the shown text of a cell SHALL become a space (approximation `copied-line-breaks`). A date whose shown text does not read back to the same day SHALL give its ISO 8601 edit text.

#### Scenario: A text that looks like a formula

- **WHEN** the user copies a cell with the text `=1+1`
- **THEN** the plain text on the clipboard is `'=1+1`

#### Scenario: A date that does not read back

- **WHEN** the user copies a cell that holds 1 June 2035 with the format `mm-dd-yy`
- **THEN** the plain text on the clipboard is `2035-06-01`, not `06-01-35`

#### Scenario: A text with a line break

- **WHEN** the user copies a cell with the text `a` and `b` on two lines
- **THEN** the plain text on the clipboard is `a b`, on one row

### Requirement: Clipboard limits

The editor SHALL refuse a copy or a paste that meets a merged range, a paste whose target leaves the grid, and a paste of more cells than the format limit. A copy of more than 16,384 cells, of an array formula, or of a data table SHALL be refused with a copy code.

#### Scenario: A merged target

- **WHEN** the target of a paste meets a merged range
- **THEN** the editor refuses the paste with `merged-range`

#### Scenario: A copy of an array formula

- **WHEN** the user copies a range that holds an array formula
- **THEN** the editor refuses the copy with `copy-array-formula`

#### Scenario: A copy of part of an array formula range without its master cell

- **WHEN** `B1` holds an array formula with the range `B1:C3`, and the user copies `B2:D4`
- **THEN** the editor refuses the copy with `copy-array-formula`, with the sheet and the array formula range `B1:C3` as the subject

### Requirement: Column and row resize

The editor SHALL set a column width or a row height from a drag of a heading edge, as one transaction on pointer up, and SHALL commit nothing for a cancelled drag or a drag that ends at the size it started with.

#### Scenario: Drag a column edge

- **WHEN** the user drags the right edge of the heading of column B to the right and releases
- **THEN** column B has a custom width that matches the new edge, and one undo restores the old width

#### Scenario: Escape during a drag

- **WHEN** the user presses Escape during a drag
- **THEN** no width changes

#### Scenario: A row drag with the whole sheet selected

- **WHEN** the whole sheet is selected and the user drags the bottom edge of a row heading
- **THEN** the editor refuses the edit with `too-many-rows`

#### Scenario: A resize under a drawing that keeps its size

- **WHEN** a `xdr:twoCellAnchor` with `@editAs="oneCell"` covers columns B to D, and the user drags the right edge of the heading of column C
- **THEN** the store refuses the edit with `resize-anchor`, with the shape name as the subject, and no width changes

#### Scenario: A resize under a control that keeps its size

- **WHEN** an embedded control anchor with no `@sizeWithCells` covers columns B to D, and the user drags the right edge of the heading of column C
- **THEN** the store refuses the edit with `resize-anchor`, and no width changes

#### Scenario: A resize on a sheet whose drawing part is kept as bytes

- **WHEN** the drawing part of a worksheet names a prefix in `mc:Ignorable` that it does not declare, so the reader keeps it as bytes, and the user drags the right edge of any column heading of that sheet
- **THEN** the store refuses the edit with `resize-anchor`, with no subject and with `causeUnreadable`, and no width changes. The message says that the editor cannot read the drawing of the sheet

#### Scenario: A column width and a call to CELL

- **WHEN** a cell formula or a defined name calls `CELL` or a macro function `GET.*`, and the user sets the width of a column
- **THEN** the cached results are marked stale, because `CELL("width")` and `GET.CELL` read the width, and the saved workbook asks for a full calculation on load

#### Scenario: A row height and a call to GET.CELL

- **WHEN** a defined name calls a macro function `GET.*`, such as `GET.CELL`, and the user sets the height of a row
- **THEN** the cached results are marked stale, because `GET.CELL` reads the row height, and the saved workbook asks for a full calculation on load

#### Scenario: A row height and a call to CELL

- **WHEN** a cell formula calls `CELL` and no formula calls a macro function `GET.*`, and the user sets the height of a row
- **THEN** the cached results stay current, because no result of `CELL` depends on a row height

### Requirement: Column and row fit to content

The editor SHALL fit a column or a row to its content on a double-click on its heading edge, as one transaction and one undo step. A selection of whole columns or rows that holds the edge SHALL fit each of them. A hidden column or row SHALL keep its size, and a fit that changes no size SHALL commit nothing.

#### Scenario: Fit a column to its widest text

- **WHEN** column B holds a text wider than the column, and the user double-clicks the right edge of the heading of column B
- **THEN** column B has the width at which the text shows in full at each zoom step from 50% and at the current zoom
- **AND** column B has `@customWidth`, and one undo restores the old width

#### Scenario: Fit a row whose height was set by hand

- **WHEN** row 1 has `@ht="60"` and `@customHeight="1"`, and the user double-clicks the bottom edge of the heading of row 1
- **THEN** row 1 has the automatic height of its content as `@ht`, without `@customHeight`, so a later font change or value edit fits it again

#### Scenario: Fit a band of whole columns with a hidden column

- **WHEN** columns B to D are selected, column C is hidden, and the user double-clicks the right edge of the heading of column D
- **THEN** columns B and D fit their content, column C keeps its size and stays hidden, and one undo restores both widths

#### Scenario: Fit a column that fits already

- **WHEN** the user double-clicks the right edge of a column heading a second time
- **THEN** no width changes, and the editor adds no undo step

### Requirement: Protection and view mode

The store SHALL refuse with `locked` an edit of this change that the protection does not allow: a paste that gives a cell a new cell format on a sheet that locks formatting cells (lock `format`), a column width or a row height on a sheet that locks it (lock `size`), and a sheet op on a workbook that locks its structure (lock `workbook`). In view mode, the editor SHALL refuse every command of this change that writes with `view-mode`, and SHALL keep copy available.

#### Scenario: A paste of a new format on a sheet that locks formatting

- **WHEN** the sheet is protected without `formatCells="0"`, and the user pastes a copy whose cell format differs from the format of the target cell
- **THEN** the store refuses the paste with `locked` and the lock `format`, and the cell keeps its value and format

#### Scenario: A paste in the same format on a sheet that locks formatting

- **WHEN** the sheet is protected without `formatCells="0"`, and the user pastes a copy of an unlocked cell into an unlocked cell with the same cell format
- **THEN** the paste applies

#### Scenario: A resize on a sheet that locks column formatting

- **WHEN** the sheet is protected without `formatColumns="0"`, and a script sets the width of a column
- **THEN** the store refuses the edit with `locked` and the lock `size`, and the column heading shows no resize edge

#### Scenario: A sheet op on a workbook that locks its structure

- **WHEN** `x:workbookProtection/@lockStructure` is true, and the user adds, renames, deletes, or moves a sheet
- **THEN** the store refuses the op with `locked` and the lock `workbook`, and the sheet tabs show the add button and the edit rows of the sheet menu as disabled

#### Scenario: View mode

- **WHEN** the editor is in view mode
- **THEN** the add button and the edit rows of the sheet menu are disabled, a heading edge starts no resize, a cut and a paste are refused with `view-mode`, and a copy applies

### Requirement: Cell protection in a cut and a paste

The store SHALL refuse with `locked` and the lock `cell` a paste into a locked cell, and a new cell format that changes the `@locked` or `@hidden` protection of a cell on any protected sheet, whatever `@formatCells` states. The editor SHALL refuse at the cut, with `locked` and the command `cut`, a cut whose source the paste could not clear: a locked cell that holds content (lock `cell`), or a sheet that locks formatting (lock `format`).

#### Scenario: A paste into a locked cell

- **WHEN** the sheet is protected, and the user pastes a copy into a locked cell
- **THEN** the store refuses the paste with `locked` and the lock `cell`, and the message names the paste

#### Scenario: A new cell format that unlocks a locked cell

- **WHEN** the sheet is protected with `formatCells="0"`, and a script gives a locked cell an unlocked cell format and then writes the cell in one transaction
- **THEN** the store refuses the transaction with `locked` and the lock `cell`, and the cell keeps its value and format

#### Scenario: A cut of a locked cell that holds content

- **WHEN** the sheet is protected, and the user cuts a locked cell that holds content
- **THEN** the editor refuses the cut with `locked`, the lock `cell`, and the command `cut`, and the cell keeps its content

#### Scenario: A cut on a sheet that locks formatting

- **WHEN** the sheet is protected without `formatCells="0"`, and the user cuts an unlocked cell
- **THEN** the editor refuses the cut with `locked`, the lock `format`, and the command `cut`

### Requirement: Cell protection in a clear

The paste of a cut SHALL clear only the source cells that it does not write. On a protected sheet, a clear of formats SHALL keep the `@locked` and `@hidden` protection of each cell, with format 0 and the protection of the cell.

#### Scenario: A cut moved over its own cells

- **WHEN** the sheet is protected with `formatCells="0"`, the user cuts the unlocked cells `B1:C1`, and pastes at `C1`
- **THEN** the paste applies, `C1:D1` hold the moved values, `B1` holds nothing, and each of the three cells stays unlocked

#### Scenario: A clear of formats keeps the protection

- **WHEN** the sheet is protected with `formatCells="0"`, and a clear of formats meets an unlocked cell whose format is not format 0
- **THEN** the cell gets format 0 with `locked="0"`, a record that the clear appends when none holds it, and the cell stays writable

#### Scenario: A clear of formats and a call to CELL

- **WHEN** `A1` holds `=CELL("format",B2)`, and a clear of formats removes the cell format of `B2`
- **THEN** the cached results are marked stale, and the saved workbook asks for a full calculation on load

### Requirement: Sheet tab moves

The sheet tabs SHALL move a sheet past the next visible tab, so a hidden sheet between two tabs never takes a move.

#### Scenario: A hidden sheet between two tabs

- **WHEN** the sheets are `A`, hidden `H`, and `B`, and the user moves `A` right
- **THEN** the sheet order is `H`, `B`, `A`
