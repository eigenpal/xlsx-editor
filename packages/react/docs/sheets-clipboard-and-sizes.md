# Sheets, clipboard, and sizes

This guide describes sheet commands, copy and paste, and column widths and row heights. For setup, see the [package README](../README.md).

## Sheets

Each sheet is an `x:sheet` in the `x:sheets` of the workbook part, and the order of those elements is the tab order. To add a sheet, click **Add a sheet** after the sheet tabs. The new sheet goes after the shown sheet. To rename a sheet, double-click its tab, or press F2 on the tab, then type the name and press Enter. Press Escape to keep the old name. If the editor refuses the name after Enter, the name stays in the field so you can fix it. If the field loses focus, a refused name ends the rename and keeps the old name.

To open the sheet actions, right-click a tab, or press Shift+F10 or the Menu key on it. The menu has **Rename**, **Delete**, **Move left**, and **Move right**. **Move left** and **Move right** move the sheet past the next tab that you can see. If the editor cannot make the change, such as a name that another sheet has, it shows the reason.

A sheet add, rename, delete, or move is one undo step. Ctrl+Z (Cmd+Z on macOS) on a focused tab undoes it. In an open sheet menu, the key closes the menu and undoes the last step.

A rename changes each formula, defined name (`x:definedName`), and hyperlink (`x:hyperlink/@location`) that names the sheet. A delete turns each reference to the sheet into `#REF!`, and removes the defined names that belong to the sheet. A rename and a delete are refused with the reason `structure-blocked` in these cases:

- The workbook holds a chart, a pivot table, or a sheet that is not a worksheet, such as a chartsheet.
- A table, conditional formatting, data validation, a drawing, or a control names the sheet.
- The workbook relates to a part that can refer to cells, such as data connections or revisions.
- For a delete: the sheet has a related part other than hyperlinks and printer settings, such as comments. A custom workbook view that shows the sheet also refuses a delete.

A `moveSheet()` or a `deleteSheet()` of a sheet that is not a worksheet, such as a chartsheet, is refused with the reason `structure-blocked`. The `about` kind of that refusal is `not-worksheet`, and its message says that the editor can move or delete only a worksheet. View mode, an open input method composition, and the workbook protection refuse such a command first, with their own reasons.

If a formula or a defined name refers to a range of sheets, such as `Sheet1:Sheet3!A1`, some sheet edits are refused with the reason `sheet-range-blocked`. A move is refused when it changes the sheets in that range, or which end comes first. A delete is refused for the first or last sheet of the range. A delete of a sheet between them keeps the reference. Charts never refuse a move. A right-click on a chartsheet tab opens no menu.

A file can give two sheets the same name in different case, such as `Data` and `DATA`. Then a delete or a row or column edit of either is refused with the reason `sheet-name-ambiguous`. So is a rename while a reference names one of them. The sheet titles in `docProps/app.xml` keep their old text.

A file can store a sheet name of more than 31 characters. Then a move or a delete of any sheet, and an insert or delete of rows or columns, is refused with `sheet-name-too-long`, which names that sheet. A rename can shorten the name.

## Copy and paste

To copy cells, select them and press Ctrl+C (Cmd+C on macOS). To cut them, press Ctrl+X. To paste, select the first cell of the target and press Ctrl+V. A paste of cells from the same editor writes their values, formulas, and cell formats, as one undo step. A copied formula changes its relative references by the distance between the source and the target. A paste of cut cells clears the source, and the cut formulas keep the cells that they refer to, also on another sheet.

Formulas that refer to the cut cells keep the old address. Hyperlinks, comments, data validation, and conditional formatting stay with the source cells. The plain text of a copy shows the number format of a true conditional formatting rule, as the grid does. When a layout of the grid runs out of its conditional formatting check budget, the notice for conditional formatting that the grid can leave out shows, and the plain text shows the cell format. A copy or a cut of a range with more conditional formatting than one copy can check is refused with the reason `copy-conditional-limit`. Copy a smaller range.

Rich text pastes as plain text, and an error value such as `#N/A` pastes as text. Text from another program pastes as typed input, one cell for each tab-separated value and one row for each line. A tab or a line break inside a copied cell becomes a space in the text that other programs paste. A date whose shown text does not read back to the same day, such as `2-Jan` or a two-digit year outside 1930 to 2029, copies its ISO 8601 text, such as `2035-06-01`.

The editor refuses these copies and pastes:

- A paste of HTML only.
- A copy or a paste that meets a merged range.
- A copy of more than 16,384 cells, of an array formula, or of a data table. The master of the array formula or data table can be inside `mc:AlternateContent`.
- A copy or a cut of a range with more conditional formatting than one copy can check, with the reason `copy-conditional-limit`.
- A copy of a formula whose text the editor cannot get, and a paste of a formula that it cannot read.
- A copy or a cut of a cell whose formula is inside `mc:AlternateContent`, with the reason `copy-wrapped-formula`. The editor shows that formula but cannot copy it.
- A paste of a formula whose moved reference leaves the sheet.
- A cut of a cell whose only value is in an `mc:AlternateContent` child, with the reason `cut-wrapped-value`. The copy cannot carry that value. A clear of the cell removes it.
- A paste of cut cells after another edit, also an undo or a redo, with the reason `cut-expired`. Cut the cells again.
- A paste of copied formulas after an insert or a delete of rows or columns on any sheet, or after a rename, a move, or a delete of a sheet that a copied formula names, with the reason `copy-expired`. A formula with a reference across a span of sheets, such as `Jan:Mar!A1`, also expires at an add, a rename, a move, or a delete of any sheet. An undo or a redo of such an edit does not bring the copy back. Copy the cells again.
- A paste after a copied cell format changed, with the reason `copy-expired`. Copy the cells again.
- On another sheet, a paste of a copied or cut formula with a name with no sheet prefix, with the reason `paste-name-scope`. This applies when the target sheet scopes the name and the source sheet does not. It also applies to a function name that the source sheet scopes.

A dashed outline marks the source cells of a pending copy or cut. The outline is not part of the workbook, and it does not print. A paste of cut cells ends it. A copy stays pending after a paste, until the next edit or Escape. An undo or a redo also ends it. Press Escape on the grid of any sheet to end the outline of a pending copy: the live region says so, and a later paste still writes the copied cells. Press Escape on the grid of any sheet to forget a pending cut: the cells stay, the live region says so, and a later paste writes only the copied text. A paste of plain text also ends a pending cut in the same way. A paste of a copy that changes its own source ends the outline, also a paste at the source cells that changes a rich text cell, an error, or a cached formula value. Escape does not forget an expired cut, so a later paste still gives the reason `cut-expired`. A copy says what it copied in the live region. After the first paste of cut cells, a second paste writes a copy. Until you copy or cut again, each later paste of an expired copy is refused with the same reason.

A copied or cut formula pasted on another sheet gives a name that the source sheet scopes the source sheet prefix. For a paste on another sheet, the copy then depends on that sheet. So such a paste after a rename or delete of that sheet is refused with the reason `copy-expired`.

## Column widths and row heights

To change a column width, drag the right edge of the column heading. To change a row height, drag the bottom edge of the row heading. A column is at most 255 characters wide, and a row at most 409 points tall. If whole columns or rows are selected and the dragged one is among them, all of them get the new size. Press Escape during the drag to cancel it. Each drag is one undo step.

A width edit writes `x:col/@width` and `@customWidth` (ECMA-376 Part 1, 18.3.1.13). A height edit writes `x:row/@ht` and `@customHeight` (ECMA-376 Part 1, 18.3.1.73).

An object that keeps its size or position, such as a picture or a control, can refuse a resize. The reason is then `resize-anchor`.

To fit a column or row to its content, double-click the same edge. A column takes the width of its widest shown text, after the number format and in the cell's font. That text must show at each zoom step from 50% and at the current zoom. So the column can be wider than the text needs at one zoom. The extra room is a few pixels for most text, and more for a long run of one repeated character, such as a line of dashes. If you fit at 50% or more and then zoom below 50%, text can be cut and a number can show `#` signs.

These cells do not count: text that wraps, rotated text, justified or distributed text, and text centered across cells. Text of a merged range over more than one column and text in a hidden row do not count either. A column with no text keeps its width. So does a column with more conditional formats than layout checks.

A row takes the height of its tallest font and of its wrapped lines. That height must hold the lines at each zoom step from 50% and at the current zoom. If you zoom below 50% after the fit, the last wrapped line can be cut. A fitted row fits again when you change its font or a value in it. A row with no cells takes the default row height of the sheet.

A row with rotated text keeps its height. Wrapped text in a merged range does not count, so a fit can cut it. A cell in a hidden column counts, as it does for the automatic height. On a sheet whose default row height is set by hand, a row also fits its content. There, a later font or value edit gives the row that default height again.

A double-click commits an open cell edit first, and the fit counts its text. If whole columns or rows are selected and the edge is among them, each one fits its own content. A hidden column or row keeps its size in every fit. Each fit is one undo step. In view mode, or on a protected sheet that does not allow the size change, a double-click fits nothing.

An object that keeps its size or position can refuse the whole fit. It refuses a fit that changes a row height, or that removes a height set by hand. For a column, only a width change counts.

## Commands on the editor instance

The ref is small. Call the sheet, clipboard, and resize commands on the editor instance that `getEditor()` returns:

<!-- sample: guide-sheet-calls -->

```tsx
const editor = ref.current?.getEditor();
editor?.addSheet('Notes');
// Column B is 1: indexes are 0-based. The width is 20 characters.
editor?.setColumnWidth(1, 1, 20);
const copied = editor?.copy();
if (copied?.ok) {
  editor?.paste({ token: copied.content.token, text: copied.content.text });
}
```

`copy()` and `cut()` return a `CopyResult`. When it is `ok`, its `content` is a `ClipboardContent`: the `text/plain` text and a token for the private type `CLIPBOARD_TYPE`. `paste()` takes a `PasteData`, which holds the untrusted data of a paste event.

`setColumnWidth()` takes at most `MAX_COLUMN_WIDTH` characters, and `setRowHeight()` at most `MAX_ROW_HEIGHT` points. One `setRowHeight()` creates at most `MAX_RESIZE_NEW_ROWS` new rows. `selectReference()` reads at most `MAX_GO_TO_LENGTH` characters. Each name is exported from `@xlsx-editor.dev/core/editor`.
