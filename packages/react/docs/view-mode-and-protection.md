# View mode and protection

This guide describes view mode, sheet protection, and workbook protection. For setup, see the [package README](../README.md).

## View mode

To show a workbook without letting the user change it, pass `mode="view"`. The default is `"edit"`. A change of `mode` applies without a reload:

<!-- sample: guide-view-mode -->

```tsx
<XlsxEditor document={bytes} mode={canEdit ? 'edit' : 'view'} style={{ height: 600 }} />
```

In view mode, the editor refuses every edit with the reason `view-mode`. Typing, F2, Backspace, Delete, and a double-click open no cell editor. The editing commands of the editor instance return the refusal, and the toolbar disables its edit controls with the reason `viewMode`. The sheet tabs disable the add button and the edit rows of the sheet menu with the same reason, and start no rename. A heading edge starts no resize. Cut and paste are refused with `view-mode`.

You can still select cells, copy, switch sheets, zoom, and save. If a cell edit is open when the mode changes to `"view"`, the editor commits it first. If the editor cannot store that edit, it stays in edit mode with the cell editor open and shows the reason. `setMode()` on the editor instance returns `false` in that case.

The `mode` prop applies when it changes. A call to `setMode()` on the editor instance stays until the next change of the prop. If the prop change to `"view"` cannot apply, it waits: the status strip shows that view mode waits, and the switch applies when the user commits or cancels the edit. A call to `setMode()` during the wait ends the wait. To show the wait in a layout of your own, read `useViewModeWaits()`, which returns `true` while the switch waits. On the editor instance, call `setMode()`, and read the mode in `snapshot().mode`.

## Sheet protection

On a protected sheet (`x:sheetProtection` with `sheet` set to true, ECMA-376 Part 1, 18.3.1.85), the editor refuses some edits with the reason `locked`. It refuses a value edit on a locked cell. It refuses a format edit unless the protection allows formatting cells (`formatCells` set to false). A paste that gives a cell another cell format is a format edit. A plain-text paste of a typed form, such as `12%`, is not refused for its format: it stores the number and keeps the cell format. A new cell format that changes the protection of a cell is refused with the rule `cell`, also when the protection allows formatting cells.

It refuses a column width unless the protection allows formatting columns (`formatColumns`), and a row height unless it allows formatting rows (`formatRows`). A heading edge whose size the protection does not allow shows no resize edge. It refuses an insert or a delete of rows or columns unless the protection allows it (`insertRows`, `deleteRows`, `insertColumns`, and `deleteColumns`). Each of these attributes is true by default, and true locks the action (`CT_SheetProtection`).

A cell is locked unless its cell format unlocks it with `x:protection/@locked` set to false (ECMA-376 Part 1, 18.8.33). A cell format with no protection of its own takes the protection of its cell style. A cell that does not exist takes the format of its row or column. When the protection allows deleting rows or columns, a deletion also deletes locked cells.

It refuses a merge or an unmerge of cells with the rule `merge`, whatever the protection allows. A cut is refused when its paste could not clear the source. That is a locked cell that holds content, or a sheet that does not allow formatting cells. The refusal has the command `cut`. A clear of formats keeps whether each cell is locked or hidden. The editor has no paste of values only, so a cut on a sheet that does not allow formatting cells is always refused (`editing.lockedRefused.cutFormats`).

Typing on a locked cell opens no cell editor, and the status strip shows that the sheet is protected. The snapshot holds what the protection allows in `protection`, a `SheetProtection`. Its fields are true when an action is allowed: `allowFormatCells`, `allowFormatColumns`, `allowFormatRows`, `allowInsertRows`, `allowDeleteRows`, `allowInsertColumns`, and `allowDeleteColumns`. `activeCellLocked` is true when the active cell is locked. `selectionContentLocked` is true when clearing the selection would change a locked cell that holds content, so Clear contents and Clear all are disabled. The editor does not enforce `protectedRanges`.

## Workbook protection

When the workbook protection (`x:workbookProtection`, ECMA-376 Part 1, 18.2.29) sets `lockStructure` to true, the editor refuses a sheet add, rename, delete, or move with `locked` and the rule `workbook`. The sheet tabs then disable those actions with the reason `workbookLocked`. The snapshot holds the workbook protection in `workbookProtection`, a `WorkbookProtection` whose `allowStructure` is false then. The `lockStructure` attribute is false by default. The editor never reads or checks the password, and it keeps the protection as the file holds it.
