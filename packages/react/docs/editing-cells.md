# Edit cells

This guide describes cell edits, typed values, the formula bar, the name box, and saves. For setup, see the [package README](../README.md).

## Edit a cell

To edit a cell, select it and type, press F2, or double-click it. Press Enter to commit the edit. Text that starts with `=` is a formula in A1 notation. The editor stores a formula (`x:f`) without a cached result (`x:v`), and sets `x:calcPr/@fullCalcOnLoad` to ask the next consumer to recalculate the workbook. If the editor cannot store an edit, it keeps your text and shows the reason.

A typed formula can call only the predefined functions of ECMA-376 Part 1, 18.17.7, in any letter case. A call to another function, such as a future function or a name with a stored-name prefix, is refused with the reason `unsupported-function`. A defined name, a table reference, or a sheet name is not a call. A formula that the workbook holds keeps its stored text, and the formula bar, the cell editor, and the announcement show it as stored, prefixes included. A save writes it unchanged. A commit of an edit to such a formula is refused with `unsupported-function`, because the text calls a prefixed name; cancel the edit to keep the stored formula.

## Typed values

Typed text in a number form stores a number: a percent (`12%`), a currency amount (`$5.50`), a grouped number (`1,234.5`), a negative in parentheses (`(100)`), a fraction with a whole part (`1 1/2`), a date (`2026-11-12`), a time (`13:45`), an elapsed time (`36:00`), or a date and a time (`2026-11-12 13:45`). A date or a time stores its serial number in the date system of the workbook (`x:workbookPr/@date1904`, ECMA-376 Part 1, 18.17.4.1). The `numberLocale` prop sets the separators and the day and month order of a date. A date with no year, such as `3/15`, takes the current year of the host clock. A cell with the General format also takes the number format of the form, in the same undo step. A cell with another number format keeps it. A cell with the Text format keeps your text, but a formula stays a formula.

On a sheet whose protection does not allow formatting cells, or in a workbook whose list of cell formats or number formats is full, or whose styles part the editor cannot read, the cell keeps its format and stores the number. So does a code of the number locale that the editor cannot read. To keep any entry as text, start it with an apostrophe (`'007`). In a date cell (`x:c/@t="d"`), ISO 8601 text such as `2024-03-01` stays a date.

When you edit the cell again, the cell editor shows a date as ISO 8601, a time on a 24-hour clock, an elapsed time in hours, and a percent with `%`. The formula bar shows the same text, also for such cells in a workbook that you open. A currency amount shows as a plain number. A plain number that you type in a percent cell is not scaled.

## Formula bar

You can also edit the active cell in the formula bar. When you type in its content field, the edit opens over the cell, and the grid shows the same text as you type. Press Enter to commit and move down, Tab to commit and move right, Shift with either to move back, Alt+Enter to add a line break, or Escape to cancel. The grid then takes the keyboard. If the editor cannot store the edit, the text and the keyboard stay in the field, and the status strip shows the reason.

A text with more than one line shows a line count. A press on it, or on the gap beside the field, puts the caret at the end of the line that the field shows. With Shift, while the field has the keyboard, the press extends the text selection to the end of that line. While an input method composes, Enter commits nothing. In view mode, the field is read-only.

## Name box

To go to a cell, type a reference in the name box and press Enter: a cell such as `B2`, a range such as `B2:C5`, whole columns such as `C:E`, whole rows such as `3:5`, or a reference with a sheet name such as `Sheet2!A1` or `'My sheet'!A1:B2`. Letters can be in any case. You can also type a defined name whose text is one reference with a sheet name, such as `Data!$A$1:$B$4`. A name that the shown sheet scopes comes before a workbook name. A sheet name before a defined name, such as `Data!Rates`, reads only the names of that sheet.

If the text names no cell or range, the name box keeps your text and the status strip shows the reason. This also happens for a defined name that names a formula, several ranges, or a range on a sheet that is not a visible worksheet, and for a sheet that is not a visible worksheet. Press Escape to drop the typed reference.

## Edit from your own chrome

Your own chrome can do the same through the editor instance:

- `setEditText(text, { composing })` makes `text` the text of the open cell edit. With no open edit, it opens one over the active cell without taking the keyboard, and it returns the refusals of `beginEdit()`, such as `locked`, `view-mode`, and `cell-alternate-content`. Pass `composing: true` while an input method composes in your field, and call it again with `false` when the composition ends, also when the field loses the keyboard or unmounts. Until then, every commit is refused with `composing`. A cancel of the edit, or a key in the cell editor, also ends it. A `text` that is not a string throws a `TypeError`.
- `commitEdit(move)` commits the open edit and moves the active cell `down`, `up`, `right`, or `left`, as Enter and Tab do. With no open edit, it only moves the active cell. When an open edit commits, or the active cell moves with no open edit, the grid takes the keyboard. A `move` that is not one of these values throws a `TypeError`.
- `selectReference(reference)` selects the range that a reference names, or returns the refusal `go-to-reference`, `go-to-name`, or `go-to-sheet`. It reads at most `MAX_GO_TO_LENGTH` (255) characters: a longer text is refused with `go-to-reference`.

The functions `editMoveOf(keys)`, `composesKey(event, composing)`, and `fieldHistoryKey(keys, typing)` give the key rules of the formula bar, so your field acts as the bar does. Their `keys` argument is a `FieldKeys`, which a DOM or React key event satisfies. `refusesEditText(code)` tells whether a refusal refuses the text of the open edit, so the field that shows it can be marked invalid. `refusesReference(code)` tells whether a refusal of `selectReference` refuses the typed reference, not the open edit. Import them, and the types `EditMove` and `EditTextOptions`, from `@xlsx-editor.dev/react`. `contentFieldText(text)` gives the start of a cell text that a content field holds: at most `MAX_CONTENT_FIELD_CHARS` (65,536) UTF-16 code units, with no half of a surrogate pair at its end. If the result is shorter than the text, make the field read-only, as the formula bar does. The edit keeps the complete text.

Inside `XlsxEditor.Root`, this field is a small form of the content field of the formula bar. It edits the active cell, commits and moves, cancels, and passes the history keys:

<!-- sample: guide-cell-field -->

```tsx
import {
  composesKey,
  editMoveOf,
  fieldHistoryKey,
  refusesEditText,
  useEditorSnapshot,
  useXlsxEditor,
} from '@xlsx-editor.dev/react';

function CellField() {
  const editor = useXlsxEditor();
  const snap = useEditorSnapshot();
  const text = snap.editText ?? snap.activeCell?.content ?? '';
  // Mark the field while the engine refuses the text of the open edit.
  const refusal = snap.editing ? snap.refusal : null;
  const invalid = refusal !== null && refusesEditText(refusal.code);
  return (
    <textarea
      rows={1}
      aria-label="Cell"
      aria-invalid={invalid || undefined}
      value={text}
      onChange={(e) => editor?.setEditText(e.currentTarget.value)}
      onKeyDown={(e) => {
        if (!editor || composesKey(e.nativeEvent, false)) return;
        const history = fieldHistoryKey(e, snap.editing);
        const move = editMoveOf(e);
        if (history) {
          e.preventDefault();
          if (history === 'undo') editor.undo();
          else editor.redo();
        } else if (move) {
          e.preventDefault();
          editor.commitEdit(move);
        } else if (e.key === 'Escape') {
          e.preventDefault();
          editor.cancelEdit();
        }
      }}
    />
  );
}
```

The sample is a text area, so it keeps the line breaks of a cell. An edit that opens clears an earlier refusal, so the mark shows only a refusal of this edit. The sample keeps the mark until the next commit or cancel; the formula bar removes it when the text changes. The sample does not report a composition or add a line break on Alt+Enter: the formula bar also calls `setEditText(text, { composing })` for each composition.

## The ref

The ref is small: `load()`, `save()`, `saveInTasks()`, `getEditor()`, `focus()`, and `snapshot()`. Call the editing commands on the editor instance that `getEditor()` returns. The ref differs from the DOCX Editor ref in these ways:

- It has no `exec()`. Each command is a typed method of `XlsxEditorInstance`, and the engine has no command union to dispatch on. To run a toolbar slot, use `useEditorCommand`.
- It has no `getDocumentHandle()`. Read the state with `snapshot()`, and edit through `getEditor()`.
- `save()` is synchronous, as in PPTX Editor. DOCX Editor's `save()` returns a promise, so the save in short tasks has its own name here, and `save()` keeps its synchronous result. `saveInTasks()` writes the same bytes in short tasks, so the page stays responsive during the save of a large workbook. It resolves to a `SavedWorkbook`, or to `null` when the save is refused.
- `load()` returns a promise of the snapshot after the open, or `null` before the editor mounts. Opening a workbook loads fonts and parses the package asynchronously, so the caller can wait for the result. If the file passes a limit of the editor, the snapshot turns `failed` with the code `package`, and `failure.limit` names the limit, such as `xml.maxBytes`. If an XML part did not read, `failure.part` names it, such as `/xl/sharedStrings.xml`. The editor then shows that the file is too large to open, with the part and the limit, as DOCX Editor does for a file that is too large. `failure.detail` is for logs only. If a styles part, a theme part, a drawing, or another relationship part passes a limit, the workbook opens without that part, and a save keeps its bytes. `snapshot().partsPastLimit` lists each such part with its `kind` and its `limit`, as a `PartPastLimit` type, and the editor shows a notice that names the part. Each message shows a part name through `displayName` from `@xlsx-editor.dev/core/editor`: it drops control characters and bidirectional format controls, and cuts the name to 64 characters. If you show `failure.part` or a `PartPastLimit` part in your own message, pass it through `displayName` too. The [workbook limits](https://github.com/eigenpal/xlsx-editor/blob/main/docs/site/content/workbook-limits.mdx) page lists the limits and their values.

## Save the workbook

To get the edited workbook, call `save()` on the ref. The [package README](../README.md#edit-and-save) has a sample.

The bytes keep the kind of the opened package: a workbook (`.xlsx`) stays a workbook, and a template (`.xltx`) stays a template. Name the file with the extension of `snapshot().workbookKind`. If the workbook already has a shared string table, text that an edit adds to it stays in the saved file after an undo of the edit, because the editor never removes items from a table that it opened. If the edit created the table, the undo removes the table, with its relationship and content type.

## Rules for open edits and loads

These rules apply while a cell edit is open, while an input method composes, and while a workbook loads:

- `save()`, `select()`, `setActiveSheet()`, `setMode('view')`, and the format, clear, insert, and delete commands commit an open cell edit first. `enterValue()` without `at`, or with the cell of the open edit, also commits it first.
- If the editor cannot store that edit, these calls do nothing. `save()` returns `null`, `select()` and `setMode()` return `false`, and the commands return the refusal. The editor shows the reason.
- A refused save keeps its refusal in `snapshot().refusal` with `command` set to `save`, and the editor shows and announces "The workbook was not saved", with the reason. A host that tracks saves reads that field. A save that is refused during a load is shown and announced while the load runs, and clears when the load settles. A good save clears an older refusal.
- While an input method composes in the cell editor or the formula bar, `snapshot().composing` is true. Each of these calls is refused with the reason `composing`, as are `undo()`, `redo()`, `commitEdit()`, and `selectReference()`, because the composed text is not converted yet. `setMode('view')` returns `false` then, with no announcement.
- If the surface closes while an input method composes, through `detach()`, `destroy()`, or an unmount of `XlsxEditor.Content`, the editor drops the edit and stores nothing. It sets `snapshot().refusal` to the code `edit-dropped` and notifies the listeners while the surface is still attached, so `onChange` and the status strip see it. After `destroy()`, no listener is left to notify.
- `snapshot().refusal` holds the last refusal until one of these events clears it:
  - An edit that the editor stores.
  - A good save.
  - A copy or a cut that succeeds.
  - A change of the selection, the sheet, or the mode.
  - A cancel of the open cell edit.
  - A cell edit that opens, unless the refusal is `edit-dropped`.
  - The start or the settle of a load.
  - A new rename in the sheet tabs, when the refusal is the reason of an earlier refused rename.

  A new refusal replaces it, and so does `edit-dropped` when the editor drops an open edit with changed text. While a load runs, each new refusal has the code `loading`. To clear the refusal from your own chrome, such as when the user closes the control that it explains, call `dismissRefusal()` on the editor instance. It changes no cell.

- An edit (a typed value or formula, or a paste), a clear, or the clear of `mergeCells({ dropValues: true })` is refused with the reason `shared-formula` in these cases. The editor keeps each formula range whole.
  - The cell is the master of, or inside, an array or data table formula range. Its master can be inside `mc:AlternateContent`. An edit of that wrapped master itself is refused with `cell-alternate-content`.
  - The cell is a shared formula master whose text cannot be read.
  - The cell is a shared formula master, and a child has its `x:f` inside `mc:AlternateContent` in its cell. The editor cannot rewrite that child. A paste over the master is refused, also when it covers that child. A clear that also covers that child is refused with `cell-alternate-content`.
  - The edit would move the formula of a shared child out of the grid.

  Typing on such a cell opens no cell editor, except when the refusal comes from a shared child that would move out of the grid. That case is refused at the commit.

  An edit to any other shared formula master writes each child's own formula first, so the children keep their formulas.

  The refusal's `subject` names the cell or the formula range that blocks the edit, such as `Sheet1!A2:A3`. When a shared child blocks the edit of its master, the subject is the shared range that holds both cells, or the child cell when that range does not hold it. When one cell's formula blocks the edit of a shared formula master, the refusal's `causeCell` names that cell, such as `Sheet1!A3`, and `masterCell` names the master, such as `Sheet1!A2`. The message names both and tells you to leave the master out. That cell is a child inside `mc:AlternateContent`, a child whose formula would move out of the grid, or the master when its text cannot be read. In the last case, `causeUnreadable` is `true`, and the message says that the formula cannot be read.

  A refused clear, paste, or merge has `command` set to `clear`, `clearAll`, `paste`, or `merge`, so its message names that command. A refused commit of an open cell edit has `command: 'commit'`. The cell editor stays open, and the message tells you to cancel the cell edit.

  A cut and its paste follow these rules:

  - A cut runs the checks of a copy first, so a cut and a copy of one range give the same copy reason. For example, a cell whose formula is inside `mc:AlternateContent` gives `copy-wrapped-formula`.
  - A cut whose source the editor cannot clear is refused at the cut, with the code of that clear.
  - Every refusal of a cut has `command: 'cut'`. `refusalText` shows its message from `editing.cutRefused.<code>`, or from `editing.cutRefusedUnnamed.<code>` when the refusal names no cell. A `shared-formula` cut with `causeCell` and `masterCell` shows `editing.formulaCellRefused.cut` or `editing.formulaUnreadableRefused.cut`. Each message says that the cells were not cut.

    `CUT_REFUSAL_CODES` in `@xlsx-editor.dev/core/editor` lists the codes that have their own cut message. A `locked` cut has a protection message, and a cut with another store code, such as `style-limit`, shows `editing.cutRefused.other`. A refusal of the open cell edit that the cut commits first is not a refusal of the cut. It has `command: 'commit'` or no command.

  - The paste of a cut clears the source cells that it does not write. So a paste over part of the source is refused with the code of that clear, such as `locked` or `shared-formula`, when that clear alone is refused.

- A cell can hold data in a form that the editor cannot update. That is other markup compatibility content, such as a formula in `mc:AlternateContent`. It is also an element in another namespace that holds `x:f`, `x:v`, or `x:is`. Typing on that cell opens no cell editor.

  A value edit and a paste onto it are refused with the reason `cell-alternate-content`. So is a clear of it while it holds a value, and a cut of it, except a cut of a cell whose formula is inside `mc:AlternateContent`. That cut is refused with `copy-wrapped-formula`, as its copy is.

  The refusal's `subject` names the sheet and the cell, such as `Data!B5`. Such a refusal of a clear, a cut, a paste, or the clear of `mergeCells({ dropValues: true })` has `command` set to `clear`, `clearAll`, `cut`, `paste`, or `merge`, so its message names that command. A refused commit of an open cell edit has `command: 'commit'`, and its message tells you to cancel the cell edit. The snapshot field `selectionContentBlocked` names the first such cell of the selection that holds a value, or is null.

- `redo()` commits no changed edit. If the open edit holds changed text, it returns the refusal `edit-pending`. An open edit with no change closes first.
- `undo()` commits no open edit. If the open edit holds changed text, `undo()` cancels the edit, the cell shows its content again, and the history does not change. If the open edit holds unchanged text, `undo()` closes it and undoes the last stored edit.
- When a new workbook starts to load, the editor drops an open edit at once with the refusal `edit-dropped`. While the load runs, the editor shows each refusal under the loading text and announces it. When the load settles, ready or failed, the other refusals clear. The refusal `edit-dropped`, and `snapshot().editDropped`, stay until the next command after the load, done or refused, an edit that the editor stores, or a selection change, or the next load, so the status message still shows that the typed text was not stored.

  `close()` drops an edit in the same way, and the empty message shows it. Until the load settles, every command is refused with the reason `loading`: the edit commands, `undo()`, and `redo()` return it, `save()` returns `null` with a refused save, `select()` and `setActiveSheet()` return `false`, `setZoom()` does nothing, and the toolbar disables its controls with the reason `loading`.

## Undo and redo

After an undo, the editor selects the selection that the edit started from, on its sheet, and scrolls its active cell into view. After a redo, it selects the selection that the edit ended with. An edit that `enterValue(text, at)` writes to another cell than the active one keeps no selection, so its undo and redo leave the selection as it is. When an undo or a redo changes a sheet that is not shown, such as a hidden sheet, or a sheet whose edit kept no selection, the editor stays on the shown sheet and announces the name of the changed sheet. A redo of a sheet rename selects on the new name, and a redo of a sheet add shows the added sheet. An undo or a redo of a sheet add or delete announces the change, the name of that sheet, and the sheet that shows.

## Screen readers

The grid has a live region for screen readers. It says the active cell or the selection when the selection moves.

When the focus enters the grid from another control, such as with Tab or F6, the live region says the active cell, or the selection with its active cell. If the step that moved the focus already says something, such as a press on a sheet tab or Enter in the formula bar, the live region says only those words. If a toolbar menu applies a format and gives the grid the focus, the live region says the change after the grid updates, or the active cell when the change keeps its shown text. A press on the grid says the cell that it selects, also when it is the active cell.

An edit, such as a number format or a typed entry, can change the shown text of the active cell. The live region then says the active cell and its new text once, after the grid updates. A format or a clear on a selection of many cells also says its active cell. A paste says the pasted selection.

When Enter or Tab moves the active cell inside a selection of many cells, the live region says the selection with the new active cell. A paste that changes cells in such a selection does the same. An undo or a redo says its step, and the active cell when the step changed its text. An edit that keeps the shown text says nothing, except a decimal step or a number format change. That edit says that the decimal places increased or decreased, or that the number format changed. A select of the same selection again also says nothing while the live region still holds its words.

The grid's accessible description says how to leave the grid: press Escape and then Tab or Shift+Tab, or press F6. F6 moves to the next region of the editor, so it leaves the grid when your layout has a toolbar, a formula bar, or a footer. When Escape arms the exit, the live region says to press Tab or Shift+Tab, and a cue over the foot of the grid shows the same words until the exit is used or canceled.

Some screen readers take Escape in their own modes. If the hint does not come, press F6, or press Home and then Shift+Tab. If your page handles Escape before the grid, such as in a `keydown` listener in the capture phase, call `preventDefault()` when you use the key. Otherwise the same Escape also arms the exit.
