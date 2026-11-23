## ADDED Requirements

### Requirement: Cell value edits through ops

The editor SHALL turn every cell edit into typed store ops, with one undo step for each user action, and SHALL refuse an edit that it cannot express.

#### Scenario: Type a number

- **WHEN** a user types `42` into `B3` and confirms
- **THEN** the store holds a numeric cell `B3` with the value `42`, and one undo restores the previous cell

#### Scenario: A leading quote

- **WHEN** a user types `'42` into `B3` and confirms, and then types a lone `'` into `B4` and confirms
- **THEN** `B3` holds the text `42`, and `B4` is empty: a lone quote stores no empty shared string item

#### Scenario: Type a number form

- **WHEN** a user types `12%`, `$5.50`, `1,234.5`, `1 1/2`, `2026-11-12`, or `13:45` into an empty cell with the General format and confirms
- **THEN** the store holds a numeric cell with no `t="s"`: 0.12, 5.5, 1234.5, 1.5, the date serial in the date system of the workbook (ECMA-376 Part 1, 18.17.4.1), or the fraction of a day
- **AND** the cell takes the number format of the form in the same transaction: a built-in format of Part 1, 18.8.30, or a custom code, such as the code of a currency amount or of an elapsed time with no seconds
- **AND** one undo restores the value and the format

#### Scenario: A number form and the cell format

- **WHEN** a user types `12%` into a cell with a number format other than General, into a cell with the Text format (`@`), and into an unlocked General cell of a sheet whose protection does not allow formatting cells
- **THEN** the first cell holds 0.12 and keeps its number format, the second cell holds the text `12%`, and the third cell holds 0.12 and keeps the General format; no edit is refused
- **AND** in a workbook whose list of cell formats holds `MAX_STYLE_RECORDS` (64,000) records, a cell holds 0.12 and keeps its format, and the edit is not refused

#### Scenario: A date in the order of the number locale

- **WHEN** a user types `11/12/2026` with a number locale whose short date (built-in format 14) orders the month first, and with one whose short date orders the day first
- **THEN** the cell holds the serial of 12 November 2026 with the first locale, and of 11 December 2026 with the second
- **AND** with a locale whose short date gives no day and month order, the text stays text, and only an ISO 8601 date reads as a date
- **AND** `1.2.10` reads as a date only with a locale whose short date uses `.`

#### Scenario: Re-edit a typed date, time, or percent

- **WHEN** a user types `2026-11-12`, `13:45`, or `12%` into a General cell, confirms, and presses F2 on the cell
- **THEN** the cell editor shows `2026-11-12`, `13:45`, or `12%`, and Enter on the unchanged text writes nothing

#### Scenario: A typed error literal

- **WHEN** a user types `#N/A` into a General cell and confirms
- **THEN** the cell holds the text `#N/A`, not an error value (`t="e"`)

#### Scenario: An edit that cannot be expressed

- **WHEN** a user edits a cell inside a preserved array formula range
- **THEN** the editor refuses the edit with a typed reason, the cell does not change, and no undo step is added

#### Scenario: Keep the children that an edit does not replace

- **WHEN** a cell holds an element in another namespace, `x:extLst`, or content that the schema does not allow, and a user edits, clears, or pastes over its value
- **THEN** the saved cell keeps each such child in its place, and the new value takes the place of the old `x:f`, `x:v`, and `x:is`
- **AND** an `mc:AlternateContent` child whose branches each hold `x:v` or `x:is` and nothing else is replaced with the value, and an empty `mc:AlternateContent` child is kept

#### Scenario: Refuse an edit of a cell with another form of its value

- **WHEN** a cell holds other markup compatibility content, such as an `x:f` in `mc:AlternateContent`, or an element in another namespace that holds `x:f`, `x:v`, or `x:is`, and a user edits or pastes over it, or clears or cuts it while it holds a value in some form
- **THEN** the editor refuses the command with `cell-alternate-content` and the sheet and cell as the subject, such as `Data!B5`, and the cell does not change
- **AND** a cut of a cell whose only value is in an `mc:AlternateContent` child is refused with `cut-wrapped-value` and the sheet and cell as the subject, because the copy cannot carry that value
- **AND** a cut of a cell whose formula is inside `mc:AlternateContent` is refused with `copy-wrapped-formula` and the sheet and cell as the subject, as its copy is
- **AND** the refusal of a clear, a cut, a paste, or a merge names that command
- **AND** the editor opens no cell editor on that cell, and Clear contents and Clear all are disabled while the selection holds it with a value

#### Scenario: Refuse the commit of an open cell edit

- **WHEN** the commit of an open cell edit is refused with `shared-formula` or `cell-alternate-content`
- **THEN** the cell editor stays open with the typed text, the cell does not change, and the refusal has the command `commit`
- **AND** the message tells the user to cancel the cell edit, and does not tell the user to edit another cell

#### Scenario: Overwrite a shared formula master

- **WHEN** `B1` is the master of a shared formula over `B1:B5` and a user types `7` into `B1`
- **THEN** `B2` to `B5` each hold their own formula text, `B1` holds `7`, and the saved part has no shared formula child without a master

#### Scenario: Clear a range

- **WHEN** a user clears `A1:C3`
- **THEN** the values and formulas of the cells in that range are removed, their cell formats stay, and one undo restores them

#### Scenario: The selection after undo and redo

- **WHEN** a user deletes the selected column `B`, selects `Z400` on another sheet, and undoes the deletion
- **THEN** the editor shows the first sheet, selects column `B` with the active cell it had before the deletion, scrolls that active cell into view, and announces the undo, the sheet, and the selection once; a redo selects the selection after the deletion in the same way, and an edit that kept no selection leaves the selection as it is
- **AND** Shift with an arrow after the undo extends the restored range from the anchor and the focus corner that it had before the deletion, and keeps its active cell
- **AND** an undo of an edit on a sheet that is not shown, when the edit kept no selection, keeps the shown sheet and its selection and announces the name of the changed sheet
- **AND** an undo or a redo with an empty history changes nothing and shows the notice "There is nothing to undo." or "There is nothing to redo." once, however often the user presses the key

#### Scenario: The history keys in an open cell editor

- **WHEN** a cell editor is open and the user presses Mod+Z, Mod+Shift+Z, or Mod+Y
- **THEN** the key goes to the input's own undo or redo: the editor stays open, does not cancel the edit, and does not step the store history
- **AND** with no cell editor open, the same keys step the store history

#### Scenario: Toolbar undo with an open cell edit

- **WHEN** a user types `x` over the cell `A1` that holds `Total`, and presses Undo in the toolbar
- **THEN** the editor cancels the edit, `A1` shows `Total`, the history does not move, and the live region announces that the cell edit was cancelled
- **AND** with an open edit whose text is unchanged, Undo closes the edit and steps the history once
- **AND** Undo is enabled while the edit text is changed or the history has a step, and Redo is disabled with the `editPending` reason while the edit text is changed; `redo()` on the editor instance refuses with `edit-pending` before any write
- **AND** typing `1` over a cell that holds `1` counts as changed text, so Undo cancels that edit and does not step an earlier edit
- **AND** while an input method composes, Undo and every command that commits the open edit are refused with `composing`, and the toolbar shows that reason; a host switch to `view` returns `false` with no announcement
- **AND** when the surface closes while an input method composes (detach, destroy, or an adapter unmount), the editor drops the edit, stores nothing, and sets the refusal `edit-dropped` in the snapshot, and the listeners see it before the surface goes
- **AND** when a new workbook starts to load, the editor drops the open edit at once with the refusal `edit-dropped`; while the load runs, the editor shows each refusal with the loading text and announces it; when the load settles, ready or failed, the other refusals clear, and `edit-dropped` (with `editDropped` in the snapshot) stays shown until the next command after the load, done or refused, an edit, or a selection change, or the next load; until the load settles, every command, a save included, is refused with the code `loading`, and a save keeps `command` set to `save`; select, sheet, and zoom changes do nothing, and the toolbar shows the `loading` reason
- **AND** Control or Command and S in the editor runs the save of the toolbar when the toolbar has an `onSave` handler, also when the save control is hidden; a refused save, from the control or the key, also during a load, shows and announces that the workbook was not saved, with the reason, and keeps the refusal with `command` set to `save` in the snapshot

#### Scenario: A number locale change during an open edit

- **WHEN** `A1` holds the number `1.5`, the user presses F2, the host changes the number locale to a decimal comma, and the user presses Enter
- **THEN** the edit reads its text in the locale that it opened with: `A1` stays the number `1.5`, and the revision does not change; the new locale applies to the next edit

#### Scenario: A typed call to a function outside the predefined functions

- **WHEN** a user types `=LET(x,1,x)` or `=_xlfn.CONCAT(A1,B1)` into a cell and confirms
- **THEN** the editor refuses the edit with `unsupported-function`, because the function is not a predefined function of ECMA-376 Part 1, 18.17.7, and the message names the refused function
- **AND** `=sum(A1:A3)` and a defined name used without `(` are accepted
- **AND** a formula that the workbook holds with a stored-name prefix keeps its stored text through save and reopen

### Requirement: Structural edits move every reference or refuse

Row and column insertion and deletion SHALL move every reference that design D4 lists together, or SHALL refuse the whole edit with a typed reason.

#### Scenario: A preserved part blocks the edit

- **WHEN** a user deletes a column that a preserved conditional formatting range crosses
- **THEN** the store refuses the deletion, names the part, and leaves the workbook unchanged

#### Scenario: A drawing part kept as bytes blocks the edit

- **WHEN** the drawing part of a worksheet names a prefix in `mc:Ignorable` that it does not declare, so the reader keeps it as bytes, and a user inserts a row on that sheet
- **THEN** the store refuses the insertion with `structure-blocked`, names the drawing part, and leaves the workbook unchanged, and the editor disables its row and column insert and delete commands with the reason

#### Scenario: A drawing relationship part kept as bytes blocks the edit

- **WHEN** the relationship part of the drawing of a worksheet is not valid XML, so the reader keeps it as bytes, and a user deletes a column on that sheet
- **THEN** the store refuses the deletion with `structure-blocked`, names the relationship part, and leaves the workbook unchanged

#### Scenario: A formula follows an inserted row

- **WHEN** `C1` holds `=SUM(A1:A10)` and a row is inserted above row 5
- **THEN** `C1` holds `=SUM(A1:A11)` and its result is stale

#### Scenario: A deleted cell under a reference

- **WHEN** `C1` holds `=A5+SUM(A4:A6)` and row 5 is deleted
- **THEN** `C1` holds `=#REF!+SUM(A4:A5)`

#### Scenario: An insertion past the last row

- **WHEN** row 1048576 holds a value and a row is inserted above row 2
- **THEN** the store refuses the insertion and the workbook does not change

#### Scenario: A reference to an empty row

- **WHEN** `C1` holds `=A900000`, no row 900000 exists in the file, and a row is inserted above row 2
- **THEN** `C1` holds `=A900001`, and no structure sized by the grid exists

### Requirement: Editing claims need save-verified fixtures

The support matrix SHALL list an editing command only when every fixture that its tests save has a save-verified record that matches the fixture hash and the serializer version (design D5a), and CI SHALL fail when a listed command does not meet this.

#### Scenario: A command with an unverified fixture

- **WHEN** the support matrix lists an editing command whose saved fixture has no matching save-verified record
- **THEN** CI fails and names the command and the fixture

### Requirement: Unreadable formula text

The editor SHALL refuse formula text that the reference tokenizer cannot read when a user enters it, and the store SHALL refuse every structural edit while the workbook holds such text from the file.

#### Scenario: A user enters an unreadable formula

- **WHEN** a user enters formula text that the tokenizer cannot read
- **THEN** the editor refuses the edit with a typed reason and the cell does not change

#### Scenario: An unreadable formula from the file

- **WHEN** the workbook holds a formula from the file that the tokenizer cannot read, and a user inserts a row
- **THEN** the store refuses the insertion, names the cell, and leaves the workbook unchanged

### Requirement: Ops carry intent

Editing ops SHALL name cells by sheet and reference and SHALL carry values, formula text, and format properties, never shared string or cell format indexes. The store SHALL resolve or append shared entries when it applies an op.

#### Scenario: Bold on a cell

- **WHEN** a user makes `B2` bold
- **THEN** the op carries the sheet, `B2`, and the bold property, and the store reuses or appends a cell format entry for it

### Requirement: View mode

The editor SHALL take a mode, `edit` or `view`, at creation and through `setMode`. In `view` mode it SHALL refuse every edit command with the typed reason `view-mode`, open no cell editor, and keep selection, sheet switching, zoom, and save available.

#### Scenario: An edit in view mode

- **WHEN** the editor is in `view` mode and a user types `7` into `B3`, or a host calls an editing command
- **THEN** no cell editor opens, the command returns the refusal `view-mode`, the workbook and the undo history do not change, and the reason is announced once

#### Scenario: The toolbar in view mode

- **WHEN** the editor is in `view` mode
- **THEN** every toolbar control that edits the workbook is disabled with the `viewMode` reason, and the save control keeps its own state

#### Scenario: Switching to view mode with an open edit

- **WHEN** a cell edit is open and the host sets the mode to `view`
- **THEN** the editor commits the open edit first, and then refuses later edits

#### Scenario: Switching to view mode with an edit the store refuses

- **WHEN** a cell edit with text that the store refuses is open and the host sets the mode to `view`
- **THEN** the editor refuses the switch and stays in `edit` mode, keeps the cell editor open with the typed text, does not change the workbook, and shows the refusal of the commit
- **AND** when the adapter's `mode` prop set the mode, the switch waits, the status strip shows that view mode waits for the cell edit, and the switch applies once the user commits or cancels the edit; a call to `setMode` on the editor instance during the wait ends the wait

#### Scenario: The mode prop of an adapter

- **WHEN** a host changes the `mode` prop of an adapter, and later calls `setMode` on the editor instance
- **THEN** the prop applies when it changes, and the instance call stays until the next change of the prop

#### Scenario: The mode before the editor mounts

- **WHEN** an adapter renders with `mode` set to `view` before the editor exists
- **THEN** the snapshot that the adapter reports has `mode` set to `view`

### Requirement: Sheet protection

When `x:sheetProtection/@sheet` of a worksheet is true, the store SHALL refuse with the typed reason `locked` a value edit on a locked cell, and each format or structural edit that the protection does not allow (ECMA-376 Part 1, 18.3.1.85). A cell SHALL be locked unless its cell format, or the cell style that it takes its protection from, unlocks it. The editor SHALL keep the protection element as the file holds it.

#### Scenario: A locked cell

- **WHEN** a sheet is protected and a user types a value into a cell whose cell format states no protection
- **THEN** no cell editor opens, the command returns the refusal `locked`, the reason is announced once, and the workbook and the undo history do not change

#### Scenario: An unlocked cell

- **WHEN** a sheet is protected and a user enters a value into a cell whose cell format sets `x:protection/@locked` to false
- **THEN** the store applies the edit as on a sheet that is not protected

#### Scenario: A cell with only a row or column format

- **WHEN** a sheet is protected and a value is entered into a cell that does not exist, in a column whose `x:col/@style` names an unlocked cell format
- **THEN** the store applies the edit, and the new cell takes the column format

#### Scenario: The cell style fallback

- **WHEN** a cell format has no `x:protection` and no `@applyProtection`, and its cell style sets `x:protection/@locked` to false
- **THEN** a cell with that format is unlocked, and a cell format with `@applyProtection` set to true and no `x:protection` is locked

#### Scenario: Format cells

- **WHEN** a sheet is protected and a user applies or clears a format
- **THEN** the store refuses it with `locked` unless `@formatCells` is false, and the toolbar disables the format and clear-format controls with the reason that formatting cells is not allowed

#### Scenario: Structural attributes

- **WHEN** a sheet is protected and a user inserts or deletes rows or columns
- **THEN** the store refuses each of the four commands with `locked` unless its attribute (`@insertRows`, `@deleteRows`, `@insertColumns`, or `@deleteColumns`) is false, and an allowed deletion deletes locked cells too

#### Scenario: The hash is kept and never checked

- **WHEN** a protected sheet with a password hash is edited where the protection allows it, and saved
- **THEN** the saved `x:sheetProtection` keeps its hash, salt, and spin count as the file held them, and the editor reads none of them

#### Scenario: The locked code

- **WHEN** the store refuses an op for the protection of its sheet
- **THEN** the refusal has the code `locked`, its message is in the shared strings, and a toolbar control that it locks shows a disabled reason by the same protection rule
- **AND** the refusal has `lock` set to the rule that refused it: `cell` for a locked cell that a value edit or a clear of values changes, `format` when the sheet does not allow formatting cells (a format, a clear of formats, or the formats part of a clear of values and formats), or `structure` for an insert or a delete that the sheet does not allow; the message names that rule and the command, so a clear of formats on unlocked cells never says that a cell is locked
