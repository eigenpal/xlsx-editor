## MODIFIED Requirements

### Requirement: Ops carry intent

Editing ops SHALL name cells by sheet and reference and SHALL carry values, formula text, and format properties, never shared string or cell format indexes. The store SHALL resolve or append shared entries when it applies an op. One op is the exception: `setCellFormatIndex` SHALL carry a cell format index, because a paste gives each target cell the cell format that the copy snapshot holds. A format edit appends a new cell format record or reuses a record that holds the same format, and it never changes the record at an index. So the record at a copied index keeps the copied format, and a paste applies the format from the time of the copy. The editor SHALL refuse the paste with `copy-expired` when the record at a copied index, its font, fill, or border, or its number format code no longer holds the copied format. That happens when an undo removes the record and a later edit puts another format at its index.

#### Scenario: Bold on a cell

- **WHEN** a user makes `B2` bold
- **THEN** the op carries the sheet, `B2`, and the bold property, and the store reuses or appends a cell format entry for it

#### Scenario: A paste of a cell format

- **WHEN** a user copies `A1` and pastes it on `B1`
- **THEN** the op `setCellFormatIndex` carries the cell format index of `A1` from the copy snapshot, and `B1` shows that cell format

#### Scenario: A format edit of the source after the copy

- **WHEN** a user copies the bold cell `A1`, makes `A1` italic, and pastes on `B1`
- **THEN** the paste applies, and `B1` is bold and not italic, as `A1` was at the copy

#### Scenario: A paste after the record at the copied index changed

- **WHEN** a user makes `A1` bold, copies `A1`, undoes the bold, makes another cell italic so that the store puts the italic format at the copied index, and pastes on `B1`
- **THEN** the editor refuses the paste with `copy-expired` and `B1` does not change

### Requirement: Sheet protection

When `x:sheetProtection/@sheet` of a worksheet is true, the store SHALL refuse with the typed reason `locked` a value edit on a locked cell, and each format, size, or structural edit that the protection does not allow (ECMA-376 Part 1, 18.3.1.85). It SHALL also refuse with `locked` a new cell format that changes the `@locked` or `@hidden` protection of a cell, whatever `@formatCells` states. A clear of formats on a protected sheet SHALL keep the `@locked` and `@hidden` protection of each cell. A cell SHALL be locked unless its cell format, or the cell style that it takes its protection from, unlocks it. The editor SHALL keep the protection element as the file holds it.

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
- **AND** when `@formatCells` is false, a new cell format that changes the `@locked` or `@hidden` protection of a cell is still refused with `locked` and the lock `cell`

#### Scenario: A clear of formats keeps the protection

- **WHEN** a sheet is protected with `formatCells="0"`, and a clear of formats meets an unlocked cell whose format is not format 0
- **THEN** the cell gets format 0 with its own protection, in a record that the clear appends when no record holds it, and the cell stays unlocked

#### Scenario: Structural attributes

- **WHEN** a sheet is protected and a user inserts or deletes rows or columns
- **THEN** the store refuses each of the four commands with `locked` unless its attribute (`@insertRows`, `@deleteRows`, `@insertColumns`, or `@deleteColumns`) is false, and an allowed deletion deletes locked cells too

#### Scenario: The hash is kept and never checked

- **WHEN** a protected sheet with a password hash is edited where the protection allows it, and saved
- **THEN** the saved `x:sheetProtection` keeps its hash, salt, and spin count as the file held them, and the editor reads none of them

#### Scenario: The locked code

- **WHEN** the store refuses an op for the protection of its sheet or its workbook
- **THEN** the refusal has the code `locked`, its message is in the shared strings, and a toolbar control that it locks shows a disabled reason by the same protection rule
- **AND** the refusal has `lock` set to the rule that refused it: `cell` for a locked cell that a value edit or a clear of values changes, for a cut of a locked cell that holds content, or for a new cell format that changes the protection of a cell; `format` when the sheet does not allow formatting cells (a format, a clear of formats, or the formats part of a clear of values and formats); `structure` for an insert or a delete that the sheet does not allow; `size` for a column width or a row height that the sheet does not allow (`@formatColumns`, `@formatRows`); or `workbook` for a sheet op when `x:workbookProtection/@lockStructure` is true
- **AND** a refusal of a format, a clear, a structure edit, a cut, or a paste has `command` set to that command: `format`, `clear`, `clearAll`, `structure`, `cut`, or `paste`; the message names the rule and the command, so a clear of formats on unlocked cells never says that a cell is locked
- **AND** a refusal of a value edit, a resize, or a sheet op has no `command`; its rule (`cell`, `size`, or `workbook`) names the message
