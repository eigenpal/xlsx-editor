## MODIFIED Requirements

### Requirement: Stale formula results are visible

The stale state SHALL be `x:calcPr/@fullCalcOnLoad="1"` (the whole workbook) or `x:sheetCalcPr/@fullCalcOnLoad="1"` (one sheet) in the tree. The first transaction that changes a cell value or the grid structure in a workbook that holds a formula cell after the transaction SHALL set `x:calcPr/@fullCalcOnLoad`. The editor snapshot SHALL mark every formula result as stale while it is set, and undo SHALL NOT clear it.

Some edits change no cell value, but they change the result of a function that reads the cell format, a size, or the sheet list. Such an edit, in a workbook that holds a formula cell, SHALL set `x:calcPr/@fullCalcOnLoad` too:

- A format edit (`applyFormat`), a clear of formats, and a paste or an op `setCellFormatIndex` that changes a cell format, when a cell formula or a defined name calls `CELL` or a macro function `GET.*`.
- A column width (`setColumnWidth`, `setColumnWidths`), when a cell formula or a defined name calls `CELL` or a macro function `GET.*`.
- A row height (`setRowHeight`, `setFittedRowHeights`), when a cell formula or a defined name calls a macro function `GET.*`. No result of `CELL` depends on a row height.
- A sheet add, rename, or move, when a cell formula or a defined name calls `INDIRECT`, `CELL`, `INFO`, `_xlfn.SHEET`, `_xlfn.SHEETS`, `_xlfn.FORMULATEXT`, or a macro function `GET.*`. A formula text that the tokenizer cannot read counts as such a call.
- A sheet add or rename that gives a sheet a name that a cell formula or a defined name already names.
- A sheet delete in a workbook that holds a formula cell.

#### Scenario: Undo keeps the stale state

- **WHEN** a host edits a value and then undoes the edit
- **THEN** the value is restored and `fullCalcOnLoad` is still `1`

#### Scenario: A value edit makes results stale

- **WHEN** a host edits a value in a workbook that has formulas
- **THEN** the snapshot reports the formula cells as stale and the adapters show the stale state

#### Scenario: A shown sheet with no formula cell

- **WHEN** the workbook sets `fullCalcOnLoad` and the shown sheet holds no formula cell
- **THEN** the snapshot reports `stale` as false for that sheet, so the adapters show no stale notice, and the saved workbook keeps `fullCalcOnLoad`

#### Scenario: Save after an edit

- **WHEN** the workbook is saved after a value edit
- **THEN** the saved workbook part has `fullCalcOnLoad="1"` on `x:calcPr`

#### Scenario: Reopen a saved edit

- **WHEN** a workbook saved after a value edit is opened again
- **THEN** every formula cell shows its cached value with the stale state, because `fullCalcOnLoad` is `1`

#### Scenario: A formula with no cached value

- **WHEN** a formula cell has no `x:v`
- **THEN** the cell shows no value and the stale state

#### Scenario: A format edit and a call to CELL

- **WHEN** `A1` holds `=CELL("format",B1)`, and the user makes `B1` bold, clears the formats of `B1`, or pastes another cell format on `B1`
- **THEN** the cached results are marked stale, because `CELL("format")` reads the cell format, and the saved workbook asks for a full calculation on load

#### Scenario: A format edit with no call to CELL or GET

- **WHEN** no cell formula or defined name calls `CELL` or a macro function `GET.*`, and the user makes a cell bold
- **THEN** the cached results stay current, because no function result depends on the cell format

#### Scenario: A column width and a call to CELL

- **WHEN** a cell formula calls `CELL`, and the user sets the width of a column
- **THEN** the cached results are marked stale, because `CELL("width")` reads the width, and the saved workbook asks for a full calculation on load

#### Scenario: A row height and a call to GET.CELL

- **WHEN** a defined name calls `GET.CELL`, and the user sets the height of a row
- **THEN** the cached results are marked stale, because `GET.CELL` can read the row height, and the saved workbook asks for a full calculation on load

#### Scenario: A row height and a call to CELL only

- **WHEN** a cell formula calls `CELL`, no formula calls a macro function `GET.*`, and the user sets the height of a row
- **THEN** the cached results stay current
