## ADDED Requirements

### Requirement: Stored values only

The system SHALL show the value that the file stores for each cell type of the schema, and SHALL NOT evaluate a formula on open.

#### Scenario: A formula shows its cached result

- **WHEN** a cell holds `x:f` and a cached `x:v`
- **THEN** the cell shows the cached value through its number format, and no formula is evaluated

#### Scenario: A shared string index out of range

- **WHEN** a cell of type `s` holds an index past the end of the shared string table
- **THEN** the cell shows no text, the problem is recorded with the cell reference, and the store keeps the cell unchanged

### Requirement: Stale formula results are visible

The stale state SHALL be `x:calcPr/@fullCalcOnLoad="1"` (the whole workbook) or `x:sheetCalcPr/@fullCalcOnLoad="1"` (one sheet) in the tree. The first transaction that changes a cell value or the grid structure in a workbook that holds a formula cell after the transaction SHALL set `x:calcPr/@fullCalcOnLoad`, the editor snapshot SHALL mark every formula result as stale while it is set, and undo SHALL NOT clear it.

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

### Requirement: Number formats

The system SHALL format numbers through the built-in number formats and through custom format codes, with a bounded parser whose time is linear in the code length.

#### Scenario: A built-in format

- **WHEN** a cell holds `1234.5` with `numFmtId="4"`
- **THEN** the cell shows `1,234.50`

#### Scenario: A hostile format code

- **WHEN** a custom format code is longer than the cap, or holds more sections or bracketed tokens than the caps allow
- **THEN** the cell falls back to the `General` format and the problem is recorded

### Requirement: Dates in both date systems

The system SHALL format date serials in the 1900 and the 1904 date systems, as `x:workbookPr/@date1904` selects.

#### Scenario: The same serial in two systems

- **WHEN** the serial `1` is formatted as a date with `date1904="0"` and with `date1904="1"`
- **THEN** the results are 1900-01-01 and 1904-01-02

### Requirement: Escaped characters and date cells

The system SHALL decode `_xHHHH_` escapes in shared and inline strings, SHALL encode characters that need them on save, and SHALL show a `t="d"` cell as its ISO 8601 date through its number format.

#### Scenario: An escaped tab

- **WHEN** a shared string holds `a_x0009_b`
- **THEN** the cell text is `a`, a tab, and `b`, and a save writes `a_x0009_b`

#### Scenario: A date cell

- **WHEN** a cell has `t="d"` and `x:v` `2024-03-01` with a date number format
- **THEN** the cell shows that date through the format

#### Scenario: A sheet-level stale state

- **WHEN** one sheet has `x:sheetCalcPr fullCalcOnLoad="1"` and the workbook has no `fullCalcOnLoad`
- **THEN** only the formula cells of that sheet are stale
