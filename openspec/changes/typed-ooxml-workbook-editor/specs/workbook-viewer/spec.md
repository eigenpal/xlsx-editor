## ADDED Requirements

### Requirement: Painted grid

The output lane SHALL paint the visible window of the grid from the layout records, with cell values, cell formats, gridlines as `x:sheetView/@showGridLines` sets them, and row and column headings, and SHALL keep the cell text accessible to assistive technology.

#### Scenario: Open a workbook

- **WHEN** a host opens a fixture workbook in the demo
- **THEN** the first sheet shows its cells, formats, merged cells, and pictures as the fixture's reference PDF shows them, within the harness thresholds

### Requirement: Sheet tabs, scrolling, and zoom

The viewer SHALL show one tab for each visible sheet in workbook order, SHALL scroll the grid with the pointer and the keyboard while painting only the visible window, and SHALL zoom from `x:sheetView/@zoomScale` and from a user action.

#### Scenario: Switch sheets

- **WHEN** a user selects the second sheet tab
- **THEN** the grid shows the second sheet, and hidden sheets have no tab

#### Scenario: Scroll a long sheet

- **WHEN** a user scrolls to row 100000 of a sheet with values in rows 1 and 100000
- **THEN** the grid shows row 100000 and paints only the rows in the window

### Requirement: Selection

The viewer SHALL select a cell or a range with the pointer and the keyboard, SHALL show the active cell's reference and content, and SHALL keep the selection inside the grid bounds.

#### Scenario: Extend a selection with the keyboard

- **WHEN** `B2` is active and a user presses Shift with the right arrow twice
- **THEN** the selection is `B2:D2` and `B2` stays the active cell

#### Scenario: Extend a range whose active cell moved inside it

- **WHEN** `B2:D4` is selected, Enter moves the active cell to `B3`, and a user presses Shift with the down arrow or Shift-clicks `D5`
- **THEN** the selection is `B2:D5` and `B3` stays the active cell; when the extended range no longer holds the active cell, the anchor `B2` becomes the active cell

### Requirement: Editing toolbar

The toolbar SHALL show the editing commands of this change through the chrome registry, with their strings from `packages/i18n`. A command that cannot apply SHALL render disabled with the engine's reason.

#### Scenario: A refused structural command

- **WHEN** the selection is a row that a preserved conditional formatting range crosses
- **THEN** the delete-row command is disabled, and its reason names the conditional formatting

#### Scenario: A locale change relabels the grid

- **WHEN** the host changes the catalog of a mounted editor
- **THEN** the grid label, the accessible stale words of each formula cell, and the cell editor label change to the strings of the new catalog once, and the live region empties
