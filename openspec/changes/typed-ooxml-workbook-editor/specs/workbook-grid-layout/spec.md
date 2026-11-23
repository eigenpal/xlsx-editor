## ADDED Requirements

### Requirement: DOM-free grid geometry

The layout lane SHALL compute column widths, row heights, and cell rectangles without the DOM, with an injected text measurer, in points.

#### Scenario: Column width from the default font

- **WHEN** a column has `width="10"` and the default font has a maximum digit width of 7 pixels
- **THEN** the column is 70 pixels wide, as ECMA-376 Part 1, 18.3.1.13 computes, and its point width uses the pixel density that the quick reference records (72 pixels for each inch, measured on the reference PDFs)

### Requirement: Virtualized layout

The layout lane SHALL answer the offset of any row or column and lay out only the visible cells, with cost that depends on the content of the file and the visible cells, not on the sheet size.

#### Scenario: Scroll to the last row

- **WHEN** the viewport scrolls to row 1048576 of a sheet with one explicit row height
- **THEN** layout computes the offset without visiting each row before it

### Requirement: Merged, hidden, and frozen regions

The layout lane SHALL lay out merged ranges as one cell rectangle, give hidden rows and columns zero size, and keep frozen rows and columns in place while the rest scrolls.

#### Scenario: A merged range

- **WHEN** `B2:D4` is merged and `B2` holds text
- **THEN** layout returns one rectangle for `B2:D4` with the text of `B2`, and no rectangles for the other cells

#### Scenario: A hidden column

- **WHEN** column `C` has `hidden="1"`
- **THEN** column `C` has zero width, and column `D` starts where column `C` starts

#### Scenario: Frozen rows

- **WHEN** a sheet has a frozen pane with `ySplit="2"` and the viewport scrolls down 50 rows
- **THEN** rows 1 and 2 keep their place, and the rows below them scroll

### Requirement: Cell text layout

The layout lane SHALL lay out cell text with its alignment, wrapping, indent, and rotation, and SHALL let unwrapped text overflow into empty neighbor cells up to the first cell that holds a value or a formula. Each rule SHALL match a fixture and its reference PDF, as the quick reference requires, except where a scenario names a unit test and the known approximations name the case.

#### Scenario: Overflow stops at a value

- **WHEN** `A1` holds long left-aligned text, `B1` is empty, and `C1` holds a value
- **THEN** the text of `A1` shows across `B1` and is clipped at the left edge of `C1`

#### Scenario: Overflow stops at a formula cell

- **WHEN** the `formula-overflow` fixture has long text in `A1` and `A3`, `B1` holds the formula `1+1` with the cached value 2, and `B3` holds a `t="str"` formula with an empty cached value
- **THEN** the text of `A1` and of `A3` is clipped at the left edge of `B1` and of `B3`, as its reference PDF shows
- **AND** the line scores of the fidelity harness check the stop by the clip edge of each overflowing line, which the known gap list names (`fixtureClipped`)

#### Scenario: Overflow stops at a formula with no cached value

- **WHEN** `A1` holds long left-aligned text and `B1` holds a formula with no cached value
- **THEN** the text of `A1` is clipped at the left edge of `B1`, and the text box of `B1` has no lines and keeps the stale mark
- **AND** a unit test covers the case, because no reference PDF shows it, and the known approximations name it

### Requirement: Drawing anchors

The layout lane SHALL place each drawing from its anchor: a two-cell anchor from its `xdr:from` and `xdr:to` cells and offsets, a one-cell anchor from its `xdr:from` cell and `xdr:ext`, and an absolute anchor from `xdr:pos` and `xdr:ext`.

#### Scenario: A picture anchored to two cells

- **WHEN** a picture has a two-cell anchor from `B2` with no offset to `D6` with no offset
- **THEN** its rectangle starts at the top-left corner of `B2` and ends at the top-left corner of `D6`
