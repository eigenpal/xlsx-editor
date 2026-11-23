## ADDED Requirements

### Requirement: Pattern fills

The output lane SHALL paint a pattern fill as the repeated 8 by 8 tile of its `ST_PatternType`, with 1 point for each tile pixel, the foreground color on the set pixels, and the background color on the others. Each tile pixel SHALL cover whole device pixels. When one tile is smaller than 8 device pixels or larger than 512 device pixels, or the canvas context cannot make or place a pattern, the fill SHALL paint the mix of its two colors in the share of set pixels.

#### Scenario: A light grid pattern

- **WHEN** a cell has `patternType="lightGrid"` with a red foreground and a yellow background
- **THEN** the cell paints the `lightGrid` tile of the reference PDF in red on yellow, not one mixed color

#### Scenario: Absent colors

- **WHEN** a pattern fill has no `x:fgColor`, or no `x:bgColor`
- **THEN** the set pixels paint black, or the other pixels paint white

#### Scenario: A low zoom

- **WHEN** one tile of 8 points is smaller than 8 device pixels
- **THEN** a `mediumGray` fill in black on white paints the color `#808080`

### Requirement: Pattern tile origin

The tile grid SHALL start at `PaintOptions.patternOrigin` when it is set, and at the sheet origin of each pane when it is not.

#### Scenario: Adjacent cells on a printed page

- **WHEN** H8 and I8 of `fills-patterns` have the same `darkTrellis` fill, and the harness paints the printed page
- **THEN** the tile grid starts at the bottom-left corner of the page, and the pattern runs across the edge between H8 and I8 with no phase step

#### Scenario: Scrolling the view

- **WHEN** no `patternOrigin` is set, and the view scrolls
- **THEN** each pattern moves with its cells

### Requirement: Gradient fills

The output lane SHALL paint a gradient over the box of its cell or merged range: the cell and its right and bottom gridline pixels. It SHALL paint a `linear` gradient in the unit square of the box, with its axis along `@degree`. Its lines of equal color SHALL join the same fractions of the box edges. A `path` gradient SHALL paint position 0 on its inner rectangle and position 1 on the box edges, along rectangles.

#### Scenario: A linear gradient at 45 degrees

- **WHEN** a cell of 100 by 40 points with a gridline of 1 point has a `linear` gradient with `degree="45"`
- **THEN** the axis runs from the top-left corner of the box of 101 by 41 points to its bottom-right corner

#### Scenario: A linear gradient at 90 degrees

- **WHEN** a cell has a `linear` gradient with `degree="90"`, red at position 0 and blue at position 1
- **THEN** the cell is red at the top and blue at the bottom

#### Scenario: A path gradient from the center

- **WHEN** a cell has `type="path"` with `left`, `right`, `top`, and `bottom` of 0.5
- **THEN** the color of position 0 is at the center of the box of the cell, and the color is equal along each rectangle around the center

#### Scenario: A path rectangle on a corner

- **WHEN** a cell has `type="path"` with all four edges 0
- **THEN** only the right and bottom parts have a gradient, and position 0 is at the top-left corner

### Requirement: Path gradient rectangle

The cascade SHALL read `@left`, `@right`, `@top`, and `@bottom` of a `path` gradient as edge positions from the left and top edges, with 0 for an absent edge. Paint SHALL apply them to the box of the cell. The cascade SHALL limit each value to 0 to 1, read text that is not a number as 0, and order a reversed pair. It SHALL record one problem `gradient-rectangle-out-of-range` for each changed rectangle, with the first change as its detail.

#### Scenario: An off-center rectangle

- **WHEN** a path gradient has `left="0.2" right="0.4" top="0.3" bottom="0.6"`
- **THEN** the inner rectangle runs from 0.2 to 0.4 of the width and from 0.3 to 0.6 of the height, with no problem

#### Scenario: A pair in reverse order

- **WHEN** a path gradient has `left="0.8"` and `right="0.2"`
- **THEN** the inner rectangle runs from 0.2 to 0.8 of the width, and the problem `gradient-rectangle-out-of-range` is recorded

#### Scenario: Values out of range and not a number

- **WHEN** a path gradient has `left="-3"` and `right="NaN"`
- **THEN** both edges are 0, and one problem is recorded, with the detail `left=-3`

### Requirement: Cell formats under cell style formats

A cell format that omits its font, fill, border, or number format SHALL resolve record 0. A cell format that omits `x:alignment` SHALL resolve the default alignment. The fixture `cell-styles` measures this for each `apply*` state under a cell style format that applies all records, and with `apply*` absent under one that applies none. A record of the cell style format SHALL NOT apply. An own `x:alignment` SHALL replace the whole alignment.

#### Scenario: An omitted font under a bold style

- **WHEN** a cell format has no `fontId` and `applyFont="1"`, and its cell style format has a bold red font
- **THEN** the cell shows font 0

#### Scenario: An omitted alignment under a centered style

- **WHEN** a cell format has no `x:alignment`, and its cell style format centers its text
- **THEN** the number in the cell is right-aligned, by the `general` alignment

#### Scenario: An alignment that sets only wrap

- **WHEN** a cell format has `<alignment wrapText="1"/>`, and its cell style format centers its text
- **THEN** the cell has the `general` horizontal alignment
