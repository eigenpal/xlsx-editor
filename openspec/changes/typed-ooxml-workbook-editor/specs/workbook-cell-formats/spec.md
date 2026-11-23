## ADDED Requirements

### Requirement: Cell format cascade

The cascade lane SHALL resolve each cell's format from `x:c/@s` through `x:cellXfs` to its number format, font, fill, border, alignment, and protection, without the DOM. An index out of range SHALL fall back to the default format and SHALL be recorded with the cell reference.

#### Scenario: A bold red cell

- **WHEN** a cell's `x:xf` names a font with `x:b` and `x:color rgb="FFFF0000"`
- **THEN** the resolved format is bold with the color `#FF0000`

#### Scenario: A style index out of range

- **WHEN** a cell has `s="999"` and the workbook has 3 cell formats
- **THEN** the cell resolves the default format, and the problem is recorded with the cell reference

### Requirement: Colors

The cascade lane SHALL resolve `CT_Color` values from `rgb`, `theme`, `indexed`, and `auto`, and SHALL apply `tint` in the HLS color space, as `reference/quick-ref/themes-colors.md` states.

#### Scenario: A theme color with a tint

- **WHEN** a fill has `theme="4" tint="0.5"` and the theme's `a:accent1` is `4472C4`
- **THEN** the resolved color is `accent1` with its lightness raised as the tint formula states, with the rounding that the quick reference records

#### Scenario: An indexed color

- **WHEN** a font has `indexed="10"` and the style part has no `x:indexedColors`
- **THEN** the resolved color is `#FF0000` from the default palette

### Requirement: Theme fonts

The cascade lane SHALL resolve a font with `x:scheme val="minor"` or `val="major"` to the theme's minor or major Latin typeface, and SHALL use `x:name/@val` when the scheme is `none` or absent.

#### Scenario: A minor theme font

- **WHEN** a font has `x:scheme val="minor"` and the theme's minor Latin typeface is `Aptos`
- **THEN** the resolved font family is `Aptos`

### Requirement: Borders and alignment

The cascade lane SHALL resolve the four cell borders with their styles and colors, and the horizontal and vertical alignment, wrapping, indent, and rotation of each cell.

#### Scenario: A thin bottom border

- **WHEN** a cell's border has `x:bottom style="thin"` with `x:color indexed="64"`
- **THEN** the resolved bottom border is thin, with the system foreground color
