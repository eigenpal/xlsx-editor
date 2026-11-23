# workbook-conditional-formatting Specification

## Purpose

Conditional formatting (ECMA-376 Part 1, 18.3.1.18 `conditionalFormatting` and 18.3.1.10 `cfRule`): how the engine reads the rules of a worksheet within limits, which rule types it evaluates on the values that the file stores, and how a differential format (`x:dxf`, 18.8.14) changes the cell format that the layout paints.

## Requirements

### Requirement: Rule reader limits

The engine SHALL read at most 4,096 `x:cfRule` elements and 65,536 rule and range pairs for each worksheet. It SHALL count a rule past these limits as skipped and SHALL NOT build it. A rule whose `@sqref` holds more than 512 ranges or more than 16,384 characters SHALL be `malformed`. The engine SHALL NOT allocate for the cells that a range names.

#### Scenario: Rules past the sheet limit

- **WHEN** a worksheet holds 4,099 `x:cfRule` elements
- **THEN** the engine builds 4,096 rules and counts 3 as skipped

#### Scenario: An @sqref with 513 ranges

- **WHEN** a rule's `@sqref` holds 513 ranges
- **THEN** the rule is not shown, with the reason `malformed`

#### Scenario: An @sqref with a whole-column range

- **WHEN** a rule's `@sqref` is `B:B`
- **THEN** the rule is not shown, with the reason `malformed`, and the sheet gets the notice

#### Scenario: A range over the whole grid

- **WHEN** a rule covers `A1:XFD1048576`
- **THEN** the engine stores the range as its four bounds and evaluates the rule for each visible cell within the check budget of the layout

### Requirement: Shown rules

The engine SHALL show a `cellIs` rule whose operands are literals: a number, a quoted string, `TRUE`, or `FALSE`. It SHALL NOT show any other rule, and it SHALL give each such rule that it builds one reason: `unsupported`, `needs-formula`, or `malformed`. A rule SHALL read the stored value of its cell, also a cached formula result, and the engine SHALL NOT evaluate a formula.

#### Scenario: A rule with a reference operand

- **WHEN** a `cellIs` rule compares with `$B$1`
- **THEN** the rule is not shown, with the reason `needs-formula`

### Requirement: Evaluated cells

Layout SHALL evaluate each visible cell that a rule covers, also a cell without an `x:c` element and the top-left cell of an empty merged range, with an empty value. A formula cell without a cached value SHALL NOT be evaluated.

#### Scenario: A formula without a cached value

- **WHEN** a formula cell has no `x:v` and a rule `lessThan 5` covers it
- **THEN** the rule does not paint the cell

#### Scenario: An out-of-date cached value

- **WHEN** a rule covers a formula cell whose cached value the first change marks out of date
- **THEN** the rule follows the cached value, and the text of the cell keeps its out-of-date mark

#### Scenario: An expression rule

- **WHEN** a rule has `type="expression"`
- **THEN** the rule is not shown, with the reason `unsupported`

### Requirement: Comparison

The engine SHALL order numbers before strings and strings before booleans. It SHALL compare strings without case. An empty value, also of a cell without an `x:c` element, SHALL be 0 against a number operand and the empty string against a string operand. An error value, an ISO date value (`t="d"`), or an empty value against a boolean operand SHALL make a rule false.

#### Scenario: A text value under a number operand

- **WHEN** B5 of `conditional-rules` holds `abc` under a rule `greaterThan 3` that sets bold
- **THEN** the value shows bold, as the reference PDF draws it

#### Scenario: A cell without a value

- **WHEN** B17 of `conditional-rules` has no `x:c` element and a rule `equal 0` that sets a yellow fill
- **THEN** the cell paints the yellow fill, as the reference PDF draws it

#### Scenario: A string with another case

- **WHEN** a cell holds `abc` and a rule tests `equal` to `"ABC"`
- **THEN** the rule is true

### Requirement: Priority walk

The engine SHALL walk the rules over a cell by ascending `@priority`, then document order. Each property SHALL take the first value that a true rule sets. A true rule with `stopIfTrue` SHALL end the walk. A rule that is not shown SHALL NOT paint and SHALL NOT end the walk. For one cell, the engine SHALL check at most 1,024 rule ranges of its row and SHALL walk at most 64 shown rules.

#### Scenario: Two true rules

- **WHEN** a true rule of priority 1 sets italic and a red font, and a true rule of priority 2 sets bold
- **THEN** the cell shows bold, italic, and red

#### Scenario: A rule that is not shown with stopIfTrue

- **WHEN** an `expression` rule of priority 1 has `stopIfTrue`, and a true `cellIs` rule of priority 2 sets bold
- **THEN** the cell shows bold

### Requirement: Differential format overlay

The engine SHALL lay the differential format of each true rule over the cell format. The font properties that `x:font` sets SHALL replace those of the cell. An `x:numFmt` SHALL replace the number format, also in the plain text of a copy and in the announcement of the active cell, except after a budget cut of the sheet, when those use the cell format. Borders, alignment, and protection SHALL NOT apply, and row heights SHALL NOT change.

#### Scenario: The rule of preserved-parts

- **WHEN** the engine lays out B1 of `preserved-parts`, which holds 5 under a rule `greaterThan 3` whose format sets bold
- **THEN** the value shows bold, as the reference PDF draws it in `Calibri-Bold`

#### Scenario: The cases of conditional-rules

- **WHEN** the engine lays out column B of `conditional-rules`
- **THEN** each cell has the fill, the face, the decoration, the color, and the number format that its reference PDF draws, in the locale of the view and within the approximation "Tint rounding"

#### Scenario: Copied and announced text

- **WHEN** a cell holds 0.5 under a true rule whose format sets the number format `0%`
- **THEN** a copy writes `50%` to the plain text of the clipboard, and the announcement of the active cell says `50%`, as the grid shows it

#### Scenario: Copied and announced text after a budget cut

- **WHEN** a layout of the shown sheet has used up its check budget, and a cell holds 0.5 under a true rule whose format sets `0%`
- **THEN** a copy writes `0.5`, and the announcement says `0.5`, the text of the cell format

### Requirement: Copy check budget

A copy SHALL charge each row of its range once in full, and each cell for its own checks, against its own check budget of 4,000,000 rule-range checks, so its result depends only on the range. A copy or a cut that runs out of that budget SHALL be refused with `copy-conditional-limit`.

#### Scenario: A copy past its check budget

- **WHEN** a copy or a cut uses up its check budget of 4,000,000 rule-range checks, with each row charged once in full
- **THEN** it is refused with `copy-conditional-limit`, and a second copy of the same range is refused too
- **AND** the snapshot does not state `conditionalHidden` for that cause, and a later copy inside the budget gives the text that the grid shows

### Requirement: Differential fill

A solid or untyped `x:patternFill` of a true rule SHALL paint its `x:bgColor`. Without `x:bgColor`, it SHALL keep the fill of the cell format, and SHALL paint white over a cell with no fill. Any other differential fill SHALL NOT change the fill.

#### Scenario: A fill with only a foreground color

- **WHEN** a true rule's format has an `x:patternFill` with only `x:fgColor`, over B23 of `conditional-rules` with its own fill and over B10 with no fill
- **THEN** B23 keeps its own fill and B10 paints white, as the reference PDF draws them

### Requirement: Toolbar state

The toolbar state SHALL show the cell format of the active cell without the differential formats of its rules, because a format command changes the cell format.

#### Scenario: Bold from a rule

- **WHEN** a true rule makes the active cell bold and its cell format is not bold
- **THEN** the bold command shows as not active

### Requirement: Toolbar hint

When a true rule changes the state of a format command at the active cell, that command SHALL carry the hint `hint.conditionalFormat`, which the chrome shows after the label in the tooltip and as the accessible description. After a layout of the shown sheet ran out of its check budget, no command SHALL carry the hint, because the grid can show the cell format where a rule is true.

#### Scenario: Bold from a rule, with a hint

- **WHEN** a true rule makes the active cell bold
- **THEN** the bold command carries the hint `hint.conditionalFormat`
- **AND** the italic and alignment commands carry no hint
- **AND** the accessible description of the bold control is the hint

#### Scenario: A rule that changes nothing

- **WHEN** the only true rule at the active cell has an empty differential format
- **THEN** no command carries the hint

#### Scenario: A rule that sets a state that the cell format has

- **WHEN** a true rule makes the active cell bold, and the cell format is bold
- **THEN** the snapshot has no `conditionalFormat`, and no command carries the hint

#### Scenario: No hint after a cut

- **WHEN** a true rule makes the active cell bold
- **AND** a layout of the shown sheet runs out of its check budget
- **THEN** the snapshot has no `conditionalFormat`, and the bold command carries no hint

### Requirement: Notice for rules that are not shown

When the shown sheet has a rule that is not shown, rules past the read limits, more shown rules or ranges than the caps of one cell or one row take, or rules in `x:extLst`, or when a layout of the shown sheet ran out of its check budget since the sheet was shown, the chrome SHALL show the notice `status.conditionalHidden`. A notice from the budget SHALL stay with its sheet, not with its index, until the shown sheet changes or another workbook opens.

#### Scenario: An expression rule

- **WHEN** the shown sheet has an `expression` rule
- **THEN** the chrome shows the notice `status.conditionalHidden`

#### Scenario: A layout past its budget

- **WHEN** one layout of the shown sheet uses up its check budget
- **THEN** the snapshot states `conditionalHidden`, so the chrome shows the notice `status.conditionalHidden`
- **AND** the notice stays when a later layout of a smaller window stays inside the budget

#### Scenario: A sheet op after a budget cut

- **WHEN** a layout of the shown sheet uses up its check budget
- **AND** a sheet op moves that sheet to another index
- **THEN** the notice stays with that sheet
- **WHEN** a sheet op then deletes that sheet, and another sheet shows at its index
- **THEN** the notice and the hints of the other sheet come from its own rules, not from the cut of the deleted sheet

#### Scenario: Only shown rules

- **WHEN** every rule of the shown sheet is a shown `cellIs` rule
- **THEN** the chrome shows no such notice

### Requirement: Operand length and extension rules

The engine SHALL give a rule with an operand longer than 8,192 characters the reason `needs-formula`. It SHALL NOT read the rules in the extension content (`x:extLst`) of a worksheet, so these rules never paint.

#### Scenario: A long string operand

- **WHEN** a `cellIs` operand is a quoted string of 8,193 characters
- **THEN** the rule is not shown, with the reason `needs-formula`
