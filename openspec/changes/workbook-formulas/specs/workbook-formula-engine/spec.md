## ADDED Requirements

### Requirement: A read-only engine

The engine SHALL read cells, sheets, and defined names only through a read-only `CellSource`. It SHALL NOT read the package, the store, a clock, or a global, and it SHALL NOT change the workbook. For each formula cell, it SHALL return a calculated value or an unsupported result with a reason from a closed list, never a guessed value. Nothing SHALL call the engine on open.

#### Scenario: A run changes nothing

- **WHEN** a test makes a full run over a `CellSource` of a fixture
- **THEN** the fixture package is unchanged, and each formula cell has a calculated value or an unsupported result

#### Scenario: A dynamic range and its reader

- **WHEN** `B1` holds `=INDIRECT("A1")`, `C1` holds `=B1+1`, and a full run calculates them
- **THEN** `B1` is unsupported with `outside-subset` and the subject `dynamic-range`, and `C1` is unsupported with `input-not-calculated` and the subject `B1`

#### Scenario: A reference to an unavailable sheet

- **WHEN** the source reports `Sheet2` as unavailable, or the sheets `Sheet2` and `SHEET2` both exist, and `Sheet1!C1` holds `=Sheet2!A1+1`
- **THEN** `C1` is unsupported with `outside-subset` and the subject `unavailable-sheet`

#### Scenario: A formula on an unavailable sheet

- **WHEN** the source reports `Sheet2` as unavailable, and `Sheet2!A1` holds `=1+1`
- **THEN** the results hold `Sheet2!A1` as unsupported with `outside-subset` and the subject `unreached-sheet`

### Requirement: Formula parser

The parser SHALL read the tokens of the existing tokenizer with the grammar and the precedence of ECMA-376 Part 1, 18.17.2.2, with left-to-right association. It SHALL use explicit stacks, never recursion on the expression tree. Text that it cannot read SHALL give `unreadable-formula`.

#### Scenario: Unary minus binds tighter than the power operator

- **WHEN** a cell holds `=-2^2`
- **THEN** the engine returns 4

#### Scenario: The power operator associates left to right

- **WHEN** a cell holds `=2^3^2`
- **THEN** the engine returns 64

#### Scenario: A long chain of unary operators

- **WHEN** a formula holds 4,000 unary minus operators before a number
- **THEN** the parser and the evaluator read it with no stack overflow

### Requirement: Value semantics

The engine SHALL compute in IEEE 754 doubles, read each number from its stored text, and write each number as the shortest text that reads back as the same double. It SHALL follow 18.17.2.2, 18.17.2.6, and 18.17.3, other than in a case that the measured record decides (requirement "Engine evidence"). A conversion or comparison that depends on the locale SHALL give `text-conversion` or `text-locale`. A difference due to doubles only SHALL be the approximation `double-arithmetic`.

#### Scenario: Division by zero

- **WHEN** a cell holds `=1/0`
- **THEN** the engine returns the error `#DIV/0!`

#### Scenario: Zero raised to zero

- **WHEN** a cell holds `=0^0`
- **THEN** the engine returns the value that the measured record holds for that cell

#### Scenario: An integer joined to text

- **WHEN** `A1` holds 12 and `B1` holds `="Total "&A1`
- **THEN** the engine returns the text `Total 12`

#### Scenario: A number with a fraction joined to text

- **WHEN** `A1` holds 1.5 and `B1` holds `="Total "&A1`
- **THEN** `B1` is unsupported with `text-conversion`

#### Scenario: An empty cell in a comparison

- **WHEN** `A1` is empty, `B1` holds `=A1=0`, and `C1` holds `=A1=""`
- **THEN** the engine returns `TRUE` for both

#### Scenario: A text comparison outside ASCII

- **WHEN** `B1` holds `="é"="É"`
- **THEN** `B1` is unsupported with `text-locale`

### Requirement: References and names

A range SHALL read only the cells that exist. A multi-cell range of one area in a function argument where one value is expected SHALL use the implicit intersection. For an operator on such a range, the measured record SHALL decide. Several areas there, and element-by-element operations, SHALL give `array-operation`. Names and sheet prefixes SHALL match by `sheetNameKey`. An undefined name SHALL be `#NAME?`, and an undefined sheet SHALL be `#REF!`.

#### Scenario: A whole column

- **WHEN** `B1` holds `=SUM(A:A)` and column A holds three numbers
- **THEN** the engine reads three cells, never the rows of the grid

#### Scenario: An operator on a range

- **WHEN** `C2` holds `=A1:A3+1`
- **THEN** the engine returns the value that the measured record holds for `C2`

#### Scenario: An element-by-element operation

- **WHEN** a cell holds `=SUM(A1:A3*2)`
- **THEN** the cell is unsupported with `outside-subset` and the subject `array-operation`

#### Scenario: A sheet-scoped name

- **WHEN** `Rate` is defined for `Sheet1` as `Sheet1!$A$1` and for the workbook as `Sheet2!$A$1`, and `Sheet1!B1` holds `=Rate*2`
- **THEN** the engine reads `Sheet1!A1`

#### Scenario: A name that holds a formula

- **WHEN** `Total` is defined as `SUM(Sheet1!$A:$A)`, and `B1` holds `=Total`
- **THEN** `B1` is unsupported with `outside-subset` and the subject `complex-name`

#### Scenario: A name that is not defined

- **WHEN** `B1` holds `=NoSuchName+1`
- **THEN** the engine returns `#NAME?`

### Requirement: The first function set

The engine SHALL calculate `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `ROUND`, `ABS`, `IF`, `AND`, `OR`, `NOT`, `IFERROR`, `CONCATENATE`, `&`, `LEN`, `LEFT`, `RIGHT`, `MID`, `UPPER`, `LOWER`, `TRIM`, `VLOOKUP`, `INDEX`, `MATCH`, `TODAY`, and `NOW`, each by its subclause of 18.17.7. `INDIRECT` and `OFFSET` SHALL give `outside-subset` with `dynamic-range`. Another function SHALL give `function-not-calculated`. No function SHALL allocate or loop by a number from the file.

#### Scenario: Rounding half away from zero

- **WHEN** a cell holds `=ROUND(2.675,2)`
- **THEN** the engine returns 2.68

#### Scenario: A branch that the evaluation does not reach

- **WHEN** a cell holds `=IF(TRUE,1,INDIRECT("A1"))`
- **THEN** the engine returns 1

#### Scenario: An exact lookup with a wildcard

- **WHEN** a cell holds `=MATCH("a*",A1:A3,0)` and `A2` holds `abc`
- **THEN** the engine returns 2, with no backtracking regular expression

#### Scenario: An approximate lookup on unsorted cells

- **WHEN** a cell holds `=VLOOKUP(5,A1:B3,2,TRUE)` and `A1:A3` holds 3, 1, and 7
- **THEN** the cell is unsupported with `unsorted-lookup`

#### Scenario: An approximate lookup on mixed keys

- **WHEN** a cell holds `=VLOOKUP("a",A1:B3,2,TRUE)` and `A1:A3` holds 1, 2, and `x`
- **THEN** the cell is unsupported with `unsorted-lookup`

#### Scenario: A length from the file

- **WHEN** a cell holds `=LEFT(A1,1E+300)`
- **THEN** the engine returns the text of `A1`

#### Scenario: No clock

- **WHEN** a run has no `now`, and a cell holds `=TODAY()`
- **THEN** the cell is unsupported with `no-clock`

### Requirement: Dependency graph and caps

A run SHALL calculate the cells that its dirty set reaches: each dirty formula cell, and each formula that depends on a dirty cell, directly or through other formulas. A full run SHALL start from every formula cell. The engine SHALL order formulas by their dependencies with explicit stacks, never by `/xl/calcChain.xml`. The graph build and each run SHALL have fixed caps in counts. A run that reaches a cap SHALL return no results.

#### Scenario: A run from one changed cell

- **WHEN** `A1` holds 1, `B1` holds `=A1+1`, `C1` holds `=5*2`, and a run starts from `A1`
- **THEN** the results hold `B1` with 2, and no result for `C1`

#### Scenario: An incremental run equals a full run

- **WHEN** a property test stores the results of a full run as the stored values, changes cells, including the target of a simple name, and runs from the changed cells
- **THEN** each reached cell has the calculated value of a full run on the changed source, or both runs give an unsupported result, with any reason
- **AND** each formula cell that the run does not reach has a full-run result equal to its stored value, or both are unsupported
- **AND** the comparison excludes each reached cell that is `blocked` in the full run and neither `circular` nor `blocked` in the incremental run

#### Scenario: A build that reaches its cap

- **WHEN** the graph build reaches its cap of references
- **THEN** `runFormulas` returns the status `limit` and no results

### Requirement: Cycles

Each cycle cell that the run reaches SHALL give `circular-reference`. A reached formula that refers to such a cell, directly or through other reached formulas, SHALL give `input-not-calculated`, also when its evaluation skips that reference. A formula SHALL read a cycle cell outside the run by its stored `x:v`, as each other cell outside the run. A cell that refers to itself is a cycle of one cell. `reachFormulas` gives the first set as `circular` and the second as `blocked`.

#### Scenario: A circular reference

- **WHEN** `A1` holds `=B1+1`, `B1` holds `=A1+1`, `C1` holds `=A1`, and a full run starts
- **THEN** `A1` and `B1` are unsupported with `circular-reference`, and `C1` with `input-not-calculated` and the subject `A1`

#### Scenario: A self-reference

- **WHEN** `A1` holds 1, `A2` holds 2, `A3` holds `=SUM(A1:A3)`, and a full run starts
- **THEN** `A3` is unsupported with `circular-reference`

#### Scenario: A cycle outside the run

- **WHEN** `A1` holds `=B1+1`, `B1` holds `=A1+1`, `C1` holds `=IF(X1>0,1,A1)`, `D1` holds `=C1+1`, `X1` holds 5, and a run starts from `X1`
- **THEN** the engine returns 1 for `C1` and 2 for `D1`, because `A1` is outside the run, so `C1` is not `blocked`, and `IF` does not read `A1`

### Requirement: Formula inputs

A formula SHALL read a reached cell by the result of the run. It SHALL read every other cell by its stored value, and a formula cell, supported or not, by its stored `x:v`. A formula cell with no stored `x:v` SHALL read as unsupported with `unknown-input`. A formula whose evaluation reads a reached cell that is unsupported SHALL be unsupported with `input-not-calculated`. The source SHALL give each cell of an array or data-table `@ref` as `unknown`.

#### Scenario: Inputs that the run does not reach

- **WHEN** `B1` and `B2` hold `=SUMIF(A1:A3,">0")`, `B1` with the stored `x:v` 7 and `B2` with no `x:v`, `C1` holds `=B1+1`, `C2` holds `=B2+1`, and a run starts from `C1` and `C2`
- **THEN** the engine returns 8 for `C1`, and `C2` is unsupported with `unknown-input` and the subject `B2`

#### Scenario: A missing cell inside an array formula range

- **WHEN** `A1` holds an array formula with `@ref="A1:A3"`, `A2` has no `x:c`, and `B1` holds `=SUM(A2:A3)`
- **THEN** `B1` is unsupported with `unknown-input`, never 0

### Requirement: Engine evidence

Each function and value rule SHALL have unit tests and cases in a measured record. A test SHALL compare the engine with that record, and a stale record SHALL fail. Where ECMA-376 does not specify a rule, or a subclause conflicts with a rule or with its own example, the cached `x:v` of the record SHALL decide, and the case file SHALL list the case.

#### Scenario: A subclause that conflicts with its own example

- **WHEN** a cell holds `=MID("abcdef",4,5)`
- **THEN** the engine returns the value that the measured record holds for that cell

#### Scenario: A stale record

- **WHEN** the calculation fixture changes and its record is not measured again
- **THEN** the test fails, and names the fixture
