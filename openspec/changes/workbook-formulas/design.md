## Context

The engine shows the value that a file stores, and never evaluates a formula (design D3 of `typed-ooxml-workbook-editor`). This change adds a formula engine and no caller. The `store` lane already holds the pieces that read formula text:

- `store/formula/tokenizer.ts`: `tokenizeFormula`, one forward pass. Its one regular expression, `ROW_WORD`, is anchored and bounded, with no backtracking. It caps the text at 8,192 code units (`MAX_FORMULA_LENGTH`), the tokens at 4,096 (`MAX_FORMULA_TOKENS`), and the nesting at 64 (`MAX_FORMULA_DEPTH`).
- `store/formula/predefined-functions.generated.ts`: `PREDEFINED_FUNCTIONS`, the names of ECMA-376 Part 1, 18.17.7.
- `store/store/shared-formulas.ts` and `moveFormula` (`store/formula/format.ts`): the text of a shared formula child, moved by its offset.
- `store/workbook/workbook-index.ts`: `sheetNameKey` folds a sheet name or a defined name for lookup. `sheetsByKey` finds the sheets of one key. `localNameKeys` gives the keys of the names that one sheet scopes.
- `cascade/values.ts`: `cellValue` reads the stored value of a cell as a `CellValue`, and `parseDouble` reads `ST_Double` text. These are in the `cascade` lane, so the engine does not import them (D8).

ECMA-376 Part 1 clauses that apply: 18.17.2 (syntax), 18.17.2.2 (operators), 18.17.2.4 (functions and multi-value arguments), 18.17.2.5 (names), 18.17.2.6 (types and values), 18.17.2.7 (single-cell and array formulas), 18.17.3 (error values), 18.17.4.1 (serial date-times), 18.17.5.1 (limits), 18.17.5.2 to 18.17.5.4 (precision, lexical form, interpretation), 18.17.6.4 (evaluation order), 18.17.7 (functions), 18.2.5 (`definedName`), and 18.5.1.2 (`table/@displayName`).

## Decisions

### D1. Engine API

The engine is a pure function of a `CellSource`, a dirty set, and options. It never reads the package, the store, a clock, or a global, and it never writes. DOCX Editor has no formula engine. The engine takes the pattern of its pure planner, `packages/core/src/layout/field-ref-refresh.ts` at `03735f11`, which returns results and never writes the tree.

```ts
// The API. Each field is readonly.
type SheetId = string; // chosen by the source
type CellRef = { sheet: SheetId; row: number; col: number };
type CellArea = { top: number; left: number; bottom: number; right: number };
// store/formula/formula-values.ts (task 3.1)
type CalculatedValue =
  | { kind: 'number'; value: number }
  | { kind: 'string'; text: string }
  | { kind: 'boolean'; value: boolean }
  | { kind: 'error'; code: FormulaError };
type FormulaType = 'normal' | 'shared' | 'array' | 'dataTable';
// store/formula-engine/source.ts (task 3.2)
type SourceCell =
  | CalculatedValue
  | { kind: 'empty' }
  | { kind: 'formula'; text: string; type: FormulaType; stored: CalculatedValue | null }
  | { kind: 'unknown' };
type FormulaCell = CellRef & { available: boolean };
type NameInput = { kind: 'defined'; text: string } | { kind: 'table' } | { kind: 'ambiguous' };
interface CellSource {
  sheet(key: string): { id: SheetId } | { unavailable: true } | undefined;
  cell(sheet: SheetId, row: number, col: number): SourceCell;
  cells(sheet: SheetId, area: CellArea): Iterable<{ row: number; col: number; input: SourceCell }>;
  name(key: string, scope: SheetId | null): NameInput | undefined;
  formulas(): Iterable<FormulaCell>;
}
type Unsupported = { kind: 'unsupported'; reason: UnsupportedReason; subject?: string };
type FormulaResult = CalculatedValue | Unsupported;
type RunOptions = { now?: number; limits?: Partial<EngineLimits> }; // now: local time as a serial date-time
type Limit = { status: 'limit'; cap: EngineCap };
// The results hold the reached cells only. The inner key is the text of `formatCellRef` (`store/cell-ref.ts`) on the 0-based row and column, such as `B1`.
type RunResult = { status: 'done'; results: Map<SheetId, Map<string, FormulaResult>> } | Limit;
declare function buildFormulaGraph(
  source: CellSource,
  limits?: Partial<EngineLimits>
): { status: 'done'; graph: FormulaGraph } | Limit;
declare function reachFormulas(
  graph: FormulaGraph,
  dirty: Iterable<CellRef>,
  limits?: Partial<EngineLimits>
):
  | { status: 'done'; order: CellRef[]; circular: CellRef[]; blocked: CellRef[]; steps: number }
  | Limit;
declare function runFormulas(
  source: CellSource,
  dirty: Iterable<CellRef>,
  options?: RunOptions
): RunResult;
```

- Indexes are 0-based, and a `CellArea` includes its edges. `cells` gives the existing cells of an area in sheet order (row, then column), and never iterates the grid. It yields the cells lazily, and includes the `@ref` positions with no `x:c`. Each yielded position counts one step, and a range read stops at its first `unknown`. `EngineCap` names the cap of D6 that a run reached. `FormulaGraph` is the graph of D6, and `reachFormulas` and `runFormulas` are in D6. Each `runFormulas` builds its own graph from the source with `buildFormulaGraph`.
- `sheet` and `name` take a `sheetNameKey` key. A `null` scope is the workbook scope. When two sheets fold to one key, `sheet` gives `{ unavailable: true }`. When two names of one scope fold to one key, `name` gives `ambiguous`. The engine never picks one.
- `formulas()` gives the place of every formula cell of every sheet, in sheet order. Each call gives a new iterable. `cell` gives its formula, and its stored `x:v` as `stored`, or `null` when it has no `x:v`. `available` is false for a sheet that the source cannot give, such as a macro sheet. Each such formula gets `outside-subset` with `unreached-sheet`.
- The source reads `x:f/@t` and `@ref`. It gives each cell inside the `@ref` of an array formula or a data table as `unknown`, with or without an `x:c`, other than the cell that holds the `x:f`. So `cells` also gives the cells of such a range that have no `x:c`, and a read never gives 0 for a missing cell there. Each such cell counts against the step cap (D6). A shared formula child gives the moved text of its master. A date cell (`@t="d"`) gives its serial date-time as a `number`, or `unknown`.
- The value kinds and the `code` of an error take the names of `CellValue` in `cascade/values.ts`. `FormulaError` is the union of the error values of 18.17.3. A cell `subject` is its A1 reference, such as `B1`. A cell on another sheet than the formula adds its sheet name as a formula writes it, such as `'Sheet 2'!B1`.
- `numberText(value)` gives the shortest decimal text that reads back as the same double, in the lexical form of 18.17.5.3.

### D2. Parser

- `parseFormula(tokens)` builds an expression tree from the tokens of `tokenizeFormula`, with no second scan of the text.
- A space token is the intersection operator only between two reference operands, as `checkFormulaGrammar` decides (`packages/core/src/store/formula/grammar.ts:122`).
- Precedence follows the table of 18.17.2.2, from highest: `:`, the union `,` inside a group, the intersection (a space), unary `-`, `%`, `^`, `*` and `/`, `+` and `-`, `&`, then the comparisons. So `-2^2` is 4. Operators of one precedence associate left to right, so `2^3^2` is 64. The measured record checks the order of union and intersection.
- The node count is at most the token count. A chain of operators can be as deep as the token count, so the parser and the evaluator use explicit stacks, never recursion on the tree.
- The parser reads a number literal as the double of 18.17.5.4. Text that it cannot read gives `unreadable-formula`.
- A function name with the prefix `_xlfn.` names the same function without the prefix. A name with the prefix `ECMA.` or `ISO.` names the predefined function of that full name (18.17.2.4), which is outside the first set.

### D3. Value semantics

- **Numbers** are IEEE 754 doubles (18.17.5.2). An input is the double of its stored text. A result is written as the shortest round-trip text (`numberText`). The engine never rounds a result at 15 significant digits. The cached `x:v` of a fixture can hold a result that is rounded near that limit. Such a difference is the named approximation `double-arithmetic` (D7): a result near zero after a subtraction, two numbers equal at 15 significant digits, a literal with more than 15 digits, and a literal halfway between two doubles, which reads to the even double and not to the smaller magnitude of 18.17.5.4. A later change measures these cases (proposal, "Later changes").
- A result that is not finite is `#NUM!` (18.17.5.4). A division by zero is `#DIV/0!` (18.17.3). 18.17.7.255 gives `#DIV/0!` for `POWER` when x is 0 and y is 0 or less. 18.17.2.2 gives no such rule for `^`, so the measured record decides `0^0` and `0^-1` (D5).
- **Empty cells** read as 0 in a number context and as `""` in a text context. A result that is a reference to an empty cell is 0.
- **Text to number** is implementation-defined (18.17.2.6). The engine converts ASCII digits with an optional leading sign. Text of ASCII letters gives `#VALUE!` (18.17.2.2). Every other text, such as `"1.5"`, `"10%"`, or a date, gives `text-conversion`.
- **Number to text** writes an integer with an absolute value below 10^15 as its digits and an optional sign. Every other number, and each logical value, gives `text-conversion`, because its form depends on the locale.
- **Logical values** read as 1 and 0 in a number context. A number is true when it is not 0 (18.17.2.2). Text in a logical context gives `text-conversion`.
- **Comparisons** order numbers before text before logical values. An empty operand compares as 0, `""`, or `FALSE`, by the type of the other operand. Two ASCII texts are equal when their ASCII upper case is equal. An ordered comparison of two ASCII texts compares their ASCII upper-case forms by code unit, as text equality does. The measured record checks it. An ordered comparison of texts needs ASCII letters, ASCII digits, and U+0020 only. Every other text comparison gives `text-locale`.
- **Errors** are the values of 18.17.3. An operator with one error operand returns it. For two different errors, 18.17.2.2 leaves the result unspecified: the engine returns the left one, and the measured record checks it.
- **Dates.** A serial date-time counts local time (18.17.4.1). The `CellSource` gives a date cell as its serial number, and the caller gives `now`. How a production source builds them is work of the next change. The test source of D7 may use `cascade/values.ts`, because the lane import check skips test files (`packages/core/src/__tests__/core-lane-imports.test.ts`). The engine itself does no calendar arithmetic.
- **Text length.** Text functions count UTF-16 code units. The engine sets no length limit for one text result, and never gives `#VALUE!` for a length. The run cap on text code units (D6) is the security bound.
- **Rounding.** `ROUND` rounds the decimal form of the value at 15 significant digits, half away from zero, so `ROUND(2.675,2)` is 2.68.

### D4. References and names

- A range reads only the existing cells, through `cells`. `SUM(A:A)` reads the cells of column A that exist.
- **Implicit intersection.** For a multi-cell range in a function argument where one value is expected, 18.17.2.4 lets the engine use the first element or the implicit intersection. The engine uses the implicit intersection for one area: the cell in the row or column of the formula cell, or `#VALUE!`. For an operator on such a range, 18.17.2.4 evaluates element by element, and 18.17.2.7 gives a single-cell formula the first value. The engine does not choose between these readings for an operator on such a range, or for a formula whose result is such a range: the measured record decides it (D5), such as `=A1:A3+1` and `=A1:A3` in `C2`. For several areas, such as `=ABS((A1:A2,B1))`, the result is unsupported (`outside-subset`, `array-operation`).
- **Array operations.** An operator other than `:`, union, and intersection on a multi-cell operand, inside an argument that takes several values, is evaluated element by element (18.17.2.4). So `=SUM(A1:A3*2)` and `=SUM(ABS({-1,-2,3}))` are unsupported (`array-operation`). An array constant where one value is expected gives its first element (18.17.2.4).
- **Parameter kinds.** Each parameter takes one value or several values. A single value in a parameter that takes several values is a 1x1 array (18.17.2.4).
- **Sheet prefixes** look up `sheet(sheetNameKey(prefix))`. An undefined sheet is `#REF!`. An unavailable sheet, or a key of two sheets, is unsupported (`outside-subset`, `unavailable-sheet`). A range whose two ends name different sheets is `sheet-range`. A workbook prefix is `external-reference`.
- **Names** compare without case (18.17.2.5), by `sheetNameKey`, as `localNameKeys` does. A name with no prefix resolves in the scope of the formula sheet (`definedName/@localSheetId`, 18.2.5) first, then in the workbook scope. A sheet-qualified name, such as `'S1'!Rate`, resolves only in the scope of that sheet, or gives `complex-name`.
- A **simple name** holds one sheet-qualified cell, range, whole-column, or whole-row reference with `$` on each row and column, such as `Sheet1!$A$1:$B$9`. Every other name gives `complex-name`. A table name shares the namespace of defined names (18.5.1.2), so it gives `structured-reference`. An `ambiguous` name gives `duplicate-name`. An undefined name is `#NAME?` (18.17.3).
- A name of the form of a cell reference in `A1` to `XFD1048576` is a cell reference (18.17.5.1). The tokenizer reads a reference only inside these bounds.
- File-keyed data, such as names and sheet keys, goes in a `Map`, never in object keys.

### D5. The first function set

- The function table is a `Map` from the folded name to an entry: the argument counts, the kind of each parameter, the arguments that it evaluates only when needed, and the implementation. Each function follows its subclause of 18.17.7, and the `&` operator follows 18.17.2.2.
- Where ECMA-376 does not specify a rule, or a subclause conflicts with D3 or with its own example, the cached `x:v` of the measured record decides, and `scripts/fixtures/calc-cases.ts` lists the case. 18.17.2.2 gives `#VALUE!` for an operand that ECMA-376 does not specify, so the record decides `=TRUE+1`, `=1<"a"`, and `="a"="A"`. 18.17.7.148 reads an empty `value` of `IFERROR` as `""`, and D3 reads it as 0: the record decides `=IFERROR(A1,0)` with an empty `A1`. 18.17.7.210 extracts the whole string past the end, and its example gives `def` for `MID("abcdef",4,5)`: the record decides that case and `=MID("abc",5,1)`. 18.17.7.53 counts text values and also ignores text in a range: the record decides `=COUNTA(A1:A3)` with text in the range. 18.17.7.246 does not specify `OR` with no logical values: the record decides `=OR(A1:A2)` over text. The record also decides `=IFERROR(1/0,A1)` with an empty `A1`, `=SUM((A1,A1))`, and `=SUM(IFERROR(A1:A3,0))`. ECMA-376 does not state how a count, position, or digits argument with a fraction becomes an integer. The record decides it, and the engine applies that rule to each such argument of the first set.
- Groups: aggregate (`SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`), math (`ROUND`, `ABS`), logical (`IF`, `AND`, `OR`, `NOT`, `IFERROR`), text (`CONCATENATE`, `&`, `LEN`, `LEFT`, `RIGHT`, `MID`, `UPPER`, `LOWER`, `TRIM`), lookup (`VLOOKUP`, `INDEX`, `MATCH`), and time (`TODAY`, `NOW`).
- `IF` and `IFERROR` evaluate only the branch that they return, with the parameter kind of their own position. So `=SUM(IF(TRUE,A1:A3))` sums the range, and `=IF(TRUE,1,INDIRECT("A1"))` is 1. The graph records the references of both branches.
- `UPPER` and `LOWER` change ASCII text only, or give `text-locale`. `TRIM` removes U+0020 only.
- `VLOOKUP` and `MATCH` with an exact match compare ASCII text without case. An exact match gives `text-locale` when the lookup value or a key that it compares is text with a character outside ASCII. An exact match skips an empty key cell, with or without an `x:c`. The measured record decides the result for an empty lookup value, in an exact and an approximate match. They read the wildcards `*`, `?`, and `~` with a linear matcher that has no backtracking regular expression and counts its steps. An approximate match needs existing lookup keys of one type, all numbers or all texts, of the type of the lookup value. They must be strictly sorted, with no empty cell between them. Other keys give `unsorted-lookup`. An approximate match sorts and searches with the comparisons of D3: numbers, then ASCII text by its upper-case form. In an approximate match, a text key or lookup value outside ASCII letters, ASCII digits, and U+0020 gives `text-locale`.
- `TODAY` is the integer part of `now`, and `NOW` is `now`. With no `now`, both give `no-clock`.
- `INDIRECT` and `OFFSET` give `outside-subset` with `dynamic-range`. Every other function outside the set gives `function-not-calculated`, never `#NAME?`.
- No function allocates or loops by a number from the file. `LEFT(A1,1E+300)` returns the text of `A1`.

### D6. Dependency graph, cycles, and bounds

- **Build.** `buildFormulaGraph(source)` reads each formula of `formulas()` once. Each single-cell reference is an entry in a `Map` by cell key. Each range is one entry in an interval list of its sheet, sorted by its first row, so a whole column costs one entry. The build also indexes the target of each simple name. It puts the target range in the same `Map` or interval list as a direct reference, with each formula that uses the name as a reader. So `reachFormulas` reaches `B1` `=Rate*2` after a change to the target of `Rate`. The build only indexes the references of each formula: the edge checks, the order, and the cycle split run in `reachFormulas` and count as steps against the step cap.
- **Order.** The engine orders the formulas to calculate by their dependencies (Kahn's algorithm, explicit queue). It never reads `/xl/calcChain.xml`, although 18.17.6.4 names the calculation chain as the order, because the chain comes from the file and can be wrong.
- **Cycles.** Kahn's algorithm runs over the reached cells and counts only the edges between them. A reached cell left after the order is in a cycle, or refers to one. A cell that refers to itself, directly, through a range, or through a simple name, is a cycle of one cell. Each cell of a cycle gets `circular-reference`. Each other cell left gets `input-not-calculated`, also when its evaluation skips that reference. A cycle cell outside the run is read by its stored value.
- **Runs.** `reachFormulas(graph, dirty)` gives the reached cells: each formula cell in `dirty`, and each formula that depends on a cell in `dirty`, directly or through other formulas. It gives `order`, the reached cells in the order of a full run. It gives `circular`, the reached cells in a cycle, and `blocked`, the reached cells that refer to a cell of `circular`, directly or through other formulas. It gives `steps`, the steps that it counted. `runFormulas(source, dirty, options)` builds the graph with `buildFormulaGraph` and calls `reachFormulas`. It evaluates the cells of `order` in that order. It gives each cell of `circular` the reason `circular-reference`, and each cell of `blocked` the reason `input-not-calculated`. Its subject follows the rule of "Inputs". It starts its step count at `steps`. A full run is `runFormulas(source, source.formulas(), options)`. The results hold the reached cells only.
- **Inputs.** A formula reads a reached cell by the result of this run. It reads every other cell by its stored value: a value cell by its value, and a formula cell, supported or not, by its `stored` value. A formula cell with no `stored` value reads as unsupported with `unknown-input`. A formula whose evaluation reads a reached cell that is unsupported is unsupported with `input-not-calculated`. The subject of an `input-not-calculated` result has one meaning, for a reached reader and for a `blocked` cell. For a reached reader, it is the first unsupported input cell that the evaluation reads. For a `blocked` cell, it is the first input cell in the formula text that is in `circular` or `blocked`. A simple name counts as its target, at the position of the name. The cells of a range count row by row. In a full run, every formula cell is reached, so a formula whose evaluation reads an unsupported formula is unsupported.
- **Equal results.** An incremental run and a full run use the same graph, the same order, and the same evaluator, and a cell that is not reached keeps its stored value. So, on the reached cells, both runs give the same calculated value. Where one gives an unsupported result, the other also gives an unsupported result, and the two count as equal. A cell can be `blocked` in a full run and calculated in an incremental run. This occurs when the cycle is outside the run: the incremental run reads the stored value of the cycle cell, or a function skips the argument that refers to it. The comparison excludes each reached cell that is in `blocked` of the full run and in neither `circular` nor `blocked` of the incremental run. The reason can differ: an unsupported result has no stored value, so an incremental run gives `unknown-input` where a full run gives `input-not-calculated`. A property test checks this comparison. The source starts with the results of a full run as the stored values. After a change, it runs from the changed cells and compares the reached cells with a full run on the changed source. It also checks that each formula cell that the incremental run does not reach has a full-run result equal to its stored value, or both are unsupported. The test inputs include a simple name, such as `B1` `=Rate*2` with a change to the target of `Rate`. The test uses fixed seeds and an inline small generator, such as `mulberry32`, with no new dependency. It runs 200 cases, one for each seed.
- **Caps.** Fixed counts in `store/formula-engine/engine-limits.ts`, never a time, so a result is the same on every machine. Each cap is one `MAX_*` constant and one `EngineCap` member, which `Limit` carries. `EngineLimits` has one camelCase field for each cap, as `ResourceLimits` in `packages/ooxml-core/src/limits.ts` does. `runFormulas`, `buildFormulaGraph`, and `reachFormulas` take the same `limits`. The engine resolves each value of `limits` as `resolveLimits` does (`packages/ooxml-core/src/limits.ts`), except for a value below 1. It floors a finite value. A value below 1 after the floor, such as 0.5, 0, or a negative value, counts as invalid. An invalid value, `NaN`, or an infinite value takes the `MAX_*` constant. Then the engine uses the smaller of the value and the `MAX_*` constant. So a value in `limits` can only lower a cap, and never raises or removes it. Starting values, which task 7.1 measures and sets:
  - `MAX_FORMULA_CELLS` (`formula-cells`, `maxFormulaCells`): at most 100,000 formula cells;
  - `MAX_BUILD_CODE_UNITS` (`build-code-units`, `maxBuildCodeUnits`) and `MAX_GRAPH_REFERENCES` (`graph-references`, `maxGraphReferences`): a build of at most 50,000,000 code units tokenized and 4,000,000 references;
  - `MAX_RUN_STEPS` (`run-steps`, `maxRunSteps`): a run of at most 10,000,000 steps, where a step is one node evaluated, one cell that a range read examines, one graph entry checked, or one code unit that a text operation reads, writes, or compares;
  - `MAX_TEXT_CODE_UNITS` (`text-code-units`, `maxTextCodeUnits`): at most 10,000,000 code units of text results in one run.
- A run that reaches a cap returns `{ status: 'limit', cap }` and no results. The steps of `reachFormulas` count against the step cap.
- **Target.** `scripts/bench/calc-graph-bench.ts`, which the `bench:calc-graph` script of `package.json` runs, measures a build at its caps and one run at the step cap for each kind of step. Task 7.1 sets the caps so that a build and a full run take at most 250 milliseconds on the CI runner image, and records the times, with a fourfold CPU slowdown, in the pull request.

### D7. Unsupported results and evidence

The closed reason list is in `store/formula-engine/reasons.ts`:

| Reason | When | Subject |
| --- | --- | --- |
| `function-not-calculated` | A function outside the first set, other than `INDIRECT` and `OFFSET` | The function name |
| `outside-subset` | A kind of "Later changes", or input that the source cannot give | `array-formula`, `data-table`, `dynamic-range`, `array-operation`, `structured-reference`, `external-reference`, `sheet-range`, `complex-name`, `duplicate-name`, `unavailable-sheet`, or `unreached-sheet` |
| `input-not-calculated` | A read of a reached cell that is unsupported, or a `blocked` cell (D6) | For a reached reader, the first unsupported input cell that the evaluation reads. For a `blocked` cell, the first input cell in the formula text that is in `circular` or `blocked` (D6) |
| `circular-reference` | A reached cell in a cycle (`circular` of D6) | None |
| `text-conversion` | A conversion of D3 that depends on the locale | None |
| `text-locale` | A text comparison or case change outside ASCII | None |
| `unsorted-lookup` | An approximate lookup over cells that are not sorted | None |
| `unknown-input` | An `unknown` input, or a formula cell that is not reached and has no stored `x:v` | The cell |
| `no-clock` | `TODAY` or `NOW` with no `now` | None |
| `unreadable-formula` | Text that the tokenizer or the parser cannot read | None |

A subject from the file is cut to 64 code units.

Evidence:

- **Unit tests** for each rule of D2 to D6, in `__tests__` next to each file.
- **Measured record.** The fixture `formula-calc` is built by `scripts/fixtures/calc-catalog.ts`, in `fixtures/calc/`, outside `FIXTURES`, because it has no reference PDF. It holds the measured formula cells and their inputs, with `fullCalcOnLoad="1"` and no cached values. `fixtures/calculated.json` records the calculated `x:v` and `x:c/@t` of each measured cell, with the SHA-256 of the fixture, the locale, and the `now` of the calculation. `scripts/fixtures/calculated.ts` reads and writes the record, as `scripts/fixtures/measured.ts` does.
- **Test against the record.** A test builds a `CellSource` over the fixture with `cellValue` (tests may import across lanes), runs a full run with the `now` of the record, and compares each cell: the same type, text, and double. `scripts/fixtures/calc-cases.ts` names each case, its task, and its expected reason or expected `double-arithmetic` difference. A case whose function group has not landed expects `function-not-calculated`. A stale record fails, and names the fixture.
- `double-arithmetic` is a constant in `store/formula-engine/reasons.ts` that the case file cites. AGENTS.md ("Fidelity") and `scripts/check-support-matrix.ts` need a support matrix row that cites each entry of `packages/core/src/layout/approximations.ts`. A `values` row must cite a fidelity result, and no calculated value shows in this change, so no row can cite the entry yet. The integration task that first shows a calculated value adds the entry to `packages/core/src/layout/approximations.ts` and `docs/site/content/known-approximations.mdx`. The same task adds the row that cites it: `values.calculated-formulas`, category `values`, status `partial`, with `approximations: ['double-arithmetic']`.

### D8. Lane and files

The engine lives in the existing DOM-free `store` lane. The integration change runs it inside `WorkbookDocumentStore.transact`, which is in `store`, and `store` imports no other lane. The `cascade` lane could import `cellValue`, but `store` cannot import `cascade`, so the integration could not call it. So `CellSource` hands the engine typed inputs, and the lane graph does not change.

- `store/formula/parser.ts` and `store/formula/formula-values.ts` (`CalculatedValue`, `FormulaError`, conversions, comparisons, `numberText`).
- `store/formula-engine/`: `source.ts` (`CellSource`, `SourceCell`), `reasons.ts`, `engine-limits.ts`, `evaluate.ts`, `range-reads.ts`, `names.ts`, `functions/` (one file for each group), `graph.ts` (`buildFormulaGraph`, `reachFormulas`), and `run.ts` (`runFormulas`).
- No new name repeats an existing file or directory name under `packages/core/src/store`, such as `store/store/calc.ts`, `store/limits.ts`, or `store/formula/references.ts`. No new type name repeats a type name that the `store` barrel exports, such as `CellInput` (`store/store/ops.ts`).
- The entry points `runFormulas`, `buildFormulaGraph`, and `reachFormulas` join the `store` barrel (`packages/core/src/store/index.ts`). These types join it too: `SheetId`, `CellRef`, `CellArea`, `CalculatedValue`, `FormulaType`, `SourceCell`, `FormulaCell`, `NameInput`, `CellSource`, `Unsupported`, `FormulaResult`, `RunOptions`, `Limit`, `RunResult`, `FormulaGraph`, `EngineLimits`, `EngineCap`, `UnsupportedReason`, and `FormulaError`. The other new names stay out of it.
- Each file stays under the 1,000-line cap.

## Risks and trade-offs

- [The engine has no caller] → Nothing in the editor changes until the integration change. The tests and the measured record are the only evidence, so a defect in the API shape shows late. Mitigation: the record test calls `runFormulas` through a `CellSource`, as the integration will.
- [Another consumer calculates a different value] → Each rule rests on ECMA-376 or on the measured record. A rule that neither covers is `double-arithmetic` or an unsupported result.
- [Many real formulas are unsupported] → Text joined to a fraction, complex names, and array formulas are common.
- [The caps stop a large workbook] → A run past a cap returns no results, never a part of them. The integration change decides what the editor shows then.
