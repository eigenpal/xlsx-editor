## Why

The editor shows only the values that a file stores. A typed formula shows no result. To calculate after an edit, the editor needs a formula engine first. This change adds that engine alone, with its tests, so that a later integration change can build on a tested engine.

## What Changes

- Parse formula text into an expression tree with the operator precedence of ECMA-376 Part 1, 18.17.2.2. The parser reads the tokens of the existing bounded tokenizer (`packages/core/src/store/formula/tokenizer.ts`).
- Define the value semantics: numbers as IEEE 754 doubles, text, logical values, error values, and empty cells, with their conversions and comparisons.
- Resolve cell references, ranges, sheet prefixes, and defined names through a read-only `CellSource`. The engine never reads the package or the store.
- Evaluate a first function set: `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `ROUND`, `ABS`, `IF`, `AND`, `OR`, `NOT`, `IFERROR`, `CONCATENATE` and the `&` operator, `LEN`, `LEFT`, `RIGHT`, `MID`, `UPPER`, `LOWER`, `TRIM`, `VLOOKUP`, `INDEX`, `MATCH`, `TODAY`, and `NOW`.
- Build a dependency graph of formula cells, order it with explicit stacks, find cycles, and stop at fixed caps.
- Run every formula, or only the formulas that a set of changed cells reaches. A formula that is not reached keeps its stored `x:v`, and its readers read that value.
- Return a calculated value or a defined "unsupported" result for each formula cell. The engine never guesses a value. Nothing in the workbook changes.

## Capabilities

### New Capabilities

- `workbook-formula-engine`: the parser, the value semantics, references and names through a read-only cell source, the first function set, the dependency graph with its caps, the unsupported result, and the evidence.

### Modified Capabilities

None. No editor, store, save, or chrome behavior changes.

## Non-goals

- Any change to the workbook, the store, the editor state, the save, the history, the snapshot, or the chrome.
- Evaluation on open. Nothing in this change calls the engine on open, in a snapshot, in layout, or in a paint.
- Each item of "Later changes". The engine returns "unsupported" for each kind in that list.

## Later changes

The engine returns "unsupported" for the kinds in this list. Each item is its own later change.

- Array formulas, dynamic arrays and spill ranges, and data tables.
- Dynamic ranges, such as `OFFSET`, `INDIRECT`, and a range operator with an operand that is not a cell reference.
- Structured references, 3-D sheet ranges (`Sheet1:Sheet3!A1`), and references to other workbooks.
- Defined names that hold a formula, a constant, a relative reference, or another name.
- Iterative calculation.
- Conversions and comparisons that depend on the locale.
- Function groups after the first set.
- A measurement of results near the 15-digit limit on fixtures with cached `x:v`, for the approximation `double-arithmetic`.

## Next change: workbook-formula-integration

Goal: a typed formula shows its result after an edit, and the save stays safe. That change writes its own design. Its scope:

- Calculation after an edit, inside `WorkbookDocumentStore.transact`.
- The dirty set of each edit, and when a stored `x:v` counts as current.
- The cell states: current, can be out of date, and cannot calculate.
- The approximation entry `double-arithmetic` and the support matrix row `values.calculated-formulas` that cites it.
- The `fullCalcOnLoad` flag rule, and the save of calculated values.
- Manual calculation mode and `x:calcPr/@fullPrecision`.
- The calculation notice, the cell marks, and the undo record.

## Impact

- The `store` lane gets the parser and the value semantics in `packages/core/src/store/formula/`, and the evaluator, the functions, and the graph in a new directory, `packages/core/src/store/formula-engine/`. The lane graph does not change.
- The fixtures get one calculation fixture, its builder, and the measured record `fixtures/calculated.json`.
- `scripts/bench/calc-graph-bench.ts` is a new benchmark script that a developer runs by hand with `bun run bench:calc-graph`.
- No public API, no editor file, no string of `packages/i18n`, and no support matrix row changes.
