Each task lands as one pull request of about 300 changed lines, with its tests. Each task stays in the `store` lane, `scripts`, or `fixtures`. No task touches the editor, the chrome, the save, or `packages/i18n`. Nothing calls the engine outside tests until the integration change.

## 1. Evidence

- [ ] 1.1 Add `store/formula-engine/reasons.ts` with the reason list and `double-arithmetic` of design D7. Add `scripts/fixtures/calc-catalog.ts` and the fixture `fixtures/calc/formula-calc.xlsx`, and read the catalog in `bun run fixtures:build`, `bun run fixtures:check`, and `bun run validate:xsd`. Include the cases of D2 to D5: precedence, union and intersection, `0^0` and `0^-1`, `=TRUE+1`, `=1<"a"`, `="a"="A"`, `="a"<"B"`, `="a b"<"ab"`, `="1"<"a"`, `=VLOOKUP("B",A1:B3,2,TRUE)`, the `double-arithmetic` cases, `#REF!`, `#NAME?`, the name scopes, `=SUM(IF(TRUE,A1:A3))`, `=SUM(ABS({-1,-2,3}))`, `=A1:A3+1` and `=A1:A3`, `=IFERROR(A1,0)` with an empty `A1`, `=MID("abcdef",4,5)`, `=MID("abc",5,1)`, `=COUNTA(A1:A3)` with text in the range, `=OR(A1:A2)` over text, `=IFERROR(1/0,A1)` with an empty `A1`, `=SUM((A1,A1))`, and `=SUM(IFERROR(A1:A3,0))`. Each case that decides between readings is on its own worksheet of the fixture, with its own inputs and formula cell, so each reading gives a different result: `=SUM(IFERROR(A1:A3,0))` in `C2` with `A1`=2, `A2`=`#DIV/0!`, and `A3`=3; `=A1:A3+1` and `=A1:A3`, each on its own worksheet with `A1`=1, `A2`=2, and `A3`=3, and each once in `C2` (a row inside the range) and once in `C5` (a row outside it); `=SUM(IF(TRUE,A1:A3))` in `C2` with `A1`=1, `A2`=2, and `A3`=4; the overlapping union `=SUM((A1,A1))` with `A1`=5; and `=VLOOKUP("B",A1:B3,2,TRUE)` in `C1` with `A1`=`a`, `A2`=`b`, `A3`=`c`, `B1`=1, `B2`=2, and `B3`=3. The empty lookup cases: `=VLOOKUP(0,A1:B3,2,FALSE)` in `C1` with a styled empty `A1`, `A2`=0, `B1`=1, and `B2`=2; `=MATCH(A1,B1:B4,0)` in `C1` with an empty `A1`, `B1`=1, `B2`=0, `B3` with the formula `=""`, and `B4`=`FALSE`, so a 0 reading gives 2, a `""` reading gives 3, a `FALSE` reading gives 4, and no match gives `#N/A`; `=MATCH(A1,B1:B3,0)` in `C1` with an empty `A1`, `B1`=1, `B2` with the formula `=""`, and `B3`=0; and `=VLOOKUP(A1,D1:E3,2,TRUE)` in `C1` with an empty `A1`, `D1:D3`=0, 1, 2, and `E1:E3`=10, 20, 30. Each is on its own worksheet. The fraction cases, each on its own worksheet: `=LEFT("abc",1.5)`, `=MID("abcdef",2.5,2)`, `=INDEX(A1:A3,1.5)` with `A1:A3`=1, 2, 3, `=ROUND(1.25,1.9)`, `=ROUND(15,-0.6)`, and `=LEFT("abc",-0.5)`. `double-arithmetic` stays in `reasons.ts` in this change. The integration task that adds it to `packages/core/src/layout/approximations.ts` also adds the row `values.calculated-formulas` that cites it (D7).
- [ ] 1.2 Add `scripts/fixtures/calculated.ts` and `scripts/fixtures/calc-cases.ts`, by design D7. Record the measured values in `fixtures/calculated.json`.

## 2. Parser

- [ ] 2.1 Add `store/formula/parser.ts` with the operators, the precedence, the number literals, and an explicit stack, by design D2. Test 4,000 unary minus operators with no stack overflow.
- [ ] 2.2 Add to the parser the references, union and intersection, array constants, function calls, and the prefixes, by design D2.

## 3. Values

- [ ] 3.1 Add `store/formula/formula-values.ts` with `CalculatedValue` as its own union, `FormulaError`, `numberText`, the conversions, and the comparisons, by design D3.
- [ ] 3.2 Add `store/formula-engine/source.ts` with `CellSource` and `SourceCell`, by design D1. Add the test `CellSource` over a fixture, with the stored `x:v` of each formula cell and each cell of an array or data-table `@ref` as `unknown`.
- [ ] 3.3 Add `store/formula-engine/engine-limits.ts` with the starting caps, by design D6. Add `store/formula-engine/evaluate.ts` with the operators, an explicit stack, and the step count against the step cap, by designs D1 and D3. Until task 6.1, each function call gives `function-not-calculated`.

## 4. References and names

- [ ] 4.1 Add `store/formula-engine/range-reads.ts`: the sparse range reads, the implicit intersection, union, intersection, sheet prefixes, the parameter kind type, and the `array-operation` results, by design D4.
- [ ] 4.2 Add `store/formula-engine/names.ts`: simple names, the sheet scope before the workbook scope, sheet-qualified names, table names, `duplicate-name`, and `#NAME?`, by design D4.

## 5. Graph and bounds

- [ ] 5.1 Add `store/formula-engine/graph.ts` with `buildFormulaGraph`, the interval lists, and the build caps, and with `reachFormulas` and its result: `order`, `circular`, `blocked`, and `steps`. The steps of `reachFormulas` count against the step cap, by design D6. `buildFormulaGraph` and `reachFormulas` take the `limits` of `runFormulas`.
- [ ] 5.2 Add `store/formula-engine/run.ts` with `runFormulas`, `unreached-sheet`, the `circular` and `blocked` cells of `reachFormulas`, a step count that starts at `steps`, the stored-value inputs, the `input-not-calculated` readers, the run caps, and the `limit` result, by designs D1 and D6. Add a test that `A3` `=SUM(A1:A3)` gives `circular-reference`. Add the property test of D6, with a simple name in its inputs, fixed seeds, an inline `mulberry32`, no new dependency, and 200 cases, and the test against the record of design D7.

## 6. Functions

Each task changes the expected values of its cases in `scripts/fixtures/calc-cases.ts`, by design D7. Each task also adds its functions to the generator of the property test of design D6.

- [ ] 6.1 Add `store/formula-engine/functions/` with the function table, `function-not-calculated`, `dynamic-range` for `INDIRECT` and `OFFSET`, and the aggregate and math groups: `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `ROUND`, and `ABS`, by design D5.
- [ ] 6.2 Add the logical group: `IF`, `AND`, `OR`, `NOT`, and `IFERROR`, with the lazy branches and the measured `IFERROR` case of D5. Add a unit test of the decision of D6: `C1` `=IF(X1>0,1,A1)` and `D1` `=C1+1`, with `A1` and `B1` in a cycle and `X1`=5. A run from `X1` gives 1 and 2. A full run gives `input-not-calculated`, with the subject `A1` for `C1` and `C1` for `D1`.
- [ ] 6.3 Add the text group: `CONCATENATE`, `&`, `LEN`, `LEFT`, `RIGHT`, `MID`, `UPPER`, `LOWER`, and `TRIM`, with the text step count, the text cap, and the measured `MID` cases of D5.
- [ ] 6.4 Add the lookup group: `VLOOKUP`, `INDEX`, and `MATCH`, with the wildcard matcher and the key check.
- [ ] 6.5 Add the time group: `TODAY` and `NOW` from `options.now`, and `no-clock`.

## 7. Benchmark

- [ ] 7.1 Add `scripts/bench/calc-graph-bench.ts` and its `bench:calc-graph` script in `package.json`: a build at its caps, and a run at the step cap for each kind of step. Set the caps of `store/formula-engine/engine-limits.ts` from its times, by design D6. Record the times in the pull request.
