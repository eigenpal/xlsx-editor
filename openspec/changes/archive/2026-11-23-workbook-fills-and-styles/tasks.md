## 1. Fixtures and measurements

- [x] 1.1 Add the fixtures `fills-patterns` and `cell-styles` (`scripts/fixtures/fill-style-cases.ts`), with their reference PDFs, save check records, and the measured record of `cell-styles`.
- [x] 1.2 Add `scripts/fidelity/pdf-fills.ts`, which reads the tiling patterns, axial shadings, and path gradient images of a reference PDF, with its test against `fills-patterns`.

## 2. Fills

- [x] 2.1 Add the tile table and paint pattern fills by design D1 and D2, with `PaintOptions.patternOrigin`, the bottom-left page corner in the harness, and `PaneLayout.sheetOrigin` on screen.
- [x] 2.2 Read the inner rectangle of a path gradient by design D4, with the problem `gradient-rectangle-out-of-range` and its tests.
- [x] 2.3 Paint linear and path gradients by design D3, with the geometry tests and the path position test against the reference PDF.

## 3. Cell style formats

- [x] 3.1 Add the test that keeps the cascade equal to the `cell-styles` measured record and reference PDF, and the reopen test of the edited copy (design D5).

## 4. Evidence

- [x] 4.1 Run the harness over every fixture, record the line counts of the new fixtures in each known-gap list, and give `fills-patterns` and `cell-styles` their own pixel limits.
- [x] 4.2 Replace the approximation "Pattern fills" with "Fill pixels", and update the support matrix rows.
- [x] 4.3 Run `openspec validate workbook-fills-and-styles --strict`.
