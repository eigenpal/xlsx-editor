## Why

The change `typed-ooxml-workbook-editor` paints solid fills only. A pattern fill paints as one mixed color, and a gradient fill has no measurement: a `path` gradient paints as a circle, and a diagonal `linear` gradient stops short of the cell corners. The first change also did not measure what a cell shows when its cell format omits a record that its cell style format sets.

## What Changes

- Paint each pattern fill as its 8 by 8 tile of 1 point pixels, from the tile table that the reference PDF of `fills-patterns` measures.
- Paint `linear` gradients in the unit square of the cell, and `path` gradients from their inner rectangle out to the cell edges, as the reference PDF draws them.
- Read the inner rectangle of a `path` gradient, and record each value that the cascade limits or reorders.
- Measure, on the fixture `cell-styles`, that a cell format which omits a record shows record 0 and never the record of its cell style format. The cascade and the format edits already follow this rule, so they do not change.
- Replace the approximation "Pattern fills" with "Fill pixels", which names the differences that remain.

## Capabilities

### New Capabilities

- `workbook-fills-and-styles`: Pattern and gradient fill paint, the inner rectangle of a path gradient, and the measured resolution of cell formats under cell style formats.

### Modified Capabilities

None. The capabilities of `typed-ooxml-workbook-editor` are not archived yet. This change extends its `workbook-cell-formats` and `workbook-grid-layout` capabilities with a new capability instead of modified ones.

## Non-goals

Conditional formatting, rotation and indent edits, and new `FormatPatch` fields belong to the later changes `workbook-conditional-formatting` and `workbook-format-commands`. Cell style names, a command that applies a cell style, table styles, theme changes, and border line styles other than `thin` are not in this change.

## Impact

The change adds code to the `cascade`, `layout`, and `output` lanes of `packages/core`, and to `scripts/fixtures/` and `scripts/fidelity/`. `ResolvedFill` grows the inner rectangle of a path gradient, `FillPaint` grows a pattern kind and the inner rectangle of its gradient kind, `PaintOptions` grows `patternOrigin`, `PaneLayout` grows `sheetOrigin`, and `CascadeProblemCode` grows `gradient-rectangle-out-of-range`. The entries `@xlsx-editor.dev/core/cascade`, `@xlsx-editor.dev/core/layout`, and `@xlsx-editor.dev/core/output` export these types, and that package is private and not published, so no published API changes. The fidelity harness gains a pixel limit for each fixture, which only tightens, and gives `fills-patterns` a limit of 0.02 % and `cell-styles` a limit of 0.4 %. The change adds two fixtures with their reference PDFs, records, and save checks, one new support matrix row, and renames one row. Collaboration stays out of scope, as design D10 of the first change states.
