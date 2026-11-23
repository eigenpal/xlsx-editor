# Design

## Context

The change `typed-ooxml-workbook-editor` built the cascade, the grid layout, the canvas painter, and the fidelity harness. Its decisions stay in force. This change was written after its code: each decision states what the code does and the measurement that it follows. The measurements were made on two new fixtures and their reference PDFs.

- `fills-patterns` (`scripts/fixtures/fill-style-cases.ts`) holds each non-solid `ST_PatternType` in three color states, the same patterns in adjacent cells and in a merged range, a linear gradient at 0 degrees over a merged range, linear gradients at 0, 45, 90, 135, and 270 degrees, path gradients with centered, corner, and off-center rectangles, and gradients with three stops.
- `cell-styles` holds 37 cell formats under three cell style formats. For each of the font, fill, border, number format, and alignment, a cell format omits the record or has its own, with `apply*` absent, `0`, and `1`. Style format 1 sets every `apply*` to `1`, and style format 2 sets every `apply*` to `0`. Style format 0 has its own fill, and one case with no `xfId` uses it.

The reference PDF draws each pattern as a tiling pattern, each linear gradient in one cell as an axial shading in the unit square of the cell, and each path gradient as an image of 1 point pixels. It draws each row of a merged range with a gradient as its own image, flipped vertically. `scripts/fidelity/pdf-fills.ts` reads the tiling patterns, the axial shadings, and the path gradient images, so tests compare the paint with the reference itself.

## Goals and non-goals

Goals: paint pattern and gradient fills as the reference PDF draws them; read the inner rectangle of a path gradient with bounded, reported values; record what a cell format under a cell style format shows.

Non-goals: conditional formatting, rotation and indent edits, and new `FormatPatch` fields (later changes); cell style names and cell style commands; table styles; theme changes; border line styles other than `thin`; a fidelity measurement of the screen profile; collaboration (first-change D10).

## Decisions

### D1. Pattern tiles

- `output/pattern-tiles.ts` holds one 8 by 8 tile for each non-solid `ST_PatternType`, with 1 point for each tile pixel. The foreground color paints the set pixels and the background color the others. An absent foreground paints black, and an absent background paints white, as the reference shows.
- `scripts/fidelity/pdf-fills.test.ts` reads the tile of each pattern in each color state from the reference PDF, and fails when the table differs.
- The tile grid of the reference PDF starts at the bottom-left corner of the page, the origin of PDF pattern space. `PaintOptions.patternOrigin` sets the origin for every pane in viewport points, and the harness passes that corner. When it is absent, each pane starts the grid at its sheet origin (`PaneLayout.sheetOrigin`), so a pattern moves with its cells when the view scrolls. Paint takes the origin within one bitmap period, so the pattern matrix keeps small numbers far down a sheet.

### D2. Pattern paint

- `output/paint-fills.ts` draws tiles into a bitmap that it creates from the document of the target canvas, never from a script global. The square bitmap holds `k` by `k` tiles, with `k` from 1 to 8 chosen for the smallest period drift for each tile (`tileBitmap`). Pattern pixel `j` covers the device pixels from `round(j × side / (8 × k))` up to `round((j + 1) × side / (8 × k))` (`pixelEdges`). So each pattern pixel covers one or more whole device pixels, and the period drifts by less than half a device pixel for each bitmap. A test checks the pixel widths at each ratio from 1 to 8 in steps of 0.001.
- When one tile is smaller than 8 device pixels, or larger than 512, the fill paints the mix of its two colors in the share of set pixels, which the tile table gives. The fill also paints the mix when the context cannot make a pattern, a bitmap, or a pattern transform.
- One paint makes each pattern once for each pane. The cache lives as long as its document, so a page that mounts and unmounts the editor keeps at most the cache limits. Bitmaps are cached for each document in a `WeakMap`, by pattern, colors, and bitmap size, with at most 256 entries and 16 Mi device pixels for each document; the entry used least recently goes first. A closed window releases its bitmaps.

### D3. Gradients

- A gradient paints over the box of its cell: the cell and the gridline pixel at its right and bottom, as a solid fill covers them. The reference PDF clips each gradient to that box: 67 by 23 points for a cell of 66 by 22 points and its 1 point gridline. So the gradient positions are those of the box.
- A `linear` gradient paints in the unit square of the box, which paint then scales to the box. Its axis runs from the center minus `(cos, sin) × (|cos| + |sin|) / 2` to the center plus that amount, so 45 degrees runs corner to corner. Lines of equal color join the same fractions of the box edges. The PDF test compares the axis at each fixture angle with the reference shading, and the size of each shading clip with the paint box from layout.
- A `path` gradient paints position 0 on its inner rectangle and position 1 on the box edges, and lines of equal color are rectangles. `pathShapes` paints, in order: the whole box in the color of position 0; the left and right parts as bands of the full height, one device pixel past the inner rectangle; then the top and bottom parts, each over the whole inner rectangle, where their gradients give position 0. So each anti-aliased edge lies over a shape of the same or a near color, and no seam shows the background. A part of zero size is left out.
- A unit test checks that the last shape over each sample point paints the position of the reference rule. The PDF test builds the shapes on the paint box from layout, which has the size of the reference image, and checks each pixel of the three red-to-blue path images within 0.02 of a position; the largest measured difference is 0.012.
- Colors interpolate in sRGB, which is the canvas default and what the reference shading samples show. The PDF test resolves the three-stop cells K18 and K20 through the cascade and paint. Each paint stop equals the reference shading color at its offset, and the shading at 0.15, 0.5, and 0.75 is within 2 per channel of the sRGB mix. Each K20 pixel is within 12 per channel of the paint color; the largest measured difference is 9.
- A merged range is one box in layout, so one gradient covers the whole range. The reference measures this only for a linear gradient at 0 degrees, because its row images are flipped.

### D4. The inner rectangle

- The cascade reads `@left`, `@right`, `@top`, and `@bottom` as the positions of the inner rectangle edges, as fractions of the width and height from the left and top edges (`GradientRect` in `cascade/cell-format.ts`). Paint applies them to the box of the cell, as D3 states. The off-center case measures this: `left="0.2" right="0.4"` paints position 0 from 0.2 to 0.4 of the width.
- An absent edge is 0, the schema default. A value outside 0 to 1 is limited to 0 to 1, and a value that is not a number becomes 0. A pair in reverse order is put in order. When any of these changes a value, the cascade records one problem `gradient-rectangle-out-of-range` for the rectangle. Its detail names the first change: the edge name and at most 16 characters of its value, or the pair. Layout records it with the cell reference, as for every cascade problem.
- These malformed values are tested in unit tests only, so no reference fixture holds content that a reader can repair.

### D5. Cell formats under cell style formats

- `cell-styles` measures that a cell format which omits its font, fill, border, or number format shows record 0. A cell format with no `x:alignment` shows the default alignment. This holds for each `apply*` state under style format 1, which applies all records, and with `apply*` absent under style format 2, which applies none. A cell format with no `xfId` and no `fillId` shows fill 0, not the fill of style format 0. No case shows the record of its cell style format.
- An own `x:alignment` that sets only `wrapText` replaces the whole alignment of the style format.
- The cascade already resolves cell formats this way (`pick` in `cascade/cell-format.ts`), and the format edits take record 0 as their base for an omitted record (`StyleRecords.xfFor`). So neither changes.
- The test `cell style formats against the measured reference` keeps the cascade equal to the measured record, and to the alignment that the reference PDF text positions show. The test `format edits under cell style formats` reopens the edited copy and checks that its edits keep record 0.

### D6. Evidence

- Both fixtures have reference PDFs, and their saved and edited copies have verified save checks. `cell-styles` has a measured record in `fixtures/measured.json`.
- Both fixtures pass the fidelity thresholds with no known gap. `fills-patterns` differs in 0.004 % of its pixels with both raster toolchains, and `cell-styles` in 0.256 % locally and 0.309 % in CI. Of the local `cell-styles` pixels, 95.1 % lie inside glyph boxes. These include text that paint draws in the canvas default font, because the painter does not set its font again after a canvas restore. That defect is in the first change, and a later fix removes it. So this score is not evidence for glyph paint, and the `cell-styles` limit is measured again after that fix. The rest lie in a band at the top of the B1 fill: the reference PDFs clip the print area 1 point inside its top and left edges, and paint does not. "Fill pixels" names this. The 76 counted `fills-patterns` pixels come from the reference raster. Poppler 26.10 draws each 8 point tile boundary of a pattern one raster pixel off at the supersample density, so the reference shows one mixed pixel row in each tile period, and paint shows none. Most of these pixels count as anti-aliasing; the 76 counted ones lie next to cell edges. CI, with Poppler 24.02, gives the same score.
- One case cell of `fills-patterns` covers about 0.3 % of its page, so the page limit of 0.5 % cannot see one wrongly painted case. `FIXTURE_PIXEL_MISMATCH` in `scripts/fidelity/thresholds.ts` gives this fixture a limit of 0.02 %, five times its score. Measured on the page, the wrong cases tried differ in 0.024 % to 0.29 %, so the smallest is about 1.2 times the limit. `fidelity:check-gaps` lets this limit only tighten, and never lets a limit of one fixture be looser than the page limit. This follows the first-change requirement that thresholds only tighten. `cell-styles` has a limit of 0.4 %, so one wrong fill case fails. The probe test checks that the cascade resolves its font and border cases one by one, because they change too few pixels for a page limit. No gate in this change checks the face that paint draws. Every other fixture keeps its score. Each known-gap list records the line counts of both fixtures. The counts come from the reference PDF text, and the CI run of the pull request confirms the CI list.
- The approximation "Pattern fills" becomes "Fill pixels" (`fill-pixels`), which names the differences that remain. The support matrix row `formatting.fills` is renamed "Fills" and cites both fill fixtures. The new row `formatting.cell-style-formats` cites `cell-styles`.

## Risks and trade-offs

- Each pane rounds its own sheet origin to a whole device pixel, so a pattern can step by one device pixel at the edge between two panes. "Fill pixels" names this.
- A gradient over a merged range, other than a linear gradient at 0 degrees, has no measurement. "Fill pixels" names this.
- The reference PDF draws a path gradient as an image of 1 point pixels, and paint draws it smooth. The positions differ by at most 0.012 on the fixture.
- The harness paints at three device pixels for each CSS pixel, so it does not see screen rounding. The unit tests cover the bitmap and fallback rules at the ratios of the screen.
- At a device pixel ratio that is not a whole number, the browser can resample the canvas, so pattern pixel edges can blur. The tile rule holds for the canvas bitmap only. No browser measurement exists.
- A view with more than 256 pattern and color pairs, or with bitmaps of more than 16 Mi device pixels in all, misses the tile cache and builds bitmaps again on each paint.
