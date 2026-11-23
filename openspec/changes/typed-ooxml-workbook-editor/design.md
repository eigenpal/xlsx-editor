## Context

DOCX Editor and PPTX Editor share one engine shape: a bounded package read, one canonical OOXML tree per part as the only source of truth, typed ops as the only write path, a DOM-free cascade and layout, a painter that never re-derives geometry, and thin React and Vue adapters over one `core` package. XLSX Editor takes the same shape. This design records only what is different for SpreadsheetML, and the decisions that the first slice needs.

SpreadsheetML differs from the other two formats in four ways that matter here:

1. The content is a sparse grid, not a flow. A sheet can name cells far apart, and a file can claim any size.
2. Many cells hold formulas with cached results. The results depend on other cells.
3. Geometry comes from column widths in character units and row heights in points, not from text flow.
4. Shared structures (the shared string table and the cell formats) are indexed from every cell.

## Goals / Non-Goals

**Goals:**

- Round-trip safety for every part, known or unknown.
- Measured fidelity for the grid, cell values, and cell formats on synthetic fixtures.
- A browser viewer checkpoint, then core cell editing, then a second hands-on checkpoint.
- The same security rules, gates, and SDK shape as DOCX Editor and PPTX Editor.

**Non-Goals:**

- Formula evaluation. No calculation engine is in this change.
- Charts, pivot tables, conditional formatting, data validation, sparklines, and slicers beyond lossless preservation.
- Print pagination, except the single-page fixture layout that the fidelity harness needs.
- Collaboration. It is its own change, which starts only after the core editing checkpoint passes (design D10).
- The automation API, the PDF converter, and the review module. Each gets its own change.

## Decisions

### D1. Package and lane layout

`packages/core` holds the lanes `store`, `cascade`, `layout`, `automation`, `output`, `editor`, and `contracts`, with the same edges and environments as PPTX Editor's lane graph. The lane graph guards (`core-lane-graph.test.ts`, `core-lane-imports.test.ts`, `browser-bundle-graph.test.ts`) are copied from PPTX Editor with their provenance headers. `packages/ooxml-core` stays private and inlined.

Alternative considered: one package per lane, as DOCX Editor had before its consolidation. Rejected: DOCX Editor moved to one `core` package to load one engine copy, and that reason applies here too.

### D2. Canonical tree and sparse cell index

Each XML part is one canonical tree, as in the other products. The worksheet tree keeps `x:row` and `x:c` elements as typed nodes in document order. A derived index maps a cell reference to its node, built from the cells that exist. Nothing in the store allocates or iterates by a range that the file names (`x:dimension/@ref`, `x:row/@spans`, `x:col/@min` and `@max`, merged ranges).

Shared strings and cell formats stay in their own parts. Cells hold indexes, as in the file. An edit that adds a string appends it to the shared string table, and saving never compacts the table in this change. A save writes `x:sst/@count` and `@uniqueCount`, and the `count` attribute of each collection in the style part, from the serialized content.

Ops express intent, never shared-part indexes. An op names a sheet and a cell by reference and carries a value, a formula text, or a format as properties (for example, bold and a fill color). The store resolves or appends the shared string and cell format entries when it applies the op.

Ops address cells by sheet and A1 reference against the revision they apply to. This is the decision for this change. If the collaboration change needs stable identities, it translates them to A1 references at its own boundary, before ops reach the store, so the ops, their values and formats, and their tests stay.

Undo records inverse ops. When the store applies a transaction, it records, for each touched node, the ops that restore its prior content, including content that the edit removes (cells, `x:mergeCell` and `x:col` entries, and the formula text that becomes `#REF!`). Undo applies the inverse ops as a new transaction. The store never restores a whole-tree snapshot.

Alternative considered: convert shared strings to inline strings on load. Rejected: it changes every cell on save and breaks the save-and-reopen digest for unedited cells.

### D3. Values without a calculation engine

The engine shows what the file stores. A formula cell shows its cached `x:v`. The engine never evaluates a formula.

An edit can change a formula's result. The engine cannot know which formulas depend on an edited cell until it has a reference reader. So, in this change:

- The stale state lives in the tree, not in the session. `x:calcPr/@fullCalcOnLoad="1"` in the workbook part makes every formula cell of the workbook stale. `x:sheetCalcPr/@fullCalcOnLoad="1"` in a worksheet makes the formula cells of that sheet stale. The editor shows each stale cached value with a stale state that the snapshot exposes and the adapters render.
- The first transaction that changes a value or the grid structure in a workbook that holds a formula cell after the transaction also sets `x:calcPr/@fullCalcOnLoad`, in the same transaction. So the state is part of the document, and a save writes it, so that a consumer recalculates. The cached `x:v` values stay in the file.
- A formula cell is an `x:c` in a sheet part with an `x:f` child, or with an `x:f` child of a branch of an `mc:AlternateContent` child. A wrapper nested in a branch counts too. The grid mark reads this rule for each cell. The flag reads it through a scan of every parsed sheet part whose root element is in the Transitional SpreadsheetML namespace. The stale notice scans the shown sheet only. Inside a cell, the scan reads only a nested `x:c` and an `mc` element.
- A sheet part that the scan cannot read counts as one with a formula cell: a parsed root in another namespace, or a part that the reader keeps as unread bytes (a macro sheet) with an XML content type or none. A chartsheet and a dialog sheet hold no cells, so they never count, parsed or not.
- The wrappers of a cell are read in document order. In each wrapper, `mc:Fallback` comes first, as the branch for a consumer that understands no `mc:Choice` (ECMA-376 Part 3), then each `mc:Choice` in document order, and a branch comes before the wrappers nested in it. The formula bar shows the text of the first `x:f`. For a shared child there, it shows the text of the master that holds its `x:f` directly, moved to the child, so the field and the grid name the same cell. The cached value is the cell's `x:v`. When the cell has none, it is the `x:v` of the formula's branch, else of the first branch with an `x:v` and no `x:f`.
- The unshare of design D4 reads only a direct `x:f`. A value edit and a clear of a cell with any formula in an `mc:AlternateContent` branch, a formula range master too, are refused with `cell-alternate-content` (`cellChildBlocker`). The array range check also reads an array or data table master in a branch, so an edit, a clear, or a copy of a cell inside its range is refused, as for a direct master. Structural edits, a sheet rename or delete, and the copy of a formula read only a direct `x:f`, so they do not rewrite or carry a wrapped formula. The unshare cannot rewrite a shared child whose `x:f` is in an `mc:AlternateContent` child of its cell, so a value edit and a clear of its direct master are refused too. A clear that also clears that child is refused with `cell-alternate-content`, as for any wrapped formula.
- Undo of the edit does not clear the attribute, because the engine cannot prove that the cached values are current again. The op that sets it is outside the undo history: the undo manager does not track it.
- A file that is opened with either attribute set is stale from the start, for the scope of that attribute. The snapshot exposes this state from task 4.3, so the viewer shows it. A saved and reopened file therefore keeps its stale state.
- A formula cell with no cached `x:v` (for example, after a formula edit) shows no value and the stale state.
- The automation API, when it exists, reports stale results as stale, never as current.

Alternative considered: show cached values as current. Rejected: it presents data that can be wrong as correct, and AGENTS.md forbids silent approximation. Alternative considered: refuse every edit in a workbook with formulas. Rejected: it blocks the core editing slice for most real workbooks.

A later change can add a bounded formula reference reader. Then only the formulas that can reach an edited cell become stale.

### D4. Structural edits and references

Row and column insertion and deletion move every reference to a cell position. This change types the following elements for structural edits, in addition to the typed nodes of the proposal: cells (`x:c/@r`, `x:row/@r`), merged ranges, column entries, row and column properties, hyperlinks (`@ref`), defined names, formula references in every sheet (including references from other sheets), shared and array formula ranges (`x:f/@ref`), `x:dimension`, selections and panes (`x:selection`, `x:pane/@topLeftCell`, `x:sheetView/@topLeftCell`), and drawing anchors (`xdr:from`, `xdr:to`). In this change:

- A bounded A1 reference tokenizer reads formula text and moves the references that it recognizes, including shared formula masters and defined names.
- Formula text that the tokenizer cannot read is refused when a user enters it. Formula text from the file that the tokenizer cannot read is preserved, and it makes every structural edit refused while it exists.
- A string argument is text, not a reference, so the tokenizer never moves it. A formula that builds a reference from text (for example, with `INDIRECT`) keeps its text. The editor marks its result stale, as for every formula.
- If a formula, a defined name, or a part holds a reference that the tokenizer cannot move with certainty, the store refuses the structural edit with a typed reason. It never writes a partly moved workbook.
- An edit that overwrites, clears, or deletes the master cell of a shared formula first writes each child of that shared formula as its own formula, then applies the edit. If the tokenizer cannot read the master text, the store refuses the edit. A save never writes a child whose master is gone.
- Shared formulas: a child cell's formula is the master's text moved by the child's offset. If an edit moves the master's references and a child's references differently, the store writes the affected children as their own formulas (it unshares them), or refuses. Array formula ranges move as a whole; an edit inside an array formula range, or an insertion that splits one, is refused.
- Drawing anchors: a `xdr:twoCellAnchor` whose `@editAs` is `twoCell` (the default) moves and resizes with its cells. A `xdr:oneCellAnchor` moves its `xdr:from` with its cell and keeps its `xdr:ext`. A `xdr:absoluteAnchor` does not change, because it has no cell position. A `xdr:twoCellAnchor` whose `@editAs` is `oneCell` or `absolute` keeps its size or its position in points, which needs column widths in the store lane; until a later change provides them, the store refuses a structural edit that moves such an anchor. A drawing part that the edit cannot read (kept as bytes, or not a `xdr:wsDr`), or one whose relationships part is kept as bytes, refuses every structural edit of its sheet with `structure-blocked`, because its anchors would stay on their old cells.
- The edge rule of an insertion: an insertion at the first row or column of a range moves the range; an insertion right after its last row or column does not grow it; an insertion strictly inside it grows it. A frozen pane band (the first `ySplit` rows or `xSplit` columns of a `frozen` or `frozenSplit` pane) follows the same rule from its first row or column: an insertion inside it grows it, and the view keeps its top-left cell. A deletion inside the band shrinks it, a split that reaches 0 is removed with its axis (and the pane when both go), `x:pane/@topLeftCell` stays after the band, and `@activePane` and each `x:selection/@pane` name only the panes that remain.
- An insertion that would push a cell, a row with properties, or a merged range past the last row (1048576) or the last column (16384) is refused. A column entry (`x:col`) whose `max` reaches the last column is cut at the last column, because no cell is lost; the same applies to row properties that reach the last row. A formula reference to a cell that an insertion pushes past the last row or column becomes `#REF!`, because the reference has no cell to name; a range whose end is pushed past it is cut at the last row or column.
- Every other element or part that can hold a reference is preserved as a generic node. If one exists in the workbook and the edit changes a position that it can refer to, the store refuses the edit and names it. This is the default for anything this list does not name, including extension content under `x:extLst` (which holds `xm:sqref` and `xm:f` references), for example conditional formatting, data validation, `x:autoFilter`, `x:sortState`, `x:protectedRanges`, `x:ignoredErrors`, page breaks, `x:cellWatches`, scenarios, custom sheet views, tables, pivot tables, charts, and comments with their legacy drawing anchors. A table or query table part on any sheet whose text names the edited sheet refuses the edit, and so does a sheet range (`First:Last!`) that spans it. A relationship of the workbook or of the package outside an allowlist refuses the edit and names its part: the workbook allows worksheets, the shared string table, styles, the theme, external links, the calculation chain, custom XML data, and cell metadata; the package allows the workbook and the core, extended, and custom properties and the thumbnail.
- A hyperlink that a deletion removes takes its relationship (`x:hyperlink/@r:id`) with it, unless another hyperlink of the sheet names it. Undo puts it back.
- A structural command on a selection of the whole sheet is refused with the reason `whole-sheet`, because it would insert or delete every row or column.
- The inverse ops of the history (`restoreRow`, `restoreStyles`, and `restoreParts`) put back content of an earlier revision with no check of what it holds, so only undo and redo apply them. `transact` refuses them, and `WorkbookOp`, the op type that `transact` takes, does not name them.
- A reference to a deleted cell becomes the error value `#REF!`. Confirm each rule of this item on a fixture before code depends on it. A range that loses some of its rows or columns shrinks. A range that loses all of them becomes `#REF!`. The same rules apply to defined names. Merged ranges and hyperlinks inside the deleted cells are removed with them. A merged range that a deletion partly covers shrinks, and is removed when one cell remains.

A refused structural edit does not lock editing. AGENTS.md says that unknown content never locks editing: cell value edits, clearing, and formatting stay available in every workbook, and only the structural command that would corrupt a reference is refused, with a reason that names the part.

### D5. Calculation chain

`/xl/calcChain.xml` is optional. A save that adds, removes, or moves a formula drops the part, its relationship, and its content type override, so that no consumer reads a stale chain. The save check confirms that a file saved this way opens with no repair prompt. If the check shows a prompt, this decision changes before editing ships.

### D5a. Save-verified records

The save check runs outside CI. It opens each copy and records any repair prompt. It writes one committed record per fixture: the SHA-256 of the saved copy and of the edited copy that the check opened, and the result for each. The package writer lives in the `store` lane (`packages/core/src/store/package-write/`) and is deterministic, so CI writes both copies again and checks that their hashes match the record. The edited copy applies the edit that the fixture builder defines for that fixture. Any change that alters the saved bytes (the writer, the reader, the store ops, a fixture, its edit, or a dependency) makes the record stale until the check runs again. Alternative considered: a serializer version hashed from the writer's source files. Rejected: the bytes also depend on the reader, the ops, and dependencies, and a source hash cannot show which bytes the check opened.

### D6. Grid geometry

Layout works in points with an injected text measurer, as in the other products. Column widths come from the formulas of ECMA-376 Part 1, 18.3.1.13, with the maximum digit width measured from the workbook's default font. The pixel rounding in those formulas needs a pixel density that the standard does not state.

The reference PDFs measure 72 pixels for each inch, as the quick reference records. Layout keeps the density in a grid profile. So a target with another density rounds to its own pixels. The reference PDFs print every row height scaled by 14/15 and rounded down. So layout has a print profile, which the harness uses, and a screen profile.

No committed reference shows the screen geometry; the known approximation "Screen geometry" names it. Every caller of layout names its profile.

The fidelity harness renders the editor's own painter (the target of D7) in a headless browser at a fixed device pixel ratio. So the measured pixels are the editor's pixels. A reference PDF can show printed widths that differ from the layout widths. Then the difference is a known gap with a measured value, not a second layout.

The grid is virtualized: layout returns the geometry of the visible rows and columns, and a prefix-sum index answers the offset of any row or column without walking the sheet. Cost depends on the rows and columns that have explicit properties, on the rows whose height comes from the fonts of their cells, and on the visible cells. Each of these is in proportion to the content of the file, never to the sheet size.

### D7. Paint target

The layout records are target-neutral. The first checkpoint painted the visible window as DOM (one SVG with text elements at the measured glyph positions), because it gives accessible text and the same selection model as the other products. A measured spike compared that with a canvas painter on a large fixture. The faster target is kept, with a text alternative of the visible window, and the decision is recorded here.

**Decision: the canvas painter, with the visible cell texts as a text alternative.** The canvas draws the window glyph by glyph at the measured positions.

A canvas holds no text. So the viewer puts a visually hidden row group of the same layout in its `grid` element (`output/accessible-grid.ts`). The group has one `row` for each visible row with text. Each cell gets one `gridcell` with the cell text, row index, and column index.

This is a text alternative for the visible window only. It does not set `aria-rowcount` or `aria-colcount`, so assistive technology does not learn the size of the sheet. It has no active-cell semantics: no focus or `aria-activedescendant` on a cell. Full grid semantics are later work of the editing tasks. The selection model does not change: it maps points to cells through the layout (`view-map.ts`).

The fidelity harness takes a screenshot of the canvas for the pixel score. For the line scores, it prints the page to PDF in Chromium. That PDF keeps the glyphs of a canvas painted during the print as text, at their painted places. The harness reads that PDF as it reads the reference PDF.

The printed PDF leaves out a glyph that a cell clip hides entirely. So a clipped line matches only by the clip rectangles of both PDFs (approximation "Line scores of clipped text", `clipped-line-scores`). The reference text runs past the reference clip on the cut side. The painted clip and the end of the painted line are within `clipEdgePt` of it. Without white space, the painted text starts the reference text on a right cut, or ends it on a left cut. The painted line holds each reference glyph whose advance starts before the clip edge, or ends past it on a left cut, and no glyph more. When the reference glyphs do not account for the text, such as a ligature, the gate fails that end as a skipped check.

The spike built a workbook of 40,000 rows and 26 columns at runtime (shared strings, numbers, a fill on every third row, a bold font and a bottom border on every fifth column, gridlines, and a frozen first row and column). The spike script was not committed, so nobody can run it again from this repository. Headless Chromium 156 painted 40 windows of 1440 x 900 CSS pixels at a device pixel ratio of 2, each after a new scroll position, with the substitute fonts. The values below are medians of the second pass, in milliseconds, from one machine. They show the relative cost of the targets, not a performance guarantee. "Frame" is the time from the start of paint to the next animation frame.

| Zoom  | Target                     | Cells with text | DOM nodes | Layout | Paint | Frame |
| ----- | -------------------------- | --------------- | --------- | ------ | ----- | ----- |
| 100 % | SVG                        | 1,150           | 3,740     | 2.7    | 27.2  | 33.9  |
| 100 % | Canvas                     | 1,150           | 1         | 2.8    | 3.3   | 14.0  |
| 100 % | Canvas and accessible rows | 1,150           | 1,198     | 2.8    | 6.7   | 14.0  |
| 50 %  | SVG                        | 2,158           | 6,986     | 5.1    | 55.6  | 72.7  |
| 50 %  | Canvas                     | 2,158           | 1         | 5.6    | 6.0   | 17.1  |
| 50 %  | Canvas and accessible rows | 2,158           | 2,243     | 5.4    | 11.7  | 19.6  |

Of the SVG paint at 100 %, building the elements took 5 ms; the browser's style and layout of the SVG took the rest. The SVG target misses a 60 Hz frame at every zoom; the canvas with accessible rows stays near it. Opening the workbook took 1.7 s for both targets, because it does not depend on the paint target.

### D8. Fidelity against reference PDFs

A reference PDF of a sheet is a printed page, so the page setup affects it. Fixtures fix the page setup so that the printed page maps to the grid. They use one page, 100 % scale, known margins, and gridlines and headings set explicitly in `x:printOptions`. The harness renders the same cell range at the same scale through the editor's own layout and paint. It then compares the printable area.

The harness lands with the first painter (task 6.2). Before it, values and layout land with their fixtures and reference PDFs. Unit tests check the layout records against values measured on those PDFs. Task 6.2 then runs the harness over every earlier fixture.

The thresholds and the gap rules are PPTX Editor's (`scripts/fidelity/thresholds.ts`), copied with a provenance header into one thresholds file. One difference from PPTX Editor comes from the canvas paint, which PPTX Editor has no rule for. A list names the lines of each fixture that a cell clip shortens on both pages, where the painted page holds only part of the reference text. A key holds the text, the baseline, and the clip edge of the reference PDF, once for each cut.

The harness reads the clip rectangles of both PDFs from their operator lists. The reference PDF clips only text that overflows its cell, and the other text of a row shares the clip of the print area. The painted page clips the text of each cell. So the rule needs each reference clip as one rectangle. Without it, a cut line stays missing, and each end of a matched line fails the gate as a coverage gap. A clip hides glyphs that the line scores cannot check. So a run fails when it clips a line more often than its list names it. A change fails the merge-base check when it names a line more often than the merge base. An owner can accept the clipped lines of a new fixture, or of a fixture with a new reference PDF.

The clip rules use `clipEdgePt` as their tolerance, apart from the line position threshold. The run fails in these cases, each a regression:

- The reference clip cuts a line, and the painted line is not cut at the same edge. This rule also holds when the texts match in full.
- The reference clip holds an end of a line whole, and the painted line runs past the painted clip there.
- The painted clip hides whole glyphs at such an end, or its edge lies in a space there.
- The painted clip is unknown at such an end.
- A painted line of a join ends inside the text of one reference item, so no clip rule checks that end. The reference PDF can also draw two cells as one item, so the remedy names both causes.

The rules apply to each end of a line of upright text where the reference clip is known. An end with an unknown reference clip fails the gate: a skipped check is a gap, never a pass. Rotated, mirrored, and vertical text has no line score, so no clip rule checks its cell. The run and the report state this gap with the count of such reference items. The rules compare only the left and right clip edges. A line of one character other than a digit has no line score, so no clip rule checks its cell. Stacked text is upright, and each of its lines holds one character, so this limit covers it. The run and the report state this gap with the count of such reference lines. Lines pair in this order. First, lines of the same text on one row, within `maxBaselineDeltaPt`, pair. Then the bullet form on the row pairs, and then the joins. Then the clipped pass and the extra-cut pass run. Then lines of the same text that are left pair on any row. Last, a clipped line that is left pairs with a cut of it on any row. Each step except the clipped pass and the extra-cut pass takes its pairs nearest pair first. So its pairs do not depend on the order of the lines on a page. The clipped pass and the extra-cut pass take the reference lines in page order. Repeated text on one row can still pair across cells. A wrong pair of clipped text fails the gate. Several painted lines can hold one reference line. Then the rules also apply to each inner end of a painted line, against the reference clip of its characters. A matched line that a clip cuts on both sides compares both edges as cuts.

Paint draws no glyph for a space, so a clip edge can lie in a space of the reference text. The painted line then ends at the glyph before that space. It is cut when all of these hold, each within the threshold:

- The painted line ends at the advance end of that reference word.
- The clip edge lies past that end.
- The clip edge lies before the advance start of the next reference word.

Each known gap records its fixture, its check, and its measured value. A gap only tightens. When a score improves beyond the tolerance of its check, the recorded value moves to the new score in the same change. A new gap needs a new fixture.

A check compares the thresholds and gaps with those at the merge base, as PPTX Editor's `fidelity:check-gaps` does. It fails when a threshold or a gap becomes looser. It also fails when a fixture that has a gap is renamed or removed. That check differs from PPTX Editor's in nine ways, in the order of its header (`scripts/fidelity/check-known-gaps.ts`):

1. A new gap of a fixture of the merge base fails, also in a change that touches no engine or harness source. So does a gap of a check that the merge base does not have: a new check lands without a gap.
2. A new threshold starts at its tightest value. A new tolerance is at most the smallest tolerance of the merge base.
3. Each threshold names its loosening direction in `THRESHOLD_DIRECTIONS`, and the entry of the merge base may not change.
4. Known gaps are kept for each raster toolchain.
5. A list names its clipped lines, which only become fewer. The first record of clipped lines is not compared, because no list of the merge base has one. After that, a list without the record counts as an empty record, and a list may not drop the record that its merge base list has. A record of a name that is not a fixture fails. A change that removes a fixture of the merge base and adds one can rename it. Then the clipped lines of the added fixture fail, also before the first record, and so do the checked lines of the removed one. A change that only removes a fixture without a known gap passes, and a later change that adds the name again adds a new fixture. This two-change path is an accepted limit, because the removal shows in review. A fixture with a known gap cannot be removed until its gap is fixed. A fixture never checks fewer lines, also not with a new reference PDF.
6. The fixtures of the change include those of the fixture catalog, not only those of the reference manifest.
7. The base commit is optional and defaults to the merge base with the main branch.
8. Each error prints one line. An error that stops the check exits with code 2, so exit code 1 always means a looser list.
9. With `--require-thresholds`, a base without `thresholds.ts` exits with code 2, because the check compared nothing. `bun run fidelity --update-gaps` passes this flag. Only a missing file counts as a base without thresholds.

As in PPTX Editor, `bun run fidelity --update-gaps` writes nothing when the new list adds a gap or a check. Here it also writes nothing when the list adds a clipped line or a value is looser, because the list records values. After a write, the run starts the merge-base check, which PPTX Editor does not do. `--accept-worse` writes such a list, for a new fixture or a new toolchain. The merge-base check then applies its own rules, also in CI. A gap of a new fixture passes there. A looser list for a new toolchain needs `XLSX_FIDELITY_ACCEPT_NEW_TOOLCHAIN`, and a new clipped line needs `XLSX_FIDELITY_ACCEPT_CLIPPED`. With fixture names, the update replaces only the entries of those fixtures. A run with tools that have no list starts one, so it needs every fixture.

As in PPTX Editor, a threshold goes away only when `RETIRED_LIMITS` names it with the reason.

The font check of a reference PDF also differs. As in PPTX Editor, the check refuses a reference PDF that draws with a font that it does not embed. It also refuses a reference PDF that draws with a family that the workbook does not name.

Both products read the family (name id 1) of each PDF font from the name table of a font file. That file has the same PostScript name. The reference fonts here are not in the repository and CI has none. So a name that no font file resolves takes its family from the name itself.

A check of the face and size of each matched line lands in its own change. Its condition holds: paint applies the font of a conditional formatting rule (`overlayDxf` in `packages/core/src/cascade/cell-format.ts`). The reference PDF of `preserved-parts` draws one value bold through a rule. The check stays open under #20. The merge-base check refuses a gap of a new check.

Scores are recorded for each font environment. One difference from PPTX Editor: a missing reference fails the gate here. The fixtures are synthetic, and each one lands with its reference PDF. Fixture builds are deterministic (fixed zip timestamps and part order), so a fixture hash identifies its reference. Print pagination, scaling, and page breaks are out of scope.

The pixel score also differs from PPTX Editor's. PPTX Editor draws its own printed PDF with the rasterizer of the reference PDFs (`scripts/fidelity/score-deck.ts`). Here the pixel score compares a screenshot of the canvas with the Poppler raster of the reference PDF. The line scores read the printed PDF.

The passing fixture nearest to the page limit (`maxPixelMismatch`, 0.5%) is `number-formats`. It scores 0.490% with Poppler 26.10.0 on the local toolchain and 0.428% with Poppler 24.02.0 on the CI runner, both with Chromium 158.0.8127.18. The page has 1,938,816 pixels, so the rest of the limit is about 195 pixels with Poppler 26.10.0 and about 1400 pixels in CI: a smaller paint change cannot fail the pixel score. With Poppler 26.10.0, all 9499 of its differing pixels lie inside the glyph boxes of the reference text (`scripts/fidelity/glyph-share.ts`). `scripts/fidelity/rasterizer-control.ts` draws the reference PDF as canvas text with pdf.js in the same Chromium. That raster differs from the editor page in 3 pixels, so the two rasterizers cause almost all of the difference (approximation "Glyph rasterization"). The CI runner has no measurement of the control.

History: earlier values of this section came from reference PDFs that embedded other faces, and from paint that drew most glyphs in the default font of the canvas. `scripts/fidelity/thresholds.ts` keeps two of them, 0.507% and 0.900%, as history, and the git history keeps the rest. This tree cannot reproduce them.

A later change switches the pixel score to a Poppler raster of the printed PDF. It can merge when every fixture, `number-formats` included, scores at or under its limit (`pixelLimit`, at most `maxPixelMismatch`) with that raster on each recorded toolchain. No run on this tree measures that raster.

### D9. Fonts and text measurement

Text is shaped and measured with the same HarfBuzz measurer as PPTX Editor, copied with provenance until it moves to `packages/ooxml-core`. Metric-compatible substitute fonts come from `@docx-editor.dev/fonts`. Reference fonts for the harness come from the machine, through an environment variable, and are never committed.

### D10. Collaboration comes after core editing

Base features come first. This change ships the viewer and core editing without collaboration, and stops at a hands-on checkpoint after editing (task 7.8). Collaboration is a separate OpenSpec change that starts only after that checkpoint passes. No task in this change waits for it.

This change keeps the store ready for it, and nothing more:

- Every edit is a typed op in a `WorkbookDocumentStore` transaction, with no second write path.
- The tree holds no session state: the stale state of D3 is an attribute in the workbook part.
- No module keeps document state in a global.

The collaboration change starts from these inputs, found during this change: positions and shared-part indexes must converge without a grid-sized structure, a merge must never need a refusal, and undo must restore moved references. That change decides the position model; this change does not.

## Risks / Trade-offs

- [Collaboration later can require store changes] → The store has one write path of typed ops and no session state, so the collaboration change adds replication around the same ops.
- [The stale rule of D3 marks every formula after one edit] → The viewer shows the cached values with a clear state, and a later reference reader narrows it.
- [Structural edits refuse often while parts stay generic] → Typed support for the parts that hold references comes in later changes; the refusal names the part.
- [Column width rounding differs by font and platform] → The harness measures it per font environment, and known gaps record each difference.
- [The paint target decision comes late] → Layout records stay target-neutral, so the decision changes only the output lane.

## Open Questions

- Which font environments the harness gates on. PPTX Editor gates on reference fonts on one developer platform and on substitute fonts in CI. Each font environment has its own recorded scores. CI gates on the substitute-font environment; a local run with the reference fonts gates on that environment. A run in an environment with no recorded scores reports a gap, never a pass.
- How a host sees stale results in its own chrome: a snapshot field, an event, or both. The adapter parity contract settles the shape.
