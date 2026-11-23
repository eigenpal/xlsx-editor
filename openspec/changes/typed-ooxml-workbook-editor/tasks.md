## 1. Repository skeleton

- [x] 1.1 Set up the workspace, tooling, gates, and CI in the shape of DOCX Editor and PPTX Editor: Bun, oxlint with the line caps and repository rules, Prettier, the pre-commit hook, the parallel test runner, Changesets, the CLA, issue forms, and the pull request template.
- [x] 1.2 Copy `@eigenpal/ooxml-core` from PPTX Editor with a provenance header on each file, regenerate the preset geometry, and keep `bun run test:docx-shared` working against a DOCX Editor checkout.
- [x] 1.3 Write `AGENTS.md`, `CONTRIBUTING.md`, `README.md`, and `docs/repo-map.md`, and the naming gate (`bun run check:terminology`).
- [x] 1.4 Copy the ECMA-376 reference files and write the SpreadsheetML and theme color quick references.
- [x] 1.5 Establish this change as the only production authority.

## 2. Store: read, model, and save

- [x] 2.1 Create `packages/core` with the lane graph and its guards, copied from PPTX Editor with provenance headers.
- [x] 2.2 Write the deterministic fixture builder (`scripts/fixtures/`), with the page setup of design D8 and a defined edit for each fixture's edited copy. Each later task adds its fixtures to it.
- [x] 2.3 Read the package: content types, relationships, the workbook part, and each sheet, shared string, style, theme, and drawing part, through `@eigenpal/ooxml-core` with its limits.
- [x] 2.4 Build the canonical tree per part with typed nodes for the workbook, sheets, rows, cells, shared strings, cell formats, and the elements of design D4, and generic nodes for everything else.
- [x] 2.5 Build the derived indexes: cell reference to cell node, sheet name to part, and shared string and cell format lookups, from the cells that exist only.
- [x] 2.6 Add `WorkbookDocumentStore.transact` with the op contract and the inverse-op undo of design D2, atomic revisions, and the refusal of a transaction that breaks a structural rule of the typed model. Include the store ops that set and clear a cell value, so tests and edited fixture copies can use them.
- [x] 2.7 Serialize normalized parts in schema order, with the counts of design D2, and copy non-XML parts byte for byte.
- [x] 2.8 Add the fidelity oracles: the canonical fingerprint and the save-and-reopen semantic digest, over every fixture.
- [x] 2.9 Add the schema gate (`bun run validate:xsd`) for saved parts, and the security tests for every limit in `AGENTS.md`, "Security".
- [x] 2.10 Inventory every SpreadsheetML part and element that the store preserves without support, and record it as deferred with no support claim.

## 3. Reference PDFs and the save check

- [x] 3.1 Commit the fixture reference PDFs with their `fixtures/references.json` entries and hashes.
- [x] 3.2 Add the save check for saved and edited fixture copies, its committed records, and the CI check of design D5a.
- [x] 3.3 Add the code list of approximations, the public list in `docs/known-approximations.md`, and the test that keeps them equal.

## 4. Values and the cascade

- [x] 4.1 Resolve cell values for every `ST_CellType`, inline strings, and `t="d"`.
- [x] 4.2 Decode and encode `_xHHHH_` escapes in strings.
- [x] 4.3 Expose the stale state of design D3 in the snapshot for a workbook that opens with `fullCalcOnLoad` set on `x:calcPr` or on any `x:sheetCalcPr`.
- [x] 4.4 Resolve cell formats: number format, font, fill, border, alignment, and protection, with out-of-range indexes falling back and recorded.
- [x] 4.5 Format numbers: the built-in formats, custom format codes with a bounded linear parser, sections, conditions, and colors.
- [x] 4.6 Format dates and times in both date systems.
- [x] 4.7 Resolve theme colors, indexed colors, `rgb`, `auto`, and `tint`, and theme fonts. Measure each "Measure" item of the quick references on a fixture, and record the result.

## 5. Grid layout

- [x] 5.0 Copy PPTX Editor's HarfBuzz text measurer with provenance, and add the substitute fonts of `@docx-editor.dev/fonts` (design D9).
- [x] 5.1 Measure the maximum digit width of the default font, and compute column widths and row heights.
- [x] 5.2 Build the virtualized row and column offset index.
- [x] 5.3 Lay out cell text: alignment, wrapping, rotation, indent, overflow into empty neighbors, and numbers that do not fit.
- [x] 5.4 Lay out merged cells, hidden rows and columns, and frozen panes.
- [x] 5.5 Place drawing anchors over the grid, and draw pictures and preset shapes through `@eigenpal/ooxml-core`.

## 6. Paint, harness, and the browser viewer checkpoint

- [x] 6.1 Paint the visible window, and run the paint-target spike of design D7.
- [x] 6.2 Port PPTX Editor's harness over the painter of 6.1: thresholds, supersampled rasters, known gaps that only tighten, and the report. Run it over every fixture of sections 4 and 5.
- [x] 6.3 Add `packages/i18n` with the strings of the viewer chrome.
- [x] 6.4 Add the React adapter: Root, Viewport, and Content, sheet tabs, scrolling, zoom, and cell and range selection with the keyboard and the pointer.
- [x] 6.5 Add a demo that opens a workbook and shows it.
- [x] 6.6 Stop for hands-on feedback. Record and resolve the blocking findings before editing. Moved into 7.8: the viewer and core editing get one hands-on checkpoint there.

## 7. Core editing

- [x] 7.1 Add the bounded A1 reference tokenizer of design D4.
- [x] 7.2 Edit a cell value from the grid: number, text, boolean, and formula text written as `x:f` with its cached value removed, through the store ops, with undo and redo, the shared formula rule of design D4, and the calculation chain dropped as design D5 states.
- [x] 7.3 Apply design D3 on edit: set `fullCalcOnLoad` in the edit transaction, and save it.
- [x] 7.4 Clear cells, and apply basic cell formats: bold, italic, font, size, color, fill, borders, alignment, and number format.
- [x] 7.5 Insert and delete rows and columns with the refusals of design D4.
- [x] 7.6 Add the chrome registry and a toolbar for the commands of this section, with their strings in `packages/i18n`.
- [x] 7.7 Add the support matrix data in `docs/site/data/` with its evidence, and the CI check that ties it to the save-verified records of design D5a.
- [x] 7.7a Add the `edit` and `view` modes of the editor: refusals with `view-mode`, the toolbar state, the adapter `mode` prop, and their tests.
- [x] 7.7b Enforce sheet protection (`x:sheetProtection`): refusals with `locked` in the store, the toolbar state and the protection notice, the snapshot `protection` and `activeCellLocked`, and their tests.
- [ ] 7.8 Stop for hands-on feedback on core editing. Record and resolve the blocking findings. The collaboration change starts only after this checkpoint.

## 8. Adapters and documentation

- [x] 8.1 Add the Vue adapter twin and the parity contract.
- [x] 8.2 Add the docs site pages in the DOCX Editor page structure, over the support matrix and the known approximations.
- [ ] 8.3 Add the release workflow, the API snapshots, and the third-party notices, and publish the first packages.
- [x] 8.4 Add the toolbar part overrides of DOCX Editor: part children that replace or hide a slot where it stands, `preset={false}`, and host groups (`Toolbar.Group`) that the bar measures and collapses with its own groups. The `hidden` list of task 7.6 stays.
