## Why

The EigenPal editor family has a DOCX editor and a PPTX editor on one architecture, and no XLSX editor. Workbooks are the third common OOXML format. Hosts that embed the family need to open, show, edit, and save them with the same engine shape, security rules, and SDK as the other two products.

One change is needed to establish one typed, tree-backed SpreadsheetML authority and prove it through a read-and-edit slice of the core workbook features before broader features start.

## What Changes

- Establish this change as the only production authority for XLSX Editor.
- Read an XLSX package through the shared bounded OPC and XML layer (`@eigenpal/ooxml-core`) into one canonical ordered tree per part, with typed nodes for the workbook, sheets, rows, cells, shared strings, cell formats, and the elements that structural edits move (design D4), and lossless generic nodes for everything else.
- Make `WorkbookDocumentStore.transact` over typed ops the only write path, and serialize normalized SpreadsheetML from the tree, gated by a structural fingerprint and a save-and-reopen digest.
- Show the values that the file stores. Never evaluate a formula on open. Show cached formula results, and make a stale result visible after an edit that can change it.
- Resolve cell formats, number formats, dates, fonts, fills, borders, alignment, and theme colors through a DOM-free cascade.
- Lay out a virtualized grid, DOM-free: column widths from the maximum digit width, row heights, hidden rows and columns, merged cells, frozen panes, text overflow, and drawing anchors.
- Measure rendering against reference PDFs of synthetic fixtures, with thresholds that only tighten and gaps that never count as passes.
- Make the first implementation checkpoint a browser-visible React viewer with cell selection, then stop for hands-on feedback before editing.
- After the checkpoint, add core cell editing: values, clearing, basic cell formats, and row and column insertion and deletion, then stop for a second hands-on checkpoint. Collaboration is a later change that starts only after core editing works.
- Inventory every other SpreadsheetML feature as preserved and deferred, without a support claim.

## Capabilities

### New Capabilities

- `workbook-canonical-tree`: One bounded package read, one canonical tree per part with typed and generic nodes, atomic typed mutation, and normalized serialization with fidelity oracles.
- `workbook-cell-values`: The value each cell shows: shared and inline strings, numbers through number formats, dates in both date systems, booleans, errors, and cached formula results with stale marking.
- `workbook-cell-formats`: The cascade from a cell to its format: number format, font, fill, border, alignment, and protection, with theme, indexed, and RGB colors, tint, and theme fonts.
- `workbook-grid-layout`: DOM-free, virtualized grid geometry and cell text layout, and the drawing anchors that depend on it.
- `workbook-fidelity-harness`: Synthetic fixtures, reference PDFs, thresholds, known gaps, and the save check.
- `workbook-viewer`: The painted grid in the browser and its chrome: sheet tabs, scrolling, zoom, selection, and the toolbar of the editing commands.
- `workbook-cell-editing`: The first editing slice, its ops, its undo and redo, and its refusals.

### Modified Capabilities

None.

## Impact

The change governs the creation of `packages/core` (`@xlsx-editor.dev/core`) and its lanes, the React and Vue adapters, `packages/i18n`, and the public documentation. It reuses `packages/ooxml-core` without changing its shared modules. Public support claims change only through the feature matrix and its evidence.
