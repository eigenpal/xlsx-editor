<p align="center">
  <a href="https://www.xlsx-editor.dev/">
    <picture>
      <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/eigenpal/xlsx-editor/main/.github/assets/readme-logo-dark.svg" />
      <source media="(prefers-color-scheme: light)" srcset="https://raw.githubusercontent.com/eigenpal/xlsx-editor/main/.github/assets/readme-logo-light.svg" />
      <img src="https://raw.githubusercontent.com/eigenpal/xlsx-editor/main/.github/assets/readme-logo-light.svg" alt="XLSX Editor by EigenPal" width="320" height="90" />
    </picture>
  </a>
</p>

# @eigenpal/ooxml-core

The format-neutral OOXML layer of the EigenPal editors. It holds the code between the bytes of an OOXML (ECMA-376) file and a bounded XML tree. It also holds the text units and the DrawingML color and geometry math. It has no SpreadsheetML, PresentationML, or WordprocessingML knowledge.

This package is private. The XLSX Editor engine inlines its code and its types at build time. Package consumers do not install it. Do not import it from an adapter package or from an example.

## What it holds

| Module | Purpose |
| --- | --- |
| `zip.ts` | Bounded zip read (entry count, total size, and ratio caps; refuses overlapping entries and bad limits) and deterministic write. `writeZip` writes in one call; `writeZipSteps` yields between slices of a long entry. Stores already-compressed media, and keeps a saved archive within the ratio cap |
| `zip-write.ts` | The archive layout and the deflate, slice by slice (`writeZipEntriesSteps`, `deflateInSlices`) |
| `opc-names.ts` | OPC part name and relationship target normalization; rejects traversal and encoded separators |
| `xml-reader.ts` | The bounded XML reader: rejects DTDs and entity declarations, caps depth and element counts |
| `xml-fast-read.ts` | A one-pass reader for the plain XML subset. It builds the bounded reader's tree, or gives the part to the bounded reader |
| `xml-text.ts` | Character data for both readers: entity decoding, attribute-value normalization, validity checks |
| `sinks.ts` | Checked XML escaping, `sanitizeHref`, CSS string escaping, inert executable content |
| `limits.ts` | Resource budgets and the hard ceilings no host can raise |
| `sha256.ts` | Synchronous SHA-256 for part and font fingerprints |
| `image-size.ts` | Bounded PNG, JPEG, GIF, and BMP header reads |
| `grapheme.ts` | Grapheme clusters (UAX #29) over `Intl.Segmenter` at the invariant locale, through a replaceable boundary |
| `word-segment.ts` | UAX #29 word segments for word steps, word deletion, and double-click |
| `caret-word.ts` | The caret word rule: the word a format at a collapsed caret reaches (`caretWordRange`) |
| `drawingml-color.ts` | The DrawingML color models and transforms (ECMA-376 Part 1, 20.1.2.3). The caller resolves scheme color names |
| `drawingml-preset-colors.ts` | `PRESET_COLORS`: the `ST_PresetColorVal` names and their sRGB values (Ecma material, see `reference/ecma-376/NOTICE`) |
| `drawingml-geometry.ts` | The DrawingML geometry evaluator (ECMA-376 Part 1, 20.1.9): guide formulas, path commands, the text rectangle, and the caps on guides and path commands |
| `drawingml-preset-geometry.generated.ts`, `drawingml-preset-geometry-types.ts` | `PRESET_SHAPES`: every ECMA-376 preset as guide formulas and paths, generated from `presetShapeDefinitions.xml` by `scripts/geometry/build-presets.ts` (Ecma material, see `reference/ecma-376/NOTICE`), and its types |

## Where the modules come from

Every module is a copy of the same file in PPTX Editor (`eigenpal/pptx-editor`, `packages/ooxml-core/`) at commit `8dc88c4d`. The first line of each file names that source. Below that line, the file is equal to its PPTX Editor source. The exceptions are `index.ts`, whose header names this product, and `drawingml-preset-geometry.generated.ts`, which this repository generates.

PPTX Editor copied most modules from DOCX Editor (`eigenpal/docx-editor`). Their second header line names the DOCX Editor path and commit. It also names the fixes that DOCX Editor does not have yet. The DrawingML color and geometry modules and `image-size.ts` are PPTX Editor's own. For the full history of each module, see the PPTX Editor README of this package.

SpreadsheetML uses DrawingML for drawings, charts, and the theme, so XLSX Editor keeps the DrawingML modules. The theme color math for cells is different. SpreadsheetML cell colors use `tint` in HLS space (ECMA-376 Part 1, 18.8.19). That math lives in the engine, not here.

## Rules

- No imports from `@xlsx-editor.dev/*` and no SpreadsheetML names. If a module needs either, it belongs in `packages/core`.
- DOM-free. `src/tsconfig.json` compiles `src/` without the DOM library.
- A fix here is a security fix for every product. Name the fix in the commit message and in the file header. Send it to the PPTX Editor and DOCX Editor teams.
- Keep each copied module equal to its source. `bun run test:docx-shared` runs DOCX Editor's tests of these modules against this package and reports any drift. It needs a local DOCX Editor checkout, and it reads that checkout without changing it.
- Put a test of one module in `src/__tests__/`. Put a test that needs a workbook, the package model, or a fixture in `packages/core`.

## License

Apache-2.0. See [LICENSE](LICENSE). The preset shape geometry and the preset color names are ECMA-376 tables. The Ecma copyright notice covers them, not Apache-2.0.
