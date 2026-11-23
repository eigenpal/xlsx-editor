# ECMA-376 reference files

This directory holds copies of files from the ECMA-376 standard (Office Open XML File Formats). The engine implements the standard from these files. They are Ecma International material, not EigenPal code, and the repository's Apache-2.0 license does not apply to them.

## Source

All files come from the free downloads at [ECMA-376 on the Ecma International site](https://ecma-international.org/publications-and-standards/standards/ecma-376/), 5th edition, December 2016.

| Directory | Source archive | Inner archive | Changes |
| --- | --- | --- | --- |
| `ecma-376/part1/drawingml-geometries/` | `ECMA-376-1_5th_edition_december_2016.zip` | `OfficeOpenXML-DrawingMLGeometries.zip` | None. Byte-identical. |
| `ecma-376/part4/transitional-schemas/` | `ECMA-376-4_5th_edition_december_2016.zip` | `OfficeOpenXML-XMLSchema-Transitional.zip` | None. Byte-identical. |
| `ecma-376/part1/predefined-functions.txt` | `ECMA-376-1_5th_edition_december_2016.zip` | Part 1 PDF, table of contents of 18.17.7 | Derived: the 356 function names of 18.17.7.1 to 18.17.7.356, one for each subclause, with a header that names the source. No other text of the standard. |
| `ecma-376/part1/schemas/` | Copied from PPTX Editor, which copied them from DOCX Editor's `reference/ecma-376/part1/schemas/` | Part 1 Strict schemas | None. The seven files are byte-identical to the 5th edition `OfficeOpenXML-XMLSchema-Strict.zip`. DOCX Editor's modified `dml-main.xsd` is not copied: use `part4/transitional-schemas/dml-main.xsd`. |

## Full text of the standard

The PDFs of Parts 1 to 4 are large, so git ignores them. Fetch them, with a text copy and a subclause index for each part:

```bash
bun run reference:fetch
```

The command needs Poppler (`pdftotext`), as the fidelity harness does. It writes, for each part N:

| File | Holds |
| --- | --- |
| `ecma-376/partN/ECMA-376-N.pdf` | The standard, as Ecma publishes it |
| `ecma-376/partN/*.zip` | The supplementary archives: Strict and Transitional schemas, RELAX NG, the DrawingML preset geometries, the SpreadsheetML styles, and the WordprocessingML art borders |
| `ecma-376/partN/text/full.txt` | The text of the PDF, with its layout |
| `ecma-376/partN/text/index.tsv` | One row for each subclause: number, title, PDF page, and line in `full.txt` |

To read a subclause, find its row in the index, and then read `full.txt` from that line:

```bash
grep -P '^18\.3\.1\.13\t' reference/ecma-376/part1/text/index.tsv
sed -n '79030,79110p' reference/ecma-376/part1/text/full.txt
```

The parts are Part 1 (fundamentals and the markup languages, 5th edition, 2016), Part 2 (Open Packaging Conventions, 5th edition, 2021), Part 3 (Markup Compatibility and Extensibility, 5th edition, 2015), and Part 4 (transitional migration features, 5th edition, 2016). Cite these editions.

`bun run reference:check-citations` checks each citation of the form "ECMA-376 Part 1, 18.3.1.13" in the tracked files against the index, and fails on a number that is not a subclause. Add `--list` to print each citation with the title of its subclause. The check needs the local index, so CI does not run it.

## Use

| Files | Used by |
| --- | --- |
| `presetShapeDefinitions.xml` | `scripts/geometry/build-presets.ts` generates `packages/ooxml-core/src/drawingml-preset-geometry.generated.ts` from it. The generated file carries the guide formulas as data. |
| `presetTextWarpDefinitions.xml` | Lookup only. |
| `part1/predefined-functions.txt` | `scripts/formula/build-functions.ts` generates `packages/core/src/store/formula/predefined-functions.generated.ts` from it, and a test keeps the two equal. A typed formula can call only these functions. |
| `part4/transitional-schemas/*.xsd` | Lookup, and the schema gate that the active change adds for saved parts. Never bundled. `sml.xsd` is the SpreadsheetML schema. |
| `part1/schemas/*.xsd` | Lookup only. No gate or build step reads them. |

The published packages contain only the generated data (`drawingml-preset-geometry.generated.ts` and `predefined-functions.generated.ts`), not these files.

## Quick references

`quick-ref/` holds short summaries written for this project, in the DOCX Editor format. They are EigenPal text, not Ecma material, so the Apache-2.0 license applies to them. Read them before you change parsing, cascade, or layout code.

| File | Covers |
| --- | --- |
| [`quick-ref/spreadsheetml.md`](quick-ref/spreadsheetml.md) | Package parts, the workbook, worksheets, cells and values, shared strings, styles, number formats, dates, column widths and row heights, merged cells, and the parts to preserve |
| [`quick-ref/themes-colors.md`](quick-ref/themes-colors.md) | The theme part, SpreadsheetML cell colors (`theme`, `indexed`, `rgb`, `tint`), and DrawingML colors in drawings and charts |

For DrawingML shapes, pictures, and text in drawings, read PPTX Editor's [`reference/quick-ref/drawingml.md`](https://github.com/eigenpal/pptx-editor/blob/main/reference/quick-ref/drawingml.md). Its PresentationML notes do not apply here.

When a quick reference and the schema disagree, the schema wins. Fix the quick reference.

## Terms

Ecma International holds the copyright. Ecma's text copyright policy permits copies of a standard, in whole or in part, and derivative works that implement the standard in a conformant product, if the copies carry the Ecma copyright notice and its Copyright License and Disclaimer, and if the content is not modified. The notice is in [`ecma-376/NOTICE`](ecma-376/NOTICE). The published `core` package carries the same notice in its `THIRD_PARTY_NOTICES.md`, and its `LICENSE` says that the Ecma terms, not Apache-2.0, apply to the ECMA-376 tables. The generated geometry file and the generated function names hold the full notice in a legal comment (`/*!`). Bundlers that keep legal comments, such as esbuild and Bun, keep the notice in the bundles of apps that use `core`. A minified Vite build removes legal comments, unless the app sets `build.rolldownOptions.output.comments.legal` to `true`.

Never edit a copied file in `ecma-376/`. To use a different version, replace the file with an unmodified Ecma copy and record its source in the table above. The derived `part1/predefined-functions.txt` changes only when its source edition changes.

## Known defect in the source

The 5th edition `presetShapeDefinitions.xml` defines `upDownArrow` twice and has no `upArrow`. The geometry evaluator derives `upArrow` as a vertical mirror of `downArrow` (`MIRRORED_PRESETS` in `packages/ooxml-core/src/drawingml-geometry.ts`).
