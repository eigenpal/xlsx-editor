# SpreadsheetML quick reference

A summary of the SpreadsheetML parts and rules that the engine reads, written for this project in the DOCX Editor format. ECMA-376 Part 1, clause 18, defines SpreadsheetML. The Transitional schema is `reference/ecma-376/part4/transitional-schemas/sml.xsd`. When this page and the schema disagree, the schema wins.

Each rule names its source. "Schema" means `sml.xsd`. "Part 1, 18.x" means a clause of ECMA-376 Part 1. The Part 1 PDF is a free download from Ecma (see `reference/README.md`). Add a clause number here only after you check it in that PDF. "Measure" means that the standard does not settle the behavior: confirm it on a fixture and its reference PDF before you rely on it, and record the result here.

## Package parts

| Part | Typical name | Content | Relationship from |
| --- | --- | --- | --- |
| Workbook | `/xl/workbook.xml` | `x:workbook`: sheets, defined names, calculation properties, views | Package (`officeDocument`) |
| Worksheet | `/xl/worksheets/sheetN.xml` | `x:worksheet`: one grid | Workbook |
| Chartsheet | `/xl/chartsheets/sheetN.xml` | `x:chartsheet`: a sheet that holds one chart | Workbook |
| Shared strings | `/xl/sharedStrings.xml` | `x:sst`: the shared string table | Workbook |
| Styles | `/xl/styles.xml` | `x:styleSheet`: number formats, fonts, fills, borders, cell formats | Workbook |
| Theme | `/xl/theme/theme1.xml` | `a:theme` (DrawingML) | Workbook |
| Calculation chain | `/xl/calcChain.xml` | `x:calcChain`: the order of the last calculation | Workbook |
| Drawing | `/xl/drawings/drawingN.xml` | `xdr:wsDr`: anchored shapes, pictures, and chart frames | Worksheet (`x:drawing`) |
| Chart | `/xl/charts/chartN.xml` | `c:chartSpace` (DrawingML Charts, Part 1, 21.2) | Drawing |
| Comments | `/xl/commentsN.xml` | `x:comments`: authors and comment text | Worksheet |
| Legacy drawing | `/xl/drawings/vmlDrawingN.vml` | VML shapes, for example the comment boxes | Worksheet (`x:legacyDrawing`) |
| Table | `/xl/tables/tableN.xml` | `x:table`: a named range with columns, a header row, and a style | Worksheet (`x:tableParts`) |
| Pivot table and cache | `/xl/pivotTables/`, `/xl/pivotCache/` | Pivot definitions and cached source data | Worksheet, workbook |
| External link | `/xl/externalLinks/externalLinkN.xml` | Cached values of another workbook | Workbook |
| Connections | `/xl/connections.xml` | External data connections | Workbook |

Part names are not fixed. Always follow the relationships, never a path pattern.

The main part content type is `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml`, and `application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml` for a template. A macro-enabled workbook or template uses another main-part content type that ECMA-376 does not define, with the same `x:workbook` root; the engine opens both and keeps the macro part as inert bytes. Treat any other main-part content type, such as an add-in, as unsupported, and refuse to edit rather than guess.

The main namespace is `http://schemas.openxmlformats.org/spreadsheetml/2006/main`. Files often declare extension namespaces with `mc:Ignorable` and carry extension data in `x:extLst`. Preserve both unchanged.

## Workbook

Children of `x:workbook`, in schema order: `fileVersion`, `fileSharing`, `workbookPr`, `workbookProtection`, `bookViews`, `sheets`, `functionGroups`, `externalReferences`, `definedNames`, `calcPr`, `oleSize`, `customWorkbookViews`, `pivotCaches`, `smartTagPr`, `smartTagTypes`, `webPublishing`, `fileRecoveryPr`, `webPublishObjects`, `extLst` (schema, `CT_Workbook`, Part 1, 18.2.27).

- `x:sheets/x:sheet` (`CT_Sheet`): `name`, `sheetId`, `state` (`ST_SheetState`: `visible` by default, `hidden`, `veryHidden`), and `r:id` to the sheet part (schema). `name`, `sheetId`, and `r:id` are required. The order of `x:sheet` elements is the tab order.
- `x:workbookPr/@date1904` (`CT_WorkbookPr`): `false` by default (schema). It selects the date system; see "Dates and times".
- `x:workbookProtection` (`CT_WorkbookProtection`, Part 1, 18.2.29): `lockStructure`, `lockWindows`, and `lockRevision`, each `false` by default (schema), and the password hash attributes.
- `x:bookViews/x:workbookView/@activeTab` and `@firstSheet` (`CT_BookView`, each `0` by default): the selected and first visible tab.
- `x:definedNames/x:definedName` (`CT_DefinedName`): a name (`@name`, required) and its formula text. `@localSheetId` scopes a name to one sheet by its index in `x:sheets`. Built-in names start with `_xlnm.`, for example `_xlnm.Print_Area`, `_xlnm.Print_Titles`, and `_xlnm._FilterDatabase`.
- `x:calcPr` (`CT_CalcPr`): `calcMode` (`ST_CalcMode`, `auto` by default), `fullCalcOnLoad` (`false` by default), `calcId`, and the iteration settings (schema). `fullCalcOnLoad="1"` asks a consumer to recalculate every formula when it opens the file.

## Worksheet

Children of `x:worksheet`, in schema order (`CT_Worksheet`, Part 1, 18.3.1.99):

`sheetPr`, `dimension`, `sheetViews`, `sheetFormatPr`, `cols`, `sheetData`, `sheetCalcPr`, `sheetProtection`, `protectedRanges`, `scenarios`, `autoFilter`, `sortState`, `dataConsolidate`, `customSheetViews`, `mergeCells`, `phoneticPr`, `conditionalFormatting`, `dataValidations`, `hyperlinks`, `printOptions`, `pageMargins`, `pageSetup`, `headerFooter`, `rowBreaks`, `colBreaks`, `customProperties`, `cellWatches`, `ignoredErrors`, `smartTags`, `drawing`, `legacyDrawing`, `legacyDrawingHF`, `drawingHF`, `picture`, `oleObjects`, `controls`, `webPublishItems`, `tableParts`, `extLst`.

A writer that inserts an element must put it in this order. An element out of order makes the part invalid.

- `x:dimension/@ref` (Part 1, 18.3.1.35): the used range, for example `A1:H40`. It is a hint, often wrong, and attacker-controlled. Never size anything from it.
- `x:sheetViews/x:sheetView` (`CT_SheetView`): `showGridLines`, `showRowColHeaders`, and `showZeros` (`true` by default), `rightToLeft` and `tabSelected` (`false` by default), `zoomScale` (`100` by default), `topLeftCell`, and the required `workbookViewId` (schema). `x:pane` (`CT_Pane`) sets frozen or split panes: `xSplit`, `ySplit` (`0` by default), `topLeftCell`, `activePane` (`ST_Pane`, `topLeft` by default), and `state` (`ST_PaneState`: `split` by default, `frozen`, `frozenSplit`). `x:selection/@pane` names a pane too (Part 1, 18.3.1.66 and 18.18.52). For a frozen pane, `xSplit` and `ySplit` count columns and rows; for a split pane, they are distances in twentieths of a point.
- `x:sheetFormatPr` (`CT_SheetFormatPr`, Part 1, 18.3.1.81): `baseColWidth` (`8` by default), `defaultColWidth`, `defaultRowHeight` (required), `customHeight`, `zeroHeight`, `thickTop`, `thickBottom`, `outlineLevelRow`, `outlineLevelCol` (schema).
- `x:cols/x:col` (`CT_Col`, Part 1, 18.3.1.13; `x:cols`, 18.3.1.17): one entry for the columns `min` to `max` (1-based, inclusive, both required): `width`, `style` (a cell format index, `0` by default), `hidden`, `customWidth`, `bestFit`, `phonetic`, `outlineLevel`, `collapsed` (schema). A worksheet can hold more than one `x:cols`.
- `x:mergeCells/x:mergeCell/@ref` (`CT_MergeCells` and `CT_MergeCell`, Part 1, 18.3.1.55 and 18.3.1.54): a merged range. `@ref` is required. The top-left cell holds the value and the format that show. Other cells in the range can hold values; preserve them.
- `x:hyperlinks/x:hyperlink` (`CT_Hyperlink`): `ref` (required), `r:id` to an external target, or `location` for a place in the workbook, `tooltip`, and `display`. Every target goes through `sanitizeHref`.
- `x:sheetProtection` (`CT_SheetProtection`, Part 1, 18.3.1.85): `sheet` turns the protection on, `false` by default. `objects`, `scenarios`, `selectLockedCells`, and `selectUnlockedCells` are `false` by default. Every other action attribute, such as `formatCells`, `insertRows`, or `deleteColumns`, is `true` by default, and `true` locks the action (schema). The engine never checks the password hash attributes.
- `x:conditionalFormatting` and `x:cfRule` (Part 1, 18.3.1.18 and 18.3.1.10): rules over the ranges of `@sqref`. A rule names a differential format by `@dxfId`. Its type is an `ST_CfType` value, such as `cellIs`. `@operator` takes an `ST_ConditionalFormattingOperator` value (Part 1, 18.18.15).

## Cells and values

`x:sheetData` holds `x:row` elements, and each row holds `x:c` cells. The sheet is sparse: a row or cell that is not in the file is empty.

- `x:row` (`CT_Row`, Part 1, 18.3.1.73): `r` (1-based row number), `spans` (a hint), `s` and `customFormat` (a row format), `ht` (height in points) and `customHeight`, `hidden`, `outlineLevel`, `collapsed`, `thickTop`, `thickBot`, and `ph` (schema). Never trust `r` or `spans` to size a loop or an allocation.
- `x:c` (`CT_Cell`, Part 1, 18.3.1.4): `r` (the cell reference, for example `C7`), `s` (the index into `x:cellXfs`, `0` by default), `t` (the cell type, `n` by default), `cm` and `vm` (cell and value metadata indexes), and `ph` (show phonetic text) (schema). Preserve `cm` and `vm` with their metadata part.
- `x:c` children: `x:f` (formula), `x:v` (value), `x:is` (inline string), `x:extLst`.

| `t` | `x:v` holds | Shows |
| --- | --- | --- |
| `n` | A number in `xsd:double` form | The number through its number format |
| `s` | A 0-based index into the shared string table | That string |
| `inlineStr` | Nothing: the text is in `x:is` | The inline string |
| `str` | A formula's string result | The string |
| `b` | `0` or `1` | `FALSE` or `TRUE` |
| `e` | An error value: `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A`. Show and preserve any other error text as it is | The error |

These are the `ST_CellType` values of the Transitional schema (Part 1, 18.18.11). Files can also hold `t="d"` with an ISO 8601 date in `x:v`. That value is not in the Transitional schema: read it, preserve it, and measure how it shows.

The column and row of a cell come from `@r` when it is present. A cell or row without `@r` follows the one before it. Two cells with the same reference, or cells out of order, are invalid input: refuse or demote, never guess.

Grid references use column letters `A` to `XFD` and rows `1` to `1048576` in files in common use. The schema does not bound them. Treat a reference past that grid as invalid input at the parse boundary.

## Formulas

- `x:f` (`CT_CellFormula`) holds the formula text without a leading `=`. `@t` (`ST_CellFormulaType`, Part 1, 18.18.6) is `normal` by default, or `shared`, `array`, or `dataTable` (schema).
- A shared formula: the first cell has `t="shared"`, `ref` (the range it covers), `si` (the group index), and the text. The other cells have `t="shared"` and `si` only. Their formula is the first cell's text with relative references moved by the offset.
- An array formula has `t="array"` and `ref`; the text is on the first cell of the range.
- `x:v` beside `x:f` is the cached result of the last calculation. The engine shows that cached value. It never evaluates a formula on open.
- `@ca="1"` marks a formula that a consumer always recalculates.
- `/xl/calcChain.xml` is optional (Part 1, 18.6). It records the order of the last calculation. A save that adds, removes, or moves formulas makes it stale. Measure: confirm on the save check that a stale or removed chain gives no repair prompt before you choose how to save it.

## Shared strings

`x:sst` holds `x:si` items, each a rich text string (`CT_Rst`): either one `x:t`, or `x:r` runs with `x:rPr` and `x:t`, plus optional phonetic runs (`x:rPh`) and `x:phoneticPr` (schema). `@count` and `@uniqueCount` are hints. Never size anything from them.

- Preserve `xml:space="preserve"` on `x:t`, or leading and trailing spaces are lost.
- SpreadsheetML string values (`ST_Xstring`, in `shared-commonSimpleTypes.xsd`) escape characters that XML cannot hold as `_xHHHH_`, for example `_x000D_` for CR. A literal `_x` sequence is itself escaped as `_x005F_`. Decode on read and encode on write.
- Run properties (`x:rPr`, `CT_RPrElt`) use font elements: `x:rFont`, `x:charset`, `x:family`, `x:b`, `x:i`, `x:strike`, `x:outline`, `x:shadow`, `x:condense`, `x:extend`, `x:color`, `x:sz`, `x:u`, `x:vertAlign`, `x:scheme` (schema). A cell font (`x:font`, `CT_Font`) has the same children, with `x:name` in place of `x:rFont`.

## Styles

Children of `x:styleSheet`, in schema order: `numFmts`, `fonts`, `fills`, `borders`, `cellStyleXfs`, `cellXfs`, `cellStyles`, `dxfs`, `tableStyles`, `colors`, `extLst` (schema, `CT_Stylesheet`).

- A cell's `@s` is an index into `x:cellXfs` (Part 1, 18.8.10). Each `x:xf` there (Part 1, 18.8.45) names `numFmtId`, `fontId`, `fillId`, `borderId`, and `xfId`, and holds `x:alignment` (`CT_CellAlignment`) and `x:protection` (`CT_CellProtection`, Part 1, 18.8.33) (schema, `CT_Xf`).
- `xfId` indexes `x:cellStyleXfs` (Part 1, 18.8.9), the formats of the named cell styles in `x:cellStyles` (Part 1, 18.8.8). The `apply*` attributes (`applyNumberFormat`, `applyFont`, `applyFill`, `applyBorder`, `applyAlignment`, and `applyProtection`, with no schema default) record whether the cell format overrides its style. Measured on fixture `colors-fonts` (`fixtures/measured.json`): a cell format with `fontId` of a bold red font and no `applyFont` shows the bold red font. The cascade uses the ids of the cell format itself.
- `x:fonts`: the default font is the font of the Normal cell style (`x:cellStyle` with `builtinId="0"`, then its `xfId` in `x:cellStyleXfs`, then that format's `fontId`). It is usually the first font. It sets the maximum digit width that column widths use.
- `x:fills`: writers keep two reserved fills first, `patternType="none"` and `patternType="gray125"`. Do not reorder them.
- `x:dxfs`: differential formats (`x:dxf`, Part 1, 18.8.14), used by conditional formatting and table styles. They hold only the properties they change.
- `x:alignment`: `horizontal` (`ST_HorizontalAlignment`, Part 1, 18.18.40: `general`, `left`, `center`, `right`, `fill`, `justify`, `centerContinuous`, `distributed`), `vertical` (`ST_VerticalAlignment`, Part 1, 18.18.88, `bottom` by default), `wrapText`, `shrinkToFit`, `indent`, `relativeIndent`, `justifyLastLine`, `textRotation` (`ST_TextRotation`: 0 to 180, or `255` for vertical text), `readingOrder` (schema).

Every index into `x:cellXfs`, `x:fonts`, `x:fills`, `x:borders`, or `x:numFmts` is attacker-controlled. An index out of range is invalid input: fall back to the default format and record it, never read past the list.

## Number formats

`x:numFmts/x:numFmt` (`CT_NumFmts` and `CT_NumFmt`, Part 1, 18.8.31 and 18.8.30) defines custom formats: `numFmtId` and `formatCode`, both required (schema). Custom IDs start at 164. IDs below 164 are built in. A file can omit them from `x:numFmts`. Measure whether a built-in ID that the file redefines keeps the built-in code; no fixture shows it yet, and the engine uses the file's code. Part 1, 18.8.30, lists the built-in formats:

| ID | Format code | ID | Format code |
| --- | --- | --- | --- |
| 0 | `General` | 18 | `h:mm AM/PM` |
| 1 | `0` | 19 | `h:mm:ss AM/PM` |
| 2 | `0.00` | 20 | `h:mm` |
| 3 | `#,##0` | 21 | `h:mm:ss` |
| 4 | `#,##0.00` | 22 | `m/d/yy h:mm` |
| 9 | `0%` | 37 | `#,##0 ;(#,##0)` |
| 10 | `0.00%` | 38 | `#,##0 ;[Red](#,##0)` |
| 11 | `0.00E+00` | 39 | `#,##0.00;(#,##0.00)` |
| 12 | `# ?/?` | 40 | `#,##0.00;[Red](#,##0.00)` |
| 13 | `# ??/??` | 45 | `mm:ss` |
| 14 | `mm-dd-yy` | 46 | `[h]:mm:ss` |
| 15 | `d-mmm-yy` | 47 | `mmss.0` |
| 16 | `d-mmm` | 48 | `##0.0E+0` |
| 17 | `mmm-yy` | 49 | `@` |

The printed table of Part 1 gives ID 47 as `mmss.0`. Fixture `number-formats` shows the minutes and seconds separated by `:`, so the engine uses `mm:ss.0` (`packages/core/src/store/number-format/builtins.ts`).

The other IDs below 164 (for example 5 to 8, 23 to 36, 41 to 44, and 50 to 163) depend on the locale. Treat them as unknown until a fixture and its reference PDF show them.

Some IDs in this table also follow the locale of the environment that shows the workbook. Measured on fixture `number-formats` (`fixtures/measured.json`), in an environment with decimal `,` and group U+0020: 14 shows `dd/mm/yyyy`, 15 shows `dd-mmm-yy`, 16 shows `dd-mmm`, 20 shows `hh:mm`, 21 shows `hh:mm:ss`, 22 shows `dd/mm/yyyy hh:mm`, and 37 to 40 show a minus sign instead of parentheses (`#,##0;-#,##0`), with no trailing space on positive values (fixtures `number-formats` and `number-formats-more`). The other IDs of the table show their ECMA-376 code. A renderer therefore needs a locale profile, and a reference PDF records the locale it printed with (`fixtures/references.json`).

A format code has up to four sections separated by `;`: positive, negative, zero, and text. A section can start with a condition (`[>=100]`) and a color (`[Red]`, `[Color10]`). Format codes are attacker-controlled: parse them with a bounded, linear parser, never with a backtracking regex, and cap their length.

## Dates and times

A date is a number formatted with a date format. The integer part counts days, and the fraction is the time of day.

The date system and the serial numbers are in Part 1, 18.17.4.1.

- 1900 date system (`date1904="0"`): serial 1 is 1900-01-01. The system counts 1900 as a leap year, so serial 60 is the date 1900-02-29, which does not exist, and from serial 61 (1900-03-01) the date is one day earlier than a plain count of days from 1900-01-01 gives.
- 1904 date system (`date1904="1"`): serial 0 is 1904-01-01.
- A negative or very large serial is not a date. Measured on fixture `number-formats-more`: a date format and an elapsed format such as `[h]:mm` show a cell filled with `#` for serial -1 and for serial 3000000.

## Column widths and row heights

Part 1, 18.3.1.13 (`col`), defines column width in characters of the maximum digit width of the default font:

```text
width = Truncate([{number of characters} * {maximum digit width} + {5 pixel padding}]
        / {maximum digit width} * 256) / 256

pixels = Truncate(((256 * {width} + Truncate(128 / {maximum digit width})) / 256)
         * {maximum digit width})
```

The maximum digit width is the widest of the digits 0 to 9 in the default font (the font of the Normal cell style), in pixels. The standard does not state the pixel density. The formulas are in pixels, so measure the digits at the font size and round as the standard does.

Measured on the reference PDFs of `basic-values`, `colors-fonts`, `number-formats`, `dates-1904`, `preserved-parts`, and `grid-structure`, except where an item names the screen profile. No committed reference shows the screen profile; the approximation "Screen geometry" names it. The engine implements these in `packages/core/src/layout/metrics.ts` and `grid-geometry.ts`:

- The grid has 72 pixels for each inch, so 1 pixel is 1 point: the reference PDFs print each column at the width that the formulas give at that density. The screen profile uses the same density.
- The maximum digit width is the widest digit advance at the font size, rounded: Carlito 11 is 6 pixels (5.58 before rounding), Carlito 20 is 10.
- Column widths 1, 2, 5, 8.43, 10, 16, 30, and 50 are 6, 12, 30, 51, 60, 96, 180, and 300 points at a maximum digit width of 6 on screen. The reference PDFs print the same widths where they have them (16 is 96 points, 18 is 108, 24 is 144, 3 is 18).
- A default column of `baseColWidth` 8 is 53 points, and of 10 is 65 points, at a maximum digit width of 6. Carlito 20 gives 87 points. The padding is 5 pixels at a maximum digit width of 6 and 7 at 10. The engine uses `2 * ceil(MDW / 4) + 1` pixels, a formula fitted to these two measured points; other widths have no measurement.
- The line height of a font is its ascent plus its descent (the OS/2 win values), each rounded to whole pixels: Carlito 11 is 10 + 3 = 13 pixels, Carlito 20 is 19 + 5 = 24.
- A row without `@ht` is the line height of its font plus 2 pixels on screen: 15 points for Carlito 11, 26 for Carlito 20 (the screen profile, and the approximation "Screen geometry"). It takes the font height even when `defaultRowHeight` says otherwise, unless `x:sheetFormatPr/@customHeight` is set.
- The reference PDFs print every row at `floor(height * 14 / 15)` whole points, where the height is `@ht` or the screen height of the row's font: `ht` 10, 20, and 33.75 print at 9, 18, and 31 points, and a default Carlito 11 row at 14. Columns print at their screen widths. The fidelity harness uses this print profile, because the reference PDFs are the ground truth. Every caller of layout names its profile.
- A row with `@ht` keeps whole pixels, rounded down: `ht` 10, 20, and 33.75 are 10, 20, and 33 points on screen (the screen profile).
- On the printed page, the grid starts at the margins (36 points for 0.5 inch margins), for cells and drawings alike. Each cell owns the gridline pixel at its left and top edges, also when gridlines do not print: the clip of cell `B8:D10` starts 1 point right of and below the cell.

- A column without an `x:col` entry uses `x:sheetFormatPr/@defaultColWidth`. Without that attribute, the width comes from `baseColWidth` (`8` characters by default) plus the padding of the default column item.
- Hidden columns and rows have no width or height in the grid.

## Text in cells

Measure these rendering rules on fixtures. The standard does not define them.

- Text that is wider than its cell can show across empty cells to its right (or left, for right-aligned text). It stops at the first cell that holds a value or a formula. The reference PDF of `formula-overflow` shows the stop at a formula with a cached number and at a `t="str"` formula whose cached value is empty. The line scores of the fidelity harness check both stops by the clip edge of each overflowing line (`fixtureClipped` in `scripts/fidelity/thresholds.ts`), and unit tests of layout guard them too. Both of those stops come from cached values: `B1` caches 2, and the empty `<v></v>` of `B3` is an empty string value. A formula with no cached value also stops it, and no reference PDF shows that case, so that rule rests on the unit tests alone, as the known approximations say.
- Wrapped text (`wrapText="1"`) breaks lines inside the cell and does not overflow.
- `horizontal="general"` aligns text left, numbers right, and booleans and errors center.
- A number that does not fit its column can show as `#` characters instead of its digits.

Measured on the reference PDFs of `basic-values`, `number-formats`, `grid-structure`, and `text-layout`, and implemented in `packages/core/src/layout/cell-text.ts` and `number-fit.ts`:

- Every glyph advance rounds to whole pixels on its own: each glyph of the PDFs starts on a whole point, and `Shared text` in Carlito 11 advances 5, 6, 5, 4, 5, 6, 2, 4, 5, 5, and 4 points.
- Text keeps a margin of `floor(MDW / 3)` pixels of its font on each side, inside the left gridline pixel that the cell owns: Carlito 11 starts 3 points inside its cell and ends 2 points inside it. Carlito 8, and text shrunk to 6 points, start 2 points inside; Carlito 20 starts 4 points inside.
- Centered text starts at `round((free - width) / 2)` inside the text area, half up: `#N/A` (22 points) in a 43 point area starts 11 points in.
- An indent level is one maximum digit width of the default font (6 points for Carlito 11).
- Bottom-aligned text has the descent of its font, in whole pixels, above the bottom of the cell. Top-aligned text has its ascent below the top gridline pixel.
- Wrapped lines are one printed row height of their font apart (14 points for Carlito 11). Centered text centers its lines with those heights, half up.
- `fill` repeats the text as often as it fits whole. `justify` spreads the spaces of each line but the last. `distributed` spreads every line and centers a line without a space. `centerContinuous` centers across the empty neighbor cells that share the alignment. `shrinkToFit` scales the font down to whole points (11 to 6).
- A row without `@ht` takes the tallest row of the fonts of its cells and their rich text runs, never less than the default row: Carlito 20 prints at 24 points, Carlito 8 at 14, and a row with a 14 point run at 17.
- A General number shows the longest rounding narrower than the text area: 1234567.891 shows `1234568` in 48 points, and `1.2E+11` and `98765.4` show in 43 points where `1.23E+11` and `98765.43` (43 and 45 points) do not. A rounding that reaches zero shows `0`.
- A formatted number that does not fit, and a General number that fits in no form, fill the text area with `#`: five `#` in 25 points.
- Measured, and different in layout by a fixed amount (the approximation "Measured text offsets"): text rotated by 90 degrees starts its baseline 53 points left of and 2 points below the PDF; text rotated clockwise by 45 degrees starts 27.8 points left and 0.8 points low; stacked text sits 2 points low and 1 point right; Carlito 8 text sits 1 point low; and a line centered in a merged range sits 1 point low.
- Measure: overflow that reaches a cell with a constant value, overflow that reaches a formula with no cached value, and a row with wrapped text and no `@ht`.

## Drawings and anchors

`xdr:wsDr` (`CT_Drawing`, `dml-spreadsheetDrawing.xsd`) holds anchors. Each anchor places one shape, picture, group, or chart frame (Part 1, 20.5).

- `xdr:twoCellAnchor`: `xdr:from` and `xdr:to`, each a cell (`xdr:col`, `xdr:row`, 0-based) and an offset in EMU (`xdr:colOff`, `xdr:rowOff`). `@editAs` (`ST_EditAs`: `twoCell` by default, `oneCell`, `absolute`) says how the object moves and resizes with cells.
- `xdr:oneCellAnchor`: `xdr:from` and an extent (`xdr:ext`).
- `xdr:absoluteAnchor`: a position (`xdr:pos`) and an extent, in EMU.

The anchor position depends on the column widths and row heights, so layout places drawings after the grid. The shapes themselves are DrawingML; `packages/ooxml-core` evaluates their geometry and colors.

## Preserve, do not run

Preserve these parts and elements unchanged, and never act on them: macros (`vbaProject.bin`), embedded controls (`x:controls`), OLE objects (`x:oleObjects`), external links, data connections, query tables, web publishing settings, printer settings, custom XML, and extension data in `x:extLst`. Never load an external target without an explicit user action.
