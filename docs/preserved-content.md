# Preserved content

XLSX Editor preserves every part and element of a workbook through open, edit, and save. It models some of them as typed nodes, and keeps the rest as generic nodes or as the bytes of the part. This page lists the content that the engine keeps without support. The engine does not show it, edit it, or act on it, and this page makes no support claim for it. One exception: the grid paints the `cellIs` rules of `conditionalFormatting` (`x:cfRule`, ECMA-376 Part 1, 18.3.1.10) with literal operands, with the differential formats of `dxfs`, as the support matrix row "Conditional formatting" states.

A test (`packages/core/src/store/__tests__/preserved-content.test.ts`) reads the children of the schema types in this list from `sml.xsd` and fails when a child is neither typed by the engine nor in this list.

## Parts kept as bytes

The engine saves these parts byte for byte. A row or column insertion or deletion would change cell positions that some of these parts refer to, so the engine refuses it while the workbook holds a chart or a pivot part, or while the edited sheet relates to a table or a comments part. A sheet rename or delete would change sheet names that some of these parts hold. So the engine refuses it while the workbook holds a chart, a pivot table, or a sheet that is not a worksheet, such as a chartsheet. It also refuses it while the workbook or the package has a relationship to another part that the engine does not rewrite. A sheet delete also refuses a sheet with a related part other than hyperlinks and printer settings. The refusal names the part, or the sheet name of a chartsheet, and other edits stay available.

| Part | Content |
| --- | --- |
| Chartsheet | `x:chartsheet`: a sheet that holds one chart |
| Chart | `c:chartSpace`: DrawingML charts (Part 1, 21.2) |
| Comments and legacy drawing | `x:comments`: cell comments, and their VML shapes |
| Table | `x:table`: a cell range with named columns, an optional header row, and a table style |
| Pivot table and pivot cache | Pivot definitions and their cached data |
| Calculation chain | `x:calcChain`: the order of the last calculation (Part 1, 18.6) |
| External link | `x:externalLink`: cached values of another workbook. Never loaded |
| Connections and query tables | `x:connections` and `x:queryTable`: external data connections. Never loaded |
| Macros, embedded controls, and OLE objects | Active content. Never run |
| Custom XML, printer settings, and document properties | Data for other consumers |
| Media of a part that the engine does not model | Images and binaries |

## Prefixes in attribute values

Markup Compatibility attributes, such as `mc:Ignorable`, name namespace prefixes in their values (ECMA-376 Part 3). An edit that keeps an element keeps these attributes and the namespace declarations of their prefixes. An edit that builds an element again keeps the declarations of that element. A column that a width edit moves out of a later `x:cols` takes the declarations of that list. If the first list declares a prefix of the worksheet again with another URI, the moved column declares the prefix of the worksheet again. If a later `x:cols` has an attribute, the engine refuses the width edit with the code `invalid-sheet`. The engine refuses an edit with the code `invalid-part` when the edit would leave a prefix in one of these attributes undeclared in a row, a cell, another child of the worksheet, a shared string, or a style record, such as a cell format, a font, or a border.

The engine checks the prefixes in `mc:Ignorable`, `mc:MustUnderstand`, `mc:Choice/@Requires`, `mc:ProcessContent`, `mc:PreserveElements`, `mc:PreserveAttributes`, and `xsi:type`. Each prefix must be declared in scope. If one of these parts names an undeclared prefix, the engine refuses to open the workbook:

- The content types part
- The package relationships
- The workbook part and its relationships
- A worksheet and its relationships

If another part that the engine reads names an undeclared prefix, the engine keeps that part as bytes and does not use it. These parts are the shared string table, the styles part, the theme, a drawing, and the relationships of every other part. If the shared string table is kept as bytes, a cell that refers to a shared string shows no text, and an edit that writes a string is refused. If the styles part is kept as bytes, every cell shows the default format.

## Formulas kept without edits

A formula (`x:f`, `CT_CellFormula`) keeps its text and its cached value (`x:v`). The engine never evaluates it. The formula type `x:f/@t` (`ST_CellFormulaType`, Part 1, 18.18.6) is `normal`, `shared`, `array`, or `dataTable`. An edit to the master cell of a shared formula writes the formula of each other cell of the shared formula into that cell first, so those cells keep their formulas. An edit to the master cell of an array or data table formula is refused, and so is an edit to a cell inside an array or data table range. A refusal keeps the formula definition that other cells use.

A row or column insertion or deletion moves the references of every formula. It refuses the edit, and names the cell, while the workbook holds formula text that the engine cannot read, or a reference that it cannot move with certainty, such as a sheet range that spans the edited sheet. An array formula range moves as a whole, and an edit that splits it is refused.

## Elements that refuse a row or column edit

On the edited sheet, every worksheet element that this page lists as generic, and every relationship to a part other than a hyperlink, a drawing, an image, a printer setting, or a header and footer drawing, refuses a row or column insertion or deletion. The refusal names the element and its part. On every other sheet, and in the workbook part, such an element refuses the edit only when its text names the edited sheet.

## Sheet protection

The editor keeps `x:sheetProtection` (`CT_SheetProtection`, ECMA-376 Part 1, 18.3.1.85) as the file holds it, and enforces it. When its `sheet` attribute is true, the editor refuses an edit with the reason `locked`:

- A value edit on a locked cell. A cell is locked unless its cell format sets `x:protection/@locked` (`CT_CellProtection`, Part 1, 18.8.33) to false. A cell format with no `x:protection` and no `applyProtection` takes the protection of its cell style in `x:cellStyleXfs`, by `xfId`. A cell that does not exist takes the format of its row or column.
- A format edit, unless `formatCells` is false. A paste that gives a cell another cell format is a format edit. A paste that keeps the cell format of each target cell is not.
- A new cell format that changes the `locked` or `hidden` protection of a cell, whatever `formatCells` states. The lock rule is `cell`. So a paste or a script cannot unlock a locked cell and then write it.
- A cut whose paste could not clear the source: a locked cell that holds content, or a sheet where `formatCells` is not false. The editor refuses the cut itself, not its paste.
- A column width, unless `formatColumns` is false, and a row height, unless `formatRows` is false. A heading edge whose size is locked shows no resize edge.
- An insertion or a deletion of rows or columns, unless `insertRows`, `deleteRows`, `insertColumns`, or `deleteColumns` is false. When `deleteRows` or `deleteColumns` is false, a deletion also deletes locked cells.
- A merge or an unmerge of cells, whatever the protection attributes state. The lock rule is `merge`.

On a protected sheet, a clear of formats keeps the `locked` and `hidden` protection of each cell. The cell gets the default cell format with its own protection. So the paste of a cut, which clears the source cells that it does not write, never changes their protection.

The `sheet` attribute is false by default. Each action attribute in this list (`formatCells`, `formatColumns`, `formatRows`, `insertColumns`, `insertRows`, `deleteColumns`, and `deleteRows`) is true by default, and true locks its action (`CT_SheetProtection` in `sml.xsd`). The editor never reads or checks the password hash, and it keeps the hash as it is. It does not enforce `protectedRanges` (`CT_ProtectedRanges`).

## Workbook protection

The editor keeps `x:workbookProtection` (`CT_WorkbookProtection`, ECMA-376 Part 1, 18.2.29) as the file holds it. When its `lockStructure` attribute is true, the editor refuses a sheet add, rename, delete, or move with the reason `locked`. The sheet tabs then disable those actions. The `lockStructure` attribute is false by default (`sml.xsd`). The editor does not enforce `lockWindows` or `lockRevision`, and it keeps the password hashes as they are.

## Workbook elements kept as generic nodes

`x:workbook` (`CT_Workbook`) children with no typed support (Part 1, 18.2.27):

- `fileVersion`
- `fileSharing`
- `workbookProtection`
- `functionGroups`
- `externalReferences`
- `oleSize`
- `customWorkbookViews`
- `pivotCaches`
- `smartTagPr`
- `smartTagTypes`
- `webPublishing`
- `fileRecoveryPr`
- `webPublishObjects`
- `extLst`

## Worksheet elements kept as generic nodes

`x:worksheet` (`CT_Worksheet`) children with no typed support (Part 1, 18.3.1.99):

- `sheetProtection`
- `protectedRanges`
- `scenarios`
- `autoFilter`
- `sortState`
- `dataConsolidate`
- `customSheetViews`
- `phoneticPr`
- `conditionalFormatting`
- `dataValidations`
- `headerFooter`
- `rowBreaks`
- `colBreaks`
- `customProperties`
- `cellWatches`
- `ignoredErrors`
- `smartTags`
- `legacyDrawing`
- `legacyDrawingHF`
- `drawingHF`
- `picture`
- `oleObjects`
- `controls`
- `webPublishItems`
- `tableParts`
- `extLst`

`x:sheetView` (`CT_SheetView`) children with no typed support:

- `pivotSelection`
- `extLst`

`x:row` (`CT_Row`) and `x:c` (`CT_Cell`) children with no typed support:

- `extLst`

`x:c` (`CT_Cell`) children in other namespaces and markup compatibility content:

- A value edit, a clear, or a paste of a cell replaces `x:f`, `x:v`, `x:is`, and an `mc:AlternateContent` child that holds only `x:v` and `x:is`.
- Every other child of `x:c` stays in its place. This includes an element in another namespace, an empty `mc:AlternateContent`, and content that the schema does not allow.
- A cell can hold another form of its value: other markup compatibility content, such as a formula in `mc:AlternateContent` or a branch with no value, or an element in another namespace with `x:f`, `x:v`, or `x:is` in it. An edit or a paste of that cell is refused with `cell-alternate-content`. So is a clear of it, or the clear of a merge with `dropValues`, while it holds a value. The message of a clear, a paste, or a merge names that command. A cut of that cell is refused with `copy-wrapped-formula` when its formula is inside `mc:AlternateContent`, as its copy is. Otherwise a cut gives the reason of the clear of its source, such as `cell-alternate-content`. The editor opens no cell editor on it. Clear contents and Clear all are disabled while the selection holds such a cell with a value.
- A clear or a cut reads only a cell that holds a value in some form. It leaves a cell with no value as it is and does not refuse it.
- A cut is refused with `cut-wrapped-value` for a cell whose only value is in an `mc:AlternateContent` child, because the copy cannot carry that value. A clear removes such a value, and a merge asks about it as about any other value.

## Style elements kept as generic nodes

`x:styleSheet` (`CT_Stylesheet`) children with no typed support:

- `dxfs`
- `tableStyles`
- `extLst`

## Shared string elements kept as generic nodes

`x:sst` (`CT_Sst`) and `x:si` (`CT_Rst`) children with no typed support:

- `rPh`
- `phoneticPr`
- `extLst`
