---
'@xlsx-editor.dev/core': minor
'@xlsx-editor.dev/i18n': patch
---

A sheet that the editor adds to a workbook with a chart, a pivot table, a chartsheet, or a data connection can now be renamed and deleted in the session, unless that content can refer to the sheet. Add `EditorSheet.editRefusal`, the reason of each sheet, which `sheetCommandState` reads; a sheet without it takes `sheetEditRefusal`, which now holds the reason of the sheets of the opened file only. The scan of a table, query table, drawing, or legacy drawing part now also reads a sheet name that a CDATA section, a comment, or a line break in an attribute splits, so a structural edit refuses in more cases.
