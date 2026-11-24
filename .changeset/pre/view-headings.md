---
'@xlsx-editor.dev/core': minor
'@xlsx-editor.dev/i18n': minor
'@xlsx-editor.dev/react': minor
'@xlsx-editor.dev/vue': minor
---

Add View > Headings, which shows or hides the row and column headings of the shown sheet and writes `x:sheetView/@showRowColHeaders`, so undo, redo, and save keep it. Add the `view.headings` slot, the `Menu.Headings` named row, the `setShowHeadings` op and editor method, `EditorSnapshot.headings`, and the `menu.headings` catalog key.
