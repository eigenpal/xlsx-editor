---
'@xlsx-editor.dev/react': patch
'@xlsx-editor.dev/vue': patch
---

`XlsxEditor.Content` (`XlsxEditorContent` in Vue) now hides the grid while no workbook is ready, so a custom layout shows only the status message when no workbook is open, during a load, and after a failed open.
