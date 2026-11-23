---
'@xlsx-editor.dev/core': minor
'@xlsx-editor.dev/i18n': minor
---

Do the protection checks of a transaction once, not for each cell write. A workbook with a stored sheet name of more than 32,767 UTF-16 code units does not open, and the refusal names the limit `workbook.maxSheetNameLength`.
