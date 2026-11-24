---
'@xlsx-editor.dev/core': minor
'@xlsx-editor.dev/i18n': minor
'@xlsx-editor.dev/react': minor
'@xlsx-editor.dev/vue': minor
---

A macro-enabled workbook (`.xlsm`) or template (`.xltm`) now opens with its macro part kept as inert bytes and refuses structure edits that macro code can see, and a host that names the saved file from `workbookKind` must also read the new `snapshot().macroEnabled` to pick the `.xlsm` or `.xltm` extension.
