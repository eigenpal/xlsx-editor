---
'@xlsx-editor.dev/core': minor
---

`readXlsx`, `readXlsxSteps`, and `openAutomationHost` now also read an `ArrayBuffer`, and an `ArrayBuffer` or a `Uint8Array` of another window. Before, they failed with the reason `zip` and the detail `inflate-error` for a valid file.
