---
'@xlsx-editor.dev/core': minor
---

A frozen band now starts at `x:sheetView/@topLeftCell`, so `FrozenPanes` has the new required `firstRow` and `firstCol` fields, and `rows` and `cols` now count the frozen lines from them, not from row 1 and column A.
