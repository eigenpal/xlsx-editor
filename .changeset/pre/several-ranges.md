---
'@xlsx-editor.dev/core': minor
'@xlsx-editor.dev/i18n': minor
'@xlsx-editor.dev/react': patch
'@xlsx-editor.dev/vue': patch
---

A selection can now hold up to `MAX_SELECTION_RANGES` ranges, from a Control or Command press, the name box, `editor.selectRanges`, or `x:selection/@sqref` on open, with every range in `snapshot.selectionRef` and in the optional `ranges` and `activeRange` fields of `GridSelection` and `HistorySelection`, where a clear or a format applies to every range, a command that acts on one range refuses with `multiple-ranges`, a range past the limit refuses with `too-many-ranges`, and `mergeOps`, `unmergeOp`, and `structureOp` throw a `RangeError` for several ranges.
