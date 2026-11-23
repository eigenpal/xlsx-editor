---
'@xlsx-editor.dev/core': minor
'@xlsx-editor.dev/react': minor
'@xlsx-editor.dev/vue': minor
---

Add the named toolbar parts (`Toolbar.Bold` and one for each slot), `icon`, `asChild`, `className`, and `overflowContent` on a slot, `Toolbar.Action`, `Toolbar.Separator`, and `priority`, `pinned`, and `overflowContent` on `Toolbar.Group`; behavior change: a host group with no `priority` now leaves the bar for the "More" panel before every built-in group, also when `after` places it at the start of the bar.
