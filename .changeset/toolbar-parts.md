---
'@xlsx-editor.dev/core': minor
'@xlsx-editor.dev/react': minor
'@xlsx-editor.dev/vue': minor
---

Add a named toolbar part for each slot, such as `Toolbar.Bold`. A slot now takes `icon`, `asChild`, `className`, and `overflowContent`. Add `Toolbar.Action` and `Toolbar.Separator`. `Toolbar.Group` now takes `priority`, `pinned`, and `overflowContent`. Behavior change: a host group with no `priority` now leaves the bar before every built-in group.
