---
'@xlsx-editor.dev/core': patch
---

An Escape that closes an open menu or toolbar list while the keyboard is in the host page now has its default prevented, so a host handler that closes its own panel on Escape can check `defaultPrevented` and keep the panel open.
