---
'@xlsx-editor.dev/i18n': minor
'@xlsx-editor.dev/core': minor
'@xlsx-editor.dev/react': minor
'@xlsx-editor.dev/vue': minor
---

A catalog with a right-to-left `_lang` or `_dir` now mirrors the chrome. The grid keeps a left-to-right layout. Add `LocaleProvider`, `useChromeTranslate`, and `useChromeDir`. Add a `t` prop to the toolbar, the formula bar, the sheet tabs, the zoom control, and the status message. `useTranslation()` now also returns `catalogue`, `lang`, and `dir`. The root's `locale` now merges onto the catalog of an outer `LocaleProvider`.
