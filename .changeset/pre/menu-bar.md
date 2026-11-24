---
'@xlsx-editor.dev/core': minor
'@xlsx-editor.dev/react': minor
'@xlsx-editor.dev/vue': minor
'@xlsx-editor.dev/i18n': minor
---

Add a menu bar with File, Edit, View, Insert, and Format menus (`XlsxEditor.Menu` in React, `XlsxEditorMenu` in Vue). The packaged editor now shows the menu bar by default. Pass `menu={false}` to keep the old layout. The default File > Open loads a file that the user picks, and it does not ask first about unsaved edits: pass `onOpen` to ask. F6 now moves the keyboard to the menu bar first. `XlsxEditorInstance` has a new `selectAll()` method. `CHROME_GROUPS`, `CHROME_SLOT_IDS`, and `ChromeSlotId` hold 12 more slots that only the menu bar shows, and `ChromeControlKind` has the new value `'open'`. Use `defaultChromeGroups()` for the groups of the preset toolbar. `ChromeGroupId` has the new values `'edit'`, `'view'`, `'sheet'`, and `'merge'`; `ChromeGroup` has the new field `contextual`, and `ChromeControl` the new field `defaultToolbar`. `ToolbarHost` has the new members `open` and `readClipboard`. `EditRefusalCode` has the new values `'clipboard-read'`, for a Paste whose clipboard read the browser refuses, and `'open-failed'`, for a picked file that the editor cannot open: the open workbook and its edits stay. `EditorFailure.code` has the new value `'read'`, for a picked file that the browser cannot read when no workbook is open. A host with an exhaustive switch over `EditorFailure.code` must handle `'read'`. The alignment rows are `menuitemradio`, and `Menu.Row` takes `selected` for a row of a one-of set.
