# @xlsx-editor.dev/i18n

## 0.1.0-beta.2

### Patch Changes

- 4b61d19: The README of each package now states that the package is a beta and is not ready for production use.

## 0.1.0-beta.1

### Minor Changes

- 2cc37c1: A catalog with a right-to-left `_lang` or `_dir` now mirrors the chrome. The grid keeps a left-to-right layout. Add `LocaleProvider`, `useChromeTranslate`, and `useChromeDir`. Add a `t` prop of type `ChromePartTranslate` to the toolbar, the formula bar, the sheet tabs, the zoom control, and the status message. `useChromeTranslate` returns a `ChromeTranslate`, which always returns a string. `useTranslation()` now also returns `catalogue`, `lang`, and `dir`. The root's `locale` now merges onto the catalog of an outer `LocaleProvider`. Add `formatMessage`, `localeDirection`, `catalogDirection`, and `chromeDirection` to `@xlsx-editor.dev/i18n`.
- 634b50e: Add the grid context menu: a secondary press, Shift+F10, or the Menu key opens clipboard, row, column, clear, merge, and number format rows, which hosts hide, replace, extend, or remove through `XlsxEditor.ContextMenu` and the `contextMenu` prop.
- fc57e49: The font color and fill color controls are now split buttons with the color picker of the DOCX and PPTX editors: Automatic or No fill, the workbook theme colors with tints and shades, the standard colors, and a custom hex color. `Toolbar.FontColor` and `Toolbar.FillColor` take `icon`. The snapshot has `themeColors`, and core exports the headless picker (`colorPickerGrid` and its helpers). The grid shows its focus ring only when the keyboard moves the focus to it, not after a click on a cell.
- ad70495: A macro-enabled workbook (`.xlsm`) or template (`.xltm`) now opens with its macro part kept as inert bytes and refuses structure edits that macro code can see, and a host that names the saved file from `workbookKind` must also read the new `snapshot().macroEnabled` to pick the `.xlsm` or `.xltm` extension.
- 5188a18: Add a menu bar with File, Edit, View, Insert, and Format menus (`XlsxEditor.Menu` in React, `XlsxEditorMenu` in Vue). The packaged editor now shows the menu bar by default. Pass `menu={false}` to keep the old layout. The default File > Open loads a file that the user picks, and it does not ask first about unsaved edits: pass `onOpen` to ask. F6 now moves the keyboard to the menu bar first. `XlsxEditorInstance` has a new `selectAll()` method. `CHROME_GROUPS`, `CHROME_SLOT_IDS`, and `ChromeSlotId` hold 12 more slots that only the menu bar shows, and `ChromeControlKind` has the new value `'open'`. Use `defaultChromeGroups()` for the groups of the preset toolbar. `ChromeGroupId` has the new values `'edit'`, `'view'`, `'sheet'`, and `'merge'`; `ChromeGroup` has the new field `contextual`, and `ChromeControl` the new field `defaultToolbar`. `ToolbarHost` has the new members `open` and `readClipboard`. `EditRefusalCode` has the new values `'clipboard-read'`, for a Paste whose clipboard read the browser refuses, and `'open-failed'`, for a picked file that the editor cannot open: the open workbook and its edits stay. `EditorFailure.code` has the new value `'read'`, for a picked file that the browser cannot read when no workbook is open. A host with an exhaustive switch over `EditorFailure.code` must handle `'read'`. The alignment rows are `menuitemradio`, and `Menu.Row` takes `selected` for a row of a one-of set.
- d558e11: Do the protection checks of a transaction once, not for each cell write. A workbook with a stored sheet name of more than 32,767 UTF-16 code units does not open, and the refusal names the limit `workbook.maxSheetNameLength`.
- ce9474a: A selection can now hold up to `MAX_SELECTION_RANGES` ranges, from a Control or Command press, the name box, `editor.selectRanges`, or `x:selection/@sqref` on open, with every range in `snapshot.selectionRef` and in the optional `ranges` and `activeRange` fields of `GridSelection` and `HistorySelection`, where a clear or a format applies to every range, a command that acts on one range refuses with `multiple-ranges`, a range past the limit refuses with `too-many-ranges`, and `mergeOps`, `unmergeOp`, and `structureOp` throw a `RangeError` for several ranges.
- 9e973a1: Add View > Headings, which shows or hides the row and column headings of the shown sheet and writes `x:sheetView/@showRowColHeaders`, so undo, redo, and save keep it. Add the `view.headings` slot, the `Menu.Headings` named row, the `setShowHeadings` op and editor method, `EditorSnapshot.headings`, and the `menu.headings` catalog key.

### Patch Changes

- ec46ef7: A sheet that the editor adds to a workbook with a chart, a pivot table, a chartsheet, or a data connection can now be renamed and deleted in the session, unless that content can refer to the sheet. Add `EditorSheet.editRefusal`, the reason of each sheet, which `sheetCommandState` reads; a sheet without it takes `sheetEditRefusal`, which now holds the reason of the sheets of the opened file only. The scan of a table, query table, drawing, or legacy drawing part now also reads a sheet name that a CDATA section, a comment, or a line break in an attribute splits, so a structural edit refuses in more cases.

## 0.1.0-beta.0

### Minor Changes

- 7c52f6d: Publish the first beta of the workbook editor engine and its React and Vue adapters.
