---
'@xlsx-editor.dev/core': minor
'@xlsx-editor.dev/react': minor
'@xlsx-editor.dev/vue': minor
'@xlsx-editor.dev/i18n': minor
---

The font color and fill color controls are now split buttons with the color picker of the DOCX and PPTX editors: Automatic or No fill, the workbook theme colors with tints and shades, the standard colors, and a custom hex color. `Toolbar.FontColor` and `Toolbar.FillColor` take `icon`. The snapshot has `themeColors`, and core exports the headless picker (`colorPickerGrid` and its helpers). The grid shows its focus ring only when the keyboard moves the focus to it, not after a click on a cell.
