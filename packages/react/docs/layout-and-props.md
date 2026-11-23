# Layout and props

This guide describes the parts of a custom layout, the props of `XlsxEditor.Root`, and the hooks of `@xlsx-editor.dev/react`. For setup, see the [package README](../README.md).

## Viewport props

`XlsxEditor.Viewport` takes three optional props:

| Prop | Type | What it does |
| --- | --- | --- |
| `style` | `CSSProperties` | Inline style of the viewport. Without a height here or in `className`, the grid has a height of 0. |
| `className` | `string` | Added after the `xlsx-editor` and `xlsx-editor-viewport` classes. |
| `label` | `string \| false` | The name of the landmark region. The default is "Workbook editor". `false` renders no region. |

## Status message

Keep `XlsxEditor.StatusMessage` in your layout. It is the only part that shows the text of the editor state: the empty and loading messages, the reason a workbook did not open, the reason an edit was refused, the notice that formula results can be out of date, and the notice for conditional formatting that the grid can leave out.

While a workbook is open, it is a strip above the grid. The strip is one line high when it is empty. A refusal shows in full on its own row, and the strip grows while the refusal shows. A notice that shows alone wraps when the strip is too narrow for it. The `xlsx-footer` class marks the footer region that F6 visits.

In development, `XlsxEditor.Root` logs one warning when the layout has `XlsxEditor.Content` and no `XlsxEditor.StatusMessage`.

## Named exports

Each part is also a named export: `XlsxEditorRoot`, `XlsxEditorViewport`, `XlsxEditorContent`, `XlsxEditorToolbar`, `XlsxEditorToolbarSlot`, `XlsxEditorToolbarGroup`, `FormulaBar`, `SheetTabs`, `ZoomControl`, and `StatusMessage`. The `XlsxEditorNamespace` type describes `XlsxEditor` with its parts, and the `XlsxEditorToolbarNamespace` type describes `XlsxEditor.Toolbar` with its parts.

## Root props

`XlsxEditor.Root` takes these props:

| Prop | Type | Default | What it does |
| --- | --- | --- | --- |
| `document` | `Uint8Array` or `ArrayBuffer` | none | The bytes of a workbook (`.xlsx`) or a template (`.xltx`). A new array or buffer opens the new workbook. A change to `undefined` closes the workbook, and the editor shows none. |
| `fonts` | `FontSource[]` or a function that resolves to them | the packaged substitute faces | The faces that measure and paint text. Different faces replace the current faces in place: the workbook, the selection, and the undo history stay. A list is compared by its faces, so an inline list with the same faces loads nothing. A function is compared by identity, so keep it stable, for example with `useCallback`. |
| `numberLocale` | `NumberLocale` | US English separators | The separators and built-in format codes of the shown values, and how typed values read. See [Number locale](#number-locale). |
| `mode` | `'edit'` or `'view'` | `'edit'` | Whether the editor takes edits. A change applies without a reload. For when a change applies and when it waits, see [View mode](view-mode-and-protection.md#view-mode). |
| `locale` | `Translations` | English | The UI strings. A missing or `null` entry falls back to English. |
| `colorMode` | `'light'`, `'dark'`, or `'system'` | `'light'` | The color theme of the chrome, as in DOCX Editor. See [Color mode](#color-mode). |
| `onReady` | `(editor: XlsxEditorInstance) => void` | none | Called once for each editor, after the grid attaches and the first open of `document` settles, as in DOCX Editor. Without `document`, it is called right after the grid attaches. A selection or a zoom set from it applies to the opened workbook. |
| `onChange` | `(snapshot: EditorSnapshot) => void` | none | Called after each change of the editor state: a content change or a view change. |
| `onFontError` | `(failure: EditorFailure) => void` | none | Called once when the fonts do not load. The editor also shows the failure, and the snapshot turns `failed` with the code `fonts`. When some packaged substitute faces do not load and the others do, the workbook opens, and the failure lists them in `substituteFailures`. A download of a packaged substitute face fails when it sends no byte for 10 seconds. After one download fails this way, the limit for the others is 2.5 seconds. All substitute face downloads of one open end after 30 seconds, so a slow network cannot hold the open. Without a handler, each failure prints one console warning. |
| `ref` | `XlsxEditorRef` | none | The imperative handle. |

`XlsxEditor` takes the same props, and also `className`, `style`, `onSave`, `onSaveStart`, and `opening`.

### Number locale

`numberLocale` sets the separators and built-in format codes of the shown values. `currencySymbol` sets the symbol of the currency entries of the list. The number locale also sets how typed values read: the separators, and the day and month order of a typed date, from built-in format 14.

A typed date and time takes built-in format 22 when format 22 is a plain date in the order of format 14 and one clock. Typed seconds add `:ss`, and a typed AM or PM replaces the clock of format 22. Otherwise it takes the code of format 14 and the typed clock, such as `dd/mm/yyyy h:mm AM/PM`, or `yyyy-mm-dd` and the clock when format 14 has no date order. Format 14 has an order only when it is a day, a month, and a year with one separator. With no order, only an ISO 8601 date reads as a date, and it takes `yyyy-mm-dd`.

A different value applies in place. An open cell edit keeps the number locale it opened with; the next edit reads the new one. The value is compared by its separators, codes, and currency symbol, so an inline object with the same value changes nothing.

The `currencySymbol` of `numberLocale` is `$` when absent, also with `decimal: ','`. A symbol has 1 to 8 UTF-16 code units. It has no control code, no U+FFFE or U+FFFF, and no lone surrogate. If the symbol does not meet these rules, the list leaves out Currency and Accounting. The symbol always goes before the number.

### Color mode

`colorMode` sets the color theme of the chrome, as in DOCX Editor. `'system'` follows the OS setting as it changes. `XlsxEditor.Viewport` carries the theme for the parts inside it. A part outside the viewport, such as a popup in `document.body`, gets the theme from the classes of `useScopeClassName()`: `'xlsx-editor '` or `'xlsx-editor dark '`. Inside the viewport, the hook returns `''`. See [Appearance](https://xlsx-editor.dev/docs/customization/appearance). The workbook canvas keeps the workbook's own colors. A change applies without a remount. To match parts that you build, read `useEditorColorMode()`, which returns `{ colorMode: 'light' | 'dark' }`.

## High contrast themes

The stylesheet supports high contrast themes with no prop. With a forced-colors theme (`@media (forced-colors: active)`), the browser replaces the chrome colors with system colors and removes shadows. The stylesheet then draws each state with system colors:

- Pressed toolbar toggles, the selected sheet tab, the selected list option, the selected **No fill**, and the hovered or focused menu row use `Highlight` with `HighlightText`. A focused menu row also gets a `HighlightText` ring inside it, so it differs from a hovered row. A hovered or focused list option, and the **More number formats** item of the number format list, get a `CanvasText` ring inside them, so they differ from the selected option.
- Other focus rings use `Highlight`. The ring of a sheet tab is inside the tab, and on the selected tab it uses `HighlightText`. The grid draws its focus ring as an outline.
- The selection and the active cell draw a `Highlight` line with a `Canvas` line inside it, so they show on any cell fill. When the grid does not have the focus, the line of the active cell is `CanvasText`.
- The selected row and column headings and the guide line of a resize drag use `Highlight`.
- Disabled controls use `GrayText`. A disabled toolbar button, field, or color trigger also keeps its fade to 35% opacity, as in the light and dark themes. A disabled control that has the focus does not fade.
- The line between two toolbar groups uses `CanvasText`.
- Color swatches and the color bar of the font color and fill color buttons keep their own colors (`forced-color-adjust: none`). The selected swatch gets a `Highlight` outline and a `Highlight` line inside its edge. A focused swatch gets a `CanvasText` ring in place of the outline, so it differs from the selected swatch. The line inside stays while the selected swatch has the focus.

The workbook canvas keeps the workbook's own colors. If you override these states, use system colors in a `forced-colors` media query of your own.

Each color swatch in the color lists is 24 by 24 CSS pixels, the minimum target size of Web Content Accessibility Guidelines (WCAG) 2.5.8. In the light and dark themes, the selected **No fill** gets the outline of a selected swatch.

## Unsaved edits

`onChange` receives the snapshot. DOCX Editor passes a document change, and PPTX Editor passes the editor instance. This editor passes the snapshot because it is the one immutable state object that chrome renders, and it carries the content revision. `snapshot.revision` is 0 after an open and grows by one with each edit, undo, or redo; a selection, zoom, or sheet change keeps it. To act on content changes only, or to track unsaved edits, compare it with the revision that you kept:

<!-- sample: react-unsaved -->

```tsx
import { useRef, useState } from 'react';
import { XlsxEditor, type XlsxEditorRef } from '@xlsx-editor.dev/react';
import '@xlsx-editor.dev/react/styles.css';

export function TrackedEditor({ bytes }: { bytes: Uint8Array }) {
  const editor = useRef<XlsxEditorRef>(null);
  const savedRevision = useRef(0);
  const [unsaved, setUnsaved] = useState(false);

  function save() {
    const output = editor.current?.save();
    // `save()` returns null before the editor mounts, when no workbook is open,
    // while a workbook loads, after a load fails, and when the editor cannot
    // store an open cell edit.
    if (!output || !editor.current) return;
    // Store `output` here. A save calls `onChange` only when it commits
    // an open cell edit, clears an earlier refusal, or is refused.
    savedRevision.current = editor.current.snapshot().revision;
    setUnsaved(false);
  }

  return (
    <>
      <button type="button" onClick={save}>
        Save
      </button>
      <span role="status">{unsaved ? 'Unsaved edits' : 'No unsaved edits'}</span>
      <XlsxEditor
        ref={editor}
        document={bytes}
        style={{ height: 600 }}
        onChange={(snapshot) => {
          // A new open starts at revision 0 with nothing unsaved. An undo also
          // increases the revision, so an undo back to the saved content still
          // counts as unsaved. An open cell edit with changed text counts too:
          // opening another workbook drops it.
          if (snapshot.status !== 'ready') savedRevision.current = 0;
          const draft = snapshot.editing && snapshot.editChanged;
          setUnsaved(snapshot.revision !== savedRevision.current || draft);
        }}
      />
    </>
  );
}
```

`onChange` runs after every change of the editor state, also a selection or a zoom. The sample sets state each time; `setUnsaved` with the same value does not render again. To skip more work, compare the value with the last one before you set state, or read it with `useEditorState` and a selector. A save with nothing to commit does not call `onChange`, so the `save` function of the sample keeps the saved revision and clears `unsaved` itself. If you store the bytes of `saveInTasks()` or `onSave`, first check that `SavedWorkbook.stillOpen` is `true`. Then set `savedRevision.current` to `SavedWorkbook.revision`, because an edit made during the save is not in the bytes. `save()` commits an open cell edit first. A new `document` does not: the editor drops an open cell edit when it opens other bytes, so ask before you open them while `editing` and `editChanged` are both true.

## Font changes

When `fonts` changes, the editor loads the new faces first. If they do not load, the old faces stay and `onFontError` is called with the cause. These are the instance methods `setFonts()` and `setNumberLocale()`.

## Differences from the DOCX Editor root

These props differ from the DOCX Editor root. There, `locale` is a BCP-47 tag for regional input, and the UI strings are the `i18n` prop. Here, `locale` holds the UI strings, as the PPTX Editor root does. The number locale is a separate prop, `numberLocale`, because the engine formats numbers from explicit separators and format codes, not from a BCP-47 tag: it ships no regional data to resolve a tag.

## Hooks

Read a slice of the editor state with `useEditorState(selector)`, which re-renders only when the slice changes. Read the whole state with `useEditorSnapshot()`, and the editor with `useXlsxEditor()`. Get the translator of the root's `locale` with `useTranslation()`, which returns `{ t }`. Read whether a change of the `mode` prop waits for a cell edit with `useViewModeWaits()`. Get the scope classes for a part outside the viewport with `useScopeClassName()`, as [Color mode](#color-mode) describes. Run code on a content change, a selection change, a refused edit, or a failed open with `useEditorEvent(event, handler)`, as [Editor events](https://xlsx-editor.dev/docs/customization/events) describes. Its `change` event fires only for content changes, while `onChange` fires after every state change.
