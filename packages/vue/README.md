<p align="center">
  <a href="https://www.xlsx-editor.dev/">
    <picture>
      <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/eigenpal/xlsx-editor/main/.github/assets/readme-logo-dark.svg" />
      <source media="(prefers-color-scheme: light)" srcset="https://raw.githubusercontent.com/eigenpal/xlsx-editor/main/.github/assets/readme-logo-light.svg" />
      <img src="https://raw.githubusercontent.com/eigenpal/xlsx-editor/main/.github/assets/readme-logo-light.svg" alt="XLSX Editor by EigenPal" width="320" height="90" />
    </picture>
  </a>
</p>

<p align="center">
  <a href="https://www.npmjs.com/package/@xlsx-editor.dev/vue"><img src="https://img.shields.io/npm/v/@xlsx-editor.dev/vue.svg?style=flat-square&color=107C41" alt="npm version" /></a>
  <a href="https://www.npmjs.com/package/@xlsx-editor.dev/vue"><img src="https://img.shields.io/npm/dm/@xlsx-editor.dev/vue.svg?style=flat-square&color=107C41" alt="npm downloads" /></a>
  <a href="https://github.com/eigenpal/xlsx-editor/blob/main/LICENSE"><img src="https://img.shields.io/badge/license-Apache_2.0-blue.svg?style=flat-square&color=107C41" alt="license" /></a>
</p>

# @xlsx-editor.dev/vue

The Vue 3 adapter for XLSX Editor. It is the twin of `@xlsx-editor.dev/react`. `XlsxEditorRoot` owns the engine instance, `XlsxEditorViewport` scopes the styles, and `XlsxEditorContent` attaches the grid. The toolbar, formula bar, sheet tabs, and zoom control read the engine through composables and hold no workbook state.

The Vue parts have the same props, markup, and behavior as the React parts. A part on React's `XlsxEditor.*` namespace is `XlsxEditor<Part>` or the part name in Vue, for example `XlsxEditorToolbar` or `FormulaBar`. A parity check maps each React prop, ref member, and export to its Vue twin.

## Install

Install the package with its peer dependencies. `@xlsx-editor.dev/core` is a peer, so one dependency tree loads one copy of the engine:

```bash
npm install @xlsx-editor.dev/vue @xlsx-editor.dev/core vue
```

Install all `@xlsx-editor.dev` packages at the same minor version. The adapter accepts only the minor of `@xlsx-editor.dev/core` that it releases with, for example `~0.1.0`. The peer range of `vue` is `^3.5.0`. The tests of this repository run with Vue 3.5.45.

The type declarations need TypeScript 5.7 or later, with `moduleResolution` set to `bundler` or `node16`.

The editor measures text with a WebAssembly text shaper. With Vite, use version 7.0.0 or later. `vite build` and the Vite 8 dev server handle the text shaper's `harfbuzz.wasm` and the font files with no setup. The Vite 7 dev server needs the `optimizeDeps` setting that the core README shows. If your bundler does not emit the files, serve them yourself. See [Serve the text shaper](https://github.com/eigenpal/xlsx-editor/blob/main/packages/core/README.md#serve-the-text-shaper) in the core README.

## Show a workbook

Import the stylesheet once, get the workbook bytes, and pass them to the editor. Import the stylesheet after your app's CSS reset. The following component fetches `/workbook.xlsx` from your app:

```vue
<script setup lang="ts">
import { onBeforeUnmount, shallowRef } from 'vue';
import { XlsxEditor } from '@xlsx-editor.dev/vue';
import '@xlsx-editor.dev/vue/styles.css';

const bytes = shallowRef<ArrayBuffer>();
const error = shallowRef<string>();
const request = new AbortController();
fetch('/workbook.xlsx', { signal: request.signal })
  .then((response) => {
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    return response.arrayBuffer();
  })
  .then((buffer) => (bytes.value = buffer))
  .catch((cause: Error) => {
    if (!request.signal.aborted) error.value = cause.message;
  });
onBeforeUnmount(() => request.abort());
</script>

<template>
  <p v-if="error" role="alert">The workbook did not load: {{ error }}</p>
  <XlsxEditor v-else :document="bytes" style="height: 600px" />
</template>
```

To open a file that the user picks, pass `await file.arrayBuffer()` as `document`.

The editor needs a browser. If your app renders on a server, render the editor only in the browser. In Nuxt, use the `<ClientOnly>` component.

Set the height on the editor itself, with `style` or `class`. The grid fills the editor and scrolls inside it. A height on a parent element is not enough: without its own height, the grid has a height of 0. In development, the editor logs a warning when the stylesheet is missing.

If a workbook does not open, the editor shows the reason in `StatusMessage`. To read the reason in code, listen to `change`. Its `status` is `failed`, and `failure.code` is `package`, `fonts`, or `no-worksheet`. `failure.detail` gives the engine's reason for your logs, as [Open a workbook](https://github.com/eigenpal/xlsx-editor/blob/main/packages/core/README.md#open-a-workbook) in the core README describes.

Without a `fonts` prop, the editor fetches the packaged substitute fonts of `@docx-editor.dev/fonts`. It never fetches anything that a workbook names.

## Build your own layout

`XlsxEditorRoot` owns the editor. `XlsxEditorViewport` scopes the styles and names the landmark region. `XlsxEditorContent` attaches the grid. Each part is also a static of `XlsxEditor`, for example `XlsxEditor.Toolbar`, as in React. Put the chrome parts where you want them, and set the height on `XlsxEditorViewport`:

```vue
<script setup lang="ts">
import {
  FormulaBar,
  SheetTabs,
  StatusMessage,
  XlsxEditorContent,
  XlsxEditorRoot,
  XlsxEditorToolbar,
  XlsxEditorViewport,
  ZoomControl,
} from '@xlsx-editor.dev/vue';

defineProps<{ bytes?: Uint8Array }>();
</script>

<template>
  <XlsxEditorRoot :document="bytes">
    <XlsxEditorViewport style="height: 600px">
      <XlsxEditorToolbar />
      <FormulaBar />
      <StatusMessage />
      <XlsxEditorContent />
      <div class="xlsx-footer">
        <SheetTabs />
        <ZoomControl />
      </div>
    </XlsxEditorViewport>
  </XlsxEditorRoot>
</template>
```

Keep `StatusMessage` in your layout. It is the only part that shows the text of the editor state, such as the reason for a refused edit. Composables run only in a descendant of `XlsxEditorRoot`.

`XlsxEditorToolbar` takes the parts `XlsxEditorToolbar.Slot` and `XlsxEditorToolbar.Group` in its default slot, as React's toolbar takes them as children. The following sample replaces the save control, removes the underline control, and adds a host group after the history group. Use the component of the sample in place of `<XlsxEditorToolbar />` in your layout, inside `XlsxEditorViewport`. `:preset="false"` renders only the default slot. The [Toolbar and keyboard](https://github.com/eigenpal/xlsx-editor/blob/main/packages/react/docs/toolbar-and-keyboard.md#toolbar-parts) guide describes the rules:

```vue
<script setup lang="ts">
import { XlsxEditorToolbar } from '@xlsx-editor.dev/vue';
// Your own components.
import SaveToDriveButton from './SaveToDriveButton.vue';
import ShareButton from './ShareButton.vue';
</script>

<template>
  <XlsxEditorToolbar>
    <XlsxEditorToolbar.Slot slot-id="file.save">
      <SaveToDriveButton />
    </XlsxEditorToolbar.Slot>
    <XlsxEditorToolbar.Slot slot-id="font.underline" hidden />
    <XlsxEditorToolbar.Group id="share" label="Share" after="history">
      <ShareButton />
    </XlsxEditorToolbar.Group>
  </XlsxEditorToolbar>
</template>
```

## Props and events

`XlsxEditor` and `XlsxEditorRoot` take these props. `XlsxEditor` passes `class`, `style`, and other attributes to the viewport.

| Prop | Type | Default | Description |
| --- | --- | --- | --- |
| `document` | `Uint8Array \| ArrayBuffer` | none | The bytes of a workbook or a template package. A new value opens that workbook, and `undefined` closes the open workbook. |
| `fonts` | `XlsxEditorFonts` | The packaged substitute faces | The faces that measure and paint text. Without it, the packaged substitute faces load. A new value replaces the faces in place. |
| `numberLocale` | `NumberLocale` | US English separators | The number locale of the shown values: separators and built-in format codes. A new value applies in place. |
| `mode` | `'edit' \| 'view'` | `'edit'` | The editing mode. A change applies without a reload. |
| `locale` | `Translations` | English | The strings of the chrome. Missing keys fall back to English. |
| `colorMode` | `'light' \| 'dark' \| 'system'` | `'light'` | The color mode of the chrome. The workbook keeps its own colors. |
| `opening` | `boolean` | `false` | `XlsxEditor` only. Set it while your app reads a new workbook, so the toolbar refuses **Save** until the read ends. |

The React callbacks are events in Vue:

| Event | Payload | When |
| --- | --- | --- |
| `ready` | `XlsxEditorInstance` | Once for each editor, after the grid attaches and the first open of `document` settles. Without `document`, right after the attach. |
| `change` | `EditorSnapshot` | After each change of the editor state. |
| `fontError` | `EditorFailure` | Once when the fonts of an open or a new `fonts` value do not load. When only some packaged substitute faces fail, the workbook opens and `substituteFailures` lists them. Without a listener, the editor logs a warning. |
| `save` | `(bytes: Uint8Array, saved: SavedWorkbook)` | When the user selects **Save** or presses Control or Command and S. `XlsxEditor` only. |
| `saveStart` | `Promise<SavedWorkbook \| null>` | When such a press starts a save, so you can show that it runs. `XlsxEditor` only. |

The **Save** control of `XlsxEditor` and of `XlsxEditorToolbar` is enabled only when you listen to `save`. Listen with `@save`. A `@save.once` listener leaves **Save** disabled, as no listener does, because a second press would have no listener. The other events take `.once`. Before you store the bytes, check `saved.stillOpen`. It is `false` when another open ended the workbook during the save:

```vue
<script setup lang="ts">
import { XlsxEditor, type SavedWorkbook } from '@xlsx-editor.dev/vue';

defineProps<{ bytes: Uint8Array }>();

function keep(edited: Uint8Array, saved: SavedWorkbook) {
  if (!saved.stillOpen) return;
  // Store or upload `edited`.
}
</script>

<template>
  <XlsxEditor :document="bytes" style="height: 600px" @save="keep" />
</template>
```

The guides of the React package describe the parts in detail. The Vue parts behave the same. Each React callback prop is an event without the `on` prefix in Vue: for example, `onSave` is `@save`. See these guides:

- [Layout and props](https://github.com/eigenpal/xlsx-editor/blob/main/packages/react/docs/layout-and-props.md)
- [Toolbar and keyboard](https://github.com/eigenpal/xlsx-editor/blob/main/packages/react/docs/toolbar-and-keyboard.md)
- [Edit cells](https://github.com/eigenpal/xlsx-editor/blob/main/packages/react/docs/editing-cells.md)
- [Format cells and change the structure](https://github.com/eigenpal/xlsx-editor/blob/main/packages/react/docs/formatting-and-structure.md)
- [Sheets, clipboard, and sizes](https://github.com/eigenpal/xlsx-editor/blob/main/packages/react/docs/sheets-clipboard-and-sizes.md)
- [View mode and protection](https://github.com/eigenpal/xlsx-editor/blob/main/packages/react/docs/view-mode-and-protection.md)

For the host page, the keyboard, and what a screen reader hears, see [Accessibility](https://github.com/eigenpal/xlsx-editor/blob/main/docs/site/content/accessibility.mdx).

## Edit and save

To edit a cell, select it and type, press F2, or double-click it. Press Enter to commit the edit. Text that starts with `=` is a formula in A1 notation. The editor stores a formula without a result and asks the next consumer to recalculate the workbook. If the editor cannot store an edit, it keeps your text and shows the reason.

To get the edited workbook in code, call `save()` on a template ref. Type the ref as `XlsxEditorRef`. The bytes keep the kind of the opened package:

```vue
<script setup lang="ts">
import { useTemplateRef } from 'vue';
import { XlsxEditor, type XlsxEditorRef } from '@xlsx-editor.dev/vue';

defineProps<{ bytes: Uint8Array }>();
const editor = useTemplateRef<XlsxEditorRef>('editor');

function keep() {
  const edited = editor.value?.save();
  // Store or upload `edited`.
}
</script>

<template>
  <button type="button" @click="keep">Keep the edits</button>
  <XlsxEditor ref="editor" :document="bytes" style="height: 600px" />
</template>
```

`save()` returns `null` before the editor mounts, when no workbook is open, and while a workbook loads. It also returns `null` after a load fails, and when the editor cannot store an open cell edit. Call the editing commands on the editor instance that `getEditor()` returns.

## View mode

To show a workbook without letting the user change it, pass `mode="view"`. The default is `"edit"`. A change of `mode` applies without a reload. In view mode, the editor refuses every edit with the reason `view-mode`. You can still select cells, copy, switch sheets, zoom, and save.

## Embed in a page

The editor shares the page with your app:

- Every rule in the stylesheet selects an `xlsx-` class first. The rules do not style elements outside the editor.
- The editor chrome sets the styles it needs. If you import the stylesheet after your app's CSS reset, the reset does not change the chrome.
- The editor can be in a container with a `transform` or `scale` property, such as a drawer or a side panel. There, toolbar lists open at their controls. A press on the grid selects the cell under the pointer. A column or row resize follows the pointer, and the sheet tab menu opens at the pointer.
- In a container that clips its overflow, toolbar lists stay in the visible part of the container.
- More than one editor can be on a page. Control+S (Command+S on macOS) saves only the editor that holds the focus.

Your own content inside the editor, such as a layout around its parts in `XlsxEditorViewport`, keeps your styles. It gets only the editor's text styles and box sizing. Content inside an element with an `xlsx-` class also gets the chrome defaults for controls. To change the chrome text styles, set them on `.xlsx-editor` after the editor stylesheet.

## Composables

| Composable | Returns |
| --- | --- |
| `useXlsxEditor()` | `UseXlsxEditorReturn`: a ref to the editor, or null before it mounts |
| `useEditorState(selector, isEqual?, deps?)` | `UseEditorStateReturn<T>`: the selected slice of the snapshot, updated only when it changes. `deps` names a source whose change reads the slice again |
| `useEditorSnapshot()` | `UseEditorSnapshotReturn`: the whole snapshot |
| `useEditorCommand(slotId)` | `UseEditorCommandReturn`: `execute`, which returns `true` when the engine ran the command and `false` on a refusal, and `isEnabled`, `isActive`, `disabledReason`, `hint`, and `value` as computed refs |
| `useTranslation()` | `UseTranslationReturn`: `t`, in the locale of `XlsxEditorRoot` |
| `useViewModeWaits()` | `UseViewModeWaitsReturn`: `true` while a change of `mode` to view mode waits for the open cell edit |
| `useEditorColorMode()` | `UseEditorColorModeReturn`: the resolved color mode of the chrome |

To build your own control, bind it to a chrome slot with `useEditorCommand`. Put the control inside `XlsxEditorRoot`. Prevent the default action of `mousedown`, so the control does not take the focus from the grid:

```vue
<script setup lang="ts">
import { useEditorCommand } from '@xlsx-editor.dev/vue';

const bold = useEditorCommand('font.bold');
</script>

<template>
  <button
    type="button"
    :aria-pressed="bold.isActive.value"
    :aria-disabled="!bold.isEnabled.value"
    :title="bold.disabledReason.value ?? 'Bold'"
    @mousedown.prevent
    @click="bold.execute()"
  >
    Bold
  </button>
</template>
```

The composables read the same core functions as the React hooks. So both adapters show the same enabled state and refusal reasons. `XlsxEditorRoot` holds one editor subscription and shares it with every composable in it. For each export, see the [API report](https://github.com/eigenpal/xlsx-editor/blob/main/docs/api/vue/index.api.md).

## Docs and demo

- [Vue adapter](https://github.com/eigenpal/xlsx-editor/blob/main/docs/site/content/vue/index.mdx): install, mount, and the map from the React API to Vue.
- Live demo: `bun run dev:vue` in the repository (`examples/vue`).

## License

Apache-2.0. See [LICENSE](LICENSE).
