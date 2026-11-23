# @xlsx-editor.dev/react

The React adapter for XLSX Editor. It renders the workbook editor of `@xlsx-editor.dev/core` with its chrome: the toolbar, the formula bar, the grid, the sheet tabs, and the zoom control. You can edit cell values and formulas in the grid and in the formula bar. The adapter holds no workbook state; the engine owns it.

`@xlsx-editor.dev/core` is a peer dependency, so one dependency tree loads one copy of the engine.

## Install

Install the package with its peer dependencies:

```bash
npm install @xlsx-editor.dev/react @xlsx-editor.dev/core react react-dom
```

Install all `@xlsx-editor.dev` packages at the same minor version. The adapter accepts only the minor of `@xlsx-editor.dev/core` that it releases with, for example `~0.1.0`. The peer range of `react` and `react-dom` is `^18.0.0 || ^19.0.0`. The tests and the consumer checks of this repository run with React 19. The consumer floor check also runs React 18.0.0.

The type declarations need TypeScript 5.7 or later, with `moduleResolution` set to `bundler` or `node16`. With React 18, use `@types/react` 18.2.6 or later.

## Show a workbook

Import the stylesheet once, get the workbook bytes, and pass them to the editor. Import the stylesheet after your app's CSS reset. The following component fetches `/workbook.xlsx` from your app. It also opens a file that the user picks, and gives the page a main landmark, a heading, and a label for the file input:

<!-- sample: react-editor -->

```tsx
import { useEffect, useState } from 'react';
import { XlsxEditor } from '@xlsx-editor.dev/react';
import '@xlsx-editor.dev/react/styles.css';

export function Editor() {
  const [bytes, setBytes] = useState<ArrayBuffer>();
  const [error, setError] = useState<string>();

  useEffect(() => {
    const request = new AbortController();
    fetch('/workbook.xlsx', { signal: request.signal })
      .then((response) => {
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        return response.arrayBuffer();
      })
      // A file that the user picked first stays open.
      .then((buffer) => setBytes((picked) => picked ?? buffer))
      .catch((cause: Error) => {
        if (!request.signal.aborted) setError(cause.message);
      });
    return () => request.abort();
  }, []);

  return (
    <main>
      <h1>Workbook</h1>
      {error ? <p role="alert">The workbook did not load: {error}</p> : null}
      <label>
        Open a workbook{' '}
        <input
          type="file"
          accept=".xlsx"
          onChange={async (e) => {
            const file = e.target.files?.[0];
            if (!file) return;
            setError(undefined);
            setBytes(await file.arrayBuffer());
          }}
        />
      </label>
      <XlsxEditor document={bytes} style={{ height: 600 }} />
    </main>
  );
}
```

To open a file that the user picks, pass `await file.arrayBuffer()` as `document`. A pick with no file keeps the open workbook, because `undefined` closes it.

Set the height on the editor itself, with `style` or `className`. The grid fills the editor and scrolls inside it. A height on a parent element is not enough: without its own height, the grid has a height of 0. To fill a parent that has a height, pass `style={{ height: '100%' }}`. In development, the editor logs a warning when the stylesheet is missing.

If a workbook does not open, the editor shows the reason in `XlsxEditor.StatusMessage`. To read the reason in code, read the snapshot that `onChange` receives. Its `status` is `failed`, `failure.code` is `package`, `fonts`, or `no-worksheet`, and `failure.detail` gives the engine's reason for your logs, as [Open a workbook](https://github.com/eigenpal/xlsx-editor/blob/main/packages/core/README.md#open-a-workbook) in the core README describes.

Without a `fonts` prop, the editor fetches the packaged substitute fonts of `@docx-editor.dev/fonts`. It never fetches anything that a workbook names.

With Vite, use version 7.0.0 or later. `vite build` and the Vite 8 dev server handle the text shaper's `harfbuzz.wasm` and the font files with no setup. The Vite 7 dev server needs the `optimizeDeps` setting that the core README shows. If your bundler does not emit the files, serve the files yourself, call `setHarfBuzzWasmUrl`, and pass the faces as the `fonts` prop, as [Serve the text shaper](https://github.com/eigenpal/xlsx-editor/blob/main/packages/core/README.md#serve-the-text-shaper) in the core README describes.

## Use with Next.js

The editor runs in the browser only. Load it inside a Client Component with `dynamic(..., { ssr: false })`, and set `moduleResolution` to `bundler` in `tsconfig.json`. No CI check builds a Next.js app, so no check covers the code in this section.

The following file holds the editor:

```tsx
// app/workbook-editor.tsx
'use client';

import { XlsxEditor } from '@xlsx-editor.dev/react';
import '@xlsx-editor.dev/react/styles.css';

export function WorkbookEditor({ bytes }: { bytes: Uint8Array }) {
  return <XlsxEditor document={bytes} style={{ height: 600 }} />;
}
```

The following file loads it in the browser. Import `WorkbookEditor` from this file in your pages:

```tsx
// app/workbook-editor-loader.tsx
'use client';

import dynamic from 'next/dynamic';

export const WorkbookEditor = dynamic(
  () => import('./workbook-editor').then((module) => module.WorkbookEditor),
  { ssr: false }
);
```

## Build your own layout

`XlsxEditor.Root` owns the editor. `XlsxEditor.Viewport` scopes the styles and names the landmark region. `XlsxEditor.Content` attaches the grid. Put the chrome parts where you want them, and set the height on `XlsxEditor.Viewport`:

<!-- sample: react-layout -->

```tsx
<XlsxEditor.Root document={bytes}>
  <XlsxEditor.Viewport style={{ height: 600 }}>
    <XlsxEditor.Toolbar />
    <XlsxEditor.FormulaBar />
    <XlsxEditor.StatusMessage />
    <XlsxEditor.Content />
    <div className="xlsx-footer">
      <XlsxEditor.SheetTabs />
      <XlsxEditor.ZoomControl />
    </div>
  </XlsxEditor.Viewport>
</XlsxEditor.Root>
```

To replace, remove, or add toolbar controls in this layout, pass parts to `XlsxEditor.Toolbar`. For the parts, see [Toolbar parts](docs/toolbar-and-keyboard.md#toolbar-parts).

Keep `XlsxEditor.StatusMessage` in your layout. It is the only part that shows the text of the editor state, such as the reason an edit was refused. `XlsxEditor` takes the props of `XlsxEditor.Root`, and also `className`, `style`, `onSave`, `onSaveStart`, and `opening`.

For each prop, the hooks, and a sample that tracks unsaved edits, see [Layout and props](docs/layout-and-props.md).

## Edit and save

To edit a cell, select it and type, press F2, or double-click it. Press Enter to commit the edit. Text that starts with `=` is a formula in A1 notation. The editor stores a formula without a result and asks the next consumer to recalculate the workbook. If the editor cannot store an edit, it keeps your text and shows the reason.

To get the edited workbook, call `save()` on the ref. The bytes keep the kind of the opened package:

```tsx
import { useRef } from 'react';
import { XlsxEditor, type XlsxEditorRef } from '@xlsx-editor.dev/react';

export function Editor({ bytes }: { bytes: Uint8Array }) {
  const ref = useRef<XlsxEditorRef>(null);
  const keep = () => {
    const edited = ref.current?.save();
    // Store or upload `edited`.
  };
  return (
    <>
      <button type="button" onClick={keep}>
        Keep the edits
      </button>
      <XlsxEditor ref={ref} document={bytes} style={{ height: 600 }} />
    </>
  );
}
```

`save()` returns `null` before the editor mounts, when no workbook is open, while a workbook loads, after a load fails, and when the editor cannot store an open cell edit. Call the editing commands on the editor instance that `getEditor()` returns.

The following guides describe each area:

- [Edit cells](docs/editing-cells.md): typed values, the formula bar, the name box, saves, and refusals.
- [Format cells and change the structure](docs/formatting-and-structure.md): cell formats, inserts, deletes, and merged cells.
- [Sheets, clipboard, and sizes](docs/sheets-clipboard-and-sizes.md): sheet commands, copy and paste, and column and row sizes.

## View mode

To show a workbook without letting the user change it, pass `mode="view"`. The default is `"edit"`. A change of `mode` applies without a reload:

```tsx
<XlsxEditor document={bytes} mode={canEdit ? 'edit' : 'view'} style={{ height: 600 }} />
```

In view mode, the editor refuses every edit with the reason `view-mode`. You can still select cells, copy, switch sheets, zoom, and save. For the rules of a mode change, sheet protection, and workbook protection, see [View mode and protection](docs/view-mode-and-protection.md).

## Embed in a page

The editor shares the page with your app:

- Every rule in the stylesheet selects an `xlsx-` class first. The rules do not style elements outside the editor.
- The editor chrome sets the styles it needs. If you import the stylesheet after your app's CSS reset, such as Tailwind preflight, the reset does not change the chrome.
- In a container with a `transform` or `scale` property that moves or scales the editor, such as a drawer or a side panel, toolbar lists open at their controls. A press on the grid selects the cell under the pointer, a column or row resize follows the pointer, and the sheet tab menu opens at the pointer.
- In a container that clips its overflow, toolbar lists stay in the visible part of the container.
- More than one editor can be on a page. Control+S (Command+S on macOS) saves only the editor that holds the focus.

Your own content inside the editor, such as a layout around its parts in `XlsxEditor.Viewport`, keeps your styles. It gets only the editor's text styles and box sizing. Content inside an element with an `xlsx-` class also gets the chrome defaults for controls. To change the chrome text styles, set them on `.xlsx-editor` after the editor stylesheet.

## Toolbar and keyboard

`XlsxEditor` shows a toolbar over the grid. Pass `onSave` to get the workbook when the user selects **Save**. For the toolbar props, the controls, the dialogs, the control state, and the keyboard, see [Toolbar and keyboard](docs/toolbar-and-keyboard.md).

## Accessibility

The editor names its parts for assistive technology and works with the keyboard alone. Your page gives the language, the title, the main landmark, the page heading, and the labels of your own controls. For the host page, the keyboard, what a screen reader hears, and the known limits, see [Accessibility](https://github.com/eigenpal/xlsx-editor/blob/main/docs/site/content/accessibility.mdx).

## License

Apache-2.0. See [LICENSE](LICENSE).
