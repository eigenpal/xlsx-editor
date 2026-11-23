# @xlsx-editor.dev/core

The XLSX Editor engine. It reads a SpreadsheetML package into one canonical tree for each part, resolves cell formats, lays out the grid, and paints it. The adapters render what this package computes; they hold no editing state.

The adapters depend on it as a peer, so one dependency tree loads one copy of the engine.

## Install

Install the package with npm:

```bash
npm install @xlsx-editor.dev/core
```

Install one copy of `@xlsx-editor.dev/core`, at the same minor version as the other `@xlsx-editor.dev` packages. Each adapter accepts only the minor of `@xlsx-editor.dev/core` that it releases with, for example `~0.1.0`. The package is ESM only and needs Node.js 20.16.0 or later in the 20.x release line, or Node.js 22.3.0 or later, for server use.

The type declarations need TypeScript 5.7 or later, with `moduleResolution` set to `bundler` or `node16`.

## Serve the text shaper

The engine measures text with HarfBuzz, which ships as `@xlsx-editor.dev/core/harfbuzz.wasm`. On disk, the file is `node_modules/@xlsx-editor.dev/core/dist/harfbuzz.wasm`. With Vite, use version 7.0.0 or later.

`vite build` and webpack 5.83.0 or later emit the file on their own, so they need no setup. A CI check builds a Vite app and a webpack app, and each output must hold the text shaper, the fonts, and the stylesheets. No CI check builds with Turbopack.

If your bundler does not emit `new URL(..., import.meta.url)` assets, serve the file yourself and call `setHarfBuzzWasmUrl` from `@xlsx-editor.dev/core/layout` before you create an editor:

```ts
import { setHarfBuzzWasmUrl } from '@xlsx-editor.dev/core/layout';

setHarfBuzzWasmUrl('/static/harfbuzz.wasm');
```

The Vite 8 dev server serves the text shaper and the fonts with no setup. The Vite 7 dev server answers these requests with HTML, so exclude both packages from dependency optimization:

<!-- sample: vite-7-dev -->

```ts
// vite.config.ts
import { defineConfig } from 'vite';

export default defineConfig({
  optimizeDeps: {
    exclude: ['@xlsx-editor.dev/core', '@docx-editor.dev/fonts'],
  },
});
```

The default substitute faces of `@docx-editor.dev/fonts` load the same way, so these bundlers do not emit them either. With such a bundler, copy the files of `node_modules/@docx-editor.dev/fonts/assets/` to a static folder, and pass the faces as the `fonts` option. Each face is a `FontSource`: a family, its style, and its bytes. Register each file under the name of its own family, as `loadSubstituteFonts()` of `@xlsx-editor.dev/core/editor` does. The engine maps each workbook family to its substitute, as the following table lists, and measures any other family with Liberation Sans.

| Files | Family | Stands in for |
| --- | --- | --- |
| `Carlito-*.ttf` | `Carlito` | Aptos, Aptos Display, Calibri, Calibri Light |
| `Caladea-*.ttf` | `Caladea` | Cambria |
| `LiberationSans-*.ttf` | `Liberation Sans` | Aptos Narrow, Arial, Arial Narrow, Helvetica, Helvetica Neue, Myriad Pro, Myriad Pro Light, Open Sans, Segoe UI |
| `LiberationSerif-*.ttf` | `Liberation Serif` | Times, Times New Roman |
| `LiberationMono-*.ttf` | `Liberation Mono` | Courier, Courier New |
| `TeXGyreAdventor-*.otf` | `TeX Gyre Adventor` | Century Gothic |

The following example serves the files from `/static/fonts/`:

```ts
import { createXlsxEditor } from '@xlsx-editor.dev/core/editor';
import type { FontSource } from '@xlsx-editor.dev/core/layout';

async function face(
  file: string,
  family: string,
  bold: boolean,
  italic: boolean
): Promise<FontSource> {
  const response = await fetch(`/static/fonts/${file}`);
  if (!response.ok) throw new Error(`${file}: HTTP ${response.status}`);
  const bytes = new Uint8Array(await response.arrayBuffer());
  return { family, bold, italic, bytes, origin: 'substitute' };
}

const editor = createXlsxEditor({
  fonts: async () => [
    await face('Carlito-Regular.ttf', 'Carlito', false, false),
    await face('Carlito-Bold.ttf', 'Carlito', true, false),
    await face('LiberationSans-Regular.ttf', 'Liberation Sans', false, false),
    // Add one entry for each file that you serve.
  ],
});
```

The sample checks `response.ok` because the editor cannot identify an error page that you pass as face bytes. If the `fonts` function rejects or resolves to no faces, the workbook does not open: `editor.snapshot().status` is `failed`, and the failure code is `fonts`.

## Entry points

| Entry | What it holds |
| --- | --- |
| `@xlsx-editor.dev/core` | The same exports as `@xlsx-editor.dev/core/editor` |
| `@xlsx-editor.dev/core/editor` | `createXlsxEditor`, the editor instance, the text helpers, and the chrome registry |
| `@xlsx-editor.dev/core/store` | Bounded package read and save, the canonical tree, the document store, and its ops |
| `@xlsx-editor.dev/core/cascade` | Cell formats, number formats, theme colors, and theme fonts |
| `@xlsx-editor.dev/core/layout` | DOM-free grid geometry and text layout |
| `@xlsx-editor.dev/core/automation` | The edits that lower to store ops |
| `@xlsx-editor.dev/core/output` | Canvas paint and pictures |
| `@xlsx-editor.dev/core/contracts` | Shared types |
| `@xlsx-editor.dev/core/styles/editor.css` | Chrome tokens and styles |
| `@xlsx-editor.dev/core/harfbuzz.wasm` | The text shaper binary, for a bundler that does not emit it |

## Use

The editor adds no styles of its own, so import the stylesheet once in your app. The stylesheet applies inside an element with the `xlsx-editor` class. Without that class, the editor has no colors, so the selection and the active cell do not show. The grid also has no height, so it does not scroll.

The following element holds the editor. Give it the `xlsx-editor` class and a height:

<!-- sample: core-host -->

```html
<div id="grid" class="xlsx-editor" style="height: 600px"></div>
```

The following example opens a workbook in that element and enters a value:

<!-- sample: core-open -->

```ts
import '@xlsx-editor.dev/core/styles/editor.css';
import { createXlsxEditor } from '@xlsx-editor.dev/core/editor';

const response = await fetch('/workbook.xlsx');
const bytes = new Uint8Array(await response.arrayBuffer());

const editor = createXlsxEditor();
editor.attach(document.getElementById('grid')!);
const opened = await editor.load(bytes);
if (opened.status === 'failed' && opened.failure) {
  const { code, detail } = opened.failure;
  throw new Error(`The workbook did not open: ${code}. ${detail ?? ''}`);
}
editor.select({ row: 1, col: 1 });
const result = editor.enterValue('42');
if (!result.ok) console.log(result.refusal.code);
```

To save the result and release the editor, call `save()` and `destroy()`. `save()` returns bytes of the opened kind: `snapshot().workbookKind` is `workbook` or `template`.

```ts
const saved = editor.save();
editor.destroy();
```

Rows and columns are 0-based in the API: `{ row: 1, col: 1 }` is cell B2. The grid fills the element and scrolls inside it. For a React UI, use `@xlsx-editor.dev/react`.

### Open a workbook

`load()` does not reject when a workbook does not open. It resolves with a snapshot whose `status` is `failed`, and `failure.code` gives the reason:

| Code           | Reason                                                          |
| -------------- | --------------------------------------------------------------- |
| `package`      | The file is not a workbook that the engine can read.            |
| `fonts`        | The text shaper or the fonts did not load.                      |
| `no-worksheet` | The workbook has no visible worksheet that the editor can show. |

`failure.detail` gives the engine's reason for your logs. For example, if the browser cannot fetch `harfbuzz.wasm`, the detail names the URL and `setHarfBuzzWasmUrl`. If the server answers with HTML, it also names the `optimizeDeps` setting of the Vite 7 dev server. Do not show the detail to users.

After a failed open, `editor.save()` returns `null`, and `editor.enterValue()` refuses the edit with the code `not-editable`.

## Use on a server

The `store`, `cascade`, `layout`, and `automation` entries need no DOM, so they run on Node.js. The following example reads a workbook and saves it again:

<!-- sample: core-server-save -->

```ts
import { readFileSync, writeFileSync } from 'node:fs';
import { readXlsx, saveXlsx } from '@xlsx-editor.dev/core/store';

const read = readXlsx(new Uint8Array(readFileSync('in.xlsx')));
if (!read.ok) throw new Error(`in.xlsx did not read: ${JSON.stringify(read.error)}`);
writeFileSync('out.xlsx', saveXlsx(read.pkg));
```

If the bytes do not read, `read.error` is a `PackageRejection`. Its `reason` is one of these values:

- `zip`: the bytes are not a zip archive, or the archive is damaged or crosses a limit. `detail` gives the zip reader's reason, such as `inflate-error`.
- `no-content-types`: the archive has no `[Content_Types].xml` part.
- `not-a-workbook`: the package has no workbook part of a known content type.
- `xml`: the XML of the part `partName` is malformed or crosses a limit.

To measure text on a server, load the text shaper and give it the face bytes. `loadSubstituteFonts()` of `@xlsx-editor.dev/core/editor` reads the packaged faces with `fetch`, which cannot read local files on Node.js, so it works in a browser only. On a server, read the files of `@docx-editor.dev/fonts` yourself, under the families of [Serve the text shaper](#serve-the-text-shaper). If your package manager does not hoist `@docx-editor.dev/fonts`, install it too.

<!-- sample: core-server-measure -->

```ts
import { readFileSync } from 'node:fs';
import { FONT_ASSET_ROOT } from '@docx-editor.dev/fonts';
import {
  createHarfBuzzMeasurer,
  type FontSource,
  gridMetrics,
  initializeHarfBuzz,
  PRINT_PROFILE,
} from '@xlsx-editor.dev/core/layout';
import { readXlsx, workbookIndex } from '@xlsx-editor.dev/core/store';

await initializeHarfBuzz();
function face(file: string, family: string, bold: boolean, italic: boolean): FontSource {
  const bytes = new Uint8Array(readFileSync(new URL(file, FONT_ASSET_ROOT)));
  return { family, bold, italic, bytes, origin: 'substitute' };
}
const measurer = createHarfBuzzMeasurer([
  face('Carlito-Regular.ttf', 'Carlito', false, false),
  face('Carlito-Bold.ttf', 'Carlito', true, false),
  face('LiberationSans-Regular.ttf', 'Liberation Sans', false, false),
]);

const read = readXlsx(new Uint8Array(readFileSync('in.xlsx')));
if (!read.ok) throw new Error(`in.xlsx did not read: ${JSON.stringify(read.error)}`);
const index = workbookIndex(read.pkg);
if (!index) throw new Error('in.xlsx has no workbook part');
const metrics = gridMetrics(read.pkg, index, measurer, PRINT_PROFILE);
console.log(metrics.maxDigitWidth, metrics.defaultRowHeight);
```

The sample passes the measurer to `gridMetrics` of `@xlsx-editor.dev/core/layout`. It gets the workbook index from `workbookIndex(pkg)` of `@xlsx-editor.dev/core/store`, which returns `undefined` when the package has no workbook part, so check for that case first. The profile is `PRINT_PROFILE` or `SCREEN_PROFILE` of the layout entry. The result gives the maximum digit width and the default row height of the workbook. The column widths and row heights come from these values, by the column width formulas of ECMA-376 Part 1, 18.3.1.13 (`x:col/@width`, and `x:sheetFormatPr/@defaultColWidth` or `@baseColWidth` for a column with no `x:col`).

## Fonts

`createXlsxEditor({ fonts })` takes font sources, or a function that resolves to them. Without `fonts`, the editor loads the packaged substitute faces of `@docx-editor.dev/fonts` (`loadSubstituteFonts` of `@xlsx-editor.dev/core/editor`). `editor.setFonts()` replaces the faces in place and resolves with the failure when the new faces do not load. `editor.fontReport()` lists the faces in force. The editor never fetches a font that a workbook names.

If no packaged substitute face loads, the workbook does not open. `editor.snapshot().status` is `failed`, the failure code is `fonts`, and `fontReport()` is empty. If only some packaged substitute faces load, the workbook opens and its text measures and paints with the faces that loaded. `fontReport().substituteFailures` then lists the faces that did not load.

## Untrusted input

Every value in a workbook is attacker-controlled. The engine treats the file that way:

- The package reader caps the decompressed size, the part count, the XML depth, and the element counts, and refuses paths with `..` or a leading `/`. XML has no DTD and no entity expansion.
- No value from a file reaches HTML: the grid paints into a canvas, and text goes through `textContent`.
- Hyperlink targets (`x:hyperlink`) go through an allowlist of URL schemes. External relationships (`TargetMode="External"`), external workbook links (`x:externalLink`), data connections (`x:connections`), and remote images never load on their own.
- Formulas (`x:f`) are never evaluated: a cell shows its cached `x:v`. Macros (`vbaProject.bin`), embedded objects (`x:oleObjects`), and controls (`x:controls`) are kept inert.
- File-supplied numbers, such as `x:dimension/@ref`, `x:row/@r`, and `x:row/@spans`, never size an allocation or a loop: the engine reads the cells that exist.

## Lanes

Each lane is a package subpath with a declared dependency edge and environment.

| Lane | Environment | May import | Holds |
| --- | --- | --- | --- |
| `contracts` | Browser | Nothing | Declaration-only public types. DOM-free types are in `neutral.ts` |
| `store` | DOM-free | Nothing | Package read and save, the canonical tree, the document store, and its ops |
| `cascade` | DOM-free | `store` | Cell formats, number formats, theme colors, and theme fonts |
| `layout` | DOM-free | `contracts`, `store`, `cascade` | Grid geometry and cell text layout, in points |
| `automation` | DOM-free | `contracts`, `store`, `cascade`, `layout` | The edits that lower to store ops |
| `output` | Browser | `store`, `cascade`, `layout` | Paint from layout records: the sheet window into a canvas, its cell texts as accessible rows, and pictures from package bytes |
| `editor` | Browser | Every other lane | The facade that the adapters render: the workbook editor and its grid surface |

Import a lane through its package subpath, for example `@xlsx-editor.dev/core/store`. Each subpath points at its build in `dist/`.

All chrome CSS and color tokens live in `src/styles/editor.css` (`@xlsx-editor.dev/core/styles/editor.css`). The adapters only import it. The painted grid is not themed.

## License

Apache-2.0. See [LICENSE](LICENSE). The package bundles five ECMA-376 tables. The Ecma copyright notice covers them, not Apache-2.0. The tables are the preset shape geometry, the preset color names, the predefined function names, the built-in number formats, and the default indexed color palette. Other third-party material keeps its own terms. See `THIRD_PARTY_NOTICES.md` in the package.
