// After pptx-editor/packages/core/src/__tests__/consumer.test-d.ts at e22f3f3b (EigenPal), which
// follows docx-editor/packages/core/src/__tests__/consumer.test-d.ts. Differences: the checks name
// the XLSX Editor public API, and the comment asks for a changeset only after the first publish,
// because AGENTS.md ("Releasing") says that no pull request has a changeset until then.

/**
 * Consumer-side type test.
 *
 * The package typechecking on its own proves only that the declarations are internally
 * consistent. It does not prove that a host can create an editor, read a snapshot, or pass the
 * editor that the React or Vue adapter hands out to code typed against core. Every check below
 * imports a public entry point by its package name, the way a host writes it. Two commands compile
 * this file:
 *
 * - `bun run typecheck`, through `tsconfig.consumer.json`, which maps the package names to the
 *   sources. Fast, and needs no build.
 * - `bun run check:consumer-types`, after `bun run build:packages`, which resolves the names
 *   through each `package.json` `exports` to the built `dist/*.d.ts`, as a host does. It sees
 *   breaks that exist only in the build (CI runs it in the build job).
 *
 * When this file fails to compile, the public types changed. `Unused '@ts-expect-error'
 * directive` means that a value the comment above it says is refused is now accepted. If the
 * change is intended, update the check here. After the first publish, also add a changeset:
 * `minor` for an addition, `major` for a break.
 */

import { createXlsxEditor as createFromRoot } from '@xlsx-editor.dev/core';
import {
  createXlsxEditor,
  type EditorSnapshot,
  type GridPoint,
  type XlsxEditorInstance,
  type XlsxEditorOptions,
} from '@xlsx-editor.dev/core/editor';
import { readXlsx, saveXlsx } from '@xlsx-editor.dev/core/store';
import { pointsToPixels, setHarfBuzzWasmUrl } from '@xlsx-editor.dev/core/layout';
import type { GridRange } from '@xlsx-editor.dev/core/contracts';
import {
  XlsxEditor,
  useEditorState,
  type XlsxEditorProps,
  type XlsxEditorRef,
} from '@xlsx-editor.dev/react';
import {
  XlsxEditor as VueXlsxEditor,
  useEditorState as useVueEditorState,
  type XlsxEditorProps as VueXlsxEditorProps,
  type XlsxEditorRef as VueXlsxEditorRef,
} from '@xlsx-editor.dev/vue';
import { createT, en, type Translations } from '@xlsx-editor.dev/i18n';
import englishOnly from '@xlsx-editor.dev/i18n/en';

// The root entry and the editor lane export one function.
const sameCreate: typeof createXlsxEditor = createFromRoot;
void sameCreate;

// Create an editor with options, open bytes, and save.
const options: XlsxEditorOptions = { fonts: [] };
const editor: XlsxEditorInstance = createXlsxEditor(options);
const opened: Promise<EditorSnapshot> = editor.load(new Uint8Array());
void opened;
const openedBuffer: Promise<EditorSnapshot> = editor.load(new ArrayBuffer(0));
void openedBuffer;
const at: GridPoint = { row: 1, col: 1 };
const moved: boolean = editor.select(at);
void moved;
const saved: Uint8Array<ArrayBuffer> | null = editor.save();
void saved;

// @ts-expect-error A grid point needs both a row and a column.
editor.select({ row: 1 });

// The store reads and writes a package without an editor.
const read = readXlsx(new Uint8Array());
if (read.ok) {
  const bytes: Uint8Array<ArrayBuffer> = saveXlsx(read.pkg);
  void bytes;
}

// Layout helpers.
const pixels: number = pointsToPixels(12);
void pixels;
setHarfBuzzWasmUrl('/static/harfbuzz.wasm');

const range: GridRange | null = null;
void range;

// The React adapter hands out the engine's own instance type: one copy of core.
const props: XlsxEditorProps = {
  document: new Uint8Array(),
  mode: 'view',
  onReady: (ready: XlsxEditorInstance) => void ready.snapshot(),
  onSave: (bytes: Uint8Array<ArrayBuffer>) => void bytes,
};
void props;
const bufferProps: XlsxEditorProps = { document: new ArrayBuffer(0) };
void bufferProps;
// @ts-expect-error The mode is 'edit' or 'view'.
const badMode: XlsxEditorProps = { mode: 'present' };
void badMode;

declare const ref: XlsxEditorRef;
const fromRef: XlsxEditorInstance | null = ref.getEditor();
void fromRef;
const snapshot: EditorSnapshot = ref.snapshot();
void snapshot;
void XlsxEditor;
void useEditorState;

// The Vue adapter hands out the same engine instance and the same bytes: one copy of core.
const vueProps: VueXlsxEditorProps = { document: new Uint8Array(), mode: 'view' };
void vueProps;
// @ts-expect-error The mode is 'edit' or 'view'.
const vueBadMode: VueXlsxEditorProps = { mode: 'present' };
void vueBadMode;
declare const vueRef: VueXlsxEditorRef;
const fromVueRef: XlsxEditorInstance | null = vueRef.getEditor();
void fromVueRef;
const vueSaved: Uint8Array<ArrayBuffer> | null = vueRef.save();
void vueSaved;
void VueXlsxEditor;
void useVueEditorState;

// The shared strings: a partial locale and the English catalog.
const partial: Translations = {};
void partial;
const t = createT(en);
void t;
void englishOnly;
