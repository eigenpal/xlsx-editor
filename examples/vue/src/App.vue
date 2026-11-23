<script setup lang="ts">
// The Vue demo: it opens the sample workbook of the React demo, or a workbook that you pick, and
// downloads the saved copy. It asks before a pick or a reload drops unsaved edits, as the React
// demo does. Only for local use: no deploy reads this directory.
import { onBeforeUnmount, onMounted, shallowRef, useTemplateRef } from 'vue';
import {
  XlsxEditor,
  type SavedWorkbook,
  type WorkbookKind,
  type XlsxEditorRef,
} from '@xlsx-editor.dev/vue';
import '@xlsx-editor.dev/vue/styles.css';
import { downloadBlob } from '../../vite/src/download.ts';

/** The extension and the media type of each package kind, as the React demo names them. */
const KIND_FILE: Readonly<Record<WorkbookKind, { ext: string; type: string }>> = {
  workbook: {
    ext: '.xlsx',
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  },
  template: {
    ext: '.xltx',
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template',
  },
};

const editor = useTemplateRef<XlsxEditorRef>('editor');
const name = shallowRef('sample');
const error = shallowRef<string>();
// True while a picked file is read: the toolbar refuses Save, so no save names the wrong file.
const opening = shallowRef(false);
// The revision of the last open or save. An edit after it is not saved to a file.
let cleanAt: number | null = null;
// The number of the newest open. An older open that ends later changes nothing.
let opens = 0;
let attached = (): void => {};
const ready = new Promise<void>((resolve) => (attached = resolve));

/** True when the open workbook has an edit after the last open or save. */
function hasUnsavedEdits(): boolean {
  const s = editor.value?.snapshot();
  if (!s || s.status !== 'ready' || cleanAt === null) return false;
  return s.revision !== cleanAt || (s.editing && s.editChanged);
}

/** Reads the bytes of `label` and opens them, unless a newer open started meanwhile. */
async function open(read: () => Promise<ArrayBuffer>, label: string, mine: number) {
  opening.value = true;
  try {
    const buffer = await read();
    await ready;
    if (mine !== opens) return;
    error.value = undefined;
    const result = await editor.value?.load(buffer);
    if (mine !== opens || !result) return;
    name.value = label;
    cleanAt = result.status === 'ready' ? result.revision : null;
  } catch (cause) {
    if (mine === opens) error.value = `${label} did not load: ${(cause as Error).message}`;
  } finally {
    if (mine === opens) opening.value = false;
  }
}

/**
 * Asks before a new workbook replaces unsaved edits. A value typed with no Enter commits first.
 * A commit that the editor refuses stops the open.
 */
function mayReplace(file: string): boolean {
  const instance = editor.value?.getEditor();
  if (instance?.snapshot().editing && !instance.commitEdit().ok) {
    error.value = `Finish or cancel the cell edit before you open ${file}.`;
    return false;
  }
  return (
    !hasUnsavedEdits() ||
    window.confirm(`Open ${file}? The edits to ${name.value} are lost unless you saved them.`)
  );
}

function pick(event: Event) {
  const input = event.target as HTMLInputElement;
  const file = input.files?.[0];
  // Empty the field, so a second pick of the same file opens it again.
  input.value = '';
  if (!file || !mayReplace(file.name)) return;
  const label = file.name.replace(/\.(xlsx|xltx)$/i, '') || 'workbook';
  void open(() => file.arrayBuffer(), label, ++opens);
}

function save(saved: Uint8Array<ArrayBuffer>, result: SavedWorkbook) {
  // Another open ended this workbook during the save: these bytes are not the open workbook.
  if (!result.stillOpen) return;
  const file = KIND_FILE[editor.value?.snapshot().workbookKind ?? 'workbook'];
  downloadBlob(new Blob([saved], { type: file.type }), `${name.value}-edited${file.ext}`);
  cleanAt = result.revision;
}

// Leaving the page or reloading it over unsaved edits asks first.
const warn = (event: BeforeUnloadEvent) => {
  if (hasUnsavedEdits()) event.preventDefault();
};
onMounted(() => window.addEventListener('beforeunload', warn));
onBeforeUnmount(() => window.removeEventListener('beforeunload', warn));

void open(
  async () => {
    const response = await fetch(`${import.meta.env.BASE_URL}sample.xlsx`);
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    return response.arrayBuffer();
  },
  'sample',
  ++opens
);
</script>

<template>
  <main class="demo">
    <header class="demo-bar">
      <h1>XLSX Editor (Vue)</h1>
      <label>
        Open a workbook
        <input type="file" accept=".xlsx,.xltx" @change="pick" />
      </label>
    </header>
    <p v-if="error" role="alert">{{ error }}</p>
    <XlsxEditor
      ref="editor"
      class="demo-editor"
      :opening="opening"
      @ready="attached"
      @save="save"
    />
  </main>
</template>
