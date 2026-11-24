// The open and save logic of the Vue demo (`session.ts`), under Bun, over a stand-in editor: the
// editor gets the bytes of the file and Save downloads the bytes that the editor saves, an open
// asks before unsaved edits are lost and only after the file passes its checks, a newer pick or a
// no stops an earlier open, a save waits for an open, and a name that the user gives in the top
// bar names the download until the next open.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import type { EditorSnapshot, SavedWorkbook } from '@xlsx-editor.dev/core/editor';
import { fixtureFile } from '../../../scripts/fixtures/catalog.ts';
import { DemoSession, MAX_FILE_BYTES, type PickedFile, type SessionEditor } from './session.ts';

const ROOT = join(import.meta.dir, '../../..');
const sampleBytes = () =>
  new Uint8Array(readFileSync(join(ROOT, 'examples/vite/public/sample.xlsx')));
const fixtureBytes = (name: string) =>
  new Uint8Array(readFileSync(join(ROOT, 'fixtures', fixtureFile(name))));

/** A stand-in editor: it keeps the loaded bytes and a revision that a test moves. */
class FakeEditor implements SessionEditor {
  readonly loaded: Uint8Array[] = [];
  revision = 0;
  status: EditorSnapshot['status'] = 'empty';
  editing = false;
  /** Whether a commit of the open cell edit succeeds. */
  commits = true;

  async load(bytes: Uint8Array) {
    this.status = 'loading';
    await new Promise((done) => setTimeout(done, 0));
    this.loaded.push(new Uint8Array(bytes));
    this.revision = 0;
    this.status = 'ready';
    return this.snapshot();
  }

  snapshot(): EditorSnapshot {
    const { status, revision, editing } = this;
    return {
      status,
      revision,
      editing,
      editChanged: false,
      activeSheet: 0,
      selection: null,
      workbookKind: 'workbook',
    } as unknown as EditorSnapshot;
  }

  async saveInTasks(): Promise<SavedWorkbook | null> {
    const bytes = this.loaded.at(-1);
    if (!bytes) return null;
    return { bytes: new Uint8Array(bytes), revision: this.revision, stillOpen: true };
  }

  getEditor() {
    const commitEdit = () => {
      if (!this.commits) return { ok: false };
      this.editing = false;
      return { ok: true };
    };
    return { commitEdit } as never;
  }
}

/** A session over a stand-in editor, with the questions and downloads that it gave. */
function session() {
  const editor = new FakeEditor();
  const asked: string[] = [];
  const downloads: { name: string; bytes: Promise<Uint8Array> }[] = [];
  const answers = { accept: true, onAsk: () => {} };
  // The page's `ready`, which a test can hold: `waiting` turns true when an open reads it.
  const gate = { ready: Promise.resolve(), waiting: false };
  const demo = new DemoSession({
    editor: () => editor,
    get ready() {
      gate.waiting = true;
      return gate.ready;
    },
    confirm: (text) => {
      asked.push(text);
      answers.onAsk();
      return answers.accept;
    },
    download: (blob, name) =>
      void downloads.push({ name, bytes: blob.arrayBuffer().then((b) => new Uint8Array(b)) }),
    focusGrid: () => {},
  });
  return { demo, editor, asked, downloads, answers, gate };
}

/** Holds the page's `ready` of `gate` until the returned call. */
function holdReady(gate: { ready: Promise<void>; waiting: boolean }): () => void {
  let release = () => {};
  gate.ready = new Promise<void>((resolve) => (release = resolve));
  gate.waiting = false;
  return release;
}

/** Waits until `check` holds, letting the session's tasks run. */
async function until(check: () => boolean): Promise<void> {
  for (let i = 0; i < 2000 && !check(); i++) await new Promise((done) => setTimeout(done, 1));
  if (!check()) throw new Error('the session did not reach the step');
}

/** A picked file whose read waits for the returned call. */
function heldFile(name: string, bytes: Uint8Array): { file: PickedFile; release: () => void } {
  let release = () => {};
  const done = new Promise<void>((resolve) => (release = resolve));
  return {
    file: { name, size: bytes.length, arrayBuffer: async () => (await done, bytes.slice().buffer) },
    release,
  };
}

const picked = (name: string, bytes: Uint8Array): PickedFile => ({
  name,
  size: bytes.length,
  arrayBuffer: async () => bytes.slice().buffer,
});

describe('the Vue demo session', () => {
  test('opens the bytes of the sample as they are, and Save downloads the bytes that the editor saves', async () => {
    const { demo, editor, downloads } = session();
    const sample = sampleBytes();
    await demo.openSample(async () => sample.slice().buffer);
    expect(demo.state.label).toBe('sample');
    expect(editor.loaded[0]).toEqual(sample);
    await demo.saveNow();
    expect(downloads.map((d) => d.name)).toEqual(['sample-edited.xlsx']);
    expect(await downloads[0]!.bytes).toEqual(sample);
  });

  test('a name that the user gives names the download, and the next open drops it', async () => {
    const { demo, downloads } = session();
    demo.rename('Ignored');
    expect(demo.state.chosen).toBe(null);
    await demo.openSample(async () => sampleBytes().slice().buffer);
    // The same name as before, and an empty name, keep the name of the file.
    demo.rename('Sample workbook');
    demo.rename('   ');
    expect(demo.state.chosen).toBe(null);
    demo.rename(' Budget 2026.XLSX ');
    expect(demo.state.chosen).toBe('Budget 2026');
    demo.rename('');
    expect(demo.state.chosen).toBe('Budget 2026');
    await demo.saveNow();
    await demo.openFile(picked('basic-values.xlsx', fixtureBytes('basic-values')));
    expect(demo.state.chosen).toBe(null);
    await demo.saveNow();
    expect(downloads.map((d) => d.name)).toEqual(['Budget 2026.xlsx', 'basic-values-edited.xlsx']);
  });

  test('a sample that does not download names no workbook and says what to do', async () => {
    const { demo } = session();
    await demo.openSample(async () => {
      throw new TypeError('the network failed');
    });
    expect(demo.state.label).toBeNull();
    expect(demo.state.error).toBe(
      'Could not open the sample workbook: the file did not download. Open a file to go on.'
    );
  });

  test('asks before unsaved edits are lost, only after the file passes its checks', async () => {
    const { demo, editor, asked, answers } = session();
    await demo.openSample(async () => sampleBytes().buffer);
    editor.revision = 3;
    expect(demo.hasUnsavedEdits()).toBe(true);
    // A file that does not open asks nothing and changes nothing.
    await demo.openFile(picked('broken.xlsx', new TextEncoder().encode('not a workbook')));
    expect(asked).toEqual([]);
    expect(demo.state.error).toContain('Could not open broken.xlsx');
    expect(demo.state.error).toContain('The workbook before it stays open.');
    expect(editor.loaded.length).toBe(1);
    // A no keeps the workbook and the edit.
    answers.accept = false;
    await demo.openFile(picked('basic-values.xlsx', fixtureBytes('basic-values')));
    expect(asked.length).toBe(1);
    expect(asked[0]).toContain('The edits to the sample workbook are lost');
    expect(editor.loaded.length).toBe(1);
    expect(demo.state.label).toBe('sample');
    expect(demo.hasUnsavedEdits()).toBe(true);
    // A yes opens the file.
    answers.accept = true;
    await demo.openFile(picked('basic-values.xlsx', fixtureBytes('basic-values')));
    expect(editor.loaded.length).toBe(2);
    expect(demo.state.label).toBe('basic-values.xlsx');
    expect(demo.hasUnsavedEdits()).toBe(false);
  });

  test('refuses a file over 64 MiB before it reads it', async () => {
    const { demo, editor } = session();
    let read = false;
    await demo.openFile({
      name: 'huge.xlsx',
      size: MAX_FILE_BYTES + 1,
      arrayBuffer: async () => ((read = true), new ArrayBuffer(0)),
    });
    expect(read).toBe(false);
    expect(demo.state.error).toBe('Could not open huge.xlsx: the file is larger than 64 MiB.');
    expect(editor.loaded.length).toBe(0);
  });

  test('a newer pick stops an earlier pick that is still in its read', async () => {
    const { demo, editor } = session();
    await demo.openSample(async () => sampleBytes().buffer);
    const earlier = heldFile('earlier.xlsx', fixtureBytes('structure'));
    const first = demo.openFile(earlier.file);
    await demo.openFile(picked('basic-values.xlsx', fixtureBytes('basic-values')));
    earlier.release();
    await first;
    expect(editor.loaded.length).toBe(2);
    expect(demo.state.label).toBe('basic-values.xlsx');
  });

  test('a no to a newer pick also stops an earlier pick that the user accepted', async () => {
    const { demo, editor, asked, answers, gate } = session();
    await demo.openSample(async () => sampleBytes().buffer);
    editor.revision = 5;
    // The earlier pick asks, is accepted, and waits before its load.
    const release = holdReady(gate);
    const earlier = demo.openFile(picked('earlier.xlsx', fixtureBytes('structure')));
    await until(() => gate.waiting);
    expect(asked.length).toBe(1);
    // The newer pick asks while the earlier pick is past its question, and the answer is no.
    answers.accept = false;
    await demo.openFile(picked('newer.xlsx', fixtureBytes('basic-values')));
    expect(asked.length).toBe(2);
    release();
    await earlier;
    // Neither pick loaded: the edit stays.
    expect(editor.loaded.length).toBe(1);
    expect(demo.state.label).toBe('sample');
    expect(demo.hasUnsavedEdits()).toBe(true);
  });

  test('a cell edit that does not commit stops the open and an earlier accepted pick', async () => {
    const { demo, editor, asked, gate } = session();
    await demo.openSample(async () => sampleBytes().buffer);
    editor.revision = 5;
    const release = holdReady(gate);
    const earlier = demo.openFile(picked('earlier.xlsx', fixtureBytes('structure')));
    await until(() => gate.waiting);
    editor.editing = true;
    editor.commits = false;
    await demo.openFile(picked('newer.xlsx', fixtureBytes('basic-values')));
    expect(demo.state.error).toBe(
      'Could not open newer.xlsx: finish or cancel the cell edit first.'
    );
    expect(asked.length).toBe(1);
    release();
    await earlier;
    expect(editor.loaded.length).toBe(1);
    expect(demo.state.label).toBe('sample');
  });

  test('a top bar Save and a toolbar Save during it download one copy', async () => {
    const { demo, editor, downloads } = session();
    await demo.openSample(async () => sampleBytes().buffer);
    const saved = (await editor.saveInTasks())!;
    await Promise.all([demo.saveNow(), demo.save(saved.bytes, saved), demo.saveNow()]);
    expect(downloads.length).toBe(1);
    // A second save of the same revision right after the first is a repeat, as a double click.
    await demo.saveNow();
    expect(downloads.length).toBe(1);
  });

  test('a save during an open is refused and downloads nothing', async () => {
    const { demo, downloads } = session();
    await demo.openSample(async () => sampleBytes().buffer);
    const later = heldFile('basic-values.xlsx', fixtureBytes('basic-values'));
    const open = demo.openFile(later.file);
    await demo.saveNow();
    expect(demo.state.error).toBe(
      'Could not save: a workbook is opening. Try again when it shows.'
    );
    expect(downloads.length).toBe(0);
    later.release();
    await open;
    expect(demo.state.label).toBe('basic-values.xlsx');
  });
});
