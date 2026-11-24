// The save in short tasks: what the bytes hold, and how it refuses.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx } from '../../../../../scripts/fixtures/builder.ts';
import type { FontSource } from '../../layout/index.ts';
import {
  readXlsx,
  sharedStringItems,
  workbookIndex,
  type WorkbookDocumentStore,
} from '../../store/index.ts';
import type { CellEditing } from '../cell-editing.ts';
import { saveInSlices } from '../editor-save.ts';
import type { SavedWorkbook } from '../editor-instance.ts';
import type { EditorSnapshot } from '../editor-model.ts';
import { createXlsxEditor } from '../create-editor.ts';
import type { XlsxEditorInstance } from '../xlsx-editor.ts';
import { runToolbarCommand, toolbarCommandState } from '../toolbar-commands.ts';

const fontsDir = join(
  dirname(Bun.resolveSync('@docx-editor.dev/fonts/package.json', import.meta.dir)),
  'assets'
);
const FONTS: FontSource[] = [
  {
    family: 'Carlito',
    bold: false,
    italic: false,
    bytes: new Uint8Array(readFileSync(join(fontsDir, 'Carlito-Regular.ttf'))),
    origin: 'substitute',
  },
];

let viewers: XlsxEditorInstance[] = [];
afterEach(() => {
  for (const v of viewers) v.destroy();
  viewers = [];
});

/** An editor on a page, as cell editing needs one. */
function mounted(): XlsxEditorInstance {
  const doc = new Window().document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.appendChild(host);
  const viewer = createXlsxEditor({ fonts: FONTS });
  viewers.push(viewer);
  viewer.attach(host);
  return viewer;
}

const file = () =>
  buildXlsx({ sheets: [{ name: 'Data', rows: '<row r="1"><c r="A1"><v>1</v></c></row>' }] });

async function opened(bytes = file()) {
  const viewer = mounted();
  await viewer.load(bytes);
  return viewer;
}

/** A workbook of 40,000 rows: its read and its save take more than one slice. */
function largeFile(name = 'Large'): Uint8Array {
  let rows = '';
  for (let r = 1; r <= 40_000; r++) rows += `<row r="${r}"><c r="A${r}"><v>${r}</v></c></row>`;
  return buildXlsx({ sheets: [{ name, rows }] });
}

/** A promise that settles in a later task, after the tasks that are queued now. */
const nextTask = () =>
  new Promise<void>((resolve) => {
    const channel = new MessageChannel();
    channel.port1.onmessage = () => {
      channel.port1.close();
      resolve();
    };
    channel.port2.postMessage(null);
  });

/** The text of each shared string of a saved workbook. */
function strings(bytes: Uint8Array): string[] {
  const read = readXlsx(bytes);
  if (!read.ok) throw new Error('read');
  const items = sharedStringItems(read.pkg, workbookIndex(read.pkg)!);
  return items.map((item) => JSON.stringify(item.children));
}

describe('saveInTasks', () => {
  test('the bytes are the bytes of save', async () => {
    const viewer = await opened();
    const saved = await viewer.saveInTasks();
    expect(saved?.bytes).toEqual(viewer.save()!);
    expect(saved?.stillOpen).toBe(true);
  });

  test('an edit made while the save runs belongs to the next save', async () => {
    const viewer = await opened(largeFile());
    expect(viewer.enterValue('before', { row: 1, col: 0 }).ok).toBe(true);
    const at = viewer.snapshot().revision;
    let ended = false;
    const saving = viewer.saveInTasks();
    void saving.then(() => (ended = true));
    // The save of a large workbook takes more than one slice: the edit comes during the save.
    await nextTask();
    expect(ended).toBe(false);
    expect(viewer.enterValue('after', { row: 2, col: 0 }).ok).toBe(true);
    const saved = await saving;
    expect(saved?.revision).toBe(at);
    expect(viewer.snapshot().revision).toBeGreaterThan(at);
    const text = strings(saved!.bytes).join();
    expect(text).toContain('before');
    expect(text).not.toContain('after');
    expect(strings(viewer.save()!).join()).toContain('after');
  });

  test('an open cell edit commits at the call, and its revision is in the result', async () => {
    const viewer = await opened();
    expect(viewer.beginEdit('typed', 'enter').ok).toBe(true);
    const saving = viewer.saveInTasks();
    expect(viewer.snapshot().editing).toBe(false);
    const saved = await saving;
    expect(saved?.revision).toBe(viewer.snapshot().revision);
    expect(strings(saved!.bytes).join()).toContain('typed');
  });

  test('a refusal is in the snapshot when the call returns', async () => {
    const viewer = await opened();
    expect(viewer.beginEdit('=1+', 'enter').ok).toBe(true);
    const saving = viewer.saveInTasks();
    expect(viewer.snapshot().refusal?.command).toBe('save');
    expect(await saving).toBeNull();
  });

  test('with no workbook open, the save is refused', async () => {
    const viewer = mounted();
    expect(await viewer.saveInTasks()).toBeNull();
    expect(viewer.snapshot().refusal?.command).toBe('save');
  });

  test('a load during the save marks the bytes as not current', async () => {
    const viewer = await opened();
    const saving = viewer.saveInTasks();
    const loading = viewer.load(file());
    expect((await saving)?.stillOpen).toBe(false);
    await loading;
  });
});

describe('load', () => {
  test('only the newest of two loads shows', async () => {
    const viewer = mounted();
    const other = buildXlsx({
      sheets: [{ name: 'Newest', rows: '<row r="1"><c r="A1"><v>2</v></c></row>' }],
    });
    const first = viewer.load(file());
    const second = viewer.load(other);
    await Promise.all([first, second]);
    expect(viewer.snapshot().sheets.map((s) => s.name)).toEqual(['Newest']);
  });

  test('a load that starts during the read of a large workbook stays', async () => {
    // A first open loads the fonts, so the read of the large workbook starts at once.
    const viewer = await opened();
    const small = buildXlsx({
      sheets: [{ name: 'Small', rows: '<row r="1"><c r="A1"><v>2</v></c></row>' }],
    });
    let firstEnded = false;
    const first = viewer.load(largeFile());
    void first.then(() => (firstEnded = true));
    // The read of the large workbook runs in short tasks, so the second load starts during it.
    await nextTask();
    expect(firstEnded).toBe(false);
    const second = viewer.load(small);
    await Promise.all([first, second]);
    expect(viewer.snapshot().sheets.map((s) => s.name)).toEqual(['Small']);
  });
});

describe('the save of the toolbar', () => {
  test('a save with no open workbook is not-editable, whatever refusal the snapshot holds', () => {
    const snapshot = {
      ...createXlsxEditor().snapshot(),
      status: 'failed',
      refusal: { code: 'formula-syntax' },
    } as EditorSnapshot;
    const editor = {
      saveInTasks: async () => null,
      snapshot: () => snapshot,
    } as unknown as XlsxEditorInstance;
    const r = runToolbarCommand(editor, 'file.save', undefined, { save: () => {} });
    expect(r.ok ? null : r.refusal.code).toBe('not-editable');
  });

  test('the host gets the bytes and the revision that they hold', async () => {
    const viewer = await opened();
    const at = viewer.snapshot().revision;
    const got = await new Promise<{ bytes: Uint8Array; revision: number }>((resolve) => {
      const save = (bytes: Uint8Array, saved: { revision: number }) =>
        resolve({ bytes, revision: saved.revision });
      expect(runToolbarCommand(viewer, 'file.save', undefined, { save }).ok).toBe(true);
    });
    expect(got.revision).toBe(at);
    expect(got.bytes.byteLength).toBeGreaterThan(0);
  });

  test('a press during a running save gives the host one delivery', async () => {
    const viewer = await opened();
    let calls = 0;
    const host = { save: () => (calls += 1) };
    expect(runToolbarCommand(viewer, 'file.save', undefined, host).ok).toBe(true);
    expect(runToolbarCommand(viewer, 'file.save', undefined, host).ok).toBe(true);
    await new Promise((resolve) => setTimeout(resolve, 50));
    expect(calls).toBe(1);
    // The next press after the delivery starts a new save.
    expect(runToolbarCommand(viewer, 'file.save', undefined, host).ok).toBe(true);
    await new Promise((resolve) => setTimeout(resolve, 50));
    expect(calls).toBe(2);
  });

  test('a new host object joins the running save, and the newest host gets the bytes', async () => {
    const viewer = await opened();
    const calls: string[] = [];
    runToolbarCommand(viewer, 'file.save', undefined, { save: () => calls.push('first') });
    runToolbarCommand(viewer, 'file.save', undefined, { save: () => calls.push('second') });
    await new Promise((resolve) => setTimeout(resolve, 50));
    expect(calls).toEqual(['second']);
  });

  test('a press after an edit during a save starts a save of the edit', async () => {
    const viewer = await opened();
    const revisions: number[] = [];
    const host = {
      save: (_: Uint8Array, saved: { revision: number }) => revisions.push(saved.revision),
    };
    runToolbarCommand(viewer, 'file.save', undefined, host);
    expect(viewer.beginEdit('2', 'enter').ok).toBe(true);
    expect(viewer.commitEdit().ok).toBe(true);
    const edited = viewer.snapshot().revision;
    runToolbarCommand(viewer, 'file.save', undefined, host);
    await new Promise((resolve) => setTimeout(resolve, 50));
    expect(revisions).toEqual([edited - 1, edited]);
  });

  test('a press on a workbook that replaced a saving one saves the new workbook', async () => {
    const viewer = await opened();
    const current: boolean[] = [];
    const host = {
      save: (_: Uint8Array, saved: { stillOpen: boolean }) => current.push(saved.stillOpen),
    };
    runToolbarCommand(viewer, 'file.save', undefined, host);
    await viewer.load(file());
    runToolbarCommand(viewer, 'file.save', undefined, host);
    await new Promise((resolve) => setTimeout(resolve, 50));
    expect(current).toContain(true);
  });

  test('a host hears of each started save once, and not of a joined press', async () => {
    const viewer = await opened();
    const started: Promise<unknown>[] = [];
    const host = { save: () => {}, saveStart: (s: Promise<unknown>) => started.push(s) };
    runToolbarCommand(viewer, 'file.save', undefined, host);
    runToolbarCommand(viewer, 'file.save', undefined, host);
    expect(started.length).toBe(1);
    await started[0];
  });

  test('a finished save of the same revision gives its own copy of the same bytes', async () => {
    const viewer = await opened();
    const first = await viewer.saveInTasks();
    const again = await viewer.saveInTasks();
    expect(again?.bytes).toEqual(first!.bytes);
    expect(again?.bytes.buffer).not.toBe(first!.bytes.buffer);
    again!.bytes.fill(0);
    expect((await viewer.saveInTasks())?.bytes).toEqual(first!.bytes);
  });

  test('a finished save of the same revision serializes nothing again', async () => {
    const read = readXlsx(file());
    if (!read.ok) throw new Error('read');
    let serialized = 0;
    const store = {
      revision: 1,
      get package() {
        serialized += 1;
        return read.pkg;
      },
    } as unknown as WorkbookDocumentStore;
    const edits = {
      commitOpenEdit: () => ({ ok: true }),
      saved: () => null,
      refuseSave: () => null,
    } as unknown as CellEditing;
    expect(await saveInSlices(store, edits, null, () => true)).not.toBeNull();
    expect(await saveInSlices(store, edits, null, () => true)).not.toBeNull();
    expect(serialized).toBe(1);
  });

  test('repeated presses during a load publish nothing after the first refusal', async () => {
    const viewer = await opened();
    const loading = viewer.load(file());
    const host = { save: () => {} };
    runToolbarCommand(viewer, 'file.save', undefined, host);
    let published = 0;
    const stop = viewer.subscribe(() => (published += 1));
    runToolbarCommand(viewer, 'file.save', undefined, host);
    runToolbarCommand(viewer, 'file.save', undefined, host);
    expect(published).toBe(0);
    stop();
    await loading;
  });

  test('an edit after a finished save gives new bytes', async () => {
    const viewer = await opened();
    const first = await viewer.saveInTasks();
    expect(viewer.enterValue('later', { row: 1, col: 0 }).ok).toBe(true);
    const next = await viewer.saveInTasks();
    expect(next?.revision).toBeGreaterThan(first!.revision);
    expect(strings(next!.bytes).join()).toContain('later');
  });

  test('older bytes never reach the host after newer bytes', async () => {
    const ends: ((saved: SavedWorkbook) => void)[] = [];
    const editor = {
      snapshot: () => ({ status: 'ready', refusal: null }),
      saveInTasks: () => new Promise<SavedWorkbook>((resolve) => ends.push(resolve)),
    } as unknown as XlsxEditorInstance;
    const delivered: number[] = [];
    const host = { save: (_b: Uint8Array, saved: SavedWorkbook) => delivered.push(saved.revision) };
    runToolbarCommand(editor, 'file.save', undefined, host);
    runToolbarCommand(editor, 'file.save', undefined, host);
    const bytes = new Uint8Array(1);
    ends[1]!({ bytes, revision: 2, stillOpen: true });
    await new Promise((resolve) => setTimeout(resolve, 0));
    ends[0]!({ bytes, revision: 1, stillOpen: true });
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(delivered).toEqual([2]);
  });

  test('a refused save does not tell the host that a save started', async () => {
    const viewer = await opened();
    expect(viewer.beginEdit('=1+', 'enter').ok).toBe(true);
    let started = 0;
    const host = { save: () => {}, saveStart: () => (started += 1) };
    expect(runToolbarCommand(viewer, 'file.save', undefined, host).ok).toBe(false);
    expect(started).toBe(0);
  });

  test('a destroyed editor marks the bytes of a running save as not current', async () => {
    const viewer = await opened();
    const saving = viewer.saveInTasks();
    viewer.destroy();
    expect((await saving)?.stillOpen).toBe(false);
  });

  test('a refused save calls the host never, and returns the refusal', async () => {
    const viewer = await opened();
    expect(viewer.beginEdit('=1+', 'enter').ok).toBe(true);
    let calls = 0;
    const r = runToolbarCommand(viewer, 'file.save', undefined, { save: () => (calls += 1) });
    expect(r.ok).toBe(false);
    await new Promise((resolve) => setTimeout(resolve, 5));
    expect(calls).toBe(0);
  });

  test('a press while the host opens a workbook starts no save work, and the control says why', async () => {
    const viewer = await opened();
    // Each start of save work, also one that a refusal would end later.
    let works = 0;
    const real = viewer.saveInTasks.bind(viewer);
    viewer.saveInTasks = () => {
      works += 1;
      return real();
    };
    let opening = true;
    let delivered = 0;
    let pending: Promise<SavedWorkbook | null> | null = null;
    const host = {
      save: () => (delivered += 1),
      saveStart: (saving: Promise<SavedWorkbook | null>) => (pending = saving),
      opening: () => opening,
    };
    expect(toolbarCommandState(viewer.snapshot(), 'file.save', host)).toMatchObject({
      enabled: false,
      disabledReason: { key: 'disabledReason.loading' },
    });
    expect(runToolbarCommand(viewer, 'file.save', undefined, host)).toEqual({
      ok: false,
      refusal: { code: 'loading', command: 'save' },
    });
    // The editor keeps the refusal, so the live region says why.
    expect(viewer.snapshot().refusal).toMatchObject({ code: 'loading', command: 'save' });
    await nextTask();
    expect([works, delivered, pending]).toEqual([0, 0, null]);
    // When the open ends, the same host saves.
    opening = false;
    expect(toolbarCommandState(viewer.snapshot(), 'file.save', host).enabled).toBe(true);
    expect(runToolbarCommand(viewer, 'file.save', undefined, host).ok).toBe(true);
    await pending;
    expect([works, delivered]).toEqual([1, 1]);
  });

  test('two presses that join one save get bytes of their own', async () => {
    const viewer = await opened();
    const [first, joined] = await Promise.all([viewer.saveInTasks(), viewer.saveInTasks()]);
    first!.bytes[0] = 0;
    expect(joined!.bytes[0]).toBe(0x50);
  });

  test('a joined press that changes its bytes does not change a later save', async () => {
    const viewer = await opened();
    const [first, joined] = await Promise.all([viewer.saveInTasks(), viewer.saveInTasks()]);
    joined!.bytes.fill(0);
    expect((await viewer.saveInTasks())?.bytes).toEqual(first!.bytes);
  });

  /** A store whose save passes the save rules, and fails later, in the steps of the write. */
  function failsInTheWrite(): WorkbookDocumentStore {
    const read = readXlsx(file());
    if (!read.ok) throw new Error('read');
    const parts = read.pkg.parts as Map<string, unknown>;
    parts.set('xl/media/broken.bin', {
      kind: 'binary',
      partName: 'xl/media/broken.bin',
      bytes: null,
    });
    return { revision: 1, package: read.pkg } as unknown as WorkbookDocumentStore;
  }

  /** The editing stand-in of a save, which keeps each refusal. */
  function refusals(): { edits: CellEditing; seen: unknown[] } {
    const seen: unknown[] = [];
    const edits = {
      commitOpenEdit: () => ({ ok: true }),
      saved: () => {},
      refuseSave: (reason: unknown) => {
        seen.push(reason);
        return null;
      },
    } as unknown as CellEditing;
    return { edits, seen };
  }

  test('a write that fails late refuses the save of the open workbook', async () => {
    const { edits, seen } = refusals();
    expect(await saveInSlices(failsInTheWrite(), edits, null, () => true)).toBeNull();
    expect(seen).toEqual([expect.objectContaining({ code: 'save-failed' })]);
  });

  test('a write that fails late refuses nothing after another workbook opened', async () => {
    const { edits, seen } = refusals();
    expect(await saveInSlices(failsInTheWrite(), edits, null, () => false)).toBeNull();
    expect(seen).toEqual([]);
  });

  test('a workbook that cannot serialize refuses the save with its own code', async () => {
    const refusals: unknown[] = [];
    const edits = {
      commitOpenEdit: () => ({ ok: true }),
      saved: () => {},
      refuseSave: (reason: unknown) => {
        refusals.push(reason);
        return null;
      },
    } as unknown as CellEditing;
    const store = { revision: 1, package: null } as unknown as WorkbookDocumentStore;
    expect(await saveInSlices(store, edits, null, () => true)).toBeNull();
    expect(refusals).toEqual([expect.objectContaining({ code: 'save-failed' })]);
  });
});
