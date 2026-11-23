// Typed editor events (`editor.on`): each event fires when its part of the snapshot moves,
// and stays silent for the rest. `change` follows the content revision and settled loads;
// `selectionChange` the shown sheet and the selection; `rejection` each new refusal object;
// `openFailure` each failed open. Nothing fires for the state at subscribe time.

import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import {
  createXlsxEditor,
  type EditorSnapshot,
  type DocumentChange,
  type XlsxEditorInstance,
} from '../index.ts';
import { flushEditor } from './flush-editor.ts';
import type { FontSource } from '../../layout/index.ts';

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

let editors: XlsxEditorInstance[] = [];
afterEach(() => {
  for (const e of editors) e.destroy();
  editors = [];
});

const workbook = () =>
  buildXlsx({
    sheets: [
      { name: 'S', rows: row(1, [c('A1', 1), c('B1', 2)]) + row(2, [c('A2', 3)]) },
      { name: 'T', rows: row(1, [c('A1', 9)]) },
    ],
  });

async function mount(): Promise<XlsxEditorInstance> {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.appendChild(host);
  const editor = createXlsxEditor({ fonts: FONTS, translate: (k) => k });
  editors.push(editor);
  editor.attach(host);
  const scroller = host.querySelector('.xlsx-scroller') as HTMLElement;
  Object.defineProperty(scroller, 'clientWidth', { get: () => 800 });
  Object.defineProperty(scroller, 'clientHeight', { get: () => 400 });
  await editor.load(workbook());
  flushEditor(editor);
  return editor;
}

/** Record every event of `editor`, each with the arguments of its handler. */
function record(editor: XlsxEditorInstance) {
  const log: string[] = [];
  const changes: DocumentChange[] = [];
  const offs = [
    editor.on('change', (change) => {
      log.push('change');
      changes.push(change);
    }),
    editor.on('selectionChange', () => log.push('selectionChange')),
    editor.on('rejection', (refusal) => log.push(`rejection:${refusal.code}`)),
    editor.on('openFailure', (failure) => log.push(`openFailure:${failure.code}`)),
  ];
  const take = () => log.splice(0);
  return { take, changes, off: () => offs.forEach((off) => off()) };
}

describe('editor.on', () => {
  test('nothing fires at subscribe, or for a notification with no change', async () => {
    const editor = await mount();
    const events = record(editor);
    expect(events.take()).toEqual([]);
    // A selection of the cell that is already selected publishes nothing new to report.
    editor.select({ row: 0, col: 0 });
    expect(events.take()).toEqual([]);
  });

  test('an edit fires change with no source; undo and redo fire it too', async () => {
    const editor = await mount();
    const events = record(editor);
    expect(editor.enterValue('7', { row: 0, col: 0 }).ok).toBe(true);
    expect(events.take()).toContain('change');
    expect(events.changes.at(-1)).toEqual({ revision: editor.snapshot().revision });
    expect(editor.undo().ok).toBe(true);
    expect(events.take()).toContain('change');
    expect(editor.redo().ok).toBe(true);
    expect(events.take()).toContain('change');
  });

  test('a selection change fires selectionChange and not change', async () => {
    const editor = await mount();
    const events = record(editor);
    editor.select({ row: 1, col: 1 });
    expect(events.take()).toEqual(['selectionChange']);
    editor.select({ row: 1, col: 1 }, { row: 2, col: 3 });
    expect(events.take()).toEqual(['selectionChange']);
  });

  test('a sheet change fires selectionChange and not change', async () => {
    const editor = await mount();
    const events = record(editor);
    expect(editor.setActiveSheet(1)).toBe(true);
    expect(events.take()).toEqual(['selectionChange']);
  });

  test('a zoom change fires neither change nor selectionChange', async () => {
    const editor = await mount();
    const events = record(editor);
    editor.setZoom(1.5);
    editor.zoomStep(1);
    expect(editor.snapshot().zoom).not.toBe(1);
    expect(events.take()).toEqual([]);
  });

  test('an open cell edit and its typed text do not fire change', async () => {
    const editor = await mount();
    const events = record(editor);
    expect(editor.beginEdit('4').ok).toBe(true);
    expect(editor.setEditText('42').ok).toBe(true);
    expect(events.take()).toEqual([]);
    editor.cancelEdit();
    expect(events.take()).toEqual([]);
  });

  test('a refusal fires rejection once; the same refusal object does not fire again', async () => {
    const editor = await mount();
    const events = record(editor);
    expect(editor.setMode('view')).toBe(true);
    expect(events.take()).toEqual([]);
    expect(editor.enterValue('5').ok).toBe(false);
    expect(events.take()).toEqual(['rejection:view-mode']);
    const refusal = editor.snapshot().refusal;
    // Other notifications keep the refusal in the snapshot: no new rejection.
    editor.setZoom(2);
    expect(editor.snapshot().refusal).toBe(refusal);
    expect(events.take()).toEqual([]);
    // A command refused again is a new refused command.
    expect(editor.enterValue('6').ok).toBe(false);
    expect(events.take()).toEqual(['rejection:view-mode']);
    // A dismissed refusal fires nothing.
    editor.dismissRefusal();
    expect(events.take()).toEqual([]);
  });

  test('a command state refusal does not fire rejection', async () => {
    const editor = await mount();
    const events = record(editor);
    // A whole-sheet selection gives `structureRefusal`, a command state, not a refused command.
    editor.select({ row: 0, col: 0 }, { row: 1048575, col: 16383 });
    expect(editor.snapshot().structureRefusal).not.toBeNull();
    expect(events.take()).toEqual(['selectionChange']);
  });

  test('a failed open fires openFailure and change with the load source', async () => {
    const editor = await mount();
    const events = record(editor);
    await editor.load(new Uint8Array([1, 2, 3, 4]));
    expect(editor.snapshot().status).toBe('failed');
    const log = events.take();
    expect(log).toContain('openFailure:package');
    expect(log.filter((e) => e === 'change')).toHaveLength(1);
    expect(events.changes.at(-1)?.source).toBe('load');
    // A later notification keeps the same failure: no new openFailure.
    editor.setZoom(2);
    expect(events.take()).not.toContain('openFailure:package');
  });

  test('a load fires change once with the load source; close fires it with close', async () => {
    const editor = await mount();
    expect(editor.enterValue('8', { row: 0, col: 0 }).ok).toBe(true);
    const events = record(editor);
    await editor.load(workbook());
    expect(events.take().filter((e) => e === 'change')).toHaveLength(1);
    expect(events.changes.at(-1)).toEqual({ source: 'load', revision: 0 });
    editor.close();
    expect(events.take()).toContain('change');
    expect(events.changes.at(-1)?.source).toBe('close');
  });

  test('a load that drops an open edit fires rejection once, not again when it settles', async () => {
    const editor = await mount();
    expect(editor.beginEdit('typed').ok).toBe(true);
    const events = record(editor);
    await editor.load(workbook());
    expect(editor.snapshot().refusal?.code).toBe('edit-dropped');
    expect(events.take().filter((e) => e.startsWith('rejection'))).toEqual([
      'rejection:edit-dropped',
    ]);
  });

  test('a refusal during the load leaves one edit-dropped event after it settles', async () => {
    const editor = await mount();
    expect(editor.beginEdit('typed').ok).toBe(true);
    const events = record(editor);
    const loading = editor.load(workbook());
    expect(editor.snapshot().status).toBe('loading');
    expect(editor.enterValue('5').ok).toBe(false);
    expect(editor.enterValue('6').ok).toBe(false);
    await loading;
    expect(editor.snapshot().refusal?.code).toBe('edit-dropped');
    // One `loading` refusal for the load, however many commands it refused.
    expect(events.take().filter((e) => e.startsWith('rejection'))).toEqual([
      'rejection:edit-dropped',
      'rejection:loading',
    ]);
  });

  test('a listener added during the load gets no edit-dropped for a drop before it', async () => {
    const editor = await mount();
    expect(editor.beginEdit('typed').ok).toBe(true);
    const loading = editor.load(workbook());
    // A `loading` refusal hides the drop in the snapshot when the listener subscribes.
    expect(editor.enterValue('5').ok).toBe(false);
    expect(editor.snapshot().refusal?.code).toBe('loading');
    const events = record(editor);
    await loading;
    expect(editor.snapshot().refusal?.code).toBe('edit-dropped');
    expect(events.take().filter((e) => e.startsWith('rejection'))).toEqual([]);
  });

  describe('a handler that throws', () => {
    const reported: unknown[] = [];
    const host = globalThis as { reportError?: (error: unknown) => void };
    let original: typeof host.reportError;
    beforeEach(() => {
      reported.length = 0;
      original = host.reportError;
      host.reportError = (error) => reported.push(error);
    });
    afterEach(() => {
      host.reportError = original;
    });

    test('does not stop a second handler, a subscriber, or the repaint', async () => {
      const editor = await mount();
      // Let the frame of the open run, so a new frame shows that the edit scheduled a repaint.
      await new Promise((resolve) => setTimeout(resolve, 50));
      const frame = () => (editor as unknown as { frame: number | null }).frame;
      expect(frame()).toBeNull();
      const boom = new Error('handler');
      editor.on('change', () => {
        throw boom;
      });
      const second: number[] = [];
      editor.on('change', (change) => second.push(change.revision));
      let notified = 0;
      editor.subscribe(() => (notified += 1));
      expect(editor.enterValue('7', { row: 0, col: 0 }).ok).toBe(true);
      expect(second).toEqual([editor.snapshot().revision]);
      expect(notified).toBeGreaterThan(0);
      expect(frame()).not.toBeNull();
      expect(reported).toContain(boom);
    });

    test('each throwing handler is reported', async () => {
      const editor = await mount();
      const first = new Error('first');
      const second = new Error('second');
      editor.on('selectionChange', () => {
        throw first;
      });
      editor.on('selectionChange', () => {
        throw second;
      });
      editor.select({ row: 3, col: 1 });
      expect(reported).toEqual([first, second]);
    });

    test('a throwing selectionChange handler does not throw out of select', async () => {
      const editor = await mount();
      editor.on('selectionChange', () => {
        throw new Error('handler');
      });
      expect(() => editor.select({ row: 4, col: 4 })).not.toThrow();
      expect(editor.snapshot().selectionRef).toBe('E5');
      expect(reported).toHaveLength(1);
    });
  });

  test('unsubscribe stops the handler', async () => {
    const editor = await mount();
    const seen: EditorSnapshot[] = [];
    const off = editor.on('selectionChange', (s) => seen.push(s));
    editor.select({ row: 2, col: 2 });
    expect(seen).toHaveLength(1);
    off();
    editor.select({ row: 3, col: 3 });
    expect(seen).toHaveLength(1);
  });
});
