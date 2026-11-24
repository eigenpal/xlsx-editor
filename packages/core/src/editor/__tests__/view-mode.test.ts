// View mode: `setMode('view')` refuses every edit command with `view-mode`, from the methods,
// the keyboard, and a double press, and writes nothing. Selection, sheet tabs, zoom, and save
// stay available. `toolbarCommandState` disables every slot but save with the mode's reason.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import {
  CHROME_SLOT_IDS,
  createXlsxEditor,
  runToolbarCommand,
  toolbarCommandState,
  type EditResult,
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

let viewers: XlsxEditorInstance[] = [];
afterEach(() => {
  for (const v of viewers) v.destroy();
  viewers = [];
});

const workbook = () =>
  buildXlsx({
    sheets: [
      { name: 'S', rows: row(1, [c('A1', 1), c('B1', 2)]) + row(2, [c('A2', 3)]) },
      { name: 'T', rows: row(1, [c('A1', 9)]) },
    ],
  });

async function mount(mode?: 'edit' | 'view') {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.appendChild(host);
  const viewer = createXlsxEditor({ fonts: FONTS, translate: (k) => k, ...(mode ? { mode } : {}) });
  viewers.push(viewer);
  viewer.attach(host);
  const scroller = host.querySelector('.xlsx-scroller') as HTMLElement;
  Object.defineProperty(scroller, 'clientWidth', { get: () => 800 });
  Object.defineProperty(scroller, 'clientHeight', { get: () => 400 });
  await viewer.load(workbook());
  flushEditor(viewer);
  const input = host.querySelector('.xlsx-cell-editor') as HTMLTextAreaElement;
  const live = host.querySelector('.xlsx-live') as HTMLElement;
  return { viewer, host, scroller, input, live };
}

const key = (el: Element, k: string, init: Record<string, unknown> = {}) =>
  el.dispatchEvent(
    new (el.ownerDocument!.defaultView as unknown as typeof globalThis).KeyboardEvent('keydown', {
      key: k,
      bubbles: true,
      cancelable: true,
      ...init,
    })
  );

const code = (r: EditResult) => (r.ok ? 'ok' : r.refusal.code);
const content = (v: XlsxEditorInstance) => v.snapshot().activeCell?.content;

describe('view mode', () => {
  test('the editor opens in edit mode, or in the mode of its options', async () => {
    expect((await mount()).viewer.snapshot().mode).toBe('edit');
    const { viewer, scroller } = await mount('view');
    expect(viewer.snapshot().mode).toBe('view');
    expect(scroller.getAttribute('aria-readonly')).toBe('true');
  });

  test('every edit command is refused with view-mode and writes nothing', async () => {
    const { viewer, live } = await mount();
    viewer.setMode('view');
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 1 });
    const commands: Record<string, () => EditResult> = {
      enterValue: () => viewer.enterValue('5'),
      beginEdit: () => viewer.beginEdit(),
      clearValues: () => viewer.clearValues(),
      clearFormats: () => viewer.clearFormats(),
      clearAll: () => viewer.clearAll(),
      applyFormat: () => viewer.applyFormat({ bold: true }),
      toggleFormat: () => viewer.toggleFormat('italic'),
      insertRows: () => viewer.insertRows(),
      deleteRows: () => viewer.deleteRows(),
      insertColumns: () => viewer.insertColumns(),
      deleteColumns: () => viewer.deleteColumns(),
      mergeCells: () => viewer.mergeCells(),
      unmergeCells: () => viewer.unmergeCells(),
      undo: () => viewer.undo(),
      redo: () => viewer.redo(),
    };
    for (const [name, run] of Object.entries(commands))
      expect({ name, code: code(run()) }).toEqual({ name, code: 'view-mode' });
    expect(viewer.snapshot().refusal?.code).toBe('view-mode');
    // Announced once: the repeated refusals do not change the live text again.
    expect(live.textContent).toBe('editing.refused.view-mode');
    expect(viewer.snapshot().editing).toBe(false);
    expect(viewer.snapshot().canUndo).toBe(false);
    expect(content(viewer)).toBe('1');
  });

  test('undo is refused even with a history, and switching back allows it', async () => {
    const { viewer } = await mount();
    expect(viewer.enterValue('7').ok).toBe(true);
    viewer.setMode('view');
    expect(code(viewer.undo())).toBe('view-mode');
    expect(content(viewer)).toBe('7');
    viewer.setMode('edit');
    expect(viewer.snapshot().mode).toBe('edit');
    expect(viewer.undo().ok).toBe(true);
    expect(content(viewer)).toBe('1');
  });

  test('typing, F2, Backspace, Delete, a toggle, and a double press start no edit', async () => {
    const { viewer, scroller, host } = await mount('view');
    for (const [k, init] of [
      ['x', {}],
      ['F2', {}],
      ['Backspace', {}],
      ['Delete', {}],
      ['b', { ctrlKey: true }],
      ['z', { ctrlKey: true }],
    ] as const) {
      key(scroller, k, init);
      expect({ k, editing: viewer.snapshot().editing }).toEqual({ k, editing: false });
      expect({ k, code: viewer.snapshot().refusal?.code }).toEqual({ k, code: 'view-mode' });
    }
    const stage = host.querySelector('.xlsx-stage') as HTMLElement;
    const W = host.ownerDocument.defaultView as unknown as typeof globalThis;
    stage.dispatchEvent(
      new W.MouseEvent('dblclick', { bubbles: true, cancelable: true, clientX: 200, clientY: 100 })
    );
    expect(viewer.snapshot().editing).toBe(false);
    expect(content(viewer)).toBe('1');
    expect(viewer.snapshot().activeFormat?.bold).toBe(false);
  });

  test('selection, arrow keys, sheet tabs, zoom, and save stay available', async () => {
    const { viewer, scroller } = await mount('view');
    key(scroller, 'ArrowRight');
    expect(viewer.snapshot().activeCell?.ref).toBe('B1');
    viewer.select({ row: 1, col: 0 });
    expect(content(viewer)).toBe('3');
    viewer.setZoom(1.5);
    expect(viewer.snapshot().zoom).toBe(1.5);
    expect(viewer.setActiveSheet(1)).toBe(true);
    expect(viewer.snapshot().activeSheet).toBe(1);
    expect(viewer.save()?.byteLength).toBeGreaterThan(0);
  });

  test('switching to view mode commits an open edit first', async () => {
    const { viewer, scroller, input } = await mount();
    key(scroller, '4');
    input.value = '42';
    viewer.setMode('view');
    expect(viewer.snapshot().editing).toBe(false);
    expect(content(viewer)).toBe('42');
    expect(input.hidden).toBe(true);
  });

  test('switching to view mode is refused while the store refuses the open edit', async () => {
    const { viewer, scroller, input } = await mount();
    const before = content(viewer);
    key(scroller, '=');
    input.value = '=1+';
    input.dispatchEvent(
      new (input.ownerDocument.defaultView as unknown as typeof globalThis).Event('input')
    );
    expect(viewer.setMode('view')).toBe(false);
    // The editor stays in edit mode, open with the typed text, and shows why.
    expect(viewer.snapshot().mode).toBe('edit');
    expect(viewer.snapshot().editing).toBe(true);
    expect(viewer.snapshot().editText).toBe('=1+');
    expect(content(viewer)).toBe(before);
    expect(viewer.snapshot().refusal).not.toBeNull();
    // A cancel closes the edit, and the switch then applies.
    viewer.cancelEdit();
    expect(viewer.setMode('view')).toBe(true);
    expect(viewer.snapshot().mode).toBe('view');
  });

  test('toolbarCommandState disables every slot that edits, and keeps the values', async () => {
    const { viewer } = await mount('view');
    const host = { save: () => undefined, open: () => undefined };
    const snap = viewer.snapshot();
    // Save, Open, Copy, Select all, and the zoom read the workbook or change only the view.
    const reads = new Set<string>([
      'file.save',
      'file.open',
      'edit.copy',
      'edit.selectAll',
      'view.zoomIn',
      'view.zoomOut',
    ]);
    for (const slot of CHROME_SLOT_IDS) {
      const state = toolbarCommandState(snap, slot, host);
      if (reads.has(slot)) {
        expect({ slot, enabled: state.enabled }).toEqual({ slot, enabled: true });
        continue;
      }
      // At 100%, Zoom to 100% has nothing to do: its own reason, not view mode.
      if (slot === 'view.zoomReset') {
        expect(state.disabledReason?.key).toBe('disabledReason.zoomActual');
        continue;
      }
      expect({ slot, enabled: state.enabled }).toEqual({ slot, enabled: false });
      expect({ slot, key: state.disabledReason?.key }).toEqual({
        slot,
        key: 'disabledReason.viewMode',
      });
    }
    expect(toolbarCommandState(snap, 'font.size', host).value).toBe('11');
    expect(code(runToolbarCommand(viewer, 'font.bold', undefined, host))).toBe('view-mode');
    expect(runToolbarCommand(viewer, 'file.save', undefined, host).ok).toBe(true);
  });
});
