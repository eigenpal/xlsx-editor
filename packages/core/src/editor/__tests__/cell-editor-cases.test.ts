// The cell editor's behavior around the edit (task 7.2): it follows its cell when the window
// scrolls and never commits to a cell that is out of view; save commits the open edit first,
// redo waits for a changed edit, and undo cancels changed text; a refusal has one
// announcement; an input method that starts on the grid composes in the editor; a shared
// formula child and a date cell show what they hold; and a double press, a closed surface,
// and a sheet without a part behave as F2, a commit, and a refusal.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import {
  createXlsxEditor,
  runToolbarCommand,
  toolbarCommandState,
  type XlsxEditorInstance,
} from '../index.ts';
import { flushEditor } from './flush-editor.ts';
import { CellEditing, type EditingHost } from '../cell-editing.ts';
import type { FontSource } from '../../layout/index.ts';
import { EN_US } from '../../cascade/index.ts';
import {
  attr,
  cellKey,
  readXlsx,
  sheetIndex,
  WorkbookDocumentStore,
  workbookIndex,
  xmlPartOf,
} from '../../store/index.ts';

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

async function mount(bytes: Uint8Array) {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.appendChild(host);
  const viewer = createXlsxEditor({ fonts: FONTS, translate: (k) => k });
  viewers.push(viewer);
  viewer.attach(host);
  const scroller = host.querySelector('.xlsx-scroller') as HTMLElement;
  Object.defineProperty(scroller, 'clientWidth', { get: () => 800 });
  Object.defineProperty(scroller, 'clientHeight', { get: () => 400 });
  await viewer.load(bytes);
  flushEditor(viewer);
  const input = host.querySelector('.xlsx-cell-editor') as HTMLTextAreaElement;
  const live = host.querySelector('.xlsx-live') as HTMLElement;
  const W = win as unknown as typeof globalThis;
  return { viewer, host, scroller, input, live, doc, W };
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

const content = (v: XlsxEditorInstance) => v.snapshot().activeCell?.content;

const sheet = () =>
  buildXlsx({
    sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 2)]) + row(2, [c('A2', 5)]) }],
  });

describe('the editor and the scroll', () => {
  test('it covers the whole cell, hides out of view, and a commit key only brings it back', async () => {
    const { viewer, scroller, input, doc, W } = await mount(sheet());
    key(scroller, '7');
    const height = input.style.minHeight;
    const top = Number.parseFloat(input.style.top);
    expect(Number.parseFloat(height)).toBeGreaterThan(0);
    // Half out of view: the editor keeps the full cell height and moves up with the cell.
    scroller.scrollTop = 8;
    scroller.dispatchEvent(new W.Event('scroll'));
    flushEditor(viewer);
    expect(input.style.minHeight).toBe(height);
    expect(Number.parseFloat(input.style.top)).toBeLessThan(top);
    expect(input.hasAttribute('data-offscreen')).toBe(false);
    // Wholly out of view: unseen, still focused, the text kept.
    scroller.scrollTop = 4000;
    scroller.dispatchEvent(new W.Event('scroll'));
    flushEditor(viewer);
    expect(input.hasAttribute('data-offscreen')).toBe(true);
    expect(doc.activeElement).toBe(input);
    expect(input.value).toBe('7');
    // Enter brings the cell back and does not commit.
    key(input, 'Enter');
    expect(viewer.snapshot().editing).toBe(true);
    expect(input.hasAttribute('data-offscreen')).toBe(false);
    expect(content(viewer)).toBe('1');
    // Now the cell shows, and Enter commits.
    key(input, 'Enter');
    expect(viewer.snapshot().editing).toBe(false);
    viewer.select({ row: 0, col: 0 });
    expect(content(viewer)).toBe('7');
  });

  test('the editor uses the font of the cell at the zoom', async () => {
    const { viewer, scroller, input } = await mount(sheet());
    key(scroller, '7');
    const size = Number.parseFloat(input.style.fontSize);
    expect(input.style.fontFamily).toContain('sans-serif');
    expect(size).toBeCloseTo((11 * 96) / 72, 3);
    viewer.cancelEdit();
    viewer.setZoom(2);
    key(scroller, '7');
    flushEditor(viewer);
    expect(Number.parseFloat(input.style.fontSize)).toBeCloseTo(size * 2, 3);
  });

  test('a new cell in a formatted row or column is typed in that format font', async () => {
    const { viewer, scroller, input } = await mount(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            cols: '<cols><col min="3" max="3" width="9" style="2" customWidth="1"/></cols>',
            rows: row(2, [c('A2', 1)], 's="1" customFormat="1"'),
          },
        ],
        styles: {
          fonts: ['<font><sz val="20"/><name val="Calibri"/></font>'],
          cellXfs: [
            '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
            '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
          ],
        },
      })
    );
    for (const at of [
      { row: 1, col: 1 },
      { row: 5, col: 2 },
    ]) {
      viewer.select(at);
      key(scroller, '7');
      expect(Number.parseFloat(input.style.fontSize)).toBeCloseTo((20 * 96) / 72, 3);
      viewer.cancelEdit();
    }
    viewer.select({ row: 5, col: 1 });
    key(scroller, '7');
    expect(Number.parseFloat(input.style.fontSize)).toBeCloseTo((11 * 96) / 72, 3);
  });
});

describe('undo, redo, and save with an open edit', () => {
  test('save commits the open edit first; undo cancels changed text; redo waits for it', async () => {
    const { viewer, scroller, input } = await mount(sheet());
    key(scroller, '9');
    input.value = '9';
    const saved = readXlsx(viewer.save()!);
    expect(viewer.snapshot().editing).toBe(false);
    expect(saved.ok).toBe(true);
    viewer.select({ row: 0, col: 0 });
    expect(content(viewer)).toBe('9');
    // Undo of typing over the cell cancels the edit, and keeps the history. Enter then moves
    // on, and the cell keeps its content.
    key(scroller, '3');
    const revision = viewer.snapshot().revision;
    expect(viewer.undo().ok).toBe(true);
    expect(viewer.snapshot().editing).toBe(false);
    expect(viewer.snapshot().revision).toBe(revision);
    key(scroller, 'Enter');
    viewer.select({ row: 0, col: 0 });
    expect(content(viewer)).toBe('9');
    // The next undo reverts the save's edit, one step.
    expect(viewer.undo().ok).toBe(true);
    expect(viewer.snapshot().editing).toBe(false);
    expect(content(viewer)).not.toBe('9');
    // Redo with an open edit of changed text is refused before any write; the redo stays.
    key(scroller, '3');
    const r = viewer.redo();
    expect(r.ok ? null : r.refusal.code).toBe('edit-pending');
    expect(viewer.snapshot().editing).toBe(true);
    expect(viewer.snapshot().canRedo).toBe(true);
    // An unchanged open edit closes, and the redo runs.
    viewer.cancelEdit();
    key(scroller, 'F2');
    expect(viewer.redo().ok).toBe(true);
    expect(viewer.snapshot().editing).toBe(false);
    expect(content(viewer)).toBe('9');
  });

  test('a refused commit refuses save, redo waits for the edit, and the editor keeps its text', async () => {
    const { viewer, scroller, input } = await mount(sheet());
    expect(viewer.enterValue('4').ok).toBe(true);
    key(scroller, '=');
    input.value = '=SUM(';
    expect(viewer.save()).toBeNull();
    const redo = viewer.redo();
    expect(redo.ok ? null : redo.refusal.code).toBe('edit-pending');
    expect(viewer.snapshot().editing).toBe(true);
    expect(input.value).toBe('=SUM(');
    expect(viewer.snapshot().canUndo).toBe(true);
  });
});

describe('the refusal announcement', () => {
  test('one channel, announced again when repeated, cleared by a new selection or sheet', async () => {
    const bytes = buildXlsx({
      sheets: [
        { name: 'S', rows: row(1, [c('A1', 1)]) },
        { name: 'T', rows: row(1, [c('A1', 2)]) },
      ],
    });
    const { viewer, live, host } = await mount(bytes);
    expect(viewer.enterValue('=A1+').ok).toBe(false);
    const first = live.textContent;
    expect(first).toBe('editing.refused.formula-syntax');
    expect(viewer.enterValue('=A1+').ok).toBe(false);
    expect(live.textContent).not.toBe(first);
    expect(live.textContent?.trim()).toBe(first);
    expect(host.querySelector('[role="alert"]')).toBeNull();
    viewer.select({ row: 1, col: 1 });
    expect(viewer.snapshot().refusal).toBeNull();
    expect(viewer.enterValue('=RC[1]').ok).toBe(false);
    expect(viewer.snapshot().refusal?.code).toBe('r1c1-formula');
    expect(viewer.setActiveSheet(1)).toBe(true);
    expect(viewer.snapshot().refusal).toBeNull();
  });
  test('load() and close() drop an open edit, and announce a changed one once', async () => {
    const bytes = buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] });
    const { viewer, live } = await mount(bytes);
    viewer.beginEdit('7');
    await viewer.load(bytes);
    expect(viewer.snapshot().editing).toBe(false);
    expect(live.textContent).toBe('editing.refused.edit-dropped');
    expect(content(viewer)).toBe('1');
    // An edit with no change is dropped with no announcement.
    live.textContent = '';
    viewer.beginEdit();
    await viewer.load(bytes);
    expect(live.textContent).toBe('');
    // A load that fails, and close(), drop a changed edit with the announcement.
    viewer.beginEdit('8');
    await viewer.load(new Uint8Array([1, 2, 3]));
    expect(viewer.snapshot().status).toBe('failed');
    expect(live.textContent).toBe('editing.refused.edit-dropped');
    await viewer.load(bytes);
    live.textContent = '';
    viewer.beginEdit('9');
    viewer.close();
    expect(live.textContent).toBe('editing.refused.edit-dropped');
  });

  test('undo of changed text announces it, and is refused while an input method composes', async () => {
    const bytes = buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] });
    const { viewer, live, input, W } = await mount(bytes);
    viewer.beginEdit('7');
    input.dispatchEvent(new W.Event('compositionstart'));
    const r = viewer.undo();
    expect(r.ok ? null : r.refusal.code).toBe('composing');
    expect(viewer.snapshot().editing).toBe(true);
    expect(input.value).toBe('7');
    input.dispatchEvent(new W.Event('compositionend'));
    expect(viewer.undo().ok).toBe(true);
    expect(viewer.snapshot().editing).toBe(false);
    expect(live.textContent).toBe('grid.typingUndone');
  });
});

describe('an input method on the grid', () => {
  test('a composition that starts on the grid moves to the editor before its text lands', async () => {
    const { viewer, scroller, input, doc, W } = await mount(sheet());
    scroller.dispatchEvent(new W.Event('compositionstart', { bubbles: true }));
    expect(viewer.snapshot().editing).toBe(true);
    expect(doc.activeElement).toBe(input);
    viewer.cancelEdit();
    // The keydown that starts a composition (key code 229) opens the editor the same way.
    key(scroller, 'Process', { keyCode: 229 });
    expect(doc.activeElement).toBe(input);
    expect(input.value).toBe('');
    input.dispatchEvent(new W.Event('compositionstart'));
    input.value = 'に';
    key(input, 'Enter', { isComposing: true });
    expect(viewer.snapshot().editing).toBe(true);
    input.value = '日本';
    input.dispatchEvent(new W.Event('compositionend'));
    key(input, 'Enter');
    viewer.select({ row: 0, col: 0 });
    expect(content(viewer)).toBe('日本');
  });
});

describe('what the editor shows', () => {
  test('a shared formula child shows its moved master formula, and F2 edits it', async () => {
    const rows = [1, 2, 3]
      .map((r) =>
        row(r, [
          c(`A${r}`, r),
          r === 1
            ? '<c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f><v>2</v></c>'
            : `<c r="B${r}"><f t="shared" si="0"/><v>${r * 2}</v></c>`,
        ])
      )
      .join('');
    const { viewer, scroller, input } = await mount(buildXlsx({ sheets: [{ name: 'S', rows }] }));
    viewer.select({ row: 2, col: 1 });
    expect(content(viewer)).toBe('=A3*2');
    key(scroller, 'F2');
    expect(input.value).toBe('=A3*2');
    // An unchanged commit writes nothing.
    key(input, 'Enter');
    expect(viewer.snapshot().canUndo).toBe(false);
  });

  test('a date cell shows ISO text, and ISO text typed in it stays a date', async () => {
    const bytes = buildXlsx({
      sheets: [{ name: 'S', rows: row(1, ['<c r="A1" s="1" t="d"><v>2024-01-02</v></c>']) }],
      styles: {
        cellXfs: [
          '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
        ],
      },
    });
    const { viewer } = await mount(bytes);
    expect(content(viewer)).toBe('2024-01-02');
    expect(viewer.enterValue('2024-03-05').ok).toBe(true);
    const saved = readXlsx(viewer.save()!);
    if (!saved.ok) throw new Error('save');
    const part = xmlPartOf(saved.pkg, workbookIndex(saved.pkg)!.sheets[0]!.partName)!;
    const a1 = sheetIndex(part).cells.get(cellKey(0, 0))!.el;
    expect([attr(a1, 't'), attr(a1, 's')]).toEqual(['d', '1']);
    expect(content(viewer)).toBe('2024-03-05');
    // Other text in a date cell is read as usual.
    expect(viewer.enterValue('soon').ok).toBe(true);
    expect(content(viewer)).toBe('soon');
  });
});

describe('a double press, a closed surface, and a sheet without a part', () => {
  test('a double press on a cell edits it in place', async () => {
    const { viewer, host, input, W } = await mount(sheet());
    const stage = host.querySelector('.xlsx-stage') as HTMLElement;
    stage.dispatchEvent(
      new W.MouseEvent('dblclick', { bubbles: true, cancelable: true, clientX: 60, clientY: 30 })
    );
    expect(viewer.snapshot().editing).toBe(true);
    expect(input.value).toBe(content(viewer)!);
  });

  test('detach commits a valid open edit, and announces a refused one', async () => {
    const { viewer, scroller, input } = await mount(sheet());
    key(scroller, '6');
    viewer.detach();
    expect(viewer.snapshot().editing).toBe(false);
    expect(content(viewer)).toBe('6');
    const second = await mount(sheet());
    key(second.scroller, '=');
    second.input.value = '=SUM(';
    second.viewer.detach();
    expect(second.viewer.snapshot().editing).toBe(false);
    expect(second.viewer.snapshot().refusal?.code).toBe('unreadable-formula');
    expect(input.isConnected).toBe(false);
  });

  test('beginEdit on a sheet without a part is refused before an editor opens', () => {
    const win = new Window();
    const doc = win.document as unknown as Document;
    const stage = doc.createElement('div');
    const pkg = readXlsx(sheet());
    if (!pkg.ok) throw new Error('read');
    const index = workbookIndex(pkg.pkg)!;
    const store = new WorkbookDocumentStore(pkg.pkg);
    const host: EditingHost = {
      viewing: () => false,
      loading: () => false,
      target: () => ({
        store,
        pkg: pkg.pkg,
        index,
        sheet: 0,
        entry: { ...index.sheets[0]!, partName: undefined },
        revision: store.revision,
        selection: {
          sheetIndex: 0,
          anchor: { row: 0, col: 0 },
          active: { row: 0, col: 0 },
          range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
        },
        focus: { row: 0, col: 0 },
      }),
      locale: () => EN_US,
      t: (k) => k,
      surface: () => null,
      say: () => {},
      silence: () => {},
      merges: () => [],
      cellFont: () => ({ family: 'sans-serif', sizePt: 11 }),
      publish: () => {},
      publishDraft: () => {},
      schedule: () => {},
      render: () => {},
      reveal: () => {},
      moveFrom: () => {},
      restore: () => {},
    };
    const editing = new CellEditing(host);
    editing.attach(stage);
    const r = editing.begin('x');
    expect(r.ok ? null : r.refusal.code).toBe('not-editable');
    expect(editing.editing).toBe(false);
    expect((stage.querySelector('textarea') as HTMLTextAreaElement).hidden).toBe(true);
  });
});

describe('the open edit keeps its locale, counts typing, and waits for an input method', () => {
  test('a new number locale applies to the next edit, not the open one', async () => {
    const { viewer, scroller } = await mount(
      buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1.5)]) }] })
    );
    viewer.select({ row: 0, col: 0 });
    key(scroller, 'F2');
    const revision = viewer.snapshot().revision;
    viewer.setNumberLocale({ ...EN_US, decimal: ',', group: '.' });
    expect(viewer.snapshot().editChanged).toBe(false);
    expect(viewer.commitEdit().ok).toBe(true);
    expect(viewer.snapshot().revision).toBe(revision);
    const cell = viewer.snapshot().activeCell;
    expect(cell?.content).toBe('1,5');
    // The next edit reads its text in the new locale.
    expect(viewer.enterValue('2,5').ok).toBe(true);
    expect(viewer.snapshot().activeCell?.content).toBe('2,5');
  });

  test("typing a cell's own content is a change, and toolbar undo cancels it", async () => {
    const { viewer, scroller, input } = await mount(sheet());
    expect(viewer.enterValue('6', { row: 1, col: 0 }).ok).toBe(true);
    viewer.select({ row: 0, col: 0 });
    const before = content(viewer);
    key(scroller, before!.charAt(0));
    input.value = before!;
    expect(viewer.snapshot().editChanged).toBe(true);
    expect(viewer.undo().ok).toBe(true);
    expect(viewer.snapshot().editing).toBe(false);
    expect(viewer.snapshot().canUndo).toBe(true);
  });

  test('while an input method composes, every command that commits is refused', async () => {
    const { viewer, scroller, input, W } = await mount(sheet());
    viewer.select({ row: 0, col: 0 });
    key(scroller, '7');
    input.dispatchEvent(new W.Event('compositionstart'));
    expect(viewer.snapshot().composing).toBe(true);
    const code = (r: { ok: boolean; refusal?: { code: string } }) =>
      r.ok ? null : r.refusal?.code;
    expect(code(viewer.toggleFormat('bold'))).toBe('composing');
    expect(code(viewer.clearValues())).toBe('composing');
    expect(code(viewer.redo())).toBe('composing');
    expect(code(viewer.undo())).toBe('composing');
    // A decimal step gives the input method first, as its state does, not the changed text.
    expect(code(runToolbarCommand(viewer, 'number.increaseDecimals'))).toBe('composing');
    expect(
      toolbarCommandState(viewer.snapshot(), 'number.increaseDecimals').disabledReason?.key
    ).toBe('disabledReason.composing');
    expect(viewer.save()).toBeNull();
    expect(viewer.setMode('view')).toBe(false);
    expect(viewer.select({ row: 2, col: 2 })).toBe(false);
    expect(toolbarCommandState(viewer.snapshot(), 'font.bold').disabledReason?.key).toBe(
      'disabledReason.composing'
    );
    expect(viewer.snapshot().editing).toBe(true);
    input.dispatchEvent(new W.Event('compositionend'));
    expect(viewer.snapshot().composing).toBe(false);
    expect(toolbarCommandState(viewer.snapshot(), 'font.bold').enabled).toBe(true);
    expect(viewer.commitEdit().ok).toBe(true);
  });
});

describe('a surface that closes while an input method composes', () => {
  test('detach drops the composing edit, announces it, and stores nothing', async () => {
    const { viewer, scroller, input, live, W } = await mount(sheet());
    viewer.select({ row: 0, col: 0 });
    key(scroller, '7');
    input.dispatchEvent(new W.Event('compositionstart'));
    input.value = 'か';
    const revision = viewer.snapshot().revision;
    // A listener sees the dropped edit while the surface is still connected.
    const seen: { code: string | undefined; connected: boolean; text: string | null }[] = [];
    viewer.subscribe(() =>
      seen.push({
        code: viewer.snapshot().refusal?.code,
        connected: live.isConnected,
        text: live.textContent,
      })
    );
    viewer.detach();
    expect(seen[0]).toEqual({
      code: 'edit-dropped',
      connected: true,
      text: 'editing.refused.edit-dropped',
    });
    expect(viewer.snapshot().refusal?.code).toBe('edit-dropped');
    expect(viewer.snapshot().editing).toBe(false);
    expect(viewer.snapshot().revision).toBe(revision);
    expect(viewer.snapshot().activeCell?.content).toBe('1');
  });
});

describe('host calls during composition and loading', () => {
  test('setMode view while an input method composes is refused with no announcement', async () => {
    const { viewer, scroller, input, live, W } = await mount(sheet());
    viewer.select({ row: 0, col: 0 });
    key(scroller, '7');
    input.dispatchEvent(new W.Event('compositionstart'));
    live.textContent = '';
    expect(viewer.setMode('view')).toBe(false);
    expect(live.textContent).toBe('');
    expect(viewer.snapshot().refusal).toBeNull();
    expect(viewer.snapshot().editing).toBe(true);
  });

  test('a load drops the open edit when it starts, and commands wait for it', async () => {
    const bytes = sheet();
    const { viewer, scroller, live } = await mount(bytes);
    viewer.select({ row: 0, col: 0 });
    key(scroller, '7');
    const loading = viewer.load(bytes);
    expect(viewer.snapshot().status).toBe('loading');
    expect(viewer.snapshot().editing).toBe(false);
    expect(live.textContent).toBe('editing.refused.edit-dropped');
    expect(viewer.snapshot().refusal?.code).toBe('edit-dropped');
    const code = (r: { ok: boolean; refusal?: { code: string } }) =>
      r.ok ? null : r.refusal?.code;
    // Every command, a save included, is refused with `loading`, and the live region says so.
    expect(viewer.snapshot().editDropped).toBe(true);
    expect(code(viewer.enterValue('5'))).toBe('loading');
    // A later refusal replaces `edit-dropped`, and the snapshot still says the edit dropped.
    expect(viewer.snapshot().editDropped).toBe(true);
    expect(live.textContent).toBe('editing.refused.loading');
    expect(code(viewer.undo())).toBe('loading');
    expect(code(viewer.redo())).toBe('loading');
    expect(code(viewer.applyFormat({ bold: true }))).toBe('loading');
    // A copy and a cut return the refusal that the snapshot shows.
    expect(code(viewer.copy())).toBe('loading');
    expect(viewer.cut()).toEqual({ ok: false, refusal: viewer.snapshot().refusal! });
    expect(viewer.snapshot().refusal).toEqual({ code: 'loading', command: 'cut' });
    expect(viewer.save()).toBeNull();
    expect(viewer.snapshot().refusal).toEqual({ code: 'loading', command: 'save' });
    // The view commands act on no workbook while another one loads.
    const zoom = viewer.snapshot().zoom;
    expect(viewer.select({ row: 2, col: 1 })).toBe(false);
    expect(viewer.setActiveSheet(0)).toBe(false);
    viewer.setZoom(2);
    expect(viewer.snapshot().zoom).toBe(zoom);
    expect(toolbarCommandState(viewer.snapshot(), 'font.bold').disabledReason?.key).toBe(
      'disabledReason.loading'
    );
    await loading;
    // The drop stays shown in the ready workbook, until the next selection.
    expect(viewer.snapshot().status).toBe('ready');
    expect(viewer.snapshot().editDropped).toBe(true);
    expect(viewer.snapshot().refusal).toEqual({ code: 'edit-dropped' });
    // The next command is refused: the drop is no longer shown, as the snapshot says.
    expect(code(viewer.undo())).toBe('nothing-to-undo');
    expect(viewer.snapshot().editDropped).toBe(false);
    viewer.select({ row: 0, col: 0 });
    expect(viewer.snapshot().editDropped).toBe(false);
    expect(viewer.snapshot().refusal).toBeNull();
    expect(viewer.snapshot().activeCell?.content).toBe('1');
  });
});
