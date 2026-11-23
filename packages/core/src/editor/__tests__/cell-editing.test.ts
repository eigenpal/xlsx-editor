// Cell editing from the grid (task 7.2) in a happy-dom window of its own: the key table, the
// cell editor projection (commit, cancel, refusal, input methods), one undo step for each
// command, and the save that carries the edits and the stale flag of design D3 (task 7.3).

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { strFromU8, unzipSync, zipSync } from 'fflate';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import {
  createXlsxEditor,
  EDIT_REFUSAL_CODES,
  parseCellInput,
  type XlsxEditorInstance,
} from '../index.ts';
import { flushEditor } from './flush-editor.ts';
import type { FontSource } from '../../layout/index.ts';
import { EN_US } from '../../cascade/index.ts';
import { MAX_COLS, readXlsx, workbookIndex } from '../../store/index.ts';

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

const fixture = (name: string) =>
  new Uint8Array(readFileSync(`${import.meta.dir}/../../../../../fixtures/${name}.xlsx`));

let viewers: XlsxEditorInstance[] = [];
afterEach(() => {
  for (const v of viewers) v.destroy();
  viewers = [];
});

/** A translator that shows the key and its parameters, so a test can read both. */
const withParams = (k: string, p?: Readonly<Record<string, string | number>>) =>
  p ? `${k} ${JSON.stringify(p)}` : k;

async function mount(bytes: Uint8Array, translate: typeof withParams = (k) => k) {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.appendChild(host);
  const viewer = createXlsxEditor({ fonts: FONTS, translate });
  viewers.push(viewer);
  viewer.attach(host);
  const scroller = host.querySelector('.xlsx-scroller') as HTMLElement;
  Object.defineProperty(scroller, 'clientWidth', { get: () => 800 });
  Object.defineProperty(scroller, 'clientHeight', { get: () => 400 });
  await viewer.load(bytes);
  flushEditor(viewer);
  const input = host.querySelector('.xlsx-cell-editor') as HTMLTextAreaElement;
  const live = host.querySelector('.xlsx-live') as HTMLElement;
  return { viewer, host, scroller, input, live, doc };
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
    sheets: [
      {
        name: 'S',
        rows:
          row(1, [c('A1', 1), c('B1', 2)]) + row(2, ['<c r="A2" s="1"><v>5</v></c>', c('B2', 6)]),
      },
    ],
    styles: {
      cellXfs: [
        '<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
      ],
    },
  });

describe('the cell editor', () => {
  test('typing starts an edit, Enter commits a number and moves down, and Mod+Z undoes it', async () => {
    const { viewer, scroller, input } = await mount(sheet());
    viewer.select({ row: 0, col: 2 });
    key(scroller, '4');
    expect(viewer.snapshot().editing).toBe(true);
    expect(input.hidden).toBe(false);
    expect(input.value).toBe('4');
    input.value = '42';
    key(input, 'Enter');
    expect(viewer.snapshot().editing).toBe(false);
    expect(viewer.snapshot().activeCell?.ref).toBe('C2');
    viewer.select({ row: 0, col: 2 });
    expect(content(viewer)).toBe('42');
    expect(viewer.snapshot().canUndo).toBe(true);
    key(scroller, 'z', { ctrlKey: true });
    expect(content(viewer)).toBe('');
    key(scroller, 'z', { metaKey: true, shiftKey: true });
    expect(content(viewer)).toBe('42');
    key(scroller, 'z', { ctrlKey: true });
    key(scroller, 'y', { ctrlKey: true });
    expect(content(viewer)).toBe('42');
  });

  test('F2 edits in place, Escape cancels, and nothing is written', async () => {
    const { viewer, scroller, input } = await mount(sheet());
    key(scroller, 'F2');
    expect(input.value).toBe('1');
    input.value = '99';
    key(input, 'Escape');
    expect(viewer.snapshot().editing).toBe(false);
    expect(content(viewer)).toBe('1');
    expect(viewer.snapshot().canUndo).toBe(false);
  });

  test('Tab commits and moves right; an arrow commits an edit that started by typing', async () => {
    const { viewer, scroller, input } = await mount(sheet());
    key(scroller, 'x');
    key(input, 'Tab');
    expect(viewer.snapshot().activeCell?.ref).toBe('B1');
    key(scroller, 'y');
    key(input, 'ArrowDown');
    expect(viewer.snapshot().activeCell?.ref).toBe('B2');
    viewer.select({ row: 0, col: 0 });
    expect(content(viewer)).toBe('x');
  });

  test('enterValue refuses a point outside the grid, and select refuses a non-finite one', async () => {
    const { viewer } = await mount(sheet());
    for (const at of [
      { row: -1, col: 0 },
      { row: 0, col: 16384 },
      { row: 1048576, col: 0 },
      { row: 0.5, col: 0 },
      { row: Number.NaN, col: 0 },
      { row: 0, col: Number.POSITIVE_INFINITY },
    ]) {
      const r = viewer.enterValue('5', at);
      expect(r.ok ? null : r.refusal.code).toBe('bad-reference');
    }
    expect(viewer.snapshot().canUndo).toBe(false);
    expect(viewer.enterValue('5', { row: 1048575, col: 16383 }).ok).toBe(true);
    expect(viewer.select({ row: Number.NaN, col: 0 })).toBe(false);
    expect(viewer.select({ row: 0, col: 0 }, { row: 0, col: Number.NEGATIVE_INFINITY })).toBe(
      false
    );
    expect(viewer.select({ row: 2.7, col: 1e9 })).toBe(true);
    expect(viewer.snapshot().selectionRef).toBe('XFD3');
    expect(viewer.select({ row: 0, col: 0 }, { row: 1e12, col: 1 })).toBe(true);
    expect(viewer.snapshot().selectionRef).toBe('A1:B1048576');
  });

  test('typing on a locked cell announces the refusal once, not for each key', async () => {
    const { viewer, scroller, live } = await mount(
      buildXlsx({
        sheets: [
          { name: 'S', rows: row(1, [c('A1', 1)]), beforeMerges: '<sheetProtection sheet="1"/>' },
        ],
      })
    );
    expect(viewer.beginEdit()).toMatchObject({
      ok: false,
      refusal: { code: 'locked', lock: 'cell' },
    });
    key(scroller, 'x');
    expect(viewer.snapshot().refusal).toMatchObject({ code: 'locked', lock: 'cell' });
    const first = live.textContent;
    expect(first).toBe('editing.refused.locked');
    key(scroller, 'y');
    key(scroller, 'z');
    expect(live.textContent).toBe(first);
    expect(viewer.snapshot().editing).toBe(false);
  });

  test('a locked format, clear, or structure command names its command, not an edit', async () => {
    const { viewer, live } = await mount(
      buildXlsx({
        sheets: [
          { name: 'S', rows: row(1, [c('A1', 1)]), beforeMerges: '<sheetProtection sheet="1"/>' },
        ],
      })
    );
    const refusal = (r: { ok: boolean; refusal?: unknown }) => (r.ok ? null : r.refusal);
    expect(refusal(viewer.applyFormat({ bold: true }))).toMatchObject({
      code: 'locked',
      command: 'format',
    });
    expect(live.textContent).toBe('editing.lockedRefused.format');
    expect(refusal(viewer.clearValues())).toMatchObject({ code: 'locked', command: 'clear' });
    expect(live.textContent).toBe('editing.lockedRefused.clear');
    expect(refusal(viewer.insertRows())).toMatchObject({ code: 'locked', command: 'structure' });
    expect(live.textContent).toBe('editing.lockedRefused.structure');
    expect(refusal(viewer.enterValue('2'))).toEqual({
      code: 'locked',
      detail: expect.anything(),
      lock: 'cell',
    });
  });

  test('on a sheet that does not allow formatting, a clear of unlocked cells names formatting, and a locked cell names the lock', async () => {
    const { viewer, live } = await mount(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            rows: row(1, [c('A1', 1, { s: 1 }), c('B1', 2)]),
            beforeMerges: '<sheetProtection sheet="1"/>',
          },
        ],
        styles: {
          cellXfs: [
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">' +
              '<protection locked="0"/></xf>',
          ],
        },
      })
    );
    const refusal = (r: { ok: boolean; refusal?: unknown }) => (r.ok ? null : r.refusal);
    // A1 is unlocked: its formats are not cleared, because formatting is not allowed.
    viewer.select({ row: 0, col: 0 });
    expect(refusal(viewer.clearFormats())).toMatchObject({
      code: 'locked',
      command: 'clear',
      lock: 'format',
    });
    expect(live.textContent).toBe('editing.lockedRefused.clearFormats');
    expect(refusal(viewer.clearAll())).toMatchObject({ command: 'clearAll', lock: 'format' });
    expect(live.textContent).toBe('editing.lockedRefused.clearAll');
    expect(refusal(viewer.applyFormat({ bold: true }))).toMatchObject({
      command: 'format',
      lock: 'format',
    });
    expect(live.textContent).toBe('editing.lockedRefused.format');
    // B1 is locked: a clear of its value names the locked cell.
    viewer.select({ row: 0, col: 1 });
    expect(refusal(viewer.clearValues())).toMatchObject({ command: 'clear', lock: 'cell' });
    expect(live.textContent).toBe('editing.lockedRefused.clear');
  });

  test('Escape after a refused commit clears the refusal and its announcement', async () => {
    const { viewer, scroller, input, live } = await mount(sheet());
    key(scroller, '=');
    input.value = '=SUM(';
    key(input, 'Enter');
    expect(viewer.snapshot().refusal).not.toBeNull();
    expect(live.textContent).not.toBe('');
    key(input, 'Escape');
    expect(viewer.snapshot().refusal).toBeNull();
    expect(live.textContent).toBe('');
  });

  test('Shift+Space selects the rows, and Mod+Space the columns, and neither types', async () => {
    const { viewer, scroller } = await mount(sheet());
    viewer.select({ row: 1, col: 1 }, { row: 2, col: 2 });
    key(scroller, ' ', { shiftKey: true });
    expect(viewer.snapshot().editing).toBe(false);
    expect(viewer.snapshot().selectionRef).toBe('A2:XFD3');
    expect(viewer.snapshot().activeCell?.ref).toBe('B2');
    viewer.select({ row: 1, col: 1 });
    key(scroller, ' ', { ctrlKey: true });
    expect(viewer.snapshot().selectionRef).toBe('B1:B1048576');
    // A plain space still starts an edit with the space.
    viewer.select({ row: 0, col: 0 });
    key(scroller, ' ');
    expect(viewer.snapshot().editText).toBe(' ');
  });

  test('close drops a pending load, an open edit, and the history', async () => {
    const { viewer, scroller } = await mount(sheet());
    expect(viewer.enterValue('5').ok).toBe(true);
    expect(viewer.snapshot().canUndo).toBe(true);
    key(scroller, 'x');
    expect(viewer.snapshot().editing).toBe(true);
    viewer.close();
    const closed = viewer.snapshot();
    expect([closed.status, closed.editing, closed.canUndo, closed.sheets.length]).toEqual([
      'empty',
      false,
      false,
      0,
    ]);
    expect(viewer.undo().ok).toBe(false);
    expect(viewer.save()).toBeNull();
    // A load that has not settled when the editor closes is dropped.
    const opening = viewer.load(sheet());
    viewer.close();
    expect((await opening).status).toBe('empty');
    expect(viewer.snapshot().status).toBe('empty');
    // The editor opens again, with no edit and no history of the closed workbook.
    expect((await viewer.load(sheet())).status).toBe('ready');
    expect([viewer.snapshot().editing, viewer.snapshot().canUndo]).toEqual([false, false]);
    expect(content(viewer)).toBe('1');
  });

  test('Enter and Tab after an edit skip a hidden row and a hidden column', async () => {
    const { viewer, scroller, input } = await mount(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            rows: row(1, [c('A1', 1)]) + row(2, [c('A2', 2)], 'hidden="1"'),
            cols: '<cols><col min="2" max="2" width="9" hidden="1" customWidth="1"/></cols>',
          },
        ],
      })
    );
    key(scroller, 'x');
    key(input, 'Enter');
    expect(viewer.snapshot().activeCell?.ref).toBe('A3');
    viewer.select({ row: 0, col: 0 });
    key(scroller, 'y');
    key(input, 'Tab');
    expect(viewer.snapshot().activeCell?.ref).toBe('C1');
  });

  test('Enter and Tab after an edit move inside a multi-cell selection and keep it', async () => {
    const { viewer, scroller, input } = await mount(sheet());
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 1 });
    key(scroller, 'x');
    key(input, 'Enter');
    expect(viewer.snapshot().activeCell?.ref).toBe('A2');
    expect(viewer.snapshot().selectionRef).toBe('A1:B2');
    key(scroller, 'y');
    key(input, 'Tab');
    expect(viewer.snapshot().activeCell?.ref).toBe('B2');
    expect(viewer.snapshot().selectionRef).toBe('A1:B2');
    key(scroller, 'z');
    key(input, 'Tab', { shiftKey: true });
    expect(viewer.snapshot().activeCell?.ref).toBe('A2');
    expect(viewer.snapshot().selectionRef).toBe('A1:B2');
    // An arrow that commits still collapses the selection, as an arrow does without an edit.
    key(scroller, 'w');
    key(input, 'ArrowDown');
    expect(viewer.snapshot().selectionRef).toBe('A3');
  });

  test('Shift+Tab in column A leaves the grid; Tab inside it moves the cell', async () => {
    const { viewer, scroller } = await mount(sheet());
    // dispatchEvent returns false when the grid calls preventDefault.
    expect(key(scroller, 'Tab', { shiftKey: true })).toBe(true);
    expect(viewer.snapshot().activeCell?.ref).toBe('A1');
    expect(key(scroller, 'Tab')).toBe(false);
    expect(viewer.snapshot().activeCell?.ref).toBe('B1');
    expect(key(scroller, 'Tab', { shiftKey: true })).toBe(false);
    expect(viewer.snapshot().activeCell?.ref).toBe('A1');
  });

  test('Tab in the last column leaves the grid', async () => {
    const { viewer, scroller } = await mount(sheet());
    viewer.select({ row: 0, col: MAX_COLS - 1 });
    expect(key(scroller, 'Tab')).toBe(true);
    expect(viewer.snapshot().activeCell?.ref).toBe('XFD1');
    expect(key(scroller, 'Tab', { shiftKey: true })).toBe(false);
    expect(viewer.snapshot().activeCell?.ref).toBe('XFC1');
  });

  test('a text that looks like a number or a formula stays text through an edit', async () => {
    const { viewer, scroller, input } = await mount(sheet());
    viewer.select({ row: 3, col: 0 });
    for (const typed of ["'123", "'=A1", "'TRUE", "''quoted"]) {
      expect(viewer.enterValue(typed).ok).toBe(true);
      expect(content(viewer)).toBe(typed);
      key(scroller, 'F2');
      expect(input.value).toBe(typed);
      key(input, 'Enter');
      viewer.select({ row: 3, col: 0 });
      expect(content(viewer)).toBe(typed);
    }
    expect(viewer.enterValue('plain').ok).toBe(true);
    expect(content(viewer)).toBe('plain');
  });

  test('a literal _xHHHH_ in formula text is escaped, and reads back as typed', async () => {
    const { viewer } = await mount(sheet());
    viewer.select({ row: 3, col: 1 });
    expect(viewer.enterValue('="_x0041_"').ok).toBe(true);
    expect(content(viewer)).toBe('="_x0041_"');
    const saved = viewer.save()!;
    expect(strFromU8(unzipSync(saved)['xl/worksheets/sheet1.xml']!)).toContain(
      '<f>&quot;_x005F_x0041_&quot;</f>'
    );
  });

  test('Enter and Tab move the active cell inside a multi-cell selection and keep it', async () => {
    const { viewer, scroller } = await mount(sheet());
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 1 });
    expect(viewer.snapshot().selectionRef).toBe('A1:B2');
    expect(key(scroller, 'Enter')).toBe(false);
    expect(viewer.snapshot().activeCell?.ref).toBe('A2');
    key(scroller, 'Enter');
    expect(viewer.snapshot().activeCell?.ref).toBe('B1');
    key(scroller, 'Tab');
    expect(viewer.snapshot().activeCell?.ref).toBe('A2');
    key(scroller, 'Tab', { shiftKey: true });
    expect(viewer.snapshot().activeCell?.ref).toBe('B1');
    expect(viewer.snapshot().selectionRef).toBe('A1:B2');
  });

  test('Tab and Enter in a selection skip the cells that a merged range covers', async () => {
    const bytes = buildXlsx({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 2)]), merges: ['B1:C1', 'A3:A4'] }],
    });
    const { viewer, scroller } = await mount(bytes);
    viewer.select({ row: 0, col: 0 }, { row: 0, col: 3 });
    expect(viewer.snapshot().selectionRef).toBe('A1:D1');
    key(scroller, 'Tab');
    expect(viewer.snapshot().activeCell?.ref).toBe('B1');
    key(scroller, 'Tab');
    expect(viewer.snapshot().activeCell?.ref).toBe('D1');
    key(scroller, 'Tab', { shiftKey: true });
    expect(viewer.snapshot().activeCell?.ref).toBe('B1');
    viewer.select({ row: 1, col: 0 }, { row: 4, col: 0 });
    expect(viewer.snapshot().selectionRef).toBe('A2:A5');
    key(scroller, 'Enter');
    expect(viewer.snapshot().activeCell?.ref).toBe('A3');
    key(scroller, 'Enter');
    expect(viewer.snapshot().activeCell?.ref).toBe('A5');
  });

  test('the stale notice shows only on a sheet with a formula cell; the flag stays', async () => {
    const bytes = buildXlsx({
      sheets: [
        { name: 'S', rows: row(1, [c('A1', 1)]) },
        { name: 'T', rows: row(1, [c('A1', 2, { f: '1+1' })]) },
      ],
      calcPr: 'calcId="191029" fullCalcOnLoad="1"',
    });
    const { viewer } = await mount(bytes);
    expect(viewer.snapshot().stale).toBe(false);
    expect(viewer.setActiveSheet(1)).toBe(true);
    expect(viewer.snapshot().stale).toBe(true);
    const saved = readXlsx(viewer.save()!);
    expect(saved.ok && workbookIndex(saved.pkg)!.fullCalcOnLoad).toBe(true);
  });

  test('a formula is stored as x:f, and the workbook becomes stale', async () => {
    const { viewer } = await mount(sheet());
    expect(viewer.snapshot().stale).toBe(false);
    viewer.select({ row: 2, col: 0 });
    expect(viewer.enterValue('=sum(a1:b2)').ok).toBe(true);
    expect(content(viewer)).toBe('=SUM(A1:B2)');
    expect(viewer.snapshot().stale).toBe(true);
    // The save carries the formula and the stale flag.
    const saved = readXlsx(viewer.save()!);
    expect(saved.ok && workbookIndex(saved.pkg)!.fullCalcOnLoad).toBe(true);
  });

  test('an unreadable formula is refused visibly and the editor keeps the text', async () => {
    const { viewer, scroller, input, live } = await mount(sheet());
    key(scroller, '=');
    input.value = '=SUM(A1';
    key(input, 'Enter');
    const snap = viewer.snapshot();
    expect(snap.editing).toBe(true);
    expect(input.value).toBe('=SUM(A1');
    expect(snap.refusal?.code).toBe('unreadable-formula');
    expect(live.textContent).toBe('editing.refused.unreadable-formula');
    expect(snap.canUndo).toBe(false);
    expect(content(viewer)).toBe('1');
    input.value = '=SUM(A1)';
    key(input, 'Enter');
    expect(viewer.snapshot().refusal).toBeNull();
    expect(viewer.snapshot().editing).toBe(false);
  });

  test('a pointer press on the grid commits the open edit', async () => {
    const { viewer, scroller, input, host } = await mount(sheet());
    key(scroller, '7');
    const stage = host.querySelector('.xlsx-stage') as HTMLElement;
    const W = host.ownerDocument.defaultView as unknown as typeof globalThis;
    stage.dispatchEvent(
      new W.PointerEvent('pointerdown', { bubbles: true, button: 0, clientX: 200, clientY: 100 })
    );
    expect(viewer.snapshot().editing).toBe(false);
    expect(input.hidden).toBe(true);
    viewer.select({ row: 0, col: 0 });
    expect(content(viewer)).toBe('7');
  });

  test('an input method composes without a commit', async () => {
    const { viewer, scroller, input } = await mount(sheet());
    key(scroller, 'Process', { isComposing: true });
    expect(viewer.snapshot().editing).toBe(true);
    expect(input.value).toBe('');
    input.dispatchEvent(
      new (input.ownerDocument.defaultView as never as typeof globalThis).Event('compositionstart')
    );
    input.value = 'か';
    key(input, 'Enter', { isComposing: true });
    expect(viewer.snapshot().editing).toBe(true);
    input.dispatchEvent(
      new (input.ownerDocument.defaultView as never as typeof globalThis).Event('compositionend')
    );
    key(input, 'Enter');
    viewer.select({ row: 0, col: 0 });
    expect(content(viewer)).toBe('か');
  });

  test('the history keys in an open editor belong to the input: no store step, no cancel', async () => {
    const { viewer, scroller, input } = await mount(sheet());
    expect(viewer.enterValue('8').ok).toBe(true);
    const revision = viewer.snapshot().revision;
    // An untouched F2 edit: Mod+Z, Mod+Shift+Z, and Mod+Y pass to the input, not prevented.
    key(scroller, 'F2');
    expect(key(input, 'z', { ctrlKey: true })).toBe(true);
    expect(key(input, 'z', { ctrlKey: true, shiftKey: true })).toBe(true);
    expect(key(input, 'y', { ctrlKey: true })).toBe(true);
    expect(viewer.snapshot().editing).toBe(true);
    expect(viewer.snapshot().revision).toBe(revision);
    viewer.cancelEdit();
    expect(content(viewer)).toBe('8');
    // With no edit open, the key reaches the store history.
    key(scroller, 'z', { ctrlKey: true });
    expect(content(viewer)).toBe('1');
  });

  test('Mod+Z after typing in an empty cell and deleting it keeps the edit and the history', async () => {
    const { viewer, scroller, input } = await mount(sheet());
    viewer.select({ row: 4, col: 4 });
    key(scroller, 'a');
    input.value = 'ab';
    input.value = '';
    expect(key(input, 'z', { ctrlKey: true })).toBe(true);
    expect(viewer.snapshot().editing).toBe(true);
    viewer.select({ row: 0, col: 0 });
    expect(content(viewer)).toBe('1');
  });

  test('F2, typing, Mod+Z, and Mod+Shift+Z keep the edit open and run no store redo', async () => {
    const { viewer, scroller, input } = await mount(sheet());
    expect(viewer.enterValue('8').ok).toBe(true);
    expect(viewer.undo().ok).toBe(true);
    expect(viewer.snapshot().canRedo).toBe(true);
    key(scroller, 'F2');
    input.value = '1!';
    key(input, 'z', { ctrlKey: true });
    key(input, 'z', { ctrlKey: true, shiftKey: true });
    expect(viewer.snapshot().editing).toBe(true);
    expect(viewer.snapshot().canRedo).toBe(true);
    viewer.cancelEdit();
    expect(content(viewer)).toBe('1');
  });

  test('Mod+Z after typing over a cell stays in the editor', async () => {
    const { viewer, scroller, input } = await mount(sheet());
    expect(viewer.enterValue('8').ok).toBe(true);
    key(scroller, 'x');
    expect(key(input, 'z', { ctrlKey: true })).toBe(true);
    expect(viewer.snapshot().editing).toBe(true);
    expect(viewer.snapshot().canUndo).toBe(true);
    expect(content(viewer)).toBe('8');
  });

  test('the editor input is named, has no autocomplete, and has an accessible name', async () => {
    const { input } = await mount(sheet());
    expect(input.getAttribute('name')).toBe('xlsx-cell-editor');
    expect(input.getAttribute('autocomplete')).toBe('off');
    expect(input.getAttribute('aria-label')).toBe('editing.cellEditor');
  });

  test('a key whose target is not the grid itself does nothing', async () => {
    const { viewer, host } = await mount(sheet());
    key(host.querySelector('.xlsx-extent')!, '5');
    expect(viewer.snapshot().editing).toBe(false);
  });
});

describe('the selection announcement', () => {
  test('a range of more than one cell announces its reference; one cell announces the cell', async () => {
    const merged = buildXlsx({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 2)]), merges: ['C3:D4'] }],
    });
    const { viewer, live } = await mount(merged, withParams);
    viewer.select({ row: 1, col: 1 }, { row: 1, col: 3 });
    expect(live.textContent).toBe('grid.selectedRange {"range":"B2:D2"}');
    viewer.select({ row: 0, col: 1 });
    expect(live.textContent).toBe('grid.activeCell {"ref":"B1","content":"2"}');
    // A merged cell is one cell, so it announces as the cell.
    viewer.select({ row: 2, col: 2 });
    expect(live.textContent).toBe('grid.activeCellEmpty {"ref":"C3","content":""}');
  });
});

describe('the announcement of the active cell', () => {
  test('a cell announces the text the grid shows, with its number format', async () => {
    const bytes = buildXlsx({
      styles: {
        cellXfs: [
          '<xf numFmtId="9" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
        ],
      },
      sheets: [{ name: 'S', rows: row(1, [c('A1', 0.5, { s: 1 })]) }],
    });
    const { viewer, live } = await mount(bytes, withParams);
    viewer.select({ row: 0, col: 0 });
    expect(live.textContent).toBe('grid.activeCell {"ref":"A1","content":"50%"}');
  });

  test('a formula cell announces its cached result, then its formula', async () => {
    const bytes = buildXlsx({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 2), c('B1', 4, { f: 'A1*2' })]) }],
    });
    const { viewer, live } = await mount(bytes, withParams);
    viewer.select({ row: 0, col: 1 });
    expect(live.textContent).toBe(
      'grid.activeFormula {"ref":"B1","content":"4","formula":"=A1*2"}'
    );
  });

  test('a stale formula announces the stale words, and one with no result says so', async () => {
    const bytes = buildXlsx({
      calcPr: 'fullCalcOnLoad="1"',
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 2), c('B1', 4, { f: 'A1*2' }), c('C1', undefined, { f: 'A1+1' })]),
        },
      ],
    });
    const { viewer, live } = await mount(bytes, withParams);
    viewer.select({ row: 0, col: 1 });
    const said = 'grid.activeFormula {"ref":"B1","content":"4","formula":"=A1*2"}';
    expect(live.textContent).toBe(`grid.staleResult ${JSON.stringify({ text: said })}`);
    viewer.select({ row: 0, col: 2 });
    const none = 'grid.activeFormulaNoResult {"ref":"C1","content":"","formula":"=A1+1"}';
    expect(live.textContent).toBe(`grid.staleResult ${JSON.stringify({ text: none })}`);
    // A value cell on the same sheet has no stale words.
    viewer.select({ row: 0, col: 0 });
    expect(live.textContent).toBe('grid.activeCell {"ref":"A1","content":"2"}');
  });
});

describe('the content revision and the workbook kind', () => {
  test('an edit, an undo, and a redo each add one; a view change adds none', async () => {
    const { viewer } = await mount(sheet());
    expect(viewer.snapshot().revision).toBe(0);
    expect(viewer.snapshot().workbookKind).toBe('workbook');
    viewer.select({ row: 1, col: 1 });
    viewer.setZoom(1.5);
    expect(viewer.beginEdit('9', 'enter').ok).toBe(true);
    expect(viewer.snapshot().revision).toBe(0);
    expect(viewer.commitEdit().ok).toBe(true);
    expect(viewer.snapshot().revision).toBe(1);
    expect(viewer.undo().ok).toBe(true);
    expect(viewer.snapshot().revision).toBe(2);
    expect(viewer.redo().ok).toBe(true);
    expect(viewer.snapshot().revision).toBe(3);
    // A new open starts again at 0.
    await viewer.load(sheet());
    expect(viewer.snapshot().revision).toBe(0);
  });

  test('setFonts replaces the faces in place, and a failed load keeps the old faces', async () => {
    const { viewer, host } = await mount(sheet());
    expect(viewer.enterValue('5').ok).toBe(true);
    const failure = await viewer.setFonts(() => Promise.reject(new Error('offline')));
    expect(failure?.code).toBe('fonts');
    expect(failure?.detail).toContain('offline');
    expect(viewer.snapshot().status).toBe('ready');
    const copies = FONTS.map((f) => ({ ...f, bytes: new Uint8Array(f.bytes) }));
    expect(await viewer.setFonts(copies)).toBeNull();
    // The same faces in a new list load nothing; a locale equal in value publishes nothing.
    expect(await viewer.setFonts(copies.map((f) => ({ ...f })))).toBeNull();
    let published = 0;
    const off = viewer.subscribe(() => (published += 1));
    viewer.setNumberLocale({ decimal: '.', group: ',' });
    expect(published).toBe(0);
    off();
    flushEditor(viewer);
    expect(viewer.snapshot().canUndo).toBe(true);
    expect(host.querySelector('canvas')).not.toBeNull();
    viewer.setNumberLocale({ decimal: ',', group: '.' });
    expect(viewer.enterValue('2,5').ok).toBe(true);
    expect(content(viewer)).toBe('2,5');
  });

  test('a good save clears an older refusal from the snapshot and the strip', async () => {
    const { viewer } = await mount(sheet());
    expect(viewer.undo().ok).toBe(false);
    expect(viewer.snapshot().refusal?.code).toBe('nothing-to-undo');
    expect(viewer.save()).not.toBeNull();
    expect(viewer.snapshot().refusal).toBeNull();
  });

  test('a template keeps its kind through an edit and a save', async () => {
    const files = unzipSync(sheet());
    const types = new TextDecoder().decode(files['[Content_Types].xml']!);
    files['[Content_Types].xml'] = new TextEncoder().encode(
      types.replace('spreadsheetml.sheet.main+xml', 'spreadsheetml.template.main+xml')
    );
    const { viewer } = await mount(zipSync(files));
    expect(viewer.snapshot().workbookKind).toBe('template');
    expect(viewer.enterValue('5').ok).toBe(true);
    const saved = readXlsx(viewer.save()!);
    if (!saved.ok) throw new Error('reopen');
    const reopened = createXlsxEditor({ fonts: FONTS });
    viewers.push(reopened);
    await reopened.load(viewer.save()!);
    expect(reopened.snapshot().workbookKind).toBe('template');
  });
});

describe('sheet names in a formula', () => {
  test('a sheet name with the shape of a cell, an R1C1 reference, or a boolean is written in quotes, and save and reopen keep it', async () => {
    // `XFE1` is past the last column, so it is no cell, and its name needs no quotes.
    const names = ['S', 'Jan', 'A1', 'TRUE', 'R1C1', 'XFE1'];
    const bytes = buildXlsx({
      sheets: names.map((name) => ({ name, rows: row(1, [c('A1', 1), c('B1', 2)]) })),
    });
    const { viewer } = await mount(bytes);
    viewer.select({ row: 0, col: 2 });
    expect(viewer.enterValue('=TRUE!A1+R1C1!A1+SUM(Jan:A1!B1)+A1!B1+Jan!B1+XFE1!A1').ok).toBe(true);
    const formula = "'TRUE'!A1+'R1C1'!A1+SUM('Jan:A1'!B1)+'A1'!B1+Jan!B1+XFE1!A1";
    viewer.select({ row: 0, col: 2 });
    expect(content(viewer)).toBe(`=${formula}`);
    const saved = viewer.save()!;
    const files = unzipSync(saved);
    expect(strFromU8(files['xl/worksheets/sheet1.xml']!)).toContain(
      `<f>${formula.replaceAll("'", '&apos;')}</f>`
    );
    const reopened = createXlsxEditor({ fonts: FONTS });
    viewers.push(reopened);
    await reopened.load(saved);
    reopened.select({ row: 0, col: 2 });
    expect(reopened.snapshot().activeCell?.content).toBe(`=${formula}`);
  });
});

describe('Delete and undo', () => {
  test('Delete clears the values of the selection, keeps the format, and undoes in one step', async () => {
    const { viewer, scroller } = await mount(sheet());
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    key(scroller, 'Delete');
    viewer.select({ row: 1, col: 0 });
    expect(content(viewer)).toBe('');
    const saved = readXlsx(viewer.save()!);
    expect(saved.ok).toBe(true);
    key(scroller, 'z', { ctrlKey: true });
    // Undo selects the range that it changed again.
    expect(viewer.snapshot().selectionRef).toBe('A1:A2');
    expect(content(viewer)).toBe('1');
    viewer.select({ row: 1, col: 0 });
    expect(content(viewer)).toBe('5');
    expect(viewer.snapshot().canUndo).toBe(false);
  });

  test('undo and redo with an empty history announce their refusal once', async () => {
    const { viewer, host } = await mount(sheet());
    const live = () => host.querySelector('.xlsx-live')?.textContent;
    viewer.setZoom(2);
    const r = viewer.undo();
    expect(r.ok ? null : r.refusal.code).toBe('nothing-to-undo');
    // The refusal replaces the old text, and a second press does not announce it again.
    expect(viewer.snapshot().refusal?.code).toBe('nothing-to-undo');
    const text = live();
    expect(text).not.toContain('200');
    viewer.undo();
    expect(live()).toBe(text);
    const redo = viewer.redo();
    expect(redo.ok ? null : redo.refusal.code).toBe('nothing-to-redo');
    expect(live()).not.toBe(text);
  });

  test('an edit in a fixture with formulas keeps the cached values and marks them stale', async () => {
    const { viewer, host } = await mount(fixture('basic-values'), withParams);
    viewer.select({ row: 1, col: 0 });
    expect(viewer.enterValue('7').ok).toBe(true);
    flushEditor(viewer);
    expect(viewer.snapshot().stale).toBe(true);
    // The accessible cell reads the cached value with the stale words.
    expect(host.querySelector('[data-cell="A5"]')?.textContent).toBe(
      'grid.staleResult {"text":"84"}'
    );
  });
});

describe('cell input', () => {
  test('reads the typed text', () => {
    expect(parseCellInput('', EN_US)).toEqual({ type: 'empty' });
    expect(parseCellInput('=A1', EN_US)).toEqual({ type: 'formula', text: 'A1' });
    expect(parseCellInput('=', EN_US)).toEqual({ type: 'string', value: '=' });
    expect(parseCellInput("'42", EN_US)).toEqual({ type: 'string', value: '42' });
    expect(parseCellInput("'", EN_US)).toEqual({ type: 'empty' });
    expect(parseCellInput('-2.5e3', EN_US)).toEqual({ type: 'number', value: -2500 });
    expect(parseCellInput('true', EN_US)).toEqual({ type: 'boolean', value: true });
    expect(parseCellInput(' 4', EN_US)).toEqual({ type: 'string', value: ' 4' });
    expect(parseCellInput('1,5', { ...EN_US, decimal: ',' })).toEqual({
      type: 'number',
      value: 1.5,
    });
    expect(parseCellInput('1.5', { ...EN_US, decimal: ',' })).toEqual({
      type: 'string',
      value: '1.5',
    });
  });

  test('every refusal code has a message in the shared catalog', () => {
    const en = JSON.parse(readFileSync(`${import.meta.dir}/../../../../i18n/en.json`, 'utf8')) as {
      editing: { refused: Record<string, string> };
    };
    expect(Object.keys(en.editing.refused).sort()).toEqual([...EDIT_REFUSAL_CODES].sort());
  });
});
