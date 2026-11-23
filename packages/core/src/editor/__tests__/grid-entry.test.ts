// The live region when the focus enters the grid from outside it (`LiveAnnouncer.entered`): it
// says the selection with its active cell, as a grid of the WAI-ARIA Authoring Practices reads
// the cell that has the focus. A step that writes the region and moves the focus says its own
// words only, and a press on the grid says only the cell that it selects.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { createXlsxEditor, type XlsxEditorInstance } from '../index.ts';
import { REANNOUNCE } from '../cell-editing.ts';
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

/** A key and its parameters, so a test can read the words of an announcement. */
const translate = (key: string, params?: Readonly<Record<string, string | number>>) =>
  params ? `${key} ${Object.values(params).join(' | ')}` : key;

/** A1 holds 0.5 and A2 holds 3, on the sheets S and T. */
const workbook = () =>
  buildXlsx({
    sheets: [
      { name: 'S', rows: row(1, [c('A1', 0.5)]) + row(2, [c('A2', 3)]) },
      { name: 'T', rows: row(1, [c('A1', 7)]) },
    ],
  });

/** Mount an editor beside a button outside it, and record each text of the live region. */
async function mount() {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const outside = doc.createElement('button');
  const host = doc.createElement('div');
  doc.body.append(outside, host);
  const viewer = createXlsxEditor({ fonts: FONTS, translate });
  viewers.push(viewer);
  viewer.attach(host);
  const scroller = host.querySelector('.xlsx-scroller') as HTMLElement;
  Object.defineProperty(scroller, 'clientWidth', { get: () => 800 });
  Object.defineProperty(scroller, 'clientHeight', { get: () => 400 });
  await viewer.load(workbook());
  flushEditor(viewer);
  const live = host.querySelector('.xlsx-live') as HTMLElement;
  let proto: object | null = live;
  let desc: PropertyDescriptor | undefined;
  while (proto && !desc) {
    desc = Object.getOwnPropertyDescriptor(proto, 'textContent');
    proto = Object.getPrototypeOf(proto);
  }
  const said: string[] = [];
  Object.defineProperty(live, 'textContent', {
    get: () => desc!.get!.call(live),
    set: (text: string) => {
      said.push(text);
      desc!.set!.call(live, text);
    },
  });
  const W = win as unknown as typeof globalThis;
  /**
   * The texts said since the last call, after the editor paints. A timer after it ends the
   * step, as the end of a task does in a browser.
   */
  const heard = async () => {
    flushEditor(viewer);
    await new Promise((done) => W.setTimeout(done, 0));
    return said.splice(0).filter((text) => text !== '');
  };
  return { viewer, heard, host, scroller, outside, live, W };
}

describe('the focus that enters the grid', () => {
  test('says the active cell, again when the words are the same', async () => {
    const { viewer, heard, scroller, outside } = await mount();
    viewer.select({ row: 0, col: 0 });
    await heard();
    outside.focus();
    scroller.focus();
    // The region holds these words from the select, so they change by one invisible character.
    expect(await heard()).toEqual([`grid.activeCell A1 | 0.5${REANNOUNCE}`]);
    outside.focus();
    scroller.focus();
    expect(await heard()).toEqual(['grid.activeCell A1 | 0.5']);
  });

  test('says a range with its active cell', async () => {
    const { viewer, heard, scroller, outside } = await mount();
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    await heard();
    outside.focus();
    scroller.focus();
    expect(await heard()).toEqual([
      'grid.selectionCell grid.selectedRange A1:A2 | grid.activeCell A1 | 0.5',
    ]);
  });

  test('says nothing more when the step that moved the focus wrote the region', async () => {
    const { viewer, heard, scroller, outside } = await mount();
    viewer.select({ row: 0, col: 0 });
    outside.focus();
    await heard();
    // A sheet tab says its sheet and gives the focus to the grid in one step.
    expect(viewer.setActiveSheet(1)).toBe(true);
    scroller.focus();
    expect(await heard()).toEqual(['grid.sheetActive T | grid.activeCell A1 | 7']);
  });

  test('a change that waits for its paint says the cell when it keeps the words', async () => {
    const { viewer, heard, scroller, outside } = await mount();
    viewer.select({ row: 0, col: 0 });
    outside.focus();
    await heard();
    // A toolbar menu applies bold and gives the grid the focus in one step. Bold keeps the words
    // of the active cell, so the paint says the cell that the focus entered.
    expect(viewer.toggleFormat('bold').ok).toBe(true);
    scroller.focus();
    expect(await heard()).toEqual([`grid.activeCell A1 | 0.5${REANNOUNCE}`]);
  });

  test('a change on a range that keeps the words says the range with its active cell', async () => {
    const { viewer, heard, scroller, outside } = await mount();
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    outside.focus();
    await heard();
    expect(viewer.toggleFormat('bold').ok).toBe(true);
    scroller.focus();
    expect(await heard()).toEqual([
      'grid.selectionCell grid.selectedRange A1:A2 | grid.activeCell A1 | 0.5',
    ]);
  });

  test('a message of the same step ends the entry', async () => {
    const { viewer, heard, scroller, outside } = await mount();
    viewer.select({ row: 0, col: 0 });
    outside.focus();
    await heard();
    expect(viewer.toggleFormat('bold').ok).toBe(true);
    scroller.focus();
    // The undo names its step, so the paint says nothing, and a later change has no entry.
    viewer.undo();
    expect(await heard()).toEqual(['grid.undoneAt grid.activeCell A1 | 0.5']);
    expect(viewer.toggleFormat('italic').ok).toBe(true);
    expect(await heard()).toEqual([]);
  });

  test('a change that waits for its paint keeps its confirmation', async () => {
    const { viewer, heard, scroller, outside } = await mount();
    viewer.select({ row: 0, col: 0 });
    outside.focus();
    await heard();
    // The format 0.0 keeps the words 0.5, so the paint confirms the change.
    expect(viewer.applyFormat({ numberFormat: { code: '0.0' } }).ok).toBe(true);
    scroller.focus();
    expect(await heard()).toEqual(['grid.numberFormatChanged']);
  });

  test('a refusal before the focus in the same task is the words of the step', async () => {
    const { viewer, heard, scroller, outside } = await mount();
    viewer.select({ row: 0, col: 0 });
    outside.focus();
    await heard();
    // A refused rename says its reason when its field loses the focus. A browser runs a
    // microtask checkpoint between the blur and the focus listeners of the same task.
    expect(viewer.renameSheet(1, 'S').ok).toBe(false);
    await Promise.resolve();
    scroller.focus();
    const refusal = await heard();
    expect(refusal).toHaveLength(1);
    expect(refusal[0]).not.toContain('grid.activeCell');
    // A later task is a new step.
    outside.focus();
    scroller.focus();
    expect(await heard()).toEqual(['grid.activeCell A1 | 0.5']);
  });

  test('a blur of the cell editor before the grid takes the focus is no entry', async () => {
    const { viewer, heard, scroller, host } = await mount();
    viewer.select({ row: 0, col: 0 });
    expect(viewer.beginEdit().ok).toBe(true);
    await heard();
    // A browser can blur the hidden cell editor first: the focus then comes from no element.
    (host.querySelector('.xlsx-cell-editor') as HTMLElement).blur();
    scroller.focus();
    expect(await heard()).toEqual([]);
  });

  test('a focus that left the grid for a control and comes back is an entry', async () => {
    const { viewer, heard, scroller, outside } = await mount();
    viewer.select({ row: 0, col: 0 });
    scroller.focus();
    await heard();
    // In one task: the grid gives the focus to a control, the control loses it to no element,
    // and the grid takes it again. Only a blur in the grid to no element marks the task.
    outside.focus();
    outside.blur();
    scroller.focus();
    expect(await heard()).toEqual([`grid.activeCell A1 | 0.5${REANNOUNCE}`]);
  });

  test('a focus from the cell editor is no entry', async () => {
    const { viewer, heard, scroller, host } = await mount();
    viewer.select({ row: 0, col: 0 });
    expect(viewer.beginEdit().ok).toBe(true);
    await heard();
    viewer.cancelEdit();
    expect(host.ownerDocument.activeElement).toBe(scroller);
    expect(await heard()).toEqual([]);
  });

  test('a press on the grid says only the cell that it selects', async () => {
    const { viewer, heard, host, scroller, outside, W } = await mount();
    viewer.select({ row: 1, col: 0 });
    outside.focus();
    await heard();
    const stage = host.querySelector('.xlsx-stage') as HTMLElement;
    stage.dispatchEvent(
      new W.PointerEvent('pointerdown', { bubbles: true, button: 0, clientX: 60, clientY: 5 })
    );
    const words = await heard();
    expect(viewer.snapshot().selection?.active).toEqual({ row: 0, col: 0 });
    expect(words).toEqual(['grid.activeCell A1 | 0.5']);
    // The press is over: a later entry says the cell.
    outside.focus();
    scroller.focus();
    expect(await heard()).toEqual([`grid.activeCell A1 | 0.5${REANNOUNCE}`]);
  });

  test('a press on the active cell from outside the grid says it again', async () => {
    const { viewer, heard, host, outside, W } = await mount();
    viewer.select({ row: 0, col: 0 });
    outside.focus();
    await heard();
    const stage = host.querySelector('.xlsx-stage') as HTMLElement;
    stage.dispatchEvent(
      new W.PointerEvent('pointerdown', { bubbles: true, button: 0, clientX: 60, clientY: 5 })
    );
    expect(await heard()).toEqual([`grid.activeCell A1 | 0.5${REANNOUNCE}`]);
    // A second press inside the grid is no entry, and the same selection says nothing.
    stage.dispatchEvent(
      new W.PointerEvent('pointerdown', { bubbles: true, button: 0, clientX: 60, clientY: 5 })
    );
    expect(await heard()).toEqual([]);
  });

  test('a tap on the grid says only the cell that it selects', async () => {
    const { viewer, heard, host, scroller, outside, W } = await mount();
    viewer.select({ row: 1, col: 0 });
    outside.focus();
    await heard();
    const stage = host.querySelector('.xlsx-stage') as HTMLElement;
    // A touch selects and focuses the grid on its release.
    for (const type of ['pointerdown', 'pointerup']) {
      stage.dispatchEvent(
        new W.PointerEvent(type, {
          bubbles: true,
          cancelable: true,
          button: 0,
          pointerId: 7,
          pointerType: 'touch',
          clientX: 60,
          clientY: 5,
        })
      );
    }
    const words = await heard();
    expect(viewer.snapshot().selection?.active).toEqual({ row: 0, col: 0 });
    expect(words).toEqual(['grid.activeCell A1 | 0.5']);
    // The tap is over: a later entry says the cell.
    outside.focus();
    scroller.focus();
    expect(await heard()).toEqual([`grid.activeCell A1 | 0.5${REANNOUNCE}`]);
  });
});
