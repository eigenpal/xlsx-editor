// The live region after an edit: when an edit changes the shown text of the active
// cell, the region says the new text once. A format change, a typed entry, an undo, and a redo
// count. An edit that leaves the shown text as it was says nothing.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import {
  createXlsxEditor,
  MAX_CONTENT_FIELD_CHARS,
  relabelEditor,
  runToolbarCommand,
  type XlsxEditorInstance,
} from '../index.ts';
import { flushEditor } from './flush-editor.ts';
import type { FontSource } from '../../layout/index.ts';
import { readXlsx, sheetByName, workbookIndex } from '../../store/index.ts';
import { EN_US } from '../../cascade/index.ts';
import { cellAnnouncement } from '../cell-announcement.ts';
import { englishText } from '../english.ts';

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

/** A1 holds 0.5, A2 holds 3, and B1 holds the text `x`. */
const sheet = () =>
  buildXlsx({
    sheets: [{ name: 'S', rows: row(1, [c('A1', 0.5), c('B1', 'x')]) + row(2, [c('A2', 3)]) }],
  });

/** The prefix of each key, so a test can switch the catalog. */
let lang = '';
afterEach(() => {
  lang = '';
});

/** A key and its parameters, so a test can read the words of an announcement. */
const translate = (key: string, params?: Readonly<Record<string, string | number>>) =>
  lang + (params ? `${key} ${Object.values(params).join(' | ')}` : key);

/** Mount an editor and record each text that the live region takes. */
async function mount(bytes = sheet()) {
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
  /** The texts said since the last call, after the editor paints. */
  const heard = () => {
    flushEditor(viewer);
    return said.splice(0).filter((text) => text !== '');
  };
  return { viewer, heard, host };
}

/** Press a key on the grid, Enter by default, with no cell editor open. */
function pressKey(host: HTMLElement, key = 'Enter') {
  const scroller = host.querySelector('.xlsx-scroller') as HTMLElement;
  const win = scroller.ownerDocument.defaultView as unknown as typeof globalThis;
  const init = { key, bubbles: true, cancelable: true };
  scroller.dispatchEvent(new win.KeyboardEvent('keydown', init));
}

describe('live region after an edit', () => {
  test('a commit that changes only the spaces of the active cell says the new text', async () => {
    const bytes = (v: string) => buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', v)]) }] });
    for (const [from, to] of [
      ['a b', 'a  b'],
      [' a', 'a'],
    ] as const) {
      const { viewer, heard } = await mount(bytes(from));
      viewer.select({ row: 0, col: 0 });
      heard();
      expect(viewer.beginEdit().ok).toBe(true);
      expect(viewer.setEditText(to).ok).toBe(true);
      heard();
      expect(viewer.commitEdit().ok).toBe(true);
      expect(heard()).toEqual([`grid.activeCell A1 | ${to}`]);
    }
  });

  test('a number format change that commits typed spaces says the new text', async () => {
    const bytes = buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 'a b')]) }] });
    const { viewer, heard } = await mount(bytes);
    viewer.select({ row: 0, col: 0 });
    heard();
    expect(viewer.beginEdit().ok).toBe(true);
    expect(viewer.setEditText('a  b').ok).toBe(true);
    heard();
    expect(viewer.applyFormat({ numberFormat: { id: 10 } }).ok).toBe(true);
    expect(heard()).toEqual(['grid.activeCell A1 | a  b']);
  });

  test('a number format change says the new shown text of the active cell once', async () => {
    const { viewer, heard } = await mount();
    viewer.select({ row: 0, col: 0 });
    heard();
    expect(viewer.applyFormat({ numberFormat: { id: 9 } }).ok).toBe(true);
    expect(heard()).toEqual(['grid.activeCell A1 | 50%']);
  });

  test('a typed entry in the active cell says its shown text once', async () => {
    const { viewer, heard } = await mount();
    viewer.select({ row: 1, col: 0 });
    heard();
    expect(viewer.enterValue('12%')).toEqual({ ok: true });
    expect(heard()).toEqual(['grid.activeCell A2 | 12%']);
  });

  test('a bold toggle or a typed entry that keeps the shown text says nothing', async () => {
    const { viewer, heard } = await mount();
    viewer.select({ row: 0, col: 0 });
    heard();
    expect(viewer.toggleFormat('bold').ok).toBe(true);
    expect(heard()).toEqual([]);
    expect(viewer.enterValue('0.5')).toEqual({ ok: true });
    expect(heard()).toEqual([]);
  });

  test('an undo and a redo say their step with the new shown text, once each', async () => {
    const { viewer, heard } = await mount();
    viewer.select({ row: 0, col: 0 });
    expect(viewer.applyFormat({ numberFormat: { id: 9 } }).ok).toBe(true);
    heard();
    expect(viewer.undo().ok).toBe(true);
    expect(heard()).toEqual(['grid.undoneAt grid.activeCell A1 | 0.5']);
    expect(viewer.redo().ok).toBe(true);
    expect(heard()).toEqual(['grid.redoneAt grid.activeCell A1 | 50%']);
  });

  test('a paste says the pasted selection once', async () => {
    const { viewer, heard } = await mount();
    viewer.select({ row: 0, col: 0 });
    heard();
    expect(viewer.paste({ text: '7' }).ok).toBe(true);
    expect(heard()).toEqual(['grid.activeCell A1 | 7']);
  });

  test('a paste that grows from one cell to a range says the range and the active cell', async () => {
    const { viewer, heard } = await mount();
    viewer.select({ row: 0, col: 0 });
    heard();
    expect(viewer.paste({ text: '1\t2\r\n3\t4' }).ok).toBe(true);
    flushEditor(viewer);
    expect(heard()).toEqual([
      'grid.selectionCell grid.selectedRange A1:B2 | grid.activeCell A1 | 1',
    ]);
  });

  test('a paste over a selection of its own shape says the active cell', async () => {
    const { viewer, heard } = await mount();
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 1 });
    heard();
    expect(viewer.paste({ text: '1\t2\r\n3\t4' }).ok).toBe(true);
    expect(heard()).toEqual([
      'grid.selectionCell grid.selectedRange A1:B2 | grid.activeCell A1 | 1',
    ]);
  });

  test('a paste over its own shape that keeps the active cell says the range and the cell', async () => {
    const { viewer, heard } = await mount();
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 1 });
    heard();
    expect(viewer.paste({ text: '0.5\t9\r\n3\t9' }).ok).toBe(true);
    expect(heard()).toEqual([
      'grid.selectionCell grid.selectedRange A1:B2 | grid.activeCell A1 | 0.5',
    ]);
  });

  test('Enter inside a range says the range and the new active cell', async () => {
    const { viewer, heard } = await mount();
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    heard();
    expect(viewer.beginEdit('5').ok).toBe(true);
    expect(viewer.commitEdit('down').ok).toBe(true);
    expect(heard()).toEqual([
      'grid.selectionCell grid.selectedRange A1:A2 | grid.activeCell A2 | 3',
    ]);
  });

  test('each Enter inside a range says the range and the new active cell', async () => {
    const { viewer, heard } = await mount();
    viewer.select({ row: 0, col: 0 }, { row: 2, col: 0 });
    heard();
    for (const text of ['5', '6']) {
      expect(viewer.beginEdit(text).ok).toBe(true);
      expect(viewer.commitEdit('down').ok).toBe(true);
    }
    flushEditor(viewer);
    const said = heard();
    expect(said).toHaveLength(2);
    expect(said[0]).toBe('grid.selectionCell grid.selectedRange A1:A3 | grid.activeCell A2 | 3');
    expect(said[1]).toBe(
      'grid.selectionCell grid.selectedRange A1:A3 | grid.activeCellEmpty A3 | '
    );
  });

  test('Enter inside a range after a format change says the new active cell', async () => {
    const { viewer, heard } = await mount();
    viewer.select({ row: 0, col: 0 }, { row: 2, col: 0 });
    expect(viewer.applyFormat({ numberFormat: { id: 9 } }).ok).toBe(true);
    flushEditor(viewer);
    heard();
    expect(viewer.commitEdit('down').ok).toBe(true);
    flushEditor(viewer);
    expect(heard()).toEqual([
      'grid.selectionCell grid.selectedRange A1:A3 | grid.activeCell A2 | 300%',
    ]);
  });

  test('each Enter key inside a range with no edit says the new active cell', async () => {
    const { viewer, heard, host } = await mount();
    viewer.select({ row: 0, col: 0 }, { row: 2, col: 0 });
    heard();
    pressKey(host);
    expect(heard()).toEqual([
      'grid.selectionCell grid.selectedRange A1:A3 | grid.activeCell A2 | 3',
    ]);
    pressKey(host);
    expect(heard()).toEqual([
      'grid.selectionCell grid.selectedRange A1:A3 | grid.activeCellEmpty A3 | ',
    ]);
  });

  test('a select that keeps the range and the active cell after Enter says nothing', async () => {
    const { viewer, heard, host } = await mount();
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    pressKey(host);
    heard();
    viewer.select({ row: 1, col: 0 }, { row: 0, col: 0 });
    expect(heard()).toEqual([]);
  });

  test('Enter inside a range after a sheet switch says the new active cell', async () => {
    const two = buildXlsx({
      sheets: [
        { name: 'S1', rows: row(1, [c('A1', 5)]) },
        { name: 'S2', rows: row(1, [c('A1', 5)]) + row(2, [c('A2', 3)]) },
      ],
    });
    const { viewer, heard, host } = await mount(two);
    expect(viewer.setActiveSheet(1)).toBe(true);
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    expect(viewer.setActiveSheet(0)).toBe(true);
    expect(viewer.setActiveSheet(1)).toBe(true);
    heard();
    pressKey(host);
    expect(heard()).toEqual([
      'grid.selectionCell grid.selectedRange A1:A2 | grid.activeCell A2 | 3',
    ]);
  });

  test('a select of the same cell during an edit of another cell says nothing', async () => {
    const { viewer, heard } = await mount();
    viewer.select({ row: 0, col: 0 });
    heard();
    expect(viewer.enterValue('9', { row: 0, col: 1 }).ok).toBe(true);
    viewer.select({ row: 0, col: 0 });
    expect(heard()).toEqual([]);
  });

  test('a new range during an edit of another cell says only the range', async () => {
    const { viewer, heard } = await mount();
    viewer.select({ row: 0, col: 0 });
    heard();
    expect(viewer.enterValue('9', { row: 0, col: 1 }).ok).toBe(true);
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    expect(heard()).toEqual(['grid.selectedRange A1:A2']);
  });

  test('each paste over the same range says the range and the active cell', async () => {
    const { viewer, heard } = await mount();
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 1 });
    heard();
    for (const n of [7, 8, 9]) {
      expect(viewer.paste({ text: `0.5\t${n}\r\n3\t${n}` }).ok).toBe(true);
      const said = heard();
      expect(said).toHaveLength(1);
      expect(said[0]).toStartWith(
        'grid.selectionCell grid.selectedRange A1:B2 | grid.activeCell A1 | 0.5'
      );
    }
  });

  test('Enter inside a range after an undo says the new active cell', async () => {
    const { viewer, heard, host } = await mount();
    viewer.select({ row: 0, col: 0 }, { row: 2, col: 0 });
    expect(viewer.toggleFormat('bold').ok).toBe(true);
    pressKey(host);
    expect(viewer.undo().ok).toBe(true);
    heard();
    pressKey(host);
    expect(heard()).toEqual([
      'grid.selectionCell grid.selectedRange A1:A3 | grid.activeCell A2 | 3',
    ]);
  });

  test('Enter inside a range after a sheet rename says the new active cell', async () => {
    const { viewer, heard, host } = await mount();
    viewer.select({ row: 0, col: 0 }, { row: 2, col: 0 });
    expect(viewer.renameSheet(0, 'T').ok).toBe(true);
    heard();
    pressKey(host);
    expect(heard()).toEqual([
      'grid.selectionCell grid.selectedRange A1:A3 | grid.activeCell A2 | 3',
    ]);
  });

  test('a select of the same cell after a refusal says the cell again', async () => {
    const { viewer, heard } = await mount();
    viewer.select({ row: 0, col: 0 });
    expect(viewer.paste({ html: true }).ok).toBe(false);
    heard();
    viewer.select({ row: 0, col: 0 });
    expect(heard()).toEqual(['grid.activeCell A1 | 0.5']);
  });

  test('an edit, a select, and an edit before one paint say each cell', async () => {
    const { viewer, heard } = await mount();
    viewer.select({ row: 0, col: 0 });
    heard();
    expect(viewer.enterValue('7').ok).toBe(true);
    viewer.select({ row: 1, col: 0 });
    expect(viewer.enterValue('8').ok).toBe(true);
    expect(heard()).toEqual(['grid.activeCell A2 | 3', 'grid.activeCell A2 | 8']);
  });

  test('Tab inside a range says the range and the new active cell', async () => {
    const { viewer, heard, host } = await mount();
    viewer.select({ row: 0, col: 0 }, { row: 0, col: 1 });
    heard();
    pressKey(host, 'Tab');
    expect(heard()).toEqual([
      'grid.selectionCell grid.selectedRange A1:B1 | grid.activeCell B1 | x',
    ]);
  });

  test('an edit after a refused paste says its words', async () => {
    const { viewer, heard } = await mount();
    viewer.select({ row: 0, col: 0 });
    heard();
    expect(viewer.applyFormat({ numberFormat: { id: 9 } }).ok).toBe(true);
    expect(viewer.paste({ html: true }).ok).toBe(false);
    expect(heard()).toEqual(['editing.refused.paste-html']);
    expect(viewer.applyFormat({ numberFormat: { id: 10 } }).ok).toBe(true);
    expect(heard()).toEqual(['grid.activeCell A1 | 50.00%']);
  });

  test('a switch to a sheet with the same selection words says the sheet', async () => {
    const two = buildXlsx({
      sheets: [
        { name: 'S1', rows: row(1, [c('A1', 5)]) },
        { name: 'S2', rows: row(1, [c('A1', 5)]) },
      ],
    });
    const { viewer, heard } = await mount(two);
    viewer.select({ row: 0, col: 0 });
    heard();
    expect(viewer.setActiveSheet(1)).toBe(true);
    expect(heard()).toEqual(['grid.sheetActive S2 | grid.activeCell A1 | 5']);
  });

  test('a paste over the same range after another message says the active cell', async () => {
    const { viewer, heard } = await mount();
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 1 });
    viewer.setZoom(1.5);
    heard();
    expect(viewer.paste({ text: '0.5\t9\r\n3\t9' }).ok).toBe(true);
    expect(heard()).toEqual([
      'grid.selectionCell grid.selectedRange A1:B2 | grid.activeCell A1 | 0.5',
    ]);
  });

  test('a load empties the live region', async () => {
    const { viewer, host } = await mount();
    viewer.select({ row: 0, col: 0 });
    flushEditor(viewer);
    await viewer.load(sheet());
    expect(host.querySelector('.xlsx-live')?.textContent).toBe('');
  });

  test('a new range right after an undo says only the range', async () => {
    const { viewer, heard } = await mount();
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    expect(viewer.applyFormat({ numberFormat: { id: 9 } }).ok).toBe(true);
    heard();
    expect(viewer.undo().ok).toBe(true);
    viewer.select({ row: 0, col: 0 }, { row: 0, col: 1 });
    const said = heard();
    expect(said.at(-1)).toBe('grid.selectedRange A1:B1');
  });

  test('Enter inside a range after a detach and an attach says the new active cell', async () => {
    const { viewer, heard, host } = await mount();
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    heard();
    viewer.detach();
    viewer.attach(host);
    flushEditor(viewer);
    pressKey(host);
    flushEditor(viewer);
    // The attach builds a new region, so the test reads it in place of the spy.
    expect(host.querySelector('.xlsx-live')?.textContent).toBe(
      'grid.selectionCell grid.selectedRange A1:A2 | grid.activeCell A2 | 3'
    );
  });

  test('Enter inside a range after a catalog switch says the new active cell', async () => {
    const { viewer, heard, host } = await mount();
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    heard();
    lang = 'de:';
    relabelEditor(viewer);
    heard();
    pressKey(host);
    expect(heard()).toEqual([
      'de:grid.selectionCell de:grid.selectedRange A1:A2 | de:grid.activeCell A2 | 3',
    ]);
  });

  test('a select of the same range again says nothing', async () => {
    const { viewer, heard } = await mount();
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 1 });
    heard();
    for (let i = 0; i < 3; i++) viewer.select({ row: 0, col: 0 }, { row: 1, col: 1 });
    flushEditor(viewer);
    expect(heard()).toEqual([]);
  });

  test('a cell editor commit says the next active cell, not the edited one', async () => {
    const { viewer, heard } = await mount();
    viewer.select({ row: 0, col: 0 });
    heard();
    expect(viewer.beginEdit('25%').ok).toBe(true);
    expect(viewer.commitEdit('down').ok).toBe(true);
    expect(heard()).toEqual(['grid.activeCell A2 | 3']);
  });

  test('a refusal after the edit, before the paint, is not replaced', async () => {
    const { viewer, heard } = await mount();
    viewer.select({ row: 0, col: 0 });
    heard();
    expect(viewer.applyFormat({ numberFormat: { id: 9 } }).ok).toBe(true);
    expect(viewer.paste({ html: true }).ok).toBe(false);
    expect(heard()).toEqual(['editing.refused.paste-html']);
  });

  test('two edits before one paint say the last words once', async () => {
    const { viewer, heard } = await mount();
    viewer.select({ row: 0, col: 0 });
    heard();
    expect(viewer.applyFormat({ numberFormat: { id: 9 } }).ok).toBe(true);
    expect(viewer.toggleFormat('bold').ok).toBe(true);
    expect(heard()).toEqual(['grid.activeCell A1 | 50%']);
  });

  test('an edit on a selection of many cells says its active cell', async () => {
    const { viewer, heard } = await mount();
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    heard();
    // A number format change says the change first, so the words never name only one cell.
    expect(viewer.applyFormat({ numberFormat: { id: 9 } }).ok).toBe(true);
    expect(heard()).toEqual(['grid.numberFormatChanged grid.activeCell A1 | 50%']);
    expect(viewer.clearValues().ok).toBe(true);
    expect(heard()).toEqual(['grid.activeCellEmpty A1 | ']);
  });

  test('an undo and a redo on a selection of many cells also say its active cell', async () => {
    const { viewer, heard } = await mount();
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    expect(viewer.applyFormat({ numberFormat: { id: 9 } }).ok).toBe(true);
    heard();
    expect(viewer.undo().ok).toBe(true);
    expect(heard()).toEqual([
      'grid.undoneAt grid.selectionCell grid.selectedRange A1:A2 | grid.activeCell A1 | 0.5',
    ]);
    expect(viewer.redo().ok).toBe(true);
    expect(heard()).toEqual([
      'grid.redoneAt grid.selectionCell grid.selectedRange A1:A2 | grid.activeCell A1 | 50%',
    ]);
  });

  test('an undo and a redo of an edit that kept no selection say the active cell', async () => {
    const { viewer, heard } = await mount();
    viewer.select({ row: 0, col: 0 });
    expect(viewer.enterValue('9', { row: 0, col: 1 })).toEqual({ ok: true });
    viewer.select({ row: 0, col: 1 });
    heard();
    expect(viewer.undo().ok).toBe(true);
    expect(heard()).toEqual(['grid.undoneAt grid.activeCell B1 | x']);
    expect(viewer.redo().ok).toBe(true);
    expect(heard()).toEqual(['grid.redoneAt grid.activeCell B1 | 9']);
  });

  test('an undo that keeps the text of the active cell of a range names only the range', async () => {
    const { viewer, heard } = await mount();
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    expect(viewer.toggleFormat('bold').ok).toBe(true);
    heard();
    expect(viewer.undo().ok).toBe(true);
    expect(heard()).toEqual(['grid.undoneAt grid.selectedRange A1:A2']);
  });

  test('an undo after a move names the active cell only when its text changed', async () => {
    const { viewer, heard } = await mount();
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    expect(viewer.toggleFormat('bold').ok).toBe(true);
    viewer.select({ row: 0, col: 1 });
    heard();
    expect(viewer.undo().ok).toBe(true);
    expect(heard()).toEqual(['grid.undoneAt grid.selectedRange A1:A2']);
  });

  test('a range undo on another sheet says the restored text of the same cell', async () => {
    const two = buildXlsx({
      sheets: [
        { name: 'S1', rows: row(1, [c('A1', 5)]) },
        { name: 'S2', rows: row(1, [c('A1', 5)]) + row(2, [c('A2', 3)]) },
      ],
    });
    const { viewer, heard } = await mount(two);
    expect(viewer.setActiveSheet(1)).toBe(true);
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    expect(viewer.applyFormat({ numberFormat: { id: 9 } }).ok).toBe(true);
    expect(viewer.setActiveSheet(0)).toBe(true);
    viewer.select({ row: 0, col: 0 });
    heard();
    expect(viewer.undo().ok).toBe(true);
    expect(heard()).toEqual([
      'grid.undoneOnSheet S2 | grid.selectionCell grid.selectedRange A1:A2 | grid.activeCell A1 | 5',
    ]);
  });

  test('an edit and a detach before one paint say nothing after the attach', async () => {
    // The detach drops the change: the attach builds a new region, which must not say it.
    const { viewer, heard, host } = await mount();
    heard();
    expect(viewer.enterValue('12%')).toEqual({ ok: true });
    viewer.detach();
    viewer.attach(host);
    flushEditor(viewer);
    // The attach builds a new region, so read it, not the recorded one.
    expect(host.querySelector('.xlsx-live')?.textContent).toBe('');
  });

  test('an edit and a load before one paint say nothing of the old workbook', async () => {
    // The region is empty after the first load, so the load empties nothing that it can see.
    const { viewer, heard } = await mount();
    heard();
    expect(viewer.enterValue('12%')).toEqual({ ok: true });
    const loaded = viewer.load(sheet());
    expect(heard()).toEqual([]);
    await loaded;
  });

  test('an edit of another cell says nothing', async () => {
    const { viewer, heard } = await mount();
    viewer.select({ row: 0, col: 0 });
    heard();
    expect(viewer.enterValue('9', { row: 0, col: 1 })).toEqual({ ok: true });
    expect(heard()).toEqual([]);
  });
});

describe('live region after a decimal step or a number format pick', () => {
  /** C1 holds the text `abc` in `0.00`; D1 and D2 hold General numbers. */
  const formats = () =>
    buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('C1', 'abc', { s: 1 }), c('D1', 1.5)]) + row(2, [c('D2', 2.25)]),
        },
      ],
      styles: {
        cellXfs: [
          '<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
        ],
      },
    });

  test('a step that leaves the words of the active cell says that the places changed', async () => {
    const { viewer, heard } = await mount(formats());
    viewer.select({ row: 0, col: 2 }, { row: 1, col: 3 });
    heard();
    expect(runToolbarCommand(viewer, 'number.increaseDecimals').ok).toBe(true);
    expect(heard()).toEqual(['grid.decimalsIncreased']);
    expect(runToolbarCommand(viewer, 'number.decreaseDecimals').ok).toBe(true);
    expect(heard()).toEqual(['grid.decimalsDecreased']);
    expect(viewer.applyFormat({ numberFormat: { id: 10 } }).ok).toBe(true);
    expect(heard()).toEqual(['grid.numberFormatChanged']);
  });

  test('a step that changes only the padding spaces of the active cell confirms', async () => {
    // A4 holds 2.25 in `0.0?`: an increase to `0.0??` adds only a padding space.
    const bytes = buildXlsx({
      sheets: [{ name: 'S', rows: row(4, [c('A4', 2.25, { s: 1 })]) }],
      styles: {
        numFmts: ['<numFmt numFmtId="164" formatCode="0.0?"/>'],
        cellXfs: [
          '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
        ],
      },
    });
    const { viewer, heard } = await mount(bytes);
    viewer.select({ row: 3, col: 0 });
    heard();
    expect(runToolbarCommand(viewer, 'number.increaseDecimals').ok).toBe(true);
    expect(heard()).toEqual(['grid.decimalsIncreased']);
  });

  test('a step that changes the words of one cell says only the new words', async () => {
    const { viewer, heard } = await mount(formats());
    viewer.select({ row: 0, col: 3 });
    heard();
    expect(runToolbarCommand(viewer, 'number.increaseDecimals').ok).toBe(true);
    expect(heard()).toEqual(['grid.activeCell D1 | 1.50']);
  });

  test('a General number in a narrow column says the text that the grid fits', async () => {
    const bytes = buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 3.14159265358979)]),
          cols: '<cols><col min="1" max="1" width="6" customWidth="1"/></cols>',
        },
      ],
    });
    const { viewer, heard } = await mount(bytes);
    viewer.select({ row: 0, col: 1 });
    heard();
    viewer.select({ row: 0, col: 0 });
    const [said] = heard();
    expect(said).not.toContain('3.14159265');
    const shown = said!.split('| ')[1]!;
    expect(shown.startsWith('3.1')).toBe(true);
    // A decrease writes one place fewer than the fitted text shows, and says that text.
    expect(runToolbarCommand(viewer, 'number.decreaseDecimals').ok).toBe(true);
    const places = shown.split('.')[1]!.length;
    expect(heard()).toEqual([`grid.activeCell A1 | ${(3.14159265358979).toFixed(places - 1)}`]);
  });

  test('a number format change says its words with the padding spaces collapsed', async () => {
    const { viewer, heard } = await mount(formats());
    viewer.select({ row: 0, col: 3 });
    heard();
    expect(viewer.applyFormat({ numberFormat: { code: '_(* 0.00_)' } }).ok).toBe(true);
    const [said] = heard();
    expect(said).toContain('1.50');
    expect(said).not.toMatch(/ {2}| $/);
  });

  test('a step on a range says the change, then the new words of the active cell', async () => {
    const { viewer, heard } = await mount(formats());
    viewer.select({ row: 0, col: 3 }, { row: 1, col: 3 });
    heard();
    expect(runToolbarCommand(viewer, 'number.increaseDecimals').ok).toBe(true);
    expect(heard()).toEqual(['grid.decimalsIncreased grid.activeCell D1 | 1.50']);
  });
});

describe('the live region with a very long text', () => {
  const long = 'a'.repeat(MAX_CONTENT_FIELD_CHARS + 5);
  const bytes = () =>
    buildXlsx({
      sheets: [
        {
          name: 'S',
          rows:
            row(1, [c('A1', long), c('B1', 1, { f: `"${'b'.repeat(MAX_CONTENT_FIELD_CHARS)}"` })]) +
            row(2, [c('A2', 'short')]),
        },
      ],
    });

  test('says the start of a long text and that it continues, within the limit', async () => {
    const { viewer, heard } = await mount(bytes());
    const all: string[] = [];
    for (const at of [
      { row: 1, col: 0 },
      { row: 0, col: 0 },
      { row: 0, col: 1 },
    ]) {
      viewer.select(at);
      all.push(...heard());
    }
    const words = 'grid.textContinues '.length;
    for (const text of all)
      expect(text.length).toBeLessThanOrEqual(MAX_CONTENT_FIELD_CHARS + words);
    expect(all).toHaveLength(3);
    expect(all[0]).toBe('grid.activeCell A2 | short');
    expect(all[1]).toBe(
      `grid.textContinues grid.activeCell A1 | ${long}`.slice(0, words + MAX_CONTENT_FIELD_CHARS)
    );
    expect(all[2]?.startsWith('grid.textContinues grid.activeFormula B1 | 1 | ="bbb')).toBe(true);
  });

  test('says a short text in full', async () => {
    const { viewer, heard } = await mount(bytes());
    viewer.select({ row: 1, col: 0 });
    expect(heard()).toEqual(['grid.activeCell A2 | short']);
  });

  test('keeps the stale mark of a long result, in the English catalog', () => {
    const cached = 'r'.repeat(MAX_CONTENT_FIELD_CHARS + 10);
    const read = readXlsx(
      buildXlsx({
        sheets: [{ name: 'S', rows: row(1, [c('A1', cached, { t: 'str', f: 'A2' })]) }],
        calcPr: 'calcId="191029" fullCalcOnLoad="1"',
      })
    );
    if (!read.ok) throw new Error('read');
    const index = workbookIndex(read.pkg)!;
    const part = sheetByName(index, 'S')!.partName;
    const said = cellAnnouncement(read.pkg, index, part, { row: 0, col: 0 }, EN_US, englishText);
    expect(said.endsWith(', and the text continues, not recalculated')).toBe(true);
    expect(said.startsWith('A1, rrr')).toBe(true);
    const words = ', and the text continues, not recalculated'.length;
    expect(said.length).toBeLessThanOrEqual(MAX_CONTENT_FIELD_CHARS + words);
  });
});
