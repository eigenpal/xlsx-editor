// Row and column insertion and deletion from the viewer (task 7.5) in a happy-dom window of its
// own: the commands act on whole-row or whole-column selections or on the active cell's row or
// column, the heading press selects whole rows, Mod+Shift+= and Mod+- act on whole rows and
// columns and zoom otherwise, an open cell edit commits first, a refusal shows through the
// refusal channel with what it names, and one undo step restores the sheet.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { createXlsxEditor, MAX_COLS, MAX_ROWS, type XlsxEditorInstance } from '../index.ts';
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

const sheet = (afterMerges = '') =>
  buildXlsx({
    sheets: [
      {
        name: 'S',
        headings: true,
        rows:
          row(1, [c('A1', 1), c('B1', 2), c('C1', 0, { f: 'SUM(A1:A3)' })]) +
          row(2, [c('A2', 3)]) +
          row(3, [c('A3', 5)]),
        afterMerges,
      },
    ],
    styles: { dxfs: ['<dxf/>'] },
  });

const translate = (key: string, params?: Readonly<Record<string, string | number>>) =>
  params?.subject !== undefined
    ? `${key} [${params.subject}]`
    : params?.name !== undefined
      ? `${key} (${params.name})`
      : key;

async function mount(bytes: Uint8Array) {
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
  const input = host.querySelector('textarea') as HTMLTextAreaElement;
  return { viewer, scroller, live, host, input };
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

const content = (v: XlsxEditorInstance, r: number, col: number) => {
  v.select({ row: r, col });
  return v.snapshot().activeCell?.content;
};

const wholeRows = (v: XlsxEditorInstance, from: number, to: number) =>
  v.select({ row: from, col: 0 }, { row: to, col: MAX_COLS - 1 });

describe('structural commands of the viewer', () => {
  test('insertRows on the active cell inserts one row above it; undo restores it', async () => {
    const { viewer } = await mount(sheet());
    viewer.select({ row: 1, col: 1 });
    expect(viewer.insertRows().ok).toBe(true);
    expect(content(viewer, 2, 0)).toBe('3');
    expect(content(viewer, 0, 2)).toBe('=SUM(A1:A4)');
    expect(viewer.snapshot().stale).toBe(true);
    expect(viewer.undo().ok).toBe(true);
    expect(content(viewer, 1, 0)).toBe('3');
    expect(content(viewer, 0, 2)).toBe('=SUM(A1:A3)');
  });

  test('deleteRows on a whole-row selection deletes those rows', async () => {
    const { viewer } = await mount(sheet());
    wholeRows(viewer, 0, 1);
    expect(viewer.deleteRows().ok).toBe(true);
    expect(content(viewer, 0, 0)).toBe('5');
    expect(content(viewer, 1, 0)).toBe('');
  });

  test('column commands act on the active cell column, also in a whole-row selection', async () => {
    const { viewer } = await mount(sheet());
    viewer.select({ row: 0, col: 1 });
    expect(viewer.insertColumns().ok).toBe(true);
    expect(content(viewer, 0, 2)).toBe('2');
    wholeRows(viewer, 0, 0);
    expect(viewer.deleteColumns().ok).toBe(true);
    expect(content(viewer, 0, 0)).toBe('');
    expect(content(viewer, 0, 1)).toBe('2');
    expect(content(viewer, 0, 2)).toBe('=SUM(#REF!)');
  });

  test('an open cell edit commits before the command', async () => {
    const { viewer, scroller } = await mount(sheet());
    viewer.select({ row: 1, col: 0 });
    key(scroller, '9');
    expect(viewer.snapshot().editing).toBe(true);
    expect(viewer.insertRows().ok).toBe(true);
    expect(viewer.snapshot().editing).toBe(false);
    expect(content(viewer, 2, 0)).toBe('9');
  });

  test('a refusal shows through the refusal channel, names the element, and keeps editing', async () => {
    const cf =
      '<conditionalFormatting sqref="A1"><cfRule type="expression" dxfId="0" priority="1"><formula>1</formula></cfRule></conditionalFormatting>';
    const { viewer, live } = await mount(sheet(cf));
    const r = viewer.deleteRows();
    expect(r.ok ? null : r.refusal.code).toBe('structure-blocked');
    const subject = 'x:conditionalFormatting (/xl/worksheets/sheet1.xml)';
    expect(viewer.snapshot().refusal).toMatchObject({ code: 'structure-blocked', subject });
    // The announcement names the content in words, never the part path.
    expect(live.textContent).toBe(
      'editing.refused.structure-blocked [structureSubject.conditional-formatting (S)]'
    );
    expect(viewer.snapshot().canUndo).toBe(false);
    expect(viewer.enterValue('4').ok).toBe(true);
    expect(viewer.snapshot().refusal).toBeNull();
  });
});

describe('keys and headings', () => {
  test('a press on a row heading selects the row; Mod+Shift+= inserts, Mod+- deletes', async () => {
    const { viewer, scroller, host } = await mount(sheet());
    const stage = host.querySelector('.xlsx-stage') as HTMLElement;
    const W = host.ownerDocument.defaultView as unknown as typeof globalThis;
    stage.dispatchEvent(
      new W.PointerEvent('pointerdown', { bubbles: true, button: 0, clientX: 4, clientY: 30 })
    );
    const range = viewer.snapshot().selection!.range;
    expect([range.start.col, range.end.col]).toEqual([0, MAX_COLS - 1]);
    const first = range.start.row;
    key(scroller, '+', { ctrlKey: true, shiftKey: true });
    expect(viewer.snapshot().zoom).toBe(1);
    expect(viewer.snapshot().canUndo).toBe(true);
    expect(content(viewer, first, 0)).toBe('');
    viewer.select({ row: first, col: 0 }, { row: first, col: MAX_COLS - 1 });
    key(scroller, '-', { metaKey: true });
    expect(viewer.snapshot().zoom).toBe(1);
    expect(content(viewer, 1, 0)).toBe('3');
  });

  test('on whole columns the keys insert and delete columns; elsewhere they zoom', async () => {
    const { viewer, scroller } = await mount(sheet());
    viewer.select({ row: 0, col: 0 }, { row: MAX_ROWS - 1, col: 0 });
    key(scroller, '=', { ctrlKey: true, shiftKey: true });
    expect(content(viewer, 0, 1)).toBe('1');
    viewer.select({ row: 0, col: 0 }, { row: MAX_ROWS - 1, col: 0 });
    key(scroller, '-', { ctrlKey: true });
    expect(content(viewer, 0, 0)).toBe('1');
    viewer.select({ row: 0, col: 0 });
    key(scroller, '-', { ctrlKey: true });
    expect(viewer.snapshot().zoom).toBeLessThan(1);
  });

  test('Mod++ without Shift inserts on whole rows, and zooms in elsewhere', async () => {
    const { viewer, scroller } = await mount(sheet());
    wholeRows(viewer, 0, 0);
    key(scroller, '+', { ctrlKey: true });
    expect(viewer.snapshot().zoom).toBe(1);
    expect(content(viewer, 1, 0)).toBe('1');
    viewer.select({ row: 0, col: 0 });
    key(scroller, '+', { ctrlKey: true });
    expect(viewer.snapshot().zoom).toBeGreaterThan(1);
  });

  test('a whole-sheet selection refuses the commands with whole-sheet', async () => {
    const { viewer, scroller, live } = await mount(sheet());
    key(scroller, 'a', { ctrlKey: true });
    for (const r of [
      viewer.insertRows(),
      viewer.deleteRows(),
      viewer.insertColumns(),
      viewer.deleteColumns(),
    ])
      expect(r.ok ? null : r.refusal.code).toBe('whole-sheet');
    expect(live.textContent?.trim()).toBe('editing.refused.whole-sheet');
    expect(viewer.snapshot().canUndo).toBe(false);
    expect(content(viewer, 0, 0)).toBe('1');
  });
});
