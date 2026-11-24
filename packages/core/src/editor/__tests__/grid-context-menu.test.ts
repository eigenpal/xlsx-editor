// The grid context menu request (`grid-context-menu.ts`): a secondary press selects what it hits
// when the selection does not hold it and asks for the menu, Shift+F10 and the Menu key ask at
// the active cell, the browser keeps its menu while no handler listens, and the default rows
// follow what the press hit. Also the placement at a point (`contextMenuPlacement`).

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import {
  contextMenuPlacement,
  createXlsxEditor,
  gridContextMenuEntries,
  type GridContextMenuRequest,
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

async function mount() {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.append(host);
  const viewer = createXlsxEditor({ fonts: FONTS });
  viewers.push(viewer);
  viewer.attach(host);
  const scroller = host.querySelector('.xlsx-scroller') as HTMLElement;
  Object.defineProperty(scroller, 'clientWidth', { get: () => 800 });
  Object.defineProperty(scroller, 'clientHeight', { get: () => 400 });
  await viewer.load(
    buildXlsx({
      sheets: [{ name: 'S', headings: true, rows: row(1, [c('A1', 1)]) + row(2, [c('A2', 2)]) }],
    })
  );
  flushEditor(viewer);
  const stage = host.querySelector('.xlsx-stage') as HTMLElement;
  const W = win as unknown as typeof globalThis;
  const rightPress = (clientX: number, clientY: number) => {
    const e = new W.MouseEvent('contextmenu', {
      bubbles: true,
      cancelable: true,
      button: 2,
      clientX,
      clientY,
    });
    stage.dispatchEvent(e);
    return e;
  };
  const key = (init: KeyboardEventInit) => {
    const e = new W.KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init });
    scroller.dispatchEvent(e);
    return e;
  };
  const requests: GridContextMenuRequest[] = [];
  const listen = () => viewer.on('contextMenu', (r) => requests.push(r));
  return { viewer, scroller, rightPress, key, requests, listen };
}

describe('the request', () => {
  test('with no handler, the browser keeps its menu and nothing is selected', async () => {
    const { viewer, rightPress } = await mount();
    viewer.select({ row: 0, col: 0 });
    expect(rightPress(150, 60).defaultPrevented).toBe(false);
    expect(viewer.snapshot().selectionRef).toBe('A1');
  });

  test('a press outside the selection selects the cell and asks at the pointer', async () => {
    const { viewer, rightPress, requests, listen } = await mount();
    listen();
    viewer.select({ row: 0, col: 0 });
    expect(rightPress(150, 60).defaultPrevented).toBe(true);
    expect(viewer.snapshot().selectionRef).not.toBe('A1');
    expect(requests).toHaveLength(1);
    const r = requests[0]!;
    expect([r.x, r.y, r.trigger, r.kind, r.sheetIndex]).toEqual([150, 60, 'pointer', 'cells', 0]);
    expect(r.range).toEqual(viewer.snapshot().selection!.range);
  });

  test('a press inside the selection keeps it', async () => {
    const { viewer, rightPress, requests, listen } = await mount();
    listen();
    viewer.select({ row: 0, col: 0 }, { row: 20, col: 10 });
    rightPress(150, 60);
    expect(viewer.snapshot().selectionRef).toBe('A1:K21');
    expect(requests[0]!.range).toEqual({ start: { row: 0, col: 0 }, end: { row: 20, col: 10 } });
  });

  test('a press on a heading selects the column, and on the corner the sheet', async () => {
    const { viewer, rightPress, requests, listen } = await mount();
    listen();
    viewer.select({ row: 0, col: 0 });
    rightPress(150, 4);
    expect(viewer.snapshot().selectionRef).toBe('B1:B1048576');
    expect(requests.at(-1)!.kind).toBe('columns');
    rightPress(2, 2);
    expect(viewer.snapshot().selectionRef).toBe('A1:XFD1048576');
    expect(requests.at(-1)!.kind).toBe('all');
  });

  test('Shift+F10 and the Menu key ask at the active cell, from the selection', async () => {
    const { viewer, key, rightPress, requests, listen } = await mount();
    listen();
    viewer.select({ row: 1, col: 0 });
    expect(key({ key: 'F10', shiftKey: true }).defaultPrevented).toBe(true);
    expect(requests.at(-1)!).toMatchObject({ trigger: 'keyboard', kind: 'cells' });
    expect(key({ key: 'F10' }).defaultPrevented).toBe(false);
    viewer.select({ row: 1, col: 0 }, { row: 1, col: 16383 });
    key({ key: 'ContextMenu' });
    expect(requests.at(-1)!.kind).toBe('rows');
    // The browser's `contextmenu` for the same key opens nothing more.
    const count = requests.length;
    expect(rightPress(150, 60).defaultPrevented).toBe(true);
    expect(requests).toHaveLength(count);
  });

  test('an unsubscribed handler gives the browser its menu back', async () => {
    const { viewer, rightPress, listen } = await mount();
    const stop = listen();
    stop();
    viewer.select({ row: 0, col: 0 });
    expect(rightPress(150, 60).defaultPrevented).toBe(false);
  });
});

describe('the default rows', () => {
  const slots = (kind: Parameters<typeof gridContextMenuEntries>[0]) =>
    gridContextMenuEntries(kind).map((e) => (e.kind === 'item' ? e.slot : '|'));

  test('cells show the clipboard, rows and columns, clear, merge, and the number format', () => {
    expect(slots('cells')).toEqual([
      'edit.cut',
      'edit.copy',
      'edit.paste',
      '|',
      'cells.insertRows',
      'cells.insertColumns',
      'cells.deleteRows',
      'cells.deleteColumns',
      '|',
      'clear.contents',
      'clear.formats',
      '|',
      'merge.merge',
      'merge.unmerge',
      '|',
      'number.custom',
    ]);
  });

  test('a row heading shows the row commands, and a column heading the column commands', () => {
    expect(slots('rows')).toContain('cells.insertRows');
    expect(slots('rows')).not.toContain('cells.insertColumns');
    expect(slots('columns')).toContain('cells.deleteColumns');
    expect(slots('columns')).not.toContain('cells.deleteRows');
    expect(slots('all')).not.toContain('merge.merge');
  });
});

describe('the placement at a point', () => {
  const at = (x: number, y: number, rtl = false) =>
    contextMenuPlacement({
      x,
      y,
      rtl,
      width: 200,
      height: 300,
      viewportWidth: 1000,
      viewportHeight: 800,
    });

  test('opens below and toward the end of the point', () => {
    expect(at(100, 100)).toEqual({ left: 100, top: 100, maxHeight: 784 });
    expect(at(500, 100, true).left).toBe(300);
  });

  test('flips above and toward the start at the edges', () => {
    expect(at(950, 700)).toMatchObject({ left: 750, top: 400 });
    expect(at(50, 100, true).left).toBe(50);
  });

  test('clamps to the margins when no side has room', () => {
    const tall = contextMenuPlacement({
      x: 10,
      y: 300,
      rtl: false,
      width: 200,
      height: 2000,
      viewportWidth: 1000,
      viewportHeight: 800,
    });
    expect(tall.top).toBe(8);
    expect(tall.maxHeight).toBe(784);
  });
});
