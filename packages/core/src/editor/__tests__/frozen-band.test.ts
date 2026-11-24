// A frozen band that does not start at row 1 and column A. With a frozen `x:pane` (`CT_Pane`,
// ECMA-376 Part 1, 18.3.1.66), `x:sheetView/@topLeftCell` (`CT_SheetView`, 18.3.1.87) is the
// first cell of the top-left pane, and `@ySplit` and `@xSplit` count the frozen rows and columns
// from it. The view map, and the headings and hit testing of an opened sheet, agree on the band.
// The engine is driven directly, in happy-dom.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { createXlsxEditor, type XlsxEditorInstance } from '../index.ts';
import { AxisIndex } from '../../layout/axis.ts';
import { resizeEdgeAt } from '../grid-resize.ts';
import { cycleInRange } from '../grid-keys.ts';
import {
  cellAtPoint,
  rangeOutline,
  rangeRect,
  rangeRects,
  scrollToShow,
  scrollToTopLeft,
  type ViewState,
} from '../view-map.ts';
import { scrollExtent } from '../sheet-facts.ts';
import { flushEditor } from './flush-editor.ts';
import type { FontSource, SheetGeometry, SheetLayout } from '../../layout/index.ts';

describe('the view map of a frozen band that starts at row 10 and column B', () => {
  const axis = {
    offsetOf: (i: number) => i * 20,
    sizeOf: () => 20,
    indexAt: (at: number) => Math.max(0, Math.floor(at / 20)),
  };
  const geometry = { columns: axis, rows: axis } as unknown as SheetGeometry;
  // Rows 10 and 11 (0-based 9 and 10) and column B (0-based 1) are frozen.
  const frozen = { rows: 2, cols: 1, firstRow: 9, firstCol: 1 };
  const view = (over: Partial<ViewState> = {}): ViewState => ({
    geometry,
    frozen,
    scrollX: 0,
    scrollY: 0,
    width: 200,
    height: 200,
    ...over,
  });
  const range = (r0: number, c0: number, r1: number, c1: number) => ({
    start: { row: r0, col: c0 },
    end: { row: r1, col: c1 },
  });

  test('a point of the band is a cell of the band, and the body starts after it', () => {
    expect(cellAtPoint(view(), 5, 5)).toEqual({ row: 9, col: 1 });
    expect(cellAtPoint(view(), 5, 25)).toEqual({ row: 10, col: 1 });
    expect(cellAtPoint(view(), 25, 45)).toEqual({ row: 11, col: 2 });
    // The body scrolls; the band stays.
    expect(cellAtPoint(view({ scrollY: 100 }), 25, 45)).toEqual({ row: 16, col: 2 });
    expect(cellAtPoint(view({ scrollY: 100 }), 5, 5)).toEqual({ row: 9, col: 1 });
  });

  test('the cells before the band do not show', () => {
    expect(rangeRects(view(), range(0, 0, 8, 0))).toEqual([]);
    expect(rangeOutline(view(), range(0, 0, 8, 0))).toEqual([]);
    // A range from row 1 to row 10 shows only its part in the band, with no top edge.
    expect(rangeRects(view(), range(0, 1, 9, 1))).toEqual([{ x: 0, y: 0, width: 20, height: 20 }]);
    const sides = rangeOutline(view(), range(0, 1, 9, 1)).map((e) => e.side);
    expect(sides.sort()).toEqual(['bottom', 'left', 'right']);
  });

  test('a range of the band and the body covers one rectangle in each pane', () => {
    expect(rangeRects(view(), range(9, 1, 11, 2))).toEqual([
      { x: 0, y: 0, width: 20, height: 40 },
      { x: 0, y: 40, width: 20, height: 20 },
      { x: 20, y: 0, width: 20, height: 40 },
      { x: 20, y: 40, width: 20, height: 20 },
    ]);
    expect(rangeRect(view(), range(9, 1, 9, 1))).toEqual({ x: 0, y: 0, width: 20, height: 20 });
    expect(rangeRect(view(), range(11, 2, 11, 2))).toEqual({
      x: 20,
      y: 40,
      width: 20,
      height: 20,
    });
  });

  test('a reveal scrolls only the body, and the scroll starts after the band', () => {
    // A cell of the band, and a cell before it, keep the scroll.
    expect(scrollToShow(view({ scrollY: 60 }), { row: 10, col: 1 })).toEqual({ x: 0, y: 60 });
    expect(scrollToShow(view({ scrollY: 60 }), { row: 2, col: 0 })).toEqual({ x: 0, y: 60 });
    // Row 12 (0-based 11) is the first row of the body: scroll 0.
    expect(scrollToShow(view({ scrollY: 60 }), { row: 11, col: 2 })).toEqual({ x: 0, y: 0 });
    expect(scrollToShow(view(), { row: 30, col: 2 })).toEqual({ x: 0, y: 240 });
    expect(scrollToTopLeft(geometry, frozen, { row: 14, col: 4 })).toEqual({ x: 40, y: 60 });
    expect(scrollToTopLeft(geometry, frozen, { row: 9, col: 1 })).toEqual({ x: 0, y: 0 });
  });

  test('the rows and columns before the band take no room in the scroll extent', () => {
    const count = { rows: 1_048_576, columns: 16_384 };
    const g = {
      rows: { ...axis, count: count.rows },
      columns: { ...axis, count: count.columns },
    };
    const layout = { geometry: g, panes: [], frozen } as unknown as SheetLayout;
    // The content ends at row 50 and column 10; the extent adds 100 rows and 20 columns.
    expect(scrollExtent(layout, { lastRow: 49, lastCol: 9 }, { row: 0, col: 0 })).toEqual({
      width: (30 - 1) * 20,
      height: (150 - 9) * 20,
    });
  });
});

describe('a frozen band with hidden rows', () => {
  // Rows of 10 points; row 4 (0-based 3) is hidden, inside the band of rows 3 to 5, and row 2
  // (0-based 1) is hidden before it.
  const rows = new AxisIndex(1_048_576, 10, [
    { first: 1, last: 1, size: 0 },
    { first: 3, last: 3, size: 0 },
  ]);
  const cols = new AxisIndex(16_384, 50, []);
  const geometry = { rows, columns: cols } as unknown as SheetGeometry;
  const view: ViewState = {
    geometry,
    frozen: { rows: 3, cols: 0, firstRow: 2, firstCol: 0 },
    scrollX: 0,
    scrollY: 0,
    width: 200,
    height: 100,
  };

  test('the band is as high as its shown rows, and the body starts after it', () => {
    expect(cellAtPoint(view, 5, 5)).toEqual({ row: 2, col: 0 });
    // Row 4 takes no room: the second band row on the window is row 5.
    expect(cellAtPoint(view, 5, 15)).toEqual({ row: 4, col: 0 });
    expect(cellAtPoint(view, 5, 25)).toEqual({ row: 5, col: 0 });
    expect(rangeRects(view, { start: { row: 2, col: 0 }, end: { row: 5, col: 0 } })).toEqual([
      { x: 0, y: 0, width: 50, height: 20 },
      { x: 0, y: 20, width: 50, height: 10 },
    ]);
    const edges = rangeOutline(view, { start: { row: 4, col: 0 }, end: { row: 6, col: 0 } });
    // The range starts in the band and ends in the body: its top edge is at the second band row.
    expect(edges.find((e) => e.side === 'top')).toMatchObject({ y: 10 });
    expect(edges.find((e) => e.side === 'bottom')).toMatchObject({ y: 40 });
  });
});

test('the resize edge at the start of the band is not the edge of a column before it', () => {
  // Column A has width 0 and lies before the band of column B: no heading shows it, so the left
  // edge of the band takes no resize. Without the band, the same edge widens column A again.
  const cols = new AxisIndex(16_384, 50, [{ first: 0, last: 0, size: 0 }]);
  const rows = new AxisIndex(1_048_576, 10, []);
  const geometry = { rows, columns: cols } as unknown as SheetGeometry;
  const box = { scale: 1, headWidth: 30, headHeight: 20, width: 400, height: 300 };
  const map = (firstCol: number, colsFrozen: number): ViewState => ({
    geometry,
    frozen: { rows: 0, cols: colsFrozen, firstRow: 0, firstCol },
    scrollX: 0,
    scrollY: 0,
    width: 370,
    height: 280,
  });
  expect(resizeEdgeAt(map(1, 1), box, 31, 5)).toBeNull();
  expect(resizeEdgeAt(map(0, 0), box, 31, 5)).toEqual({ axis: 'columns', index: 0 });
});

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

const editors: XlsxEditorInstance[] = [];
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
});

/**
 * An editor on a sheet whose view was frozen while it was scrolled: rows `top` and `top + 1`
 * (10 and 11 by default) and column B are frozen, and the body starts at column C, two rows
 * below `top`. The container is 800 by 400 pixels.
 */
async function offsetEditor(headings: boolean, top = 10) {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.append(host);
  const editor = createXlsxEditor({ fonts: FONTS });
  editors.push(editor);
  editor.attach(host);
  await editor.load(
    buildXlsx({
      sheets: [
        {
          name: 'S',
          rows:
            row(1, [c('A1', 'Above')]) +
            row(top, [c(`B${top}`, 'Band')]) +
            row(top + 2, [c(`C${top + 2}`, 1)]),
          headings,
          viewAttrs: `topLeftCell="B${top}"`,
          viewChildren:
            `<pane xSplit="1" ySplit="2" topLeftCell="C${top + 2}" activePane="bottomRight" state="frozen"/>` +
            `<selection pane="bottomRight" activeCell="C${top + 2}" sqref="C${top + 2}"/>`,
        },
      ],
    })
  );
  const scroller = host.querySelector<HTMLElement>('.xlsx-scroller')!;
  Object.defineProperty(scroller, 'clientWidth', { get: () => 800 });
  Object.defineProperty(scroller, 'clientHeight', { get: () => 400 });
  flushEditor(editor);
  return { editor, host, W: win as unknown as typeof globalThis };
}

const px = (v: string) => parseFloat(v);

test('the headings show the frozen rows 10 and 11 and column B first, then the body', async () => {
  const { host } = await offsetEditor(true);
  const order = (sel: string, key: 'row' | 'col') =>
    [...host.querySelectorAll<HTMLElement>(sel)]
      .sort(
        (a, b) =>
          px(a.style[key === 'row' ? 'top' : 'left']) - px(b.style[key === 'row' ? 'top' : 'left'])
      )
      .map((h) => h.textContent);
  const rows = order('.xlsx-row-heads .xlsx-head', 'row');
  expect(rows.slice(0, 4)).toEqual(['10', '11', '12', '13']);
  expect(rows).not.toContain('1');
  const cols = order('.xlsx-column-heads .xlsx-head', 'col');
  expect(cols.slice(0, 3)).toEqual(['B', 'C', 'D']);
  expect(cols).not.toContain('A');
  // The first frozen row heading starts at the top of the grid.
  const first = host.querySelector<HTMLElement>('.xlsx-row-heads [data-row="9"]')!;
  expect(px(first.style.top)).toBe(0);
});

test('a press on the frozen band selects a cell of the band, and below it a cell of the body', async () => {
  const { editor, host, W } = await offsetEditor(false);
  const stage = host.querySelector('.xlsx-stage') as HTMLElement;
  const press = (x: number, y: number) =>
    stage.dispatchEvent(
      new W.PointerEvent('pointerdown', { bubbles: true, button: 0, clientX: x, clientY: y })
    );
  press(5, 5);
  expect(editor.snapshot().selection?.active).toEqual({ row: 9, col: 1 });
  press(5, 25);
  expect(editor.snapshot().selection?.active).toEqual({ row: 10, col: 1 });
  // Rows 10 and 11 are 30 points (40 pixels) high: a press just below them is in row 12.
  press(5, 50);
  expect(editor.snapshot().selection?.active).toEqual({ row: 11, col: 1 });
});

test('the keys do not move the active cell into the rows and columns before the band', async () => {
  const { editor, host, W } = await offsetEditor(false);
  const scroller = host.querySelector('.xlsx-scroller') as HTMLElement;
  const key = (k: string, init: Record<string, unknown> = {}) =>
    scroller.dispatchEvent(
      new W.KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true, ...init })
    );
  editor.select({ row: 9, col: 1 });
  key('ArrowUp');
  key('ArrowLeft');
  expect(editor.snapshot().selection?.active).toEqual({ row: 9, col: 1 });
  // The first shown cell is the first cell of the band.
  editor.select({ row: 30, col: 5 });
  key('Home', { ctrlKey: true });
  expect(editor.snapshot().selection?.active).toEqual({ row: 9, col: 1 });
  editor.select({ row: 11, col: 5 });
  key('PageUp');
  expect(editor.snapshot().selection?.active).toEqual({ row: 9, col: 5 });
  // Home goes to the first column of the band, not to column A.
  editor.select({ row: 11, col: 5 });
  key('Home');
  expect(editor.snapshot().selection?.active).toEqual({ row: 11, col: 1 });
});

test('the keys stop at a band that starts past the bounded search of hidden lines', async () => {
  // Row 5000 is past the 4096 lines that a key searches for a shown line.
  const { editor, host, W } = await offsetEditor(false, 5000);
  const scroller = host.querySelector('.xlsx-scroller') as HTMLElement;
  const key = (k: string, init: Record<string, unknown> = {}) =>
    scroller.dispatchEvent(
      new W.KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true, ...init })
    );
  editor.select({ row: 4999, col: 1 });
  key('ArrowUp');
  expect(editor.snapshot().selection?.active).toEqual({ row: 4999, col: 1 });
  key('ArrowUp', { ctrlKey: true });
  expect(editor.snapshot().selection?.active).toEqual({ row: 4999, col: 1 });
  editor.select({ row: 5100, col: 5 });
  key('Home', { ctrlKey: true });
  expect(editor.snapshot().selection?.active).toEqual({ row: 4999, col: 1 });
  editor.select({ row: 5001, col: 5 });
  key('PageUp');
  expect(editor.snapshot().selection?.active).toEqual({ row: 4999, col: 5 });
});

test('a selection of the whole sheet takes its active cell in the band', async () => {
  const { editor } = await offsetEditor(false);
  editor.selectAll();
  expect(editor.snapshot().selection?.active).toEqual({ row: 9, col: 1 });
  // A selection whose range holds no cell that shows keeps its active cell.
  editor.select({ row: 0, col: 0 });
  expect(editor.snapshot().selection?.active).toEqual({ row: 0, col: 0 });
});

test('a file with no selection opens with the first cell of the band active', async () => {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.append(host);
  const editor = createXlsxEditor({ fonts: FONTS });
  editors.push(editor);
  editor.attach(host);
  await editor.load(
    buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: row(10, [c('B10', 'Band')]),
          viewAttrs: 'topLeftCell="B10"',
          viewChildren: '<pane xSplit="1" ySplit="2" topLeftCell="C12" state="frozen"/>',
        },
      ],
    })
  );
  expect(editor.snapshot().selection?.active).toEqual({ row: 9, col: 1 });
});

test('Enter in a selection of whole columns wraps to the first row of a band past row 4096', () => {
  const hidden = { row: (i: number) => i < 4999, col: () => false, firstRow: 4999, firstCol: 0 };
  const range = { start: { row: 0, col: 1 }, end: { row: 5001, col: 2 } };
  expect(cycleInRange(range, { row: 5001, col: 1 }, false, 1, hidden)).toEqual({
    row: 4999,
    col: 2,
  });
  // A range wholly before the band keeps its own rows.
  const before = { start: { row: 0, col: 1 }, end: { row: 1, col: 1 } };
  expect(cycleInRange(before, { row: 1, col: 1 }, false, 1, BEFORE_BAND)).toEqual({
    row: 0,
    col: 1,
  });
});

const BEFORE_BAND = { row: () => false, col: () => false, firstRow: 4999, firstCol: 0 };

test('from a cell before a band past row 4096, a step down goes to the first row of the band', async () => {
  const { editor, host, W } = await offsetEditor(false, 5000);
  const scroller = host.querySelector('.xlsx-scroller') as HTMLElement;
  const key = (k: string, init: Record<string, unknown> = {}) =>
    scroller.dispatchEvent(
      new W.KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true, ...init })
    );
  editor.select({ row: 0, col: 1 });
  key('ArrowDown');
  expect(editor.snapshot().selection?.active).toEqual({ row: 4999, col: 1 });
  editor.select({ row: 0, col: 1 });
  key('ArrowDown', { ctrlKey: true });
  expect(editor.snapshot().selection?.active).toEqual({ row: 4999, col: 1 });
  // A selection of the whole sheet moves its focus with its active cell, so Shift+ArrowRight
  // extends the range from the band.
  editor.selectAll();
  key('ArrowRight', { shiftKey: true });
  expect(editor.snapshot().selection?.range).toEqual({
    start: { row: 4999, col: 1 },
    end: { row: 4999, col: 2 },
  });
});

test('a step from A1 lands in the band on both axes, and a range from the API moves its active cell', async () => {
  const { editor, host, W } = await offsetEditor(false);
  const scroller = host.querySelector('.xlsx-scroller') as HTMLElement;
  editor.select({ row: 0, col: 0 });
  scroller.dispatchEvent(
    new W.KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true, cancelable: true })
  );
  expect(editor.snapshot().selection?.active).toEqual({ row: 9, col: 1 });
  // A1:D20 holds cells of the band: its active cell moves to B10, and the range stays.
  editor.select({ row: 0, col: 0 }, { row: 19, col: 3 });
  const sel = editor.snapshot().selection!;
  expect(sel.active).toEqual({ row: 9, col: 1 });
  expect(sel.range).toEqual({ start: { row: 0, col: 0 }, end: { row: 19, col: 3 } });
});

test('a cell of the band that a merged range covers does not become the active cell', async () => {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.append(host);
  const editor = createXlsxEditor({ fonts: FONTS });
  editors.push(editor);
  editor.attach(host);
  await editor.load(
    buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: row(9, [c('B9', 'Merged')]),
          merges: ['B9:B11'],
          viewAttrs: 'topLeftCell="B10"',
          viewChildren: '<pane xSplit="1" ySplit="2" topLeftCell="C12" state="frozen"/>',
        },
      ],
    })
  );
  // B10 is the first cell of the band, but B9:B11 covers it: the active cell stays.
  editor.selectAll();
  expect(editor.snapshot().selection?.active).toEqual({ row: 0, col: 0 });
  editor.select({ row: 0, col: 0 }, { row: 19, col: 3 });
  expect(editor.snapshot().selection?.active).toEqual({ row: 0, col: 0 });
});

test('the body opens at the top-left cell of the pane, after a band that starts at row 10', async () => {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.append(host);
  const editor = createXlsxEditor({ fonts: FONTS });
  editors.push(editor);
  editor.attach(host);
  await editor.load(
    buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: row(10, [c('A10', 'Band')]),
          headings: true,
          viewAttrs: 'topLeftCell="A10"',
          viewChildren:
            '<pane ySplit="2" topLeftCell="A40" activePane="bottomLeft" state="frozen"/>' +
            '<selection pane="bottomLeft" activeCell="A40" sqref="A40"/>',
        },
      ],
    })
  );
  const scroller = host.querySelector<HTMLElement>('.xlsx-scroller')!;
  Object.defineProperty(scroller, 'clientWidth', { get: () => 800 });
  Object.defineProperty(scroller, 'clientHeight', { get: () => 400 });
  // The opening scroll applies on the paint after the first layout.
  flushEditor(editor);
  flushEditor(editor);
  const rows = [...host.querySelectorAll<HTMLElement>('.xlsx-row-heads .xlsx-head')]
    .sort((a, b) => px(a.style.top) - px(b.style.top))
    .map((h) => h.textContent);
  expect(rows.slice(0, 4)).toEqual(['10', '11', '40', '41']);
});

test('a redo takes its active cell into the band that the edit moved', async () => {
  const { editor } = await offsetEditor(false);
  // A1:D5 holds no cell of the band at row 10, so A1 stays the active cell.
  editor.select({ row: 0, col: 0 }, { row: 4, col: 3 });
  expect(editor.deleteRows().ok).toBe(true);
  expect(editor.undo().ok).toBe(true);
  expect(editor.snapshot().selection?.active).toEqual({ row: 0, col: 0 });
  // After the redo, the band starts at row 5, inside A1:D5: B5 takes the active cell.
  expect(editor.redo().ok).toBe(true);
  expect(editor.snapshot().selection?.active).toEqual({ row: 4, col: 1 });
});

test('a file selection of A1 opens with the first cell of the band active', async () => {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.append(host);
  const editor = createXlsxEditor({ fonts: FONTS });
  editors.push(editor);
  editor.attach(host);
  await editor.load(
    buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: row(10, [c('B10', 'Band')]),
          viewAttrs: 'topLeftCell="B10"',
          viewChildren:
            '<pane xSplit="1" ySplit="2" topLeftCell="C12" state="frozen"/>' +
            '<selection activeCell="A1" sqref="A1 C12:D13"/>',
        },
      ],
    })
  );
  const sel = editor.snapshot().selection!;
  expect(sel.active).toEqual({ row: 9, col: 1 });
  expect(sel.range).toEqual({ start: { row: 9, col: 1 }, end: { row: 9, col: 1 } });
});

test('a file range that holds the band keeps its range, with its active cell in the band', async () => {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.append(host);
  const editor = createXlsxEditor({ fonts: FONTS });
  editors.push(editor);
  editor.attach(host);
  await editor.load(
    buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: row(10, [c('B10', 'Band')]),
          viewAttrs: 'topLeftCell="B10"',
          viewChildren:
            '<pane xSplit="1" ySplit="2" topLeftCell="C12" state="frozen"/>' +
            '<selection activeCell="A1" sqref="A1:D20"/>',
        },
      ],
    })
  );
  const sel = editor.snapshot().selection!;
  expect(sel.active).toEqual({ row: 9, col: 1 });
  expect(sel.range).toEqual({ start: { row: 0, col: 0 }, end: { row: 19, col: 3 } });
});
