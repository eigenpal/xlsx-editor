// The column and row resize of the editor (design D3 of the change `workbook-structure-editing`)
// in a happy-dom window of its own: `setColumnWidth` and `setRowHeight` write the file and one
// undo step puts it back; a drag of a heading edge shows the resize cursor and a guide line,
// commits one transaction on pointer up, and the layout shows the dragged size; Escape and
// pointer cancel cancel; a press with no movement commits nothing and selects nothing; a
// selection of whole columns resizes every selected column; a drag with an open cell edit
// commits the edit and the size; and a column of width 0 can be dragged wider again.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import {
  createXlsxEditor,
  MAX_COLUMN_WIDTH,
  MAX_RESIZE_NEW_ROWS,
  MAX_ROW_HEIGHT,
  MAX_ROWS,
  refusalText,
  type XlsxEditorInstance,
} from '../index.ts';
import { flushEditor } from './flush-editor.ts';
import {
  attr,
  child,
  children,
  readXlsx,
  sheetByName,
  workbookIndex,
  xmlPartOf,
} from '../../store/index.ts';
import type { FontSource } from '../../layout/index.ts';
import { hiddenLines } from '../grid-resize.ts';

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

const COLS = '<cols><col min="2" max="3" width="12" style="0" customWidth="1"/></cols>';

const workbook = (cols = COLS) =>
  buildXlsx({
    sheets: [
      {
        name: 'S',
        headings: true,
        rows: row(1, [c('A1', 1), c('B1', 2)]) + row(2, [c('A2', 3)]),
        cols,
      },
    ],
  });

async function mount(cols?: string, bytes = workbook(cols)) {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.appendChild(host);
  const viewer = createXlsxEditor({ fonts: FONTS });
  viewers.push(viewer);
  viewer.attach(host);
  const scroller = host.querySelector('.xlsx-scroller') as HTMLElement;
  Object.defineProperty(scroller, 'clientWidth', { get: () => 800 });
  Object.defineProperty(scroller, 'clientHeight', { get: () => 400 });
  await viewer.load(bytes);
  flushEditor(viewer);
  const stage = host.querySelector('.xlsx-stage') as HTMLElement;
  const W = win as unknown as typeof globalThis;
  const pointer = (type: string, clientX: number, clientY: number) =>
    stage.dispatchEvent(
      new W.PointerEvent(type, { bubbles: true, button: 0, pointerId: 1, clientX, clientY })
    );
  const escape = () =>
    doc.dispatchEvent(new W.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
  return { viewer, host, stage, pointer, escape, scroller };
}

const num = (v: string) => Number.parseFloat(v);

/** The box of a heading in stage pixels: its far edge and its size. */
function heading(host: HTMLElement, axis: 'col' | 'row', index: number) {
  const container = host.querySelector(
    axis === 'col' ? '.xlsx-column-heads' : '.xlsx-row-heads'
  ) as HTMLElement;
  const el = container.querySelector(`[data-${axis}="${index}"]`) as HTMLElement;
  if (axis === 'col') {
    const start = num(container.style.left) + num(el.style.left);
    return { edge: start + num(el.style.width), size: num(el.style.width), mid: start + 5 };
  }
  const start = num(container.style.top) + num(el.style.top);
  return { edge: start + num(el.style.height), size: num(el.style.height), mid: start + 5 };
}

/** The `x:col` and `x:row` elements of the saved workbook. */
function saved(viewer: XlsxEditorInstance) {
  const r = readXlsx(viewer.save()!);
  if (!r.ok) throw new Error('reopen');
  const entry = sheetByName(workbookIndex(r.pkg)!, 'S')!;
  const root = xmlPartOf(r.pkg, entry.partName)!.root;
  return {
    cols: children(child(root, 'x:cols'), 'x:col').map(
      (el) =>
        `${attr(el, 'min')}-${attr(el, 'max')}:${attr(el, 'width')}:${attr(el, 'style') ?? '-'}`
    ),
    rows: children(child(root, 'x:sheetData'), 'x:row').map(
      (el) => `${attr(el, 'r')}:${attr(el, 'ht') ?? '-'}`
    ),
  };
}

describe('the resize limits', () => {
  test('the editor entry exports the limits that the editor instance names', () => {
    expect([MAX_COLUMN_WIDTH, MAX_ROW_HEIGHT, MAX_RESIZE_NEW_ROWS]).toEqual([255, 409, 16_384]);
  });
});

describe('the editor instance', () => {
  test('setColumnWidth and setRowHeight take 0-based indexes, and undo puts them back', async () => {
    const { viewer } = await mount();
    const before = saved(viewer);
    expect(viewer.setColumnWidth(2, 3, 20).ok).toBe(true);
    expect(saved(viewer).cols).toEqual(['2-2:12:0', '3-3:20:0', '4-4:20:-']);
    expect(viewer.setRowHeight(2, 2, 30).ok).toBe(true);
    expect(saved(viewer).rows).toEqual(['1:-', '2:-', '3:30']);
    expect(viewer.undo().ok).toBe(true);
    expect(viewer.undo().ok).toBe(true);
    expect(saved(viewer)).toEqual(before);
  });

  test('a refused size shows through the refusal channel', async () => {
    const { viewer } = await mount();
    const r = viewer.setColumnWidth(0, 0, 300);
    expect(r.ok ? null : r.refusal.code).toBe('bad-size');
    expect(viewer.snapshot().refusal?.code).toBe('bad-size');
    const h = viewer.setRowHeight(-1, 0, 10);
    expect(h.ok ? null : h.refusal.code).toBe('bad-reference');
    expect(viewer.snapshot().canUndo).toBe(false);
  });
});

describe('the resize drag', () => {
  test('a drag of a column edge commits one width that the layout shows', async () => {
    const { viewer, host, stage, pointer } = await mount();
    const b = heading(host, 'col', 1);
    const selection = viewer.snapshot().selection;
    pointer('pointermove', b.edge - 2, 5);
    expect(stage.dataset.resize).toBe('columns');
    pointer('pointerdown', b.edge + 2, 5);
    pointer('pointermove', b.edge + 32, 5);
    const guide = stage.querySelector('.xlsx-resize-guide') as HTMLElement;
    expect(num(guide.style.left)).toBe(Math.round(b.edge + 30) - 1);
    pointer('pointerup', b.edge + 32, 5);
    expect(stage.querySelector('.xlsx-resize-guide')).toBeNull();
    expect(viewer.snapshot().selection).toEqual(selection);
    expect(saved(viewer).cols[0]).toMatch(/^2-2:[0-9.]+:0$/);
    flushEditor(viewer);
    expect(heading(host, 'col', 1).size).toBe(b.size + 30);
    expect(heading(host, 'col', 2).size).toBe(b.size);
    expect(viewer.undo().ok).toBe(true);
    expect(viewer.snapshot().canUndo).toBe(false);
    flushEditor(viewer);
    expect(heading(host, 'col', 1).size).toBe(b.size);
  });

  test('a press on the left part of an edge resizes the column before it', async () => {
    const { viewer, host, pointer } = await mount();
    const a = heading(host, 'col', 0);
    pointer('pointerdown', a.edge + 2, 5);
    pointer('pointermove', a.edge - 8, 5);
    pointer('pointerup', a.edge - 8, 5);
    flushEditor(viewer);
    expect(heading(host, 'col', 0).size).toBe(a.size - 10);
  });

  test('a drag of a row edge commits a height in points', async () => {
    const { viewer, host, stage, pointer } = await mount();
    const r = heading(host, 'row', 1);
    pointer('pointermove', 5, r.edge);
    expect(stage.dataset.resize).toBe('rows');
    pointer('pointerdown', 5, r.edge);
    pointer('pointermove', 5, r.edge + 20);
    pointer('pointerup', 5, r.edge + 20);
    flushEditor(viewer);
    expect(heading(host, 'row', 1).size).toBe(r.size + 20);
    expect(saved(viewer).rows[1]).toBe(`2:${((r.size + 20) * 72) / 96}`);
  });

  test('at a zoom of 2, a drag commits the size that the layout shows', async () => {
    const { viewer, host, pointer } = await mount();
    viewer.setZoom(2);
    flushEditor(viewer);
    const b = heading(host, 'col', 1);
    pointer('pointerdown', b.edge, 5);
    pointer('pointermove', b.edge + 40, 5);
    pointer('pointerup', b.edge + 40, 5);
    flushEditor(viewer);
    expect(heading(host, 'col', 1).size).toBe(b.size + 40);
    const r = heading(host, 'row', 1);
    pointer('pointerdown', 5, r.edge);
    pointer('pointermove', 5, r.edge + 30);
    pointer('pointerup', 5, r.edge + 30);
    flushEditor(viewer);
    expect(heading(host, 'row', 1).size).toBe(r.size + 30);
    // At zoom 2, a CSS pixel is half a pixel of zoom 1, so 30 pixels are 7.5 points more.
    expect(saved(viewer).rows[1]).toBe(`2:${((r.size + 30) / 2) * 0.75}`);
  });

  test('the guide line stops at the widest column and the tallest row', async () => {
    const { viewer, host, stage, pointer } = await mount();
    const b = heading(host, 'col', 1);
    pointer('pointerdown', b.edge, 5);
    pointer('pointermove', b.edge + 5000, 5);
    const guide = stage.querySelector('.xlsx-resize-guide') as HTMLElement;
    const widest = num(guide.style.left);
    pointer('pointermove', b.edge + 9000, 5);
    expect(num(guide.style.left)).toBe(widest);
    expect(widest).toBeLessThan(b.edge + 5000);
    pointer('pointerup', b.edge + 9000, 5);
    expect(saved(viewer).cols[0]).toBe('2-2:255:0');
    flushEditor(viewer);
    const r = heading(host, 'row', 1);
    pointer('pointerdown', 5, r.edge);
    pointer('pointermove', 5, r.edge + 5000);
    const tallest = num((stage.querySelector('.xlsx-resize-guide') as HTMLElement).style.top);
    expect(tallest).toBe(Math.round(r.edge - r.size + (409 * 4) / 3) - 1);
    pointer('pointerup', 5, r.edge + 5000);
    expect(saved(viewer).rows[1]).toBe('2:409');
  });

  test('a row edge drag with the whole sheet selected is refused with too-many-rows', async () => {
    const { viewer, host, pointer } = await mount();
    viewer.select({ row: 0, col: 0 }, { row: MAX_ROWS - 1, col: 16_383 });
    const r = heading(host, 'row', 1);
    pointer('pointerdown', 5, r.edge);
    pointer('pointermove', 5, r.edge + 10);
    pointer('pointerup', 5, r.edge + 10);
    expect(viewer.snapshot().refusal?.code).toBe('too-many-rows');
    expect(viewer.snapshot().canUndo).toBe(false);
  });

  test('a drag below 0 sets a size of 0', async () => {
    const { viewer, host, pointer } = await mount();
    const b = heading(host, 'col', 1);
    pointer('pointerdown', b.edge, 5);
    pointer('pointermove', b.edge - 500, 5);
    pointer('pointerup', b.edge - 500, 5);
    expect(saved(viewer).cols[0]).toBe('2-2:0:0');
  });

  test('Escape and pointer cancel cancel the drag; a drag that ends at its start commits nothing', async () => {
    const { viewer, host, stage, pointer, escape } = await mount();
    const b = heading(host, 'col', 1);
    const selection = viewer.snapshot().selection;
    pointer('pointerdown', b.edge, 5);
    pointer('pointermove', b.edge + 40, 5);
    escape();
    expect(stage.querySelector('.xlsx-resize-guide')).toBeNull();
    pointer('pointerup', b.edge + 40, 5);
    pointer('pointerdown', b.edge, 5);
    pointer('pointermove', b.edge + 40, 5);
    pointer('pointercancel', b.edge + 40, 5);
    pointer('pointerup', b.edge + 40, 5);
    pointer('pointerdown', b.edge, 5);
    pointer('pointerup', b.edge, 5);
    // A drag that moves away and comes back to its start commits nothing.
    pointer('pointerdown', b.edge, 5);
    pointer('pointermove', b.edge + 40, 5);
    pointer('pointermove', b.edge, 5);
    pointer('pointerup', b.edge, 5);
    expect(viewer.snapshot().canUndo).toBe(false);
    expect(viewer.snapshot().selection).toEqual(selection);
    expect(stage.dataset.resize).toBeUndefined();
  });

  test('a press inside a heading away from its edges still selects the column', async () => {
    const { viewer, host, pointer } = await mount();
    pointer('pointerdown', heading(host, 'col', 1).mid + 10, 5);
    pointer('pointerup', heading(host, 'col', 1).mid + 10, 5);
    const range = viewer.snapshot().selection!.range;
    expect([range.start.col, range.end.col, range.end.row]).toEqual([1, 1, MAX_ROWS - 1]);
  });

  test('inside a host that scales the editor, a press and a drag read stage pixels', async () => {
    const { viewer, host, stage, pointer } = await mount();
    // A host `transform: scale(2, 1.5)` at the window's origin: the stage shows at twice its
    // width and one and a half times its height.
    Object.defineProperty(stage, 'offsetWidth', { get: () => 400 });
    Object.defineProperty(stage, 'offsetHeight', { get: () => 300 });
    stage.getBoundingClientRect = () =>
      ({ left: 0, top: 0, width: 800, height: 450, right: 800, bottom: 450 }) as DOMRect;
    const b = heading(host, 'col', 1);
    pointer('pointerdown', (b.mid + 10) * 2, 10);
    pointer('pointerup', (b.mid + 10) * 2, 10);
    expect(viewer.snapshot().selection!.range.start.col).toBe(1);
    // A drag of 60 window pixels is 30 stage pixels.
    pointer('pointermove', b.edge * 2, 10);
    expect(stage.dataset.resize).toBe('columns');
    pointer('pointerdown', b.edge * 2, 10);
    pointer('pointermove', b.edge * 2 + 60, 10);
    pointer('pointerup', b.edge * 2 + 60, 10);
    flushEditor(viewer);
    expect(heading(host, 'col', 1).size).toBe(b.size + 30);
    // A press in a body cell: cell B2, at its stage position times the scale of each axis.
    const c1 = heading(host, 'col', 1);
    const r = heading(host, 'row', 1);
    pointer('pointerdown', (c1.mid + 10) * 2, (r.mid + 3) * 1.5);
    pointer('pointerup', (c1.mid + 10) * 2, (r.mid + 3) * 1.5);
    expect(viewer.snapshot().selection!.range.start).toEqual({ row: 1, col: 1 });
    // A drag of a row edge: 30 window pixels are 20 stage pixels.
    pointer('pointermove', 10, r.edge * 1.5);
    expect(stage.dataset.resize).toBe('rows');
    pointer('pointerdown', 10, r.edge * 1.5);
    pointer('pointermove', 10, r.edge * 1.5 + 30);
    pointer('pointerup', 10, r.edge * 1.5 + 30);
    flushEditor(viewer);
    expect(heading(host, 'row', 1).size).toBe(r.size + 20);
  });

  test('a drag inside a selection of whole columns resizes every selected column', async () => {
    const { viewer, host, pointer } = await mount();
    viewer.select({ row: 0, col: 0 }, { row: MAX_ROWS - 1, col: 2 });
    const b = heading(host, 'col', 1);
    pointer('pointerdown', b.edge, 5);
    pointer('pointermove', b.edge + 12, 5);
    pointer('pointerup', b.edge + 12, 5);
    const cols = saved(viewer).cols;
    // Columns 1 to 3 in one op: a new x:col for A, and the x:col of B:C with its style.
    expect(cols.map((c) => c.split(':')[0])).toEqual(['1-1', '2-3']);
    expect(new Set(cols.map((c) => c.split(':')[1])).size).toBe(1);
    expect(cols[1]!.endsWith(':0')).toBe(true);
    expect(viewer.undo().ok).toBe(true);
    expect(viewer.snapshot().canUndo).toBe(false);
    // A selection of cells does not widen the band.
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 2 });
    flushEditor(viewer);
    const b2 = heading(host, 'col', 1);
    pointer('pointerdown', b2.edge, 5);
    pointer('pointermove', b2.edge + 12, 5);
    pointer('pointerup', b2.edge + 12, 5);
    expect(saved(viewer).cols.map((c) => c.split(':')[0])).toEqual(['2-2', '3-3']);
  });
});

describe('a change of the view during a drag', () => {
  const two = buildXlsx({
    sheets: [
      { name: 'S', headings: true, rows: row(1, [c('A1', 1)]), cols: COLS },
      { name: 'T', headings: true, rows: row(1, [c('A1', 1)]) },
    ],
  });

  test('a change of the shown sheet cancels the drag, and no sheet takes its size', async () => {
    const { viewer, host, stage, pointer } = await mount(undefined, two);
    const b = heading(host, 'col', 1);
    pointer('pointerdown', b.edge, 5);
    pointer('pointermove', b.edge + 40, 5);
    expect(viewer.setActiveSheet(1)).toBe(true);
    expect(stage.querySelector('.xlsx-resize-guide')).toBeNull();
    flushEditor(viewer);
    pointer('pointermove', b.edge + 60, 5);
    pointer('pointerup', b.edge + 60, 5);
    expect(viewer.snapshot().canUndo).toBe(false);
  });

  test('a change of the store revision cancels the drag', async () => {
    const { viewer, host, stage, pointer } = await mount();
    const b = heading(host, 'col', 1);
    pointer('pointerdown', b.edge, 5);
    pointer('pointermove', b.edge + 40, 5);
    expect(viewer.enterValue('7', { row: 4, col: 0 }).ok).toBe(true);
    expect(stage.querySelector('.xlsx-resize-guide')).toBeNull();
    pointer('pointerup', b.edge + 40, 5);
    expect(viewer.undo().ok).toBe(true);
    expect(viewer.snapshot().canUndo).toBe(false);
  });

  test('a load during a drag cancels it, and the new workbook takes no size', async () => {
    const { viewer, host, stage, pointer } = await mount();
    const b = heading(host, 'col', 1);
    pointer('pointerdown', b.edge, 5);
    pointer('pointermove', b.edge + 40, 5);
    await viewer.load(workbook());
    flushEditor(viewer);
    expect(stage.isConnected ? stage.querySelector('.xlsx-resize-guide') : null).toBeNull();
    pointer('pointermove', b.edge + 60, 5);
    pointer('pointerup', b.edge + 60, 5);
    expect(viewer.snapshot().canUndo).toBe(false);
    expect(saved(viewer).cols).toEqual(['2-3:12:0']);
  });

  test('a detach during a drag cancels it, and Escape after it reaches the page', async () => {
    const { viewer, host, stage, pointer, escape } = await mount();
    const b = heading(host, 'col', 1);
    pointer('pointerdown', b.edge, 5);
    pointer('pointermove', b.edge + 40, 5);
    viewer.detach();
    expect(stage.querySelector('.xlsx-resize-guide')).toBeNull();
    let reached = false;
    stage.ownerDocument.addEventListener('keydown', () => (reached = true));
    escape();
    expect(reached).toBe(true);
    viewer.attach(host);
    flushEditor(viewer);
    const again = host.querySelector('.xlsx-stage') as HTMLElement;
    again.dispatchEvent(
      new (again.ownerDocument.defaultView as unknown as typeof globalThis).PointerEvent(
        'pointerup',
        { bubbles: true, button: 0, pointerId: 1, clientX: b.edge + 40, clientY: 5 }
      )
    );
    expect(viewer.snapshot().canUndo).toBe(false);
  });

  test('a zoom change during a drag cancels it', async () => {
    const { viewer, host, stage, pointer } = await mount();
    const b = heading(host, 'col', 1);
    pointer('pointerdown', b.edge, 5);
    pointer('pointermove', b.edge + 40, 5);
    viewer.setZoom(2);
    // The guide goes with the zoom, before the next pointer move or paint.
    expect(stage.querySelector('.xlsx-resize-guide')).toBeNull();
    expect(stage.dataset.resize).toBeUndefined();
    flushEditor(viewer);
    pointer('pointermove', b.edge + 60, 5);
    expect(stage.querySelector('.xlsx-resize-guide')).toBeNull();
    pointer('pointerup', b.edge + 60, 5);
    expect(viewer.snapshot().canUndo).toBe(false);
  });

  test('a switch to view mode during a drag cancels it, and its guide goes', async () => {
    const { viewer, host, stage, pointer } = await mount();
    const b = heading(host, 'col', 1);
    pointer('pointerdown', b.edge, 5);
    pointer('pointermove', b.edge + 40, 5);
    expect(viewer.setMode('view')).toBe(true);
    expect(stage.querySelector('.xlsx-resize-guide')).toBeNull();
    pointer('pointerup', b.edge + 40, 5);
    expect(viewer.snapshot().canUndo).toBe(false);
  });

  test('a scroll keeps the guide at the dragged edge', async () => {
    const { viewer, host, stage, pointer, scroller } = await mount();
    const b = heading(host, 'col', 1);
    pointer('pointerdown', b.edge, 5);
    pointer('pointermove', b.edge + 30, 5);
    scroller.scrollLeft = 20;
    scroller.dispatchEvent(new stage.ownerDocument.defaultView!.Event('scroll'));
    flushEditor(viewer);
    const moved = heading(host, 'col', 1);
    expect(moved.edge).toBeLessThan(b.edge);
    const guide = stage.querySelector('.xlsx-resize-guide') as HTMLElement;
    expect(num(guide.style.left)).toBe(Math.round(moved.edge + 30) - 1);
    pointer('pointerup', b.edge + 30, 5);
    flushEditor(viewer);
    expect(heading(host, 'col', 1).size).toBe(b.size + 30);
  });
});

describe('a drag with an open cell edit', () => {
  test('commits the edit and the new width in one press', async () => {
    const { viewer, host, pointer } = await mount();
    const b = heading(host, 'col', 1);
    viewer.select({ row: 3, col: 1 });
    expect(viewer.beginEdit('abc').ok).toBe(true);
    pointer('pointerdown', b.edge + 2, 5);
    pointer('pointermove', b.edge + 32, 5);
    pointer('pointerup', b.edge + 32, 5);
    flushEditor(viewer);
    expect(heading(host, 'col', 1).size).toBe(b.size + 30);
    viewer.select({ row: 3, col: 1 });
    expect(viewer.snapshot().activeCell?.content).toBe('abc');
  });
});

describe('a column of width 0', () => {
  const zero = (hidden: string) =>
    '<cols><col min="2" max="2" width="0" customWidth="1"' +
    hidden +
    '/><col min="3" max="3" width="12" customWidth="1"/></cols>';

  test('a press right of its edge drags it wider; a press left of it drags the column before', async () => {
    const { viewer, host, pointer } = await mount(zero(''));
    const a = heading(host, 'col', 0);
    pointer('pointerdown', a.edge + 2, 5);
    pointer('pointermove', a.edge + 42, 5);
    pointer('pointerup', a.edge + 42, 5);
    expect(saved(viewer).cols[0]).toMatch(/^2-2:[1-9][0-9.]*:-$/);
    expect(viewer.undo().ok).toBe(true);
    flushEditor(viewer);
    pointer('pointerdown', a.edge - 2, 5);
    pointer('pointermove', a.edge + 18, 5);
    pointer('pointerup', a.edge + 18, 5);
    flushEditor(viewer);
    expect(heading(host, 'col', 0).size).toBe(a.size + 20);
    expect(saved(viewer).cols[1]).toBe('2-2:0:-');
  });

  test('a hidden column is not hit: the press resizes the column before it', async () => {
    const { viewer, host, pointer } = await mount(zero(' hidden="1"'));
    const a = heading(host, 'col', 0);
    pointer('pointerdown', a.edge + 2, 5);
    pointer('pointermove', a.edge + 22, 5);
    pointer('pointerup', a.edge + 22, 5);
    flushEditor(viewer);
    expect(heading(host, 'col', 0).size).toBe(a.size + 20);
  });
});

describe('a part with two x:cols elements', () => {
  test('a hidden column of the second x:cols is not hit', async () => {
    const cols =
      '<cols><col min="1" max="1" width="9" customWidth="1"/></cols>' +
      '<cols><col min="2" max="2" width="0" customWidth="1" hidden="1"/></cols>';
    const { viewer } = await mount(cols);
    const r = readXlsx(viewer.save()!);
    if (!r.ok) throw new Error('reopen');
    const entry = sheetByName(workbookIndex(r.pkg)!, 'S')!;
    const hidden = hiddenLines(xmlPartOf(r.pkg, entry.partName));
    expect(hidden('columns', 1)).toBe(true);
    expect(hidden('columns', 0)).toBe(false);
  });
});

describe('a resize that an anchored object refuses', () => {
  const anchored = (name: string) =>
    buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 1)]),
          drawing: {
            anchors:
              '<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
              '<xdr:to><xdr:col>3</xdr:col><xdr:colOff>5</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>5</xdr:rowOff></xdr:to>' +
              `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="2" name="${name}"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>`,
          },
        },
      ],
    });

  const en = JSON.parse(readFileSync(`${import.meta.dir}/../../../../i18n/en.json`, 'utf8'));
  /** The English text of a key, with its parameters. */
  const translate = (key: string, params?: Readonly<Record<string, string | number>>) => {
    let text: unknown = en;
    for (const part of key.split('.')) text = (text as Record<string, unknown>)[part];
    return String(text).replace(/\{(\w+)\}/g, (all, k: string) => String(params?.[k] ?? all));
  };

  async function refusalFor(name: string) {
    const win = new Window();
    const host = (win.document as unknown as Document).createElement('div');
    (win.document as unknown as Document).body.appendChild(host);
    const viewer = createXlsxEditor({ fonts: FONTS, translate });
    viewers.push(viewer);
    viewer.attach(host);
    await viewer.load(anchored(name));
    const r = viewer.setColumnWidth(2, 2, 30);
    return { r, live: host.querySelector('[aria-live]')?.textContent ?? '' };
  }

  test('the texts of too-many-rows and sheet-name-ambiguous', () => {
    expect(translate('editing.refused.too-many-rows')).toBe(
      'The row height was not applied: the selection has too many rows with no content or size. Select fewer rows.'
    );
    expect(translate('editing.refused.sheet-name-ambiguous', { subject: 'Data' })).toContain(
      'two sheets have the name Data, also when case is ignored,'
    );
  });

  test('names the shape by its name, with resize-anchor', async () => {
    const { r, live } = await refusalFor('Logo');
    expect(r.ok ? null : [r.refusal.code, r.refusal.subject]).toEqual(['resize-anchor', 'Logo']);
    expect(live).toContain('Logo keeps its size or position');
  });

  test('names a shape with no name by a generic label, not by its part', async () => {
    const { r, live } = await refusalFor('');
    expect(r.ok ? null : [r.refusal.code, r.refusal.subject]).toEqual([
      'resize-anchor',
      'an object on the sheet',
    ]);
    expect(live).not.toContain('/xl/drawings');
    expect(live).toContain('an object on the sheet keeps its size or position');
  });
});

describe('a heading edge in view mode', () => {
  test('shows no resize cursor, starts no drag, and draws no guide', async () => {
    const { viewer, host, stage, pointer } = await mount();
    viewer.setMode('view');
    const before = saved(viewer);
    const b = heading(host, 'col', 1);
    pointer('pointermove', b.edge - 2, 5);
    expect(stage.dataset.resize).toBeUndefined();
    pointer('pointerdown', b.edge, 5);
    pointer('pointermove', b.edge + 30, 5);
    expect(stage.querySelector('.xlsx-resize-guide')).toBeNull();
    pointer('pointerup', b.edge + 30, 5);
    const r = heading(host, 'row', 1);
    pointer('pointerdown', 5, r.edge);
    pointer('pointermove', 5, r.edge + 20);
    expect(stage.querySelector('.xlsx-resize-guide')).toBeNull();
    pointer('pointerup', 5, r.edge + 20);
    expect(viewer.snapshot().canUndo).toBe(false);
    expect(saved(viewer)).toEqual(before);
    // Back in edit mode, the edge is hit again.
    viewer.setMode('edit');
    pointer('pointermove', b.edge - 2, 5);
    expect(stage.dataset.resize).toBe('columns');
  });
});

describe('a resize on a protected sheet', () => {
  const protectedBook = (attrs = '') =>
    buildXlsx({
      sheets: [
        {
          name: 'S',
          headings: true,
          rows: row(1, [c('A1', 1), c('B1', 2)]) + row(2, [c('A2', 3)]),
          cols: COLS,
          beforeMerges: `<sheetProtection sheet="1" ${attrs}/>`,
        },
      ],
    });

  test('a locked size shows no resize edge, and a drag writes nothing', async () => {
    const { viewer, host, stage, pointer } = await mount(undefined, protectedBook());
    const before = saved(viewer);
    const b = heading(host, 'col', 1);
    pointer('pointermove', b.edge - 2, 5);
    expect(stage.dataset.resize).toBeUndefined();
    pointer('pointerdown', b.edge, 5);
    pointer('pointermove', b.edge + 30, 5);
    expect(stage.querySelector('.xlsx-resize-guide')).toBeNull();
    pointer('pointerup', b.edge + 30, 5);
    const r = heading(host, 'row', 1);
    pointer('pointermove', 5, r.edge - 2);
    expect(stage.dataset.resize).toBeUndefined();
    pointer('pointerdown', 5, r.edge);
    pointer('pointermove', 5, r.edge + 20);
    pointer('pointerup', 5, r.edge + 20);
    expect(viewer.snapshot().canUndo).toBe(false);
    expect(saved(viewer)).toEqual(before);
    // The command is refused with the size lock and its message.
    const w = viewer.setColumnWidth(1, 1, 20);
    expect(w.ok ? null : [w.refusal.code, w.refusal.lock]).toEqual(['locked', 'size']);
    expect(refusalText(viewer.snapshot().refusal!, (key) => key)).toBe(
      'editing.lockedRefused.size'
    );
  });

  test('a sheet that allows only column widths shows only the column edges', async () => {
    const { host, stage, pointer } = await mount(undefined, protectedBook('formatColumns="0"'));
    const b = heading(host, 'col', 1);
    pointer('pointermove', b.edge - 2, 5);
    expect(stage.dataset.resize).toBe('columns');
    const r = heading(host, 'row', 1);
    pointer('pointermove', 5, r.edge - 2);
    expect(stage.dataset.resize).toBeUndefined();
  });

  test('formatColumns="0" allows the width, and formatRows="0" the height', async () => {
    const { viewer } = await mount(undefined, protectedBook('formatColumns="0"'));
    expect(viewer.setColumnWidth(1, 1, 20).ok).toBe(true);
    const h = viewer.setRowHeight(0, 0, 30);
    expect(h.ok ? null : h.refusal.lock).toBe('size');
    expect(viewer.snapshot().protection).toMatchObject({
      allowFormatColumns: true,
      allowFormatRows: false,
    });
    const rows = await mount(undefined, protectedBook('formatRows="0"'));
    expect(rows.viewer.setRowHeight(0, 0, 30).ok).toBe(true);
  });
});
