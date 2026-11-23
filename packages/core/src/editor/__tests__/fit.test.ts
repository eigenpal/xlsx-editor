// The fit of a column or row to its content on a double-click on a heading edge (`grid-fit.ts`),
// in a happy-dom window of its own. A column takes the width of `fitColumnWidths` at the zoom
// steps and the current zoom, written with `@customWidth`. A row takes the height of
// `fitRowHeights` without `@customHeight`, because a fit is not a height set by hand. One undo
// restores the sizes. A selection of whole columns fits each column of the band. An open cell
// edit commits first, and its text counts. View mode and a press away from an edge fit nothing.

import { afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { createXlsxEditor, MAX_COLS, MAX_ROWS, type XlsxEditorInstance } from '../index.ts';
import { flushEditor } from './flush-editor.ts';
import { FIT_ZOOMS } from '../grid-fit.ts';
import {
  attr,
  child,
  children,
  readXlsx,
  sheetByName,
  workbookIndex,
  xmlPartOf,
  type XlsxPackage,
} from '../../store/index.ts';
import {
  createHarfBuzzMeasurer,
  fitColumnWidths,
  fitRowHeights,
  initializeHarfBuzz,
  zoomProfile,
  type FontSource,
  type TextMeasurer,
} from '../../layout/index.ts';

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

let measurer: TextMeasurer;
beforeAll(async () => {
  await initializeHarfBuzz();
  measurer = createHarfBuzzMeasurer(FONTS);
});

let viewers: XlsxEditorInstance[] = [];
afterEach(() => {
  for (const v of viewers) v.destroy();
  viewers = [];
});

const LONG = 'a text that is wider than the column';

const workbook = () =>
  buildXlsx({
    sheets: [
      {
        name: 'S',
        headings: true,
        rows: row(1, [c('A1', 'a'), c('B1', LONG), c('C1', 'abc')]) + row(2, [c('B2', 'x')]),
        cols: '<cols><col min="3" max="3" width="40" customWidth="1"/></cols>',
      },
    ],
  });

async function mount() {
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
  await viewer.load(workbook());
  flushEditor(viewer);
  const stage = host.querySelector('.xlsx-stage') as HTMLElement;
  const W = win as unknown as typeof globalThis;
  const double = (clientX: number, clientY: number) =>
    stage.dispatchEvent(new W.MouseEvent('dblclick', { bubbles: true, clientX, clientY }));
  const pointer = (type: string, clientX: number, clientY: number) =>
    stage.dispatchEvent(
      new W.PointerEvent(type, { bubbles: true, button: 0, pointerId: 1, clientX, clientY })
    );
  return { viewer, host, double, pointer };
}

const num = (v: string) => Number.parseFloat(v);

/** The far edge of a heading in stage pixels, and its size. */
function heading(host: HTMLElement, axis: 'col' | 'row', index: number) {
  const container = host.querySelector(
    axis === 'col' ? '.xlsx-column-heads' : '.xlsx-row-heads'
  ) as HTMLElement;
  const el = container.querySelector(`[data-${axis}="${index}"]`) as HTMLElement;
  const start =
    axis === 'col'
      ? num(container.style.left) + num(el.style.left)
      : num(container.style.top) + num(el.style.top);
  const size = num(axis === 'col' ? el.style.width : el.style.height);
  return { edge: start + size, size, mid: start + size / 2 };
}

/** The saved and reopened package, and its worksheet root. */
function reopened(viewer: XlsxEditorInstance) {
  const r = readXlsx(viewer.save()!);
  if (!r.ok) throw new Error('reopen');
  const entry = sheetByName(workbookIndex(r.pkg)!, 'S')!;
  return { pkg: r.pkg, root: xmlPartOf(r.pkg, entry.partName)!.root };
}

/** The `x:col` elements of a worksheet root, as `min-max:width:customWidth`. */
const colsOf = (root: ReturnType<typeof reopened>['root']) =>
  children(child(root, 'x:cols'), 'x:col').map(
    (el) => `${attr(el, 'min')}-${attr(el, 'max')}:${attr(el, 'width')}:${attr(el, 'customWidth')}`
  );

const options = { profile: zoomProfile(1), zooms: FIT_ZOOMS };
const fitOf = (pkg: XlsxPackage, first: number, last: number) =>
  fitColumnWidths(pkg, 'S', first, last, { ...options, measurer })!;

describe('a double-click on a column heading edge', () => {
  test('fits the column to its widest text, as one undo step that save keeps', async () => {
    const { viewer, host, double } = await mount();
    const before = reopened(viewer);
    const width = fitOf(before.pkg, 1, 1).get(1)!;
    const b = heading(host, 'col', 1);
    double(b.edge - 1, 5);
    flushEditor(viewer);
    expect(heading(host, 'col', 1).size).toBeGreaterThan(b.size);
    const after = reopened(viewer);
    expect(colsOf(after.root)).toContain(`2-2:${width}:1`);
    // A second fit finds the same width, so a second double-click adds no undo step.
    expect(fitOf(after.pkg, 1, 1).get(1)).toBe(width);
    flushEditor(viewer);
    double(heading(host, 'col', 1).edge - 1, 5);
    expect(viewer.undo().ok).toBe(true);
    expect(viewer.snapshot().canUndo).toBe(false);
    flushEditor(viewer);
    expect(heading(host, 'col', 1).size).toBe(b.size);
  });

  test('fits at the current zoom as well as at each zoom step from 50%', async () => {
    const { viewer, host, double } = await mount();
    viewer.setZoom(0.85);
    flushEditor(viewer);
    const before = reopened(viewer);
    const zooms = [...FIT_ZOOMS, 0.85];
    const width = fitColumnWidths(before.pkg, 'S', 1, 1, { ...options, zooms, measurer })!.get(1)!;
    double(heading(host, 'col', 1).edge - 1, 5);
    expect(colsOf(reopened(viewer).root)).toContain(`2-2:${width}:1`);
  });

  test('narrows a wide column with short text', async () => {
    const { viewer, host, double } = await mount();
    const c = heading(host, 'col', 2);
    double(c.edge - 1, 5);
    flushEditor(viewer);
    expect(heading(host, 'col', 2).size).toBeLessThan(c.size);
  });

  test('fits each column of a selection of whole columns that holds the edge, in one step', async () => {
    const { viewer, host, double } = await mount();
    viewer.select({ row: 0, col: 1 }, { row: MAX_ROWS - 1, col: 2 });
    const before = reopened(viewer);
    const widths = fitOf(before.pkg, 1, 2);
    double(heading(host, 'col', 2).edge - 1, 5);
    const cols = colsOf(reopened(viewer).root);
    expect(cols).toContain(`2-2:${widths.get(1)}:1`);
    expect(cols).toContain(`3-3:${widths.get(2)}:1`);
    expect(viewer.undo().ok).toBe(true);
    expect(viewer.snapshot().canUndo).toBe(false);
  });

  test('commits an open cell edit first, and fits its text', async () => {
    const { viewer, host, double } = await mount();
    viewer.select({ row: 4, col: 0 });
    expect(viewer.beginEdit(LONG).ok).toBe(true);
    const a = heading(host, 'col', 0);
    double(a.edge - 1, 5);
    flushEditor(viewer);
    viewer.select({ row: 4, col: 0 });
    expect(viewer.snapshot().activeCell?.content).toBe(LONG);
    expect(heading(host, 'col', 0).size).toBeGreaterThan(a.size);
  });

  test('fits the edge of the press that committed an open edit, before the grid paints again', async () => {
    const { viewer, host, double, pointer } = await mount();
    viewer.select({ row: 4, col: 0 });
    expect(viewer.beginEdit(LONG).ok).toBe(true);
    const a = heading(host, 'col', 0);
    // The full double-click: the first press commits the edit and clears the paint, so the
    // double-click finds the edge that the press found.
    pointer('pointerdown', a.edge - 1, 5);
    pointer('pointerup', a.edge - 1, 5);
    double(a.edge - 1, 5);
    flushEditor(viewer);
    viewer.select({ row: 4, col: 0 });
    expect(viewer.snapshot().activeCell?.content).toBe(LONG);
    expect(heading(host, 'col', 0).size).toBeGreaterThan(a.size);
  });
});

describe('a double-click on a row heading edge', () => {
  test('fits the row height without customHeight, and one undo restores it', async () => {
    const { viewer, host, double } = await mount();
    viewer.setRowHeight(0, 0, 60);
    flushEditor(viewer);
    const before = reopened(viewer);
    const height = fitRowHeights(before.pkg, 'S', 0, 0, { ...options, measurer })!.get(0)!;
    double(5, heading(host, 'row', 0).edge - 1);
    const r1 = children(child(reopened(viewer).root, 'x:sheetData'), 'x:row')[0];
    expect([attr(r1, 'ht'), attr(r1, 'customHeight')]).toEqual([String(height), undefined]);
    expect(viewer.undo().ok).toBe(true);
    const back = children(child(reopened(viewer).root, 'x:sheetData'), 'x:row')[0];
    expect([attr(back, 'ht'), attr(back, 'customHeight')]).toEqual(['60', '1']);
    expect(viewer.snapshot().canUndo).toBe(true);
    expect(viewer.undo().ok).toBe(true);
    expect(viewer.snapshot().canUndo).toBe(false);
  });

  test('fits each row of a selection of whole rows that holds the edge, in one step', async () => {
    const { viewer, host, double } = await mount();
    viewer.setRowHeight(0, 1, 60);
    flushEditor(viewer);
    viewer.select({ row: 0, col: 0 }, { row: 1, col: MAX_COLS - 1 });
    double(5, heading(host, 'row', 1).edge - 1);
    const rows = children(child(reopened(viewer).root, 'x:sheetData'), 'x:row');
    expect(rows.map((r) => attr(r, 'customHeight'))).toEqual([undefined, undefined]);
    expect(rows.map((r) => attr(r, 'ht'))).not.toContain('60');
    expect(viewer.undo().ok).toBe(true);
    const back = children(child(reopened(viewer).root, 'x:sheetData'), 'x:row');
    expect(back.map((r) => attr(r, 'ht'))).toEqual(['60', '60']);
  });
});

describe('a double-click that fits nothing', () => {
  test('in view mode, and away from an edge', async () => {
    const { viewer, host, double } = await mount();
    const b = heading(host, 'col', 1);
    double(b.mid, 5);
    expect(viewer.snapshot().canUndo).toBe(false);
    viewer.setMode('view');
    double(b.edge - 1, 5);
    expect(viewer.snapshot().canUndo).toBe(false);
    expect(colsOf(reopened(viewer).root)).toEqual(['3-3:40:1']);
  });
});
