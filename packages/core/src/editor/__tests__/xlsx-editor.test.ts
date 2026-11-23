// The workbook editor in a happy-dom window of its own: it opens fixtures, shows one tab for
// each visible sheet, paints only the visible window, and moves the selection with the keyboard
// and the pointer inside the grid bounds.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { runInNewContext } from 'node:vm';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { createXlsxEditor, type XlsxEditorInstance } from '../index.ts';
import { flushEditor } from './flush-editor.ts';
import { stepAnnouncement } from '../selection.ts';
import type { FontSource } from '../../layout/index.ts';

const fontsDir = join(
  dirname(Bun.resolveSync('@docx-editor.dev/fonts/package.json', import.meta.dir)),
  'assets'
);
const FONTS: FontSource[] = [
  ['Carlito', 'Carlito-Regular.ttf', false, false],
  ['Carlito', 'Carlito-Bold.ttf', true, false],
  ['Liberation Sans', 'LiberationSans-Regular.ttf', false, false],
].map(([family, file, bold, italic]) => ({
  family: family as string,
  bold: bold as boolean,
  italic: italic as boolean,
  bytes: new Uint8Array(readFileSync(join(fontsDir, file as string))),
  origin: 'substitute',
}));

const fixture = (name: string) =>
  new Uint8Array(readFileSync(`${import.meta.dir}/../../../../../fixtures/${name}.xlsx`));

let viewers: XlsxEditorInstance[] = [];
afterEach(() => {
  for (const v of viewers) v.destroy();
  viewers = [];
});

function mount(width = 800, height = 400) {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.appendChild(host);
  const viewer = createXlsxEditor({
    fonts: FONTS,
    translate: (k, p) => `${k}${p ? JSON.stringify(p) : ''}`,
  });
  viewers.push(viewer);
  viewer.attach(host);
  const scroller = host.querySelector('.xlsx-scroller') as HTMLElement;
  Object.defineProperty(scroller, 'clientWidth', { get: () => width });
  Object.defineProperty(scroller, 'clientHeight', { get: () => height });
  return { viewer, host, scroller, doc };
}

const key = (el: HTMLElement, k: string, init: Record<string, boolean> = {}) =>
  el.dispatchEvent(
    new (el.ownerDocument.defaultView as unknown as typeof globalThis).KeyboardEvent('keydown', {
      key: k,
      bubbles: true,
      cancelable: true,
      ...init,
    })
  );

describe('workbook editor', () => {
  test('opens a fixture and paints its values in the window', async () => {
    const { viewer, host } = mount();
    const snap = await viewer.load(fixture('basic-values'));
    expect(snap.status).toBe('ready');
    expect(snap.sheets.map((s) => s.name)).toEqual(['Values']);
    flushEditor(viewer);
    expect(host.querySelector('[data-cell="A1"]')?.textContent).toBe('Shared text');
    // The fixture hides the headings (`showRowColHeaders="0"`), and the editor follows it.
    expect((host.querySelector('.xlsx-column-heads') as HTMLElement).hidden).toBe(true);
    expect(viewer.snapshot().activeCell).toEqual({ ref: 'A1', content: 'Shared text' });
  });

  test('with no fonts option, the packaged substitute faces load and the workbook opens', async () => {
    const viewer = createXlsxEditor();
    viewers.push(viewer);
    const snap = await viewer.load(fixture('basic-values'));
    expect(snap.failure).toBeNull();
    expect(snap.status).toBe('ready');
  });

  test('with no translate option, the grid labels are English', async () => {
    const win = new Window();
    const doc = win.document as unknown as Document;
    const host = doc.createElement('div');
    doc.body.appendChild(host);
    const viewer = createXlsxEditor({ fonts: FONTS });
    viewers.push(viewer);
    viewer.attach(host);
    await viewer.load(buildXlsx({ sheets: [{ name: 'Data', rows: row(1, [c('A1', 1)]) }] }));
    flushEditor(viewer);
    viewer.setZoom(2);
    const grid = host.querySelector('[role="grid"]')!;
    expect(grid.getAttribute('aria-label')).toBe('Worksheet Data');
    expect(host.querySelector('.xlsx-live')?.textContent).toBe('Zoom 200%');
  });

  test('resolves with the failure `package` for a detached ArrayBuffer', async () => {
    const { viewer } = mount();
    const buffer = new ArrayBuffer(8);
    structuredClone(buffer, { transfer: [buffer] });
    const snap = await viewer.load(buffer);
    expect(snap.status).toBe('failed');
    expect(snap.failure?.code).toBe('package');
  });

  test('opens a workbook from an ArrayBuffer', async () => {
    const { viewer } = mount();
    const bytes = fixture('basic-values');
    const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
    const snap = await viewer.load(buffer as ArrayBuffer);
    expect(snap.status).toBe('ready');
  });

  test('opens a workbook from an ArrayBuffer of another realm', async () => {
    const { viewer } = mount();
    const bytes = fixture('basic-values');
    const buffer = runInNewContext(`new ArrayBuffer(${bytes.byteLength})`) as ArrayBuffer;
    expect(buffer instanceof ArrayBuffer).toBe(false);
    new Uint8Array(buffer).set(bytes);
    const snap = await viewer.load(buffer);
    expect(snap.status).toBe('ready');
  });

  test('opens a workbook from a SharedArrayBuffer', async () => {
    const { viewer } = mount();
    const bytes = fixture('basic-values');
    const shared = new SharedArrayBuffer(bytes.byteLength);
    new Uint8Array(shared).set(bytes);
    const snap = await viewer.load(shared as unknown as ArrayBuffer);
    expect(snap.status).toBe('ready');
  });
  test('a file that is not a workbook fails with a code', async () => {
    const { viewer } = mount();
    const snap = await viewer.load(new Uint8Array([1, 2, 3]));
    expect(snap.status).toBe('failed');
    expect(snap.failure?.code).toBe('package');
  });

  test('a failed load closes the previous workbook and saves nothing', async () => {
    const { viewer, host } = mount();
    await viewer.load(fixture('basic-values'));
    expect(viewer.save()).not.toBeNull();
    await viewer.load(new Uint8Array([1, 2, 3]));
    expect(viewer.snapshot().status).toBe('failed');
    expect(viewer.save()).toBeNull();
    // The previous workbook is closed: no tabs, no active cell, no painted cell.
    const snap = viewer.snapshot();
    expect(snap.sheets).toEqual([]);
    expect(snap.activeCell).toBeNull();
    expect(snap.canUndo).toBe(false);
    flushEditor(viewer);
    expect(host.querySelector('[data-cell="A1"]')).toBeNull();
  });

  test('hidden sheets have no tab, and the second tab shows the second sheet', async () => {
    const { viewer, host } = mount();
    await viewer.load(
      buildXlsx({
        sheets: [
          { name: 'One', rows: row(1, [c('A1', 'first')]) },
          { name: 'Hidden', state: 'hidden', rows: row(1, [c('A1', 'secret')]) },
          { name: 'Two', rows: row(1, [c('A1', 'second')]) },
        ],
      })
    );
    const tabs = viewer.snapshot().sheets.filter((s) => s.visible);
    expect(tabs.map((s) => s.name)).toEqual(['One', 'Two']);
    expect(viewer.setActiveSheet(1)).toBe(false);
    expect(viewer.setActiveSheet(2)).toBe(true);
    flushEditor(viewer);
    expect(host.querySelector('[data-cell="A1"]')?.textContent).toBe('second');
  });

  test('a sheet change announces the sheet and its active cell', async () => {
    const { viewer, host } = mount();
    await viewer.load(
      buildXlsx({
        sheets: [
          { name: 'One', rows: row(1, [c('A1', 'first')]) },
          { name: 'Two', rows: row(1, [c('A1', 'second')]) },
        ],
      })
    );
    viewer.setActiveSheet(1);
    expect(host.querySelector('.xlsx-live')?.textContent).toBe(
      'grid.sheetActive{"sheet":"Two","selection":"grid.activeCell{\\"ref\\":\\"A1\\",\\"content\\":\\"second\\"}"}'
    );
  });

  test('Shift with the right arrow twice from B2 selects B2:D2 and keeps B2 active', async () => {
    const { viewer, scroller } = mount();
    await viewer.load(fixture('basic-values'));
    flushEditor(viewer);
    viewer.select({ row: 1, col: 1 });
    key(scroller, 'ArrowRight', { shiftKey: true });
    key(scroller, 'ArrowRight', { shiftKey: true });
    const s = viewer.snapshot();
    expect(s.selectionRef).toBe('B2:D2');
    expect(s.activeCell?.ref).toBe('B2');
  });

  test('the selection stays inside the grid bounds', async () => {
    const { viewer, scroller } = mount();
    await viewer.load(fixture('basic-values'));
    flushEditor(viewer);
    key(scroller, 'ArrowUp');
    key(scroller, 'ArrowLeft');
    expect(viewer.snapshot().selectionRef).toBe('A1');
    viewer.select({ row: 2_000_000, col: 99_999 });
    expect(viewer.snapshot().selectionRef).toBe('XFD1048576');
  });

  test('Home, Page Down, Mod+Home, Mod+End, Alt+Page Down, and Mod+A move and select', async () => {
    const { viewer, scroller } = mount();
    await viewer.load(
      buildXlsx({
        sheets: [
          { name: 'One', rows: row(1, [c('A1', 1)]) + row(5, [c('D5', 4)]) },
          { name: 'Two', rows: row(1, [c('A1', 2)]) },
        ],
      })
    );
    flushEditor(viewer);
    viewer.select({ row: 2, col: 2 });
    key(scroller, 'Home');
    expect(viewer.snapshot().selectionRef).toBe('A3');
    key(scroller, 'PageDown');
    const paged = viewer.snapshot().activeCell!.ref;
    expect(Number(paged.slice(1))).toBeGreaterThan(3);
    key(scroller, 'End', { ctrlKey: true });
    expect(viewer.snapshot().selectionRef).toBe('D5');
    key(scroller, 'Home', { ctrlKey: true, shiftKey: true });
    expect(viewer.snapshot().selectionRef).toBe('A1:D5');
    key(scroller, 'Home', { ctrlKey: true });
    expect(viewer.snapshot().selectionRef).toBe('A1');
    key(scroller, 'a', { ctrlKey: true });
    expect(viewer.snapshot().selectionRef).toBe('A1:XFD1048576');
    key(scroller, 'PageDown', { altKey: true });
    expect(viewer.snapshot().activeSheet).toBe(1);
  });

  test('Page Down, Page Up, and Mod+End skip hidden rows and columns, as the arrows do', async () => {
    const { viewer, scroller } = mount();
    let rows = row(1, [c('A1', 1)]);
    for (let r = 2; r <= 60; r += 1) rows += row(r, [], 'hidden="1"');
    rows += row(61, [c('A61', 2)]) + row(62, [c('D62', 3)], 'hidden="1"');
    await viewer.load(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            rows,
            cols: '<cols><col min="4" max="4" width="9" hidden="1" customWidth="1"/></cols>',
          },
        ],
      })
    );
    flushEditor(viewer);
    viewer.select({ row: 0, col: 0 });
    key(scroller, 'PageDown');
    expect(viewer.snapshot().selectionRef).toBe('A61');
    key(scroller, 'PageUp');
    expect(viewer.snapshot().selectionRef).toBe('A1');
    // The last content is D62, in a hidden row and a hidden column: Mod+End lands on C61.
    key(scroller, 'End', { ctrlKey: true });
    expect(viewer.snapshot().selectionRef).toBe('C61');
  });

  test('Home and Mod+Home land on the first shown column and row', async () => {
    const { viewer, scroller } = mount();
    await viewer.load(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            rows: row(1, [c('B1', 1)], 'hidden="1"') + row(3, [c('C3', 2)]),
            cols: '<cols><col min="1" max="1" width="9" hidden="1" customWidth="1"/></cols>',
          },
        ],
      })
    );
    flushEditor(viewer);
    viewer.select({ row: 2, col: 2 });
    key(scroller, 'Home');
    expect(viewer.snapshot().selectionRef).toBe('B3');
    key(scroller, 'Home', { ctrlKey: true });
    expect(viewer.snapshot().selectionRef).toBe('B2');
  });

  test('Mod+arrow leaves hidden rows and columns out of the data run', async () => {
    const { viewer, scroller } = mount();
    const cells = (r: number) => [c(`A${r}`, r), c(`B${r}`, r), c(`C${r}`, r), c(`D${r}`, r)];
    await viewer.load(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            // Row 5 and column D are hidden at the edge of the data; row 3 is hidden inside it.
            rows:
              row(1, cells(1)) +
              row(2, cells(2)) +
              row(3, cells(3), 'hidden="1"') +
              row(4, cells(4)) +
              row(5, cells(5), 'hidden="1"'),
            cols: '<cols><col min="4" max="4" width="9" hidden="1" customWidth="1"/></cols>',
          },
        ],
      })
    );
    flushEditor(viewer);
    viewer.select({ row: 0, col: 0 });
    key(scroller, 'ArrowDown', { ctrlKey: true });
    expect(viewer.snapshot().selectionRef).toBe('A4');
    // Past the hidden row 5, no value is left: the jump goes to the sheet edge.
    key(scroller, 'ArrowDown', { ctrlKey: true });
    expect(viewer.snapshot().selectionRef).toBe('A1048576');
    key(scroller, 'ArrowUp', { ctrlKey: true });
    expect(viewer.snapshot().selectionRef).toBe('A4');
    key(scroller, 'ArrowUp', { ctrlKey: true });
    expect(viewer.snapshot().selectionRef).toBe('A1');
    key(scroller, 'ArrowRight', { ctrlKey: true });
    expect(viewer.snapshot().selectionRef).toBe('C1');
    key(scroller, 'ArrowRight', { ctrlKey: true });
    expect(viewer.snapshot().selectionRef).toBe('XFD1');
    key(scroller, 'ArrowLeft', { ctrlKey: true });
    expect(viewer.snapshot().selectionRef).toBe('C1');
  });

  test('Enter inside a tall range and Shift+arrow on a band scroll the moved cell into view', async () => {
    const { viewer, scroller } = mount();
    await viewer.load(fixture('basic-values'));
    flushEditor(viewer);
    viewer.select({ row: 0, col: 0 }, { row: 199, col: 0 });
    flushEditor(viewer);
    const far = scroller.scrollTop;
    expect(far).toBeGreaterThan(0);
    key(scroller, 'Enter');
    flushEditor(viewer);
    expect(viewer.snapshot().activeCell?.ref).toBe('A2');
    expect(scroller.scrollTop).toBeLessThan(far);
    // Mod+A keeps the view where it is.
    const before = scroller.scrollTop;
    key(scroller, 'a', { ctrlKey: true });
    flushEditor(viewer);
    expect(scroller.scrollTop).toBe(before);
  });

  test('Enter that commits an edit in the last shown row scrolls the next cell into view', async () => {
    const { viewer, scroller, host } = mount();
    await viewer.load(fixture('basic-values'));
    flushEditor(viewer);
    // The last row that the window shows whole, before any scroll.
    let row = 0;
    while (scroller.scrollTop === 0 && row < 100) {
      key(scroller, 'ArrowDown');
      flushEditor(viewer);
      row += 1;
    }
    key(scroller, 'Home', { ctrlKey: true });
    flushEditor(viewer);
    for (let k = 0; k < row - 1; k += 1) key(scroller, 'ArrowDown');
    flushEditor(viewer);
    expect(scroller.scrollTop).toBe(0);
    key(scroller, 'x');
    key(host.querySelector('textarea')!, 'Enter');
    flushEditor(viewer);
    expect(viewer.snapshot().activeCell?.ref).toBe(`A${row + 1}`);
    expect(scroller.scrollTop).toBeGreaterThan(0);
    expect(host.querySelector('.xlsx-active-cell')).not.toBeNull();
  });

  test('a press on a column heading selects the column, and on the corner the sheet', async () => {
    const { viewer, host } = mount();
    await viewer.load(
      buildXlsx({ sheets: [{ name: 'S', headings: true, rows: row(1, [c('A1', 1)]) }] })
    );
    flushEditor(viewer);
    const stage = host.querySelector('.xlsx-stage') as HTMLElement;
    const W = host.ownerDocument.defaultView as unknown as typeof globalThis;
    const press = (clientX: number, clientY: number) =>
      stage.dispatchEvent(
        new W.PointerEvent('pointerdown', { bubbles: true, button: 0, clientX, clientY })
      );
    press(150, 4);
    expect(viewer.snapshot().selectionRef).toBe('B1:B1048576');
    press(2, 2);
    expect(viewer.snapshot().selectionRef).toBe('A1:XFD1048576');
  });

  test('a touch pans the sheet: a tap selects on release, a moved touch selects nothing', async () => {
    const { viewer, host } = mount();
    await viewer.load(
      buildXlsx({ sheets: [{ name: 'S', headings: true, rows: row(1, [c('A1', 1)]) }] })
    );
    flushEditor(viewer);
    const stage = host.querySelector('.xlsx-stage') as HTMLElement;
    const W = host.ownerDocument.defaultView as unknown as typeof globalThis;
    const touch = (type: string, clientX: number, clientY: number) => {
      const e = new W.PointerEvent(type, {
        bubbles: true,
        cancelable: true,
        button: 0,
        pointerId: 7,
        pointerType: 'touch',
        clientX,
        clientY,
      });
      stage.dispatchEvent(e);
      return e;
    };
    viewer.select({ row: 0, col: 0 });
    // The press leaves the default (the pan) alone and selects nothing yet.
    expect(touch('pointerdown', 150, 60).defaultPrevented).toBe(false);
    expect(viewer.snapshot().selectionRef).toBe('A1');
    touch('pointerup', 150, 60);
    const tapped = viewer.snapshot().selectionRef;
    expect(tapped).not.toBe('A1');
    // A touch that moves is a pan.
    viewer.select({ row: 0, col: 0 });
    touch('pointerdown', 150, 60);
    touch('pointermove', 150, 160);
    touch('pointerup', 150, 160);
    expect(viewer.snapshot().selectionRef).toBe('A1');
    // A cancelled touch selects nothing.
    touch('pointerdown', 150, 60);
    touch('pointercancel', 150, 60);
    touch('pointerup', 150, 60);
    expect(viewer.snapshot().selectionRef).toBe('A1');
  });

  test('a heading, corner, or Mod+A selection clears the old refusal and is announced', async () => {
    const { viewer, host, scroller } = mount();
    await viewer.load(
      buildXlsx({ sheets: [{ name: 'S', headings: true, rows: row(1, [c('A1', 1)]) }] })
    );
    flushEditor(viewer);
    const live = host.querySelector('.xlsx-live') as HTMLElement;
    const stage = host.querySelector('.xlsx-stage') as HTMLElement;
    const W = host.ownerDocument.defaultView as unknown as typeof globalThis;
    const press = (clientX: number, clientY: number) =>
      stage.dispatchEvent(
        new W.PointerEvent('pointerdown', { bubbles: true, button: 0, clientX, clientY })
      );
    // In view mode an edit is refused, and the refusal stays until the next selection.
    viewer.setMode('view');
    const refuse = () => {
      expect(viewer.enterValue('2').ok).toBe(false);
      expect(viewer.snapshot().refusal).not.toBeNull();
    };
    refuse();
    press(150, 4);
    expect(viewer.snapshot().refusal).toBeNull();
    expect(live.textContent).toBe('grid.selectedColumn{"first":"B"}');
    refuse();
    press(2, 2);
    expect(viewer.snapshot().refusal).toBeNull();
    expect(live.textContent).toBe('grid.selectedSheet');
    viewer.select({ row: 0, col: 0 });
    refuse();
    key(scroller, 'a', { ctrlKey: true });
    expect(viewer.snapshot().refusal).toBeNull();
    expect(live.textContent).toBe('grid.selectedSheet');
    // Whole rows and columns announce in words; another range by its reference.
    viewer.select({ row: 1, col: 0 }, { row: 3, col: 16383 });
    expect(live.textContent).toBe('grid.selectedRows{"first":2,"last":4}');
    viewer.select({ row: 0, col: 1 }, { row: 1048575, col: 2 });
    expect(live.textContent).toBe('grid.selectedColumns{"first":"B","last":"C"}');
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 1 });
    expect(live.textContent).toBe('grid.selectedRange{"range":"A1:B2"}');
  });

  test('a long sheet scrolls to row 100000 and paints only the rows in the window', async () => {
    const { viewer, host, scroller } = mount(800, 400);
    await viewer.load(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            headings: true,
            rows: row(1, [c('A1', 'top')]) + row(100000, [c('A100000', 'bottom')]),
          },
        ],
      })
    );
    flushEditor(viewer);
    key(scroller, 'ArrowDown', { ctrlKey: true });
    expect(viewer.snapshot().activeCell?.ref).toBe('A100000');
    flushEditor(viewer);
    expect(host.querySelector('[data-cell="A100000"]')?.textContent).toBe('bottom');
    const rows = host.querySelectorAll('.xlsx-row-heads .xlsx-head');
    expect(rows.length).toBeGreaterThan(5);
    expect(rows.length).toBeLessThan(40);
  });

  test('zoom starts from the sheet view and follows the user', async () => {
    const { viewer } = mount();
    await viewer.load(
      buildXlsx({
        sheets: [{ name: 'S', viewAttrs: 'zoomScale="150"', rows: row(1, [c('A1', 1)]) }],
      })
    );
    expect(viewer.snapshot().zoom).toBe(1.5);
    viewer.zoomStep(1);
    expect(viewer.snapshot().zoom).toBe(1.75);
    viewer.setZoom(10);
    expect(viewer.snapshot().zoom).toBe(4);
  });

  test('Mod+wheel zooms one step for each 100 pixels of deltaY, and a 0 delta does nothing', async () => {
    const { viewer, scroller } = mount();
    await viewer.load(buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] }));
    const W = (scroller.ownerDocument.defaultView as unknown as typeof globalThis).WheelEvent;
    const wheel = (deltaY: number) => {
      const e = new W('wheel', { deltaY, cancelable: true });
      // The test window's wheel event does not read `ctrlKey` from its init.
      Object.defineProperty(e, 'ctrlKey', { value: true });
      scroller.dispatchEvent(e);
    };
    expect(viewer.snapshot().zoom).toBe(1);
    wheel(0);
    expect(viewer.snapshot().zoom).toBe(1);
    for (let i = 0; i < 9; i += 1) wheel(-10);
    expect(viewer.snapshot().zoom).toBe(1);
    wheel(-10);
    expect(viewer.snapshot().zoom).toBe(1.1);
    wheel(100);
    expect(viewer.snapshot().zoom).toBe(1);
  });

  test('zoom, undo, and redo are announced, once each, with the selection they restore', async () => {
    const { viewer, host } = mount();
    await viewer.load(
      buildXlsx({
        sheets: [
          { name: 'S', rows: row(1, [c('A1', 1)]) },
          { name: 'T', rows: row(1, [c('A1', 2)]) },
        ],
      })
    );
    const live = () => host.querySelector('.xlsx-live')?.textContent;
    viewer.setZoom(2);
    expect(live()).toBe('grid.zoom{"percent":200}');
    viewer.select({ row: 0, col: 0 });
    expect(viewer.enterValue('5').ok).toBe(true);
    const a1 = (content: string) => `grid.activeCell{"ref":"A1","content":"${content}"}`;
    expect(viewer.undo().ok).toBe(true);
    expect(live()).toBe(`grid.undoneAt${JSON.stringify({ selection: a1('1') })}`);
    expect(viewer.redo().ok).toBe(true);
    expect(live()).toBe(`grid.redoneAt${JSON.stringify({ selection: a1('5') })}`);
    // The same text again is announced again, with one invisible character.
    expect(viewer.toggleFormat('bold').ok).toBe(true);
    expect(viewer.toggleFormat('bold').ok).toBe(true);
    expect(viewer.undo().ok).toBe(true);
    expect(viewer.undo().ok).toBe(true);
    expect(live()).toBe(`grid.undoneAt${JSON.stringify({ selection: a1('5') })}\u00a0`);
    // An undo on another sheet shows that sheet and names it in the one announcement.
    expect(viewer.setActiveSheet(1)).toBe(true);
    expect(viewer.undo().ok).toBe(true);
    expect(viewer.snapshot().activeSheet).toBe(0);
    expect(viewer.snapshot().selectionRef).toBe('A1');
    expect(live()).toBe(`grid.undoneOnSheet${JSON.stringify({ sheet: 'S', selection: a1('1') })}`);
    // A value written to another cell keeps no selection: undo leaves the selection alone.
    viewer.select({ row: 4, col: 4 });
    expect(viewer.enterValue('9', { row: 9, col: 9 }).ok).toBe(true);
    expect(viewer.undo().ok).toBe(true);
    expect(viewer.snapshot().selectionRef).toBe('E5');
    expect(live()).toBe('grid.undone');
    // The same edit, undone from another sheet: the shown sheet and its selection stay, and the
    // announcement names the sheet that the undo changed.
    expect(viewer.enterValue('9', { row: 9, col: 9 }).ok).toBe(true);
    expect(viewer.setActiveSheet(1)).toBe(true);
    viewer.select({ row: 2, col: 2 });
    expect(viewer.undo().ok).toBe(true);
    expect(viewer.snapshot().activeSheet).toBe(1);
    expect(viewer.snapshot().selectionRef).toBe('C3');
    expect(live()).toBe('grid.undoneSheet{"sheet":"S"}');
    expect(viewer.redo().ok).toBe(true);
    expect(live()).toBe('grid.redoneSheet{"sheet":"S"}');
  });

  test('a formula shows its text in the active cell content', async () => {
    const { viewer } = mount();
    await viewer.load(fixture('basic-values'));
    viewer.select({ row: 4, col: 0 });
    expect(viewer.snapshot().activeCell).toEqual({ ref: 'A5', content: '=A2*2' });
  });
});

describe('the grid focus ring', () => {
  const css = readFileSync(`${import.meta.dir}/../../styles/editor.css`, 'utf8');
  const rule = (selector: string) => {
    const at = css.indexOf(`\n${selector} {`);
    expect(at).toBeGreaterThanOrEqual(0);
    return css.slice(at, css.indexOf('}', at));
  };

  test('the ring is an overlay above the stage and the cell editor, shown with the focus', () => {
    const overlay = rule('.xlsx-surface::after');
    expect(overlay).toContain('position: absolute;');
    expect(overlay).toContain('pointer-events: none;');
    const ringZ = Number(/z-index: (\d+);/.exec(overlay)![1]);
    const editorZ = Number(/z-index: (\d+);/.exec(rule('.xlsx-cell-editor'))![1]);
    expect(ringZ).toBeGreaterThan(editorZ);
    expect(rule('.xlsx-surface:focus-within::after')).toContain('var(--doc-focus-ring)');
  });

  test('the active cell border differs when the grid has no focus', () => {
    expect(rule('.xlsx-active-cell')).toContain('var(--doc-primary)');
    expect(rule('.xlsx-surface:not(:focus-within) .xlsx-active-cell')).toContain(
      'var(--doc-text-muted)'
    );
  });

  test('the surface that carries the ring holds the grid', async () => {
    const { viewer, host } = mount();
    await viewer.load(buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] }));
    flushEditor(viewer);
    const surface = host.querySelector('.xlsx-surface')!;
    expect(surface.querySelector('.xlsx-scroller')).not.toBeNull();
    expect(surface.querySelector('.xlsx-active-cell')).not.toBeNull();
    // Editable content of the host page does not reach the surface.
    expect(surface.getAttribute('contenteditable')).toBe('false');
  });

  test('the grid states the sheet size, since it paints only the visible cells', () => {
    const { host } = mount();
    const grid = host.querySelector('[role="grid"]')!;
    expect(grid.getAttribute('aria-rowcount')).toBe('1048576');
    expect(grid.getAttribute('aria-colcount')).toBe('16384');
  });
});

describe('undo and redo restore the selection of their command', () => {
  /** A sheet with values in A1:C5, scrolled far away from them by a selection. */
  async function scrolledAway(band: (scroller: HTMLElement) => void) {
    const { viewer, scroller } = mount();
    let rows = '';
    for (let r = 1; r <= 5; r += 1) rows += row(r, [c(`A${r}`, r), c(`B${r}`, r), c(`C${r}`, r)]);
    await viewer.load(buildXlsx({ sheets: [{ name: 'S', rows }] }));
    flushEditor(viewer);
    viewer.select({ row: 1, col: 1 });
    band(scroller);
    flushEditor(viewer);
    return { viewer, scroller };
  }
  // A reveal of B2 from far below and to the right puts B2 at the top-left corner of the
  // window: one column and one row in. A reveal of the band end would scroll far past them.
  const nearB2 = (scroller: HTMLElement) => {
    expect(scroller.scrollLeft).toBeLessThan(100);
    expect(scroller.scrollTop).toBeLessThan(40);
  };
  const away = (viewer: XlsxEditorInstance, scroller: HTMLElement) => {
    viewer.select({ row: 400, col: 60 });
    flushEditor(viewer);
    expect([scroller.scrollLeft, scroller.scrollTop].every((n) => n > 0)).toBe(true);
  };

  test('undo of a whole-column command reveals the active cell, not the end of the band', async () => {
    const { viewer, scroller } = await scrolledAway((s) => key(s, ' ', { ctrlKey: true }));
    expect(viewer.snapshot().selectionRef).toBe('B1:B1048576');
    expect(viewer.deleteColumns().ok).toBe(true);
    away(viewer, scroller);
    expect(viewer.undo().ok).toBe(true);
    flushEditor(viewer);
    expect(viewer.snapshot().selectionRef).toBe('B1:B1048576');
    // The band keeps its active cell, which the keyboard left at B2.
    expect(viewer.snapshot().activeCell?.ref).toBe('B2');
    nearB2(scroller);
    // Redo takes the selection after the command, revealed the same way.
    away(viewer, scroller);
    expect(viewer.redo().ok).toBe(true);
    flushEditor(viewer);
    expect(viewer.snapshot().selectionRef).toBe('B1:B1048576');
    nearB2(scroller);
  });

  test('undo of a whole-row command reveals the active cell, not the end of the row', async () => {
    const { viewer, scroller } = await scrolledAway((s) => key(s, ' ', { shiftKey: true }));
    expect(viewer.snapshot().selectionRef).toBe('A2:XFD2');
    expect(viewer.deleteRows().ok).toBe(true);
    away(viewer, scroller);
    expect(viewer.undo().ok).toBe(true);
    flushEditor(viewer);
    expect(viewer.snapshot().selectionRef).toBe('A2:XFD2');
    expect(viewer.snapshot().activeCell?.ref).toBe('B2');
    nearB2(scroller);
  });

  test('undo of a whole-sheet command reveals the first cell of the sheet', async () => {
    const { viewer, scroller } = await scrolledAway((s) => key(s, 'a', { ctrlKey: true }));
    expect(viewer.clearValues().ok).toBe(true);
    away(viewer, scroller);
    expect(viewer.undo().ok).toBe(true);
    flushEditor(viewer);
    expect(viewer.snapshot().activeCell?.ref).toBe('A1');
    expect(viewer.snapshot().activeCell?.content).toBe('1');
    expect([scroller.scrollLeft, scroller.scrollTop]).toEqual([0, 0]);
  });
});

describe('the selection after undo extends as before the edit', () => {
  test('Shift+Arrow extends a restored range and a restored band from their focus corner', async () => {
    const { viewer, scroller } = mount();
    await viewer.load(
      buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 2), c('C1', 3)]) }] })
    );
    flushEditor(viewer);
    // A range from B2 to C3, with C3 as the focus corner.
    viewer.select({ row: 1, col: 1 }, { row: 2, col: 2 });
    expect(viewer.toggleFormat('bold').ok).toBe(true);
    viewer.select({ row: 9, col: 9 });
    expect(viewer.undo().ok).toBe(true);
    expect(viewer.snapshot().selectionRef).toBe('B2:C3');
    key(scroller, 'ArrowDown', { shiftKey: true });
    expect(viewer.snapshot().selectionRef).toBe('B2:C4');
    // Column B as a band, extended to column C before the edit.
    viewer.select({ row: 1, col: 1 });
    key(scroller, ' ', { ctrlKey: true });
    key(scroller, 'ArrowRight', { shiftKey: true });
    expect(viewer.snapshot().selectionRef).toBe('B1:C1048576');
    expect(viewer.deleteColumns().ok).toBe(true);
    viewer.select({ row: 9, col: 9 });
    const u = viewer.undo();
    expect(u.ok ? 'ok' : u.refusal).toBe('ok');
    expect(viewer.snapshot().selectionRef).toBe('B1:C1048576');
    key(scroller, 'ArrowRight', { shiftKey: true });
    expect(viewer.snapshot().selectionRef).toBe('B1:D1048576');
  });
});

describe('select with an open cell edit', () => {
  test('commits the edit first, and a refused commit keeps the selection', async () => {
    const { viewer } = mount();
    await viewer.load(buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] }));
    viewer.select({ row: 0, col: 0 });
    expect(viewer.beginEdit('7').ok).toBe(true);
    expect(viewer.select({ row: 4, col: 4 })).toBe(true);
    expect(viewer.snapshot().editing).toBe(false);
    viewer.select({ row: 0, col: 0 });
    expect(viewer.snapshot().activeCell?.content).toBe('7');
    // Text that the store refuses keeps the editor open and the selection where it is.
    expect(viewer.beginEdit('=SUM(').ok).toBe(true);
    expect(viewer.select({ row: 4, col: 4 })).toBe(false);
    expect(viewer.snapshot().editing).toBe(true);
    expect(viewer.snapshot().selectionRef).toBe('A1');
    viewer.cancelEdit();
  });

  test('typed text without Enter shows as a changed draft, and load drops it', async () => {
    const { viewer, scroller } = mount();
    const bytes = buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] });
    await viewer.load(bytes);
    viewer.select({ row: 0, col: 0 });
    key(scroller, '5');
    // A host reads these two flags to ask before it opens other bytes.
    expect(viewer.snapshot().editing).toBe(true);
    expect(viewer.snapshot().editChanged).toBe(true);
    expect(viewer.snapshot().revision).toBe(0);
    await viewer.load(bytes);
    expect(viewer.snapshot().editing).toBe(false);
    expect(viewer.snapshot().activeCell?.content).toBe('1');
  });

  test('undo with an open edit of unchanged text undoes a step on another sheet', async () => {
    const { viewer, host } = mount();
    await viewer.load(
      buildXlsx({
        sheets: [
          { name: 'S', rows: row(1, [c('A1', 1)]) },
          { name: 'T', rows: row(1, [c('A1', 2)]) },
        ],
      })
    );
    viewer.select({ row: 0, col: 0 });
    expect(viewer.enterValue('5').ok).toBe(true);
    expect(viewer.setActiveSheet(1)).toBe(true);
    expect(viewer.beginEdit().ok).toBe(true);
    expect(viewer.snapshot().editChanged).toBe(false);
    expect(viewer.undo().ok).toBe(true);
    // The unchanged edit closes and adds no step, so the undo reverts the edit on S and shows S.
    expect(viewer.snapshot().editing).toBe(false);
    expect(viewer.snapshot().activeSheet).toBe(0);
    expect(viewer.snapshot().activeCell?.content).toBe('1');
    const a1 = JSON.stringify({ ref: 'A1', content: '1' });
    expect(host.querySelector('.xlsx-live')?.textContent).toBe(
      `grid.undoneOnSheet${JSON.stringify({ sheet: 'S', selection: `grid.activeCell${a1}` })}`
    );
  });

  test('enterValue commits an open edit of its cell first', async () => {
    const { viewer } = mount();
    await viewer.load(buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] }));
    viewer.select({ row: 0, col: 0 });
    expect(viewer.beginEdit('abc', 'enter').ok).toBe(true);
    expect(viewer.enterValue('x').ok).toBe(true);
    expect(viewer.snapshot().editing).toBe(false);
    expect(viewer.snapshot().activeCell?.content).toBe('x');
    // The open edit is its own undo step, so its text is not lost.
    expect(viewer.undo().ok).toBe(true);
    expect(viewer.snapshot().activeCell?.content).toBe('abc');
    // A write to another cell leaves an open edit open.
    expect(viewer.beginEdit('def', 'enter').ok).toBe(true);
    expect(viewer.enterValue('y', { row: 4, col: 4 }).ok).toBe(true);
    expect(viewer.snapshot().editing).toBe(true);
    viewer.cancelEdit();
  });
});

describe('the announcement of an undo or a redo', () => {
  test('names a sheet that is not shown, also without a selection', () => {
    const t = (k: string, p?: Record<string, string>) => `${k}${p ? JSON.stringify(p) : ''}`;
    expect(stepAnnouncement(t, 'undo', undefined, undefined)).toBe('grid.undone');
    expect(stepAnnouncement(t, 'redo', undefined, 'Hidden')).toBe(
      'grid.redoneSheet{"sheet":"Hidden"}'
    );
    expect(stepAnnouncement(t, 'undo', 'A1', 'S')).toBe(
      'grid.undoneOnSheet{"sheet":"S","selection":"A1"}'
    );
  });
});

/** Dispatch pointer events with button 0 on the stage of `host`, as one pointer. */
function stagePointer(host: HTMLElement, pointerId: number) {
  const stage = host.querySelector('.xlsx-stage') as HTMLElement;
  const W = host.ownerDocument.defaultView as unknown as typeof globalThis;
  return (type: string, clientX: number, clientY: number, shiftKey = false) =>
    stage.dispatchEvent(
      new W.PointerEvent(type, { bubbles: true, button: 0, pointerId, clientX, clientY, shiftKey })
    );
}

describe('Shift extends the selection and keeps the active cell that the range holds', () => {
  async function inRange() {
    const { viewer, scroller, host } = mount();
    await viewer.load(
      buildXlsx({ sheets: [{ name: 'S', headings: true, rows: row(1, [c('A1', 1)]) }] })
    );
    flushEditor(viewer);
    // B2:D4, and Enter moves the active cell inside it to B3.
    viewer.select({ row: 1, col: 1 }, { row: 3, col: 3 });
    key(scroller, 'Enter');
    expect(viewer.snapshot().activeCell?.ref).toBe('B3');
    return { viewer, scroller, host };
  }

  test('Enter, then Shift+Arrow', async () => {
    const { viewer, scroller } = await inRange();
    key(scroller, 'ArrowDown', { shiftKey: true });
    expect(viewer.snapshot().selectionRef).toBe('B2:D5');
    expect(viewer.snapshot().activeCell?.ref).toBe('B3');
    // A range that no longer holds the active cell takes the anchor as the active cell.
    for (let i = 0; i < 4; i += 1) key(scroller, 'ArrowUp', { shiftKey: true });
    expect(viewer.snapshot().selectionRef).toBe('B1:D2');
    expect(viewer.snapshot().activeCell?.ref).toBe('B2');
  });

  test('undo, then Shift+Arrow, with an active cell that is not the anchor', async () => {
    const { viewer, scroller } = await inRange();
    expect(viewer.toggleFormat('bold').ok).toBe(true);
    viewer.select({ row: 9, col: 9 });
    expect(viewer.undo().ok).toBe(true);
    expect(viewer.snapshot().selectionRef).toBe('B2:D4');
    expect(viewer.snapshot().activeCell?.ref).toBe('B3');
    key(scroller, 'ArrowRight', { shiftKey: true });
    expect(viewer.snapshot().selectionRef).toBe('B2:E4');
    expect(viewer.snapshot().activeCell?.ref).toBe('B3');
  });

  test('Shift+click', async () => {
    const { viewer, host } = await inRange();
    const stage = host.querySelector('.xlsx-stage') as HTMLElement;
    const W = host.ownerDocument.defaultView as unknown as typeof globalThis;
    const press = (shiftKey: boolean) =>
      stage.dispatchEvent(
        new W.PointerEvent('pointerdown', {
          bubbles: true,
          button: 0,
          clientX: 500,
          clientY: 300,
          shiftKey,
        })
      );
    press(true);
    const ref = viewer.snapshot().selectionRef ?? '';
    expect(ref.startsWith('B2:')).toBe(true);
    expect(viewer.snapshot().activeCell?.ref).toBe('B3');
    // The far corner is the pressed cell.
    press(false);
    expect(ref.endsWith(`:${viewer.snapshot().selectionRef}`)).toBe(true);
  });

  test('Shift+click on a heading, and a drag after a Shift+click', async () => {
    const { viewer, host } = await inRange();
    const pointer = stagePointer(host, 3);
    // A Shift press on a column heading right of D extends B2:D4 to whole columns from B. The
    // press is inside the heading, away from the resize band of its edges.
    pointer('pointerdown', 420, 4, true);
    expect(viewer.snapshot().selectionRef?.startsWith('B1:')).toBe(true);
    expect(viewer.snapshot().selectionRef?.endsWith('1048576')).toBe(true);
    expect(viewer.snapshot().activeCell?.ref).toBe('B3');
    pointer('pointerup', 420, 4);
    // Back to B2:D4 with B3 active, then a Shift press on a cell and a drag over the grid.
    viewer.select({ row: 1, col: 1 }, { row: 3, col: 3 });
    const scroller = host.querySelector('.xlsx-scroller') as HTMLElement;
    key(scroller, 'Enter');
    pointer('pointerdown', 500, 300, true);
    pointer('pointermove', 450, 260);
    const ref = viewer.snapshot().selectionRef ?? '';
    expect(ref.startsWith('B2:')).toBe(true);
    expect(ref).not.toBe('B2:D4');
    expect(viewer.snapshot().activeCell?.ref).toBe('B3');
    pointer('pointerup', 450, 260);
  });

  test('Shift and a press on a row heading keep the active cell only when the band holds it', async () => {
    const { viewer, host } = await inRange();
    const pointer = stagePointer(host, 4);
    // Rows from 2 down to a row below the range: B3 is in the band, so it stays active.
    // The press at y 310 is on row 15, clear of the resize edge at y 300.
    pointer('pointerdown', 10, 310, true);
    pointer('pointerup', 10, 310);
    expect(viewer.snapshot().selectionRef).toBe('A2:XFD15');
    expect(viewer.snapshot().activeCell?.ref).toBe('B3');
    // Back to B2:D4, where Enter moves the active cell from B2 to B3. Then rows 1 to 2: the
    // band does not hold B3, so the anchor B2 becomes the active cell.
    viewer.select({ row: 1, col: 1 }, { row: 3, col: 3 });
    key(host.querySelector('.xlsx-scroller') as HTMLElement, 'Enter');
    pointer('pointerdown', 10, 30, true);
    pointer('pointerup', 10, 30);
    expect(viewer.snapshot().selectionRef).toBe('A1:XFD2');
    expect(viewer.snapshot().activeCell?.ref).toBe('B2');
  });
});
