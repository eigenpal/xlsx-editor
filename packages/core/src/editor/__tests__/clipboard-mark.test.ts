// The dashed outline of a pending copy or cut: where it shows, what ends it, and that it stays
// out of the workbook. The clipboard snapshot is its only state.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import {
  buildXlsx,
  c,
  row,
  type FixtureWorkbook,
} from '../../../../../scripts/fixtures/builder.ts';
import { createXlsxEditor, type XlsxEditorInstance } from '../index.ts';
import { rangeOutline, type ViewState } from '../view-map.ts';
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

const hosts = new WeakMap<XlsxEditorInstance, HTMLElement>();

async function mount(wb: FixtureWorkbook) {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.appendChild(host);
  const viewer = createXlsxEditor({ fonts: FONTS, translate: (key) => key });
  viewers.push(viewer);
  viewer.attach(host);
  const scroller = host.querySelector('.xlsx-scroller') as HTMLElement;
  Object.defineProperty(scroller, 'clientWidth', { get: () => 800 });
  Object.defineProperty(scroller, 'clientHeight', { get: () => 400 });
  await viewer.load(buildXlsx(wb));
  flushEditor(viewer);
  hosts.set(viewer, host);
  return viewer;
}

const BOOK: FixtureWorkbook = {
  sheets: [
    { name: 'S', rows: row(1, [c('A1', 1), c('B1', 2)]) },
    { name: 'T', rows: row(1, [c('A1', 3)]) },
  ],
};

/** The sides of the outline that the grid shows now. */
function sides(v: XlsxEditorInstance): string[] {
  flushEditor(v);
  const edges = hosts.get(v)!.querySelectorAll<HTMLElement>('.xlsx-overlay .xlsx-pending-edge');
  return [...edges].map((e) => e.dataset.side ?? '').sort();
}

const WHOLE = ['bottom', 'left', 'right', 'top'];

/** Press Escape on the grid. With `used`, the page takes the key first, as a message does. */
function escape(v: XlsxEditorInstance, used = false, repeat = false): boolean {
  const host = hosts.get(v)!;
  if (used)
    host.addEventListener('keydown', (e) => e.preventDefault(), { capture: true, once: true });
  const scroller = host.querySelector('.xlsx-scroller') as HTMLElement;
  const W = scroller.ownerDocument.defaultView as unknown as typeof globalThis;
  const e = new W.KeyboardEvent('keydown', {
    key: 'Escape',
    bubbles: true,
    cancelable: true,
    repeat,
  });
  scroller.dispatchEvent(e);
  return e.defaultPrevented;
}

/** Press Tab on the grid. False when the browser takes the key and leaves the grid. */
function tab(v: XlsxEditorInstance): boolean {
  const scroller = hosts.get(v)!.querySelector('.xlsx-scroller') as HTMLElement;
  const W = scroller.ownerDocument.defaultView as unknown as typeof globalThis;
  const e = new W.KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true });
  scroller.dispatchEvent(e);
  return e.defaultPrevented;
}

function copied(v: XlsxEditorInstance, mode: 'copy' | 'cut' = 'copy') {
  v.select({ row: 0, col: 0 });
  const r = mode === 'cut' ? v.cut() : v.copy();
  if (!r.ok) throw new Error(r.refusal.code);
  return r.content;
}

describe('the outline of a pending copy', () => {
  test('shows around the source, and stays after a paste until the next edit', async () => {
    const v = await mount(BOOK);
    expect(sides(v)).toEqual([]);
    const { token } = copied(v);
    expect(sides(v)).toEqual(WHOLE);
    v.select({ row: 3, col: 0 });
    expect(v.paste({ token }).ok).toBe(true);
    expect(sides(v)).toEqual(WHOLE);
    expect(v.enterValue('9', { row: 5, col: 0 }).ok).toBe(true);
    expect(sides(v)).toEqual([]);
    // A paste after that edit does not bring the outline back.
    v.select({ row: 4, col: 0 });
    expect(v.paste({ token }).ok).toBe(true);
    expect(sides(v)).toEqual([]);
  });

  test('Escape ends only the outline, says so, and the next Escape goes on', async () => {
    const v = await mount(BOOK);
    const { token } = copied(v);
    expect(escape(v)).toBe(true);
    expect(sides(v)).toEqual([]);
    expect(hosts.get(v)!.querySelector('.xlsx-live')?.textContent).toBe('grid.copyUnmarked');
    expect(escape(v)).toBe(false);
    // The copy stays, so the paste writes the copied cell and not the text of the clipboard.
    v.select({ row: 3, col: 0 });
    expect(v.paste({ token, text: 'x' }).ok).toBe(true);
    expect(v.snapshot().activeCell?.content).toBe('1');
    expect(sides(v)).toEqual([]);
  });

  test('scenario "The next Escape goes on": the second Escape arms the exit', async () => {
    const v = await mount(BOOK);
    copied(v);
    v.select({ row: 0, col: 0 });
    expect(escape(v)).toBe(true);
    expect(escape(v)).toBe(false);
    expect(hosts.get(v)!.querySelector('.xlsx-live')?.textContent).toBe('grid.exitHint');
    expect(tab(v)).toBe(false);
    expect(v.snapshot().selection?.active).toEqual({ row: 0, col: 0 });
  });

  test('scenario "Escape on another sheet after a copy": the outline ends', async () => {
    const v = await mount(BOOK);
    const { token } = copied(v);
    expect(v.setActiveSheet(1)).toBe(true);
    v.select({ row: 0, col: 0 });
    expect(escape(v)).toBe(true);
    expect(hosts.get(v)!.querySelector('.xlsx-live')?.textContent).toBe('grid.copyUnmarked');
    expect(escape(v)).toBe(false);
    expect(v.setActiveSheet(0)).toBe(true);
    expect(sides(v)).toEqual([]);
    v.select({ row: 3, col: 0 });
    expect(v.paste({ token, text: 'x' }).ok).toBe(true);
    expect(v.snapshot().activeCell?.content).toBe('1');
  });

  test('scenario "Escape on another sheet after a cut": the cut ends', async () => {
    const v = await mount(BOOK);
    const { token } = copied(v, 'cut');
    expect(v.setActiveSheet(1)).toBe(true);
    v.select({ row: 0, col: 0 });
    expect(escape(v)).toBe(true);
    expect(hosts.get(v)!.querySelector('.xlsx-live')?.textContent).toBe('grid.cutEnded');
    expect(escape(v)).toBe(false);
    expect(v.setActiveSheet(0)).toBe(true);
    expect(sides(v)).toEqual([]);
    v.select({ row: 0, col: 2 });
    expect(v.paste({ token, text: 'x' }).ok).toBe(true);
    expect(v.snapshot().activeCell?.content).toBe('x');
    v.select({ row: 0, col: 0 });
    expect(v.snapshot().activeCell?.content).toBe('1');
  });

  test('a formula copied before Escape pastes as a formula', async () => {
    const v = await mount({
      sheets: [
        {
          name: 'Sales',
          rows: row(5, [
            c('B5', 1),
            c('C5', 2),
            c('D5', 3),
            c('E5', 4),
            c('F5', 10, { f: 'SUM(B5:E5)' }),
          ]),
        },
      ],
    });
    v.select({ row: 4, col: 5 });
    const r = v.copy();
    if (!r.ok) throw new Error(r.refusal.code);
    expect(escape(v)).toBe(true);
    v.select({ row: 19, col: 9 });
    expect(v.paste({ token: r.content.token, text: '10' }).ok).toBe(true);
    expect(v.snapshot().activeCell?.content).toBe('=SUM(F20:I20)');
  });

  test('a copy after Escape still expires when a sheet that it names changes', async () => {
    const v = await mount({
      sheets: [
        { name: 'S', rows: row(1, [c('A1', 3, { f: 'T!A1' })]) },
        { name: 'T', rows: row(1, [c('A1', 3)]) },
      ],
    });
    const { token } = copied(v);
    expect(escape(v)).toBe(true);
    expect(v.renameSheet(1, 'U').ok).toBe(true);
    expect(v.setActiveSheet(0)).toBe(true);
    v.select({ row: 3, col: 0 });
    const r = v.paste({ token, text: 'x' });
    expect(r.ok ? null : r.refusal.code).toBe('copy-expired');
    expect(v.snapshot().activeCell?.content).toBe('');
  });

  test('a copy after Escape does not reach another workbook', async () => {
    const v = await mount(BOOK);
    const { token } = copied(v);
    expect(escape(v)).toBe(true);
    await v.load(buildXlsx(BOOK));
    flushEditor(v);
    v.select({ row: 3, col: 0 });
    expect(v.paste({ token, text: 'x' }).ok).toBe(true);
    expect(v.snapshot().activeCell?.content).toBe('x');
  });

  test('an Escape that the page used leaves it, and the next Escape ends it', async () => {
    const v = await mount(BOOK);
    const { token } = copied(v);
    escape(v, true);
    expect(sides(v)).toEqual(WHOLE);
    expect(escape(v)).toBe(true);
    expect(sides(v)).toEqual([]);
    v.select({ row: 3, col: 0 });
    expect(v.paste({ token }).ok).toBe(true);
    expect(v.snapshot().activeCell?.content).toBe('1');
  });

  test('a paste over the source ends it', async () => {
    const v = await mount({
      sheets: [
        { name: 'S', rows: row(1, [c('A1', 1), c('B1', 2)]) + row(2, [c('A2', 3), c('B2', 4)]) },
      ],
    });
    v.select({ row: 0, col: 0 }, { row: 1, col: 1 });
    const r = v.copy();
    if (!r.ok) throw new Error(r.refusal.code);
    expect(sides(v)).toEqual(WHOLE);
    v.select({ row: 1, col: 1 });
    expect(v.paste({ token: r.content.token }).ok).toBe(true);
    expect(v.snapshot().activeCell?.content).toBe('1');
    expect(sides(v)).toEqual([]);
  });

  test('a paste over the source that changes no source cell keeps it', async () => {
    const v = await mount({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) + row(2, [c('A2', 1)]) }],
    });
    v.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    const r = v.copy();
    if (!r.ok) throw new Error(r.refusal.code);
    v.select({ row: 1, col: 0 });
    expect(v.paste({ token: r.content.token }).ok).toBe(true);
    expect(sides(v)).toEqual(WHOLE);
    v.select({ row: 2, col: 0 });
    expect(v.snapshot().activeCell?.content).toBe('1');
  });

  test('a paste on the source itself and a paste beside it keep it', async () => {
    const v = await mount(BOOK);
    const { token } = copied(v);
    v.select({ row: 0, col: 0 });
    expect(v.paste({ token }).ok).toBe(true);
    expect(sides(v)).toEqual(WHOLE);
    v.select({ row: 0, col: 1 });
    expect(v.paste({ token }).ok).toBe(true);
    expect(sides(v)).toEqual(WHOLE);
  });

  test('a paste at the same cells of another sheet keeps it', async () => {
    const v = await mount(BOOK);
    const { token } = copied(v);
    expect(v.setActiveSheet(1)).toBe(true);
    v.select({ row: 0, col: 0 });
    expect(v.paste({ token }).ok).toBe(true);
    expect(v.setActiveSheet(0)).toBe(true);
    expect(sides(v)).toEqual(WHOLE);
  });

  test('an undo and a redo end it', async () => {
    const v = await mount(BOOK);
    expect(v.enterValue('5', { row: 6, col: 0 }).ok).toBe(true);
    copied(v);
    expect(sides(v)).toEqual(WHOLE);
    expect(v.undo().ok).toBe(true);
    expect(sides(v)).toEqual([]);
    copied(v);
    expect(v.redo().ok).toBe(true);
    expect(sides(v)).toEqual([]);
  });

  test('shows only on the sheet of the source, and comes back with that sheet', async () => {
    const v = await mount(BOOK);
    copied(v);
    expect(v.setActiveSheet(1)).toBe(true);
    expect(sides(v)).toEqual([]);
    expect(v.setActiveSheet(0)).toBe(true);
    expect(sides(v)).toEqual(WHOLE);
  });

  test('a paste on another sheet keeps it on the source sheet', async () => {
    const v = await mount(BOOK);
    const { token } = copied(v);
    expect(v.setActiveSheet(1)).toBe(true);
    v.select({ row: 3, col: 0 });
    expect(v.paste({ token }).ok).toBe(true);
    expect(sides(v)).toEqual([]);
    expect(v.setActiveSheet(0)).toBe(true);
    expect(sides(v)).toEqual(WHOLE);
  });

  test('a load of another workbook ends it', async () => {
    const v = await mount(BOOK);
    copied(v);
    await v.load(buildXlsx(BOOK));
    expect(sides(v)).toEqual([]);
  });

  test('stays out of the saved workbook and the canvas', async () => {
    const v = await mount(BOOK);
    const before = v.save();
    const canvas = hosts.get(v)!.querySelector('.xlsx-grid');
    copied(v);
    expect(sides(v)).toEqual(WHOLE);
    expect(canvas?.querySelector('.xlsx-pending-edge')).toBeNull();
    expect(v.save()).toEqual(before);
  });
});

describe('the outline of a pending cut', () => {
  test('a paste of the cut ends it', async () => {
    const v = await mount(BOOK);
    const { token } = copied(v, 'cut');
    expect(sides(v)).toEqual(WHOLE);
    v.select({ row: 3, col: 0 });
    expect(v.paste({ token }).ok).toBe(true);
    expect(sides(v)).toEqual([]);
    // The second paste of the cut is a copy, with no outline.
    v.select({ row: 4, col: 0 });
    expect(v.paste({ token }).ok).toBe(true);
    expect(sides(v)).toEqual([]);
  });

  test('Escape, an edit, and a plain paste each end it', async () => {
    const v = await mount(BOOK);
    copied(v, 'cut');
    expect(escape(v)).toBe(true);
    expect(sides(v)).toEqual([]);
    copied(v, 'cut');
    expect(v.enterValue('5', { row: 6, col: 0 }).ok).toBe(true);
    expect(sides(v)).toEqual([]);
    copied(v, 'cut');
    v.select({ row: 3, col: 0 });
    expect(v.paste({ text: 'plain' }).ok).toBe(true);
    expect(sides(v)).toEqual([]);
  });

  test('Escape forgets the cut: the cells stay, and the paste takes the plain lane', async () => {
    const v = await mount(BOOK);
    const { token } = copied(v, 'cut');
    expect(escape(v)).toBe(true);
    expect(hosts.get(v)!.querySelector('.xlsx-live')?.textContent).toBe('grid.cutEnded');
    v.select({ row: 3, col: 0 });
    expect(v.paste({ token, text: 'x' }).ok).toBe(true);
    expect(v.snapshot().activeCell?.content).toBe('x');
    v.select({ row: 0, col: 0 });
    expect(v.snapshot().activeCell?.content).toBe('1');
  });

  test('an Escape that the page used leaves it, and the next Escape ends it', async () => {
    const v = await mount(BOOK);
    copied(v, 'cut');
    escape(v, true);
    expect(sides(v)).toEqual(WHOLE);
    expect(escape(v)).toBe(true);
    expect(sides(v)).toEqual([]);
  });

  test('a refused paste keeps it', async () => {
    const v = await mount(BOOK);
    const { token } = copied(v, 'cut');
    v.setMode('view');
    expect(v.paste({ token }).ok).toBe(false);
    expect(sides(v)).toEqual(WHOLE);
  });
});

describe('rangeOutline', () => {
  const axis = {
    offsetOf: (i: number) => i * 20,
    sizeOf: () => 20,
    indexAt: (at: number) => Math.floor(at / 20),
  };
  const view = (over: Partial<ViewState>): ViewState =>
    ({
      geometry: { columns: axis, rows: axis },
      frozen: { rows: 0, cols: 0, firstRow: 0, firstCol: 0 },
      scrollX: 0,
      scrollY: 0,
      width: 200,
      height: 200,
      ...over,
    }) as unknown as ViewState;
  const range = (r0: number, c0: number, r1: number, c1: number) => ({
    start: { row: r0, col: c0 },
    end: { row: r1, col: c1 },
  });
  const named = (edges: ReturnType<typeof rangeOutline>) =>
    edges.map((e) => `${e.side} ${e.x},${e.y} ${e.length}`).sort();

  test('gives the four edges of a range in the body', () => {
    expect(named(rangeOutline(view({}), range(1, 1, 2, 2)))).toEqual([
      'bottom 20,60 40',
      'left 20,20 40',
      'right 60,20 40',
      'top 20,20 40',
    ]);
  });

  test('follows the scroll', () => {
    expect(named(rangeOutline(view({ scrollX: 10, scrollY: 5 }), range(1, 1, 1, 1)))).toEqual([
      'bottom 10,35 20',
      'left 10,15 20',
      'right 30,15 20',
      'top 10,15 20',
    ]);
  });

  test('leaves open an edge that the window cuts', () => {
    // Scrolled past the top and the left of the range, and past the end of the window.
    expect(named(rangeOutline(view({ scrollX: 30, scrollY: 30 }), range(1, 1, 20, 20)))).toEqual(
      []
    );
    const sides = rangeOutline(view({ scrollY: 30 }), range(1, 1, 3, 2)).map((e) => e.side);
    expect(sides).not.toContain('top');
    expect(sides).toContain('bottom');
  });

  test('draws no line on the seam of a frozen pane', () => {
    const frozen: ViewState['frozen'] = { rows: 2, cols: 0, firstRow: 0, firstCol: 0 };
    const edges = named(rangeOutline(view({ frozen }), range(1, 0, 3, 0)));
    expect(edges).toEqual([
      'bottom 0,80 20',
      'left 0,20 20',
      'left 0,40 40',
      'right 20,20 20',
      'right 20,40 40',
      'top 0,20 20',
    ]);
  });

  test('keeps the frozen part still, and hides what scrolls under the pane', () => {
    const frozen: ViewState['frozen'] = { rows: 2, cols: 0, firstRow: 0, firstCol: 0 };
    const edges = rangeOutline(view({ frozen, scrollY: 200 }), range(1, 0, 3, 0));
    expect(named(edges)).toEqual(['left 0,20 20', 'right 20,20 20', 'top 0,20 20']);
  });

  test('gives nothing for a range out of the window', () => {
    expect(rangeOutline(view({}), range(50, 50, 51, 51))).toEqual([]);
  });
});

describe('a copy whose formulas can refer to other cells after an edit', () => {
  const FORMULA: FixtureWorkbook = {
    sheets: [
      { name: 'S', rows: row(5, [c('A5', 7), c('B5', 7, { f: '$A$5' })]) },
      { name: 'Data', rows: row(1, [c('A1', 1)]) },
      { name: 'Other', rows: row(1, [c('A1', 2)]) },
    ],
  };
  const NAMED: FixtureWorkbook = {
    sheets: [
      { name: 'S', rows: row(1, [c('A1', 1, { f: 'Data!A1' })]) },
      { name: 'Data', rows: row(1, [c('A1', 1)]) },
      { name: 'Other', rows: row(1, [c('A1', 2)]) },
    ],
  };

  function copyAt(v: XlsxEditorInstance, r: number, col: number, mode: 'copy' | 'cut' = 'copy') {
    v.select({ row: r, col });
    const out = mode === 'cut' ? v.cut() : v.copy();
    if (!out.ok) throw new Error(out.refusal.code);
    return out.content.token;
  }

  /** Paste at `D20`, and state the refusal code. The cell must stay empty after a refusal. */
  function refusalAtD20(v: XlsxEditorInstance, token: string): string | null {
    v.select({ row: 19, col: 3 });
    const r = v.paste({ token, text: 'plain' });
    return r.ok ? null : r.refusal.code;
  }

  const d20 = (v: XlsxEditorInstance) => {
    v.select({ row: 19, col: 3 });
    return v.snapshot().activeCell?.content ?? '';
  };

  const insertAbove = (v: XlsxEditorInstance) => {
    v.select({ row: 0, col: 0 });
    expect(v.insertRows().ok).toBe(true);
  };

  test('a row insert ends the copy: the paste writes nothing, also the second time', async () => {
    const v = await mount(FORMULA);
    const token = copyAt(v, 4, 1);
    expect(sides(v)).toEqual(WHOLE);
    insertAbove(v);
    expect(sides(v)).toEqual([]);
    expect(refusalAtD20(v, token)).toBe('copy-expired');
    expect(d20(v)).toBe('');
    expect(refusalAtD20(v, token)).toBe('copy-expired');
    expect(d20(v)).toBe('');
  });

  test('a column insert and a delete of rows or columns end the copy', async () => {
    for (const edit of ['insertColumns', 'deleteRows', 'deleteColumns'] as const) {
      const v = await mount(FORMULA);
      const token = copyAt(v, 4, 1);
      v.select({ row: 0, col: 0 });
      expect(v[edit]().ok).toBe(true);
      expect(refusalAtD20(v, token)).toBe('copy-expired');
      expect(d20(v)).toBe('');
    }
  });

  test('an undo of a row insert that came before the copy ends the copy', async () => {
    const v = await mount(FORMULA);
    insertAbove(v);
    const token = copyAt(v, 5, 1);
    expect(v.undo().ok).toBe(true);
    expect(sides(v)).toEqual([]);
    expect(refusalAtD20(v, token)).toBe('copy-expired');
    expect(d20(v)).toBe('');
  });

  test('a redo of a row insert ends the copy', async () => {
    const v = await mount(FORMULA);
    insertAbove(v);
    expect(v.undo().ok).toBe(true);
    const token = copyAt(v, 4, 1);
    expect(v.redo().ok).toBe(true);
    expect(sides(v)).toEqual([]);
    expect(refusalAtD20(v, token)).toBe('copy-expired');
    expect(d20(v)).toBe('');
  });

  test('an insert and its undo do not bring the copy back', async () => {
    const v = await mount(FORMULA);
    const token = copyAt(v, 4, 1);
    insertAbove(v);
    expect(v.undo().ok).toBe(true);
    expect(sides(v)).toEqual([]);
    expect(refusalAtD20(v, token)).toBe('copy-expired');
    expect(d20(v)).toBe('');
  });

  test('a row insert ends a pending cut, and the source stays', async () => {
    const v = await mount(FORMULA);
    const token = copyAt(v, 4, 1, 'cut');
    insertAbove(v);
    expect(sides(v)).toEqual([]);
    expect(refusalAtD20(v, token)).toBe('cut-expired');
    expect(d20(v)).toBe('');
    v.select({ row: 5, col: 1 });
    expect(v.snapshot().activeCell?.content).toBe('=$A$6');
  });

  test('a rename, a move, and a delete of a named sheet end the copy', async () => {
    const edits: ((v: XlsxEditorInstance) => { ok: boolean })[] = [
      (v) => v.renameSheet(1, 'Renamed'),
      (v) => v.moveSheet(1, 2),
      (v) => v.deleteSheet(1),
    ];
    for (const edit of edits) {
      const v = await mount(NAMED);
      const token = copyAt(v, 0, 0);
      expect(edit(v).ok).toBe(true);
      expect(v.setActiveSheet(0)).toBe(true);
      expect(refusalAtD20(v, token)).toBe('copy-expired');
      expect(d20(v)).toBe('');
    }
  });

  test('a rename of a sheet that no copied formula names keeps the copy', async () => {
    const v = await mount(NAMED);
    const token = copyAt(v, 0, 0);
    expect(v.renameSheet(2, 'Renamed').ok).toBe(true);
    expect(v.setActiveSheet(0)).toBe(true);
    expect(refusalAtD20(v, token)).toBeNull();
    expect(d20(v)).toBe('=Data!D20');
  });

  describe('with a reference across a span of sheets', () => {
    const SPAN: FixtureWorkbook = {
      sheets: [
        { name: 'S', rows: row(1, [c('A1', 3, { f: 'SUM(Jan:Mar!A1)' })]) },
        { name: 'Jan', rows: row(1, [c('A1', 1)]) },
        { name: 'Feb', rows: row(1, [c('A1', 1)]) },
        { name: 'Mar', rows: row(1, [c('A1', 1)]) },
        { name: 'Other', rows: row(1, [c('A1', 2)]) },
      ],
    };

    test('a delete or a rename inside the span and any sheet move end the copy', async () => {
      const edits: ((v: XlsxEditorInstance) => { ok: boolean })[] = [
        (v) => v.moveSheet(4, 0),
        (v) => v.deleteSheet(2),
        (v) => v.renameSheet(2, 'Renamed'),
      ];
      for (const edit of edits) {
        const v = await mount(SPAN);
        const token = copyAt(v, 0, 0);
        expect(edit(v).ok).toBe(true);
        expect(v.setActiveSheet(0)).toBe(true);
        expect(sides(v)).toEqual([]);
        expect(refusalAtD20(v, token)).toBe('copy-expired');
        expect(d20(v)).toBe('');
      }
    });

    test('an undo of the delete does not bring the copy back', async () => {
      const v = await mount(SPAN);
      const token = copyAt(v, 0, 0);
      expect(v.deleteSheet(2).ok).toBe(true);
      expect(v.undo().ok).toBe(true);
      expect(v.setActiveSheet(0)).toBe(true);
      expect(refusalAtD20(v, token)).toBe('copy-expired');
    });

    test('a span in another workbook does not end the copy', async () => {
      const v = await mount({
        sheets: [
          { name: 'S', rows: row(1, [c('A1', 0, { f: 'SUM([1]Jan:Mar!A1)' })]) },
          { name: 'Other', rows: '' },
          { name: 'More', rows: '' },
        ],
        externalReferences: '<externalReference r:id="rIdX1"/>',
        extraParts: [
          [
            '/xl/externalLinks/externalLink1.xml',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml',
            new TextEncoder().encode(
              '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>'
            ),
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink',
          ],
        ],
      });
      const token = copyAt(v, 0, 0);
      expect(v.moveSheet(1, 2).ok).toBe(true);
      expect(v.setActiveSheet(0)).toBe(true);
      expect(refusalAtD20(v, token)).toBeNull();
    });
  });

  test('a row insert keeps a copy of values', async () => {
    const v = await mount(FORMULA);
    const token = copyAt(v, 4, 0);
    insertAbove(v);
    expect(refusalAtD20(v, token)).toBeNull();
    expect(d20(v)).toBe('7');
  });

  test('a new copy after the refusal pastes', async () => {
    const v = await mount(FORMULA);
    const stale = copyAt(v, 4, 1);
    insertAbove(v);
    expect(refusalAtD20(v, stale)).toBe('copy-expired');
    const token = copyAt(v, 5, 1);
    expect(refusalAtD20(v, token)).toBeNull();
    expect(d20(v)).toBe('=$A$6');
  });
});

describe('a paste of a copy at its own source', () => {
  const pasteAtSource = (v: XlsxEditorInstance) => {
    const { token } = copied(v);
    expect(sides(v)).toEqual(WHOLE);
    v.select({ row: 0, col: 0 });
    expect(v.paste({ token }).ok).toBe(true);
    return sides(v);
  };

  test('keeps the outline when the cells do not change', async () => {
    expect(pasteAtSource(await mount(BOOK))).toEqual(WHOLE);
  });

  test('ends the outline when a rich text cell becomes plain text', async () => {
    const v = await mount({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 0, { t: 's' })]) }],
      rawStrings: ['<si><r><rPr><b/></rPr><t>Bold</t></r></si>'],
    });
    expect(pasteAtSource(v)).toEqual([]);
  });

  test('ends the outline when an error value becomes text', async () => {
    const v = await mount({
      sheets: [{ name: 'S', rows: row(1, [c('A1', '#N/A', { t: 'e' })]) }],
    });
    expect(pasteAtSource(v)).toEqual([]);
  });

  test('ends the outline when a formula loses its cached value', async () => {
    const v = await mount({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 3, { f: '1+2' })]) }],
    });
    expect(pasteAtSource(v)).toEqual([]);
  });
});

describe('a held Escape', () => {
  test('does not end the outline of a pending copy', async () => {
    const v = await mount(BOOK);
    copied(v);
    escape(v, false, true);
    expect(sides(v)).toEqual(WHOLE);
    expect(escape(v)).toBe(true);
    expect(sides(v)).toEqual([]);
  });
});
