// A selection of several ranges (`x:selection/@sqref`, `ST_Sqref`, ECMA-376 Part 1, 18.18.76; the
// `selection` element, Part 1, 18.3.1.78): the open reads `@sqref` and `@activeCellId` with caps,
// a press with the platform modifier adds a range and its drag extends it, Shift extends the
// range of the active cell, Enter and Tab go through every range, the paint marks every range,
// a clear and a format patch apply to every range in one undo step, and the commands that act on
// one range refuse with `multiple-ranges`.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx, c, row, type FixtureSheet } from '../../../../../scripts/fixtures/builder.ts';
import {
  MAX_ROWS,
  MAX_SELECTION_RANGES,
  createXlsxEditor,
  mergeCommandState,
  toolbarCommandState,
  type XlsxEditorInstance,
} from '../index.ts';
import { flushEditor } from './flush-editor.ts';
import { formulaSlice } from '../chrome-status.ts';
import { readSqref, unionCover } from '../selection-ranges.ts';
import { historyGridSelection, openingSelection } from '../selection.ts';
import { cycleSelection } from '../grid-keys.ts';
import {
  WorkbookDocumentStore,
  attr,
  child,
  children,
  readXlsx,
  workbookIndex,
  xmlPartOf,
} from '../../store/index.ts';
import { mergeOps, structureOp, unmergeOp } from '../../automation/index.ts';
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

const APPLE_AGENT = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36';

async function mount(sheet: Partial<FixtureSheet> = {}, apple = false) {
  const win = apple
    ? new Window({ settings: { navigator: { userAgent: APPLE_AGENT } } })
    : new Window();
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
  Object.defineProperty(scroller, 'clientWidth', { get: () => 800 });
  Object.defineProperty(scroller, 'clientHeight', { get: () => 400 });
  await viewer.load(
    buildXlsx({
      sheets: [
        {
          name: 'S',
          headings: true,
          rows: row(1, [c('A1', 1), c('B1', 2)]) + row(5, [c('E5', 5)]),
          ...sheet,
        },
      ],
    })
  );
  flushEditor(viewer);
  return { viewer, host, scroller };
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

/** The client point of the middle of a cell, read from the painted headings. */
function cellPoint(host: HTMLElement, rowIndex: number, col: number) {
  const colHead = host.querySelector(`.xlsx-column-heads [data-col="${col}"]`) as HTMLElement;
  const rowHead = host.querySelector(`.xlsx-row-heads [data-row="${rowIndex}"]`) as HTMLElement;
  const heads = host.querySelector('.xlsx-column-heads') as HTMLElement;
  const rows = host.querySelector('.xlsx-row-heads') as HTMLElement;
  const n = (v: string) => Number.parseFloat(v);
  return {
    x: n(heads.style.left) + n(colHead.style.left) + n(colHead.style.width) / 2,
    y: n(rows.style.top) + n(rowHead.style.top) + n(rowHead.style.height) / 2,
  };
}

/** Pointer events on the stage, as one mouse pointer. */
function pointer(host: HTMLElement) {
  const stage = host.querySelector('.xlsx-stage') as HTMLElement;
  const W = host.ownerDocument.defaultView as unknown as typeof globalThis;
  return (
    type: 'pointerdown' | 'pointermove' | 'pointerup',
    at: { x: number; y: number },
    mods: { ctrlKey?: boolean; metaKey?: boolean; shiftKey?: boolean } = {}
  ) =>
    stage.dispatchEvent(
      new W.PointerEvent(type, {
        bubbles: true,
        button: 0,
        pointerId: 1,
        pointerType: 'mouse',
        clientX: at.x,
        clientY: at.y,
        ...mods,
      })
    );
}

/** Each stored cell of the saved sheet as `ref=s`, `s` being its cell format index. */
function savedStyles(viewer: XlsxEditorInstance): Record<string, string> {
  const r = readXlsx(viewer.save()!);
  if (!r.ok) throw new Error('reopen');
  const index = workbookIndex(r.pkg)!;
  const root = xmlPartOf(r.pkg, index.sheets[0]!.partName)!.root;
  const out: Record<string, string> = {};
  for (const rowEl of children(child(root, 'x:sheetData'), 'x:row'))
    for (const cell of children(rowEl, 'x:c')) out[attr(cell, 'r')!] = attr(cell, 's') ?? '0';
  return out;
}

const range = (r0: number, c0: number, r1 = r0, c1 = c0) => ({
  start: { row: r0, col: c0 },
  end: { row: r1, col: c1 },
});

describe('the open reads x:selection/@sqref and @activeCellId', () => {
  test('several ranges open selected, with the active cell in its range', async () => {
    const { viewer } = await mount({
      viewChildren: '<selection activeCell="E5" activeCellId="1" sqref="B2:C4 E5:F8 H2"/>',
    });
    const snap = viewer.snapshot();
    expect(snap.selectionRef).toBe('B2:C4 E5:F8 H2');
    expect(snap.activeCell?.ref).toBe('E5');
    expect(snap.selection?.activeRange).toBe(1);
    expect(snap.selection!.ranges).toHaveLength(3);
  });

  test('the name box shows the active cell of several ranges, and the range of one', async () => {
    const { viewer } = await mount({
      viewChildren: '<selection activeCell="E5" activeCellId="1" sqref="B2:C4 E5:F8"/>',
    });
    expect(formulaSlice(viewer.snapshot()).ref).toBe('E5');
    viewer.select({ row: 1, col: 1 }, { row: 2, col: 2 });
    expect(formulaSlice(viewer.snapshot()).ref).toBe('B2:C3');
  });

  test('with no @activeCell, the first corner of the range of @activeCellId is active', async () => {
    const one = await mount({ viewChildren: '<selection sqref="C3:D4"/>' });
    expect(one.viewer.snapshot().selectionRef).toBe('C3:D4');
    expect(one.viewer.snapshot().activeCell?.ref).toBe('C3');
    const two = await mount({ viewChildren: '<selection activeCellId="1" sqref="B2 E5:F8"/>' });
    expect(two.viewer.snapshot().selectionRef).toBe('B2 E5:F8');
    expect(two.viewer.snapshot().activeCell?.ref).toBe('E5');
  });

  test('an @activeCellId out of range leaves @activeCell to find the range', async () => {
    const { viewer } = await mount({
      viewChildren: '<selection activeCell="H2" activeCellId="9" sqref="B2:C4 E5:F8 H2"/>',
    });
    expect(viewer.snapshot().selection?.activeRange).toBe(2);
    expect(viewer.snapshot().activeCell?.ref).toBe('H2');
  });

  test('an @activeCellId past a range that does not read still finds the active range', async () => {
    // `A:A` does not read as a cell range, so the ranges after it move down one place.
    const { viewer } = await mount({
      viewChildren: '<selection activeCell="B2" activeCellId="1" sqref="A:A B2 H2"/>',
    });
    expect(viewer.snapshot().activeCell?.ref).toBe('B2');
    expect(viewer.snapshot().selectionRef).toBe('B2 H2');
  });

  test('a frozen sheet opens with the ranges of the selection of its active pane', async () => {
    const { viewer } = await mount({
      viewChildren:
        '<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>' +
        '<selection pane="topRight" activeCell="B1" sqref="B1"/>' +
        '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>' +
        '<selection pane="bottomRight" activeCell="F6" activeCellId="1" sqref="C3:D4 F6"/>',
    });
    expect(viewer.snapshot().selectionRef).toBe('C3:D4 F6');
    expect(viewer.snapshot().activeCell?.ref).toBe('F6');
  });

  test('an active cell that no range holds selects that cell alone', async () => {
    const { viewer } = await mount({
      viewChildren: '<selection activeCell="Z9" sqref="B2:C4 E5:F8"/>',
    });
    expect(viewer.snapshot().selectionRef).toBe('Z9');
  });

  test('one range opens as that range', async () => {
    const { viewer } = await mount({
      viewChildren: '<selection activeCell="C3" sqref="B2:D5"/>',
    });
    expect(viewer.snapshot().selectionRef).toBe('B2:D5');
    expect(viewer.snapshot().selection?.ranges).toBeUndefined();
    expect(viewer.snapshot().activeCell?.ref).toBe('C3');
  });

  test('readSqref caps the ranges and skips text that does not read', () => {
    const many = Array.from({ length: 5000 }, (_, i) => `A${i + 1}`).join(' ');
    expect(readSqref(many)).toHaveLength(MAX_SELECTION_RANGES);
    expect(readSqref('B2 nonsense A1:B2:C3 $C$3 D4:E5')).toEqual([range(1, 1), range(3, 3, 4, 4)]);
    expect(readSqref(' '.repeat(100_000) + 'A1')).toEqual([]);
    expect(readSqref(undefined)).toEqual([]);
    // A value cut in the middle of a reference leaves that reference out.
    // The cut at 16,384 characters falls inside `A1:B23`, which would read as `A1:B2`.
    expect(readSqref(`${' '.repeat(16_379)}A1:B23`)).toEqual([]);
    // Exactly 16,384 characters: nothing is cut, and the reference reads whole.
    expect(readSqref(`${' '.repeat(16_378)}A1:B23`)).toEqual([range(0, 0, 22, 1)]);
  });

  test('a range that meets a merged range grows to cover it', () => {
    const merges = [{ firstRow: 3, lastRow: 4, firstCol: 0, lastCol: 1 }];
    const sel = openingSelection(
      0,
      { active: { row: 0, col: 0 }, ranges: [range(0, 0), range(4, 1)], activeRange: undefined },
      merges
    );
    expect(sel.ranges).toEqual([range(0, 0), range(3, 0, 4, 1)]);
    expect(sel.activeRange).toBe(0);
  });

  test('many merged ranges read fewer ranges, and keep the active range', () => {
    const merges = Array.from({ length: 500_000 }, (_, i) => ({
      firstRow: 2000 + i,
      lastRow: 2000 + i,
      firstCol: 0,
      lastCol: 1,
    }));
    const ranges = Array.from({ length: 10 }, (_, i) => range(i, 5));
    const sel = openingSelection(0, { active: { row: 9, col: 5 }, ranges, activeRange: 9 }, merges);
    expect(sel.ranges).toBeUndefined();
    expect(sel.range).toEqual(range(9, 5));
    expect(sel.active).toEqual({ row: 9, col: 5 });
  });
});

describe('the pointer adds ranges with the platform modifier', () => {
  test('Control and a press adds a cell; its drag extends the new range', async () => {
    const { viewer, host } = await mount();
    const p = pointer(host);
    p('pointerdown', cellPoint(host, 1, 1));
    p('pointerup', cellPoint(host, 1, 1));
    p('pointerdown', cellPoint(host, 4, 4), { ctrlKey: true });
    p('pointermove', cellPoint(host, 6, 5), { ctrlKey: true });
    p('pointerup', cellPoint(host, 6, 5), { ctrlKey: true });
    expect(viewer.snapshot().selectionRef).toBe('B2 E5:F7');
    expect(viewer.snapshot().activeCell?.ref).toBe('E5');
    // Shift extends the range of the active cell, and keeps the other range.
    p('pointerdown', cellPoint(host, 7, 6), { shiftKey: true });
    p('pointerup', cellPoint(host, 7, 6), { shiftKey: true });
    expect(viewer.snapshot().selectionRef).toBe('B2 E5:G8');
    // A plain press starts one range again.
    p('pointerdown', cellPoint(host, 2, 2));
    expect(viewer.snapshot().selectionRef).toBe('C3');
  });

  test('a press past MAX_SELECTION_RANGES is refused with too-many-ranges, and adds nothing', async () => {
    const { viewer, host } = await mount();
    const full = Array.from({ length: MAX_SELECTION_RANGES }, (_, i) => ({
      anchor: { row: 100 + i, col: 0 },
    }));
    expect(viewer.selectRanges(full)).toBe(true);
    const before = viewer.snapshot().selectionRef;
    const p = pointer(host);
    p('pointerdown', cellPoint(host, 1, 1), { ctrlKey: true });
    p('pointerup', cellPoint(host, 1, 1), { ctrlKey: true });
    expect(viewer.snapshot().selectionRef).toBe(before);
    expect(viewer.snapshot().refusal?.code).toBe('too-many-ranges');
  });

  test('Control and a press on a heading adds the band', async () => {
    const { viewer, host } = await mount();
    const p = pointer(host);
    p('pointerdown', cellPoint(host, 0, 0));
    p('pointerup', cellPoint(host, 0, 0));
    const colHead = cellPoint(host, 0, 2);
    p('pointerdown', { x: colHead.x, y: 4 }, { ctrlKey: true });
    p('pointerup', { x: colHead.x, y: 4 }, { ctrlKey: true });
    expect(viewer.snapshot().selectionRef).toBe('A1 C1:C1048576');
    flushEditor(viewer);
    const heads = [...host.querySelectorAll('.xlsx-column-heads .xlsx-head-selected')];
    expect(heads.map((h) => (h as HTMLElement).dataset.col)).toEqual(['0', '2']);
  });

  test('on an Apple platform Command adds a range, and Control does not', async () => {
    const { viewer, host } = await mount({}, true);
    const p = pointer(host);
    p('pointerdown', cellPoint(host, 1, 1));
    p('pointerup', cellPoint(host, 1, 1));
    p('pointerdown', cellPoint(host, 5, 5), { metaKey: true });
    p('pointerup', cellPoint(host, 5, 5), { metaKey: true });
    expect(viewer.snapshot().selectionRef).toBe('B2 F6');
    // Control and a press is the context-menu gesture there: with a menu, it selects nothing.
    const off = viewer.on('contextMenu', () => {});
    p('pointerdown', cellPoint(host, 3, 3), { ctrlKey: true });
    p('pointerup', cellPoint(host, 3, 3), { ctrlKey: true });
    expect(viewer.snapshot().selectionRef).toBe('B2 F6');
    // With no menu to open, the press selects the cell, and adds no range.
    off();
    p('pointerdown', cellPoint(host, 3, 3), { ctrlKey: true });
    p('pointerup', cellPoint(host, 3, 3), { ctrlKey: true });
    expect(viewer.snapshot().selectionRef).toBe('D4');
  });

  test('a right press inside any range keeps every range, and outside selects the cell', async () => {
    const { viewer, host } = await mount();
    // With a menu handler, the engine selects before it asks for the menu.
    viewer.on('contextMenu', () => {});
    viewer.selectRanges([
      { anchor: { row: 0, col: 0 }, focus: { row: 1, col: 1 } },
      { anchor: { row: 4, col: 4 } },
    ]);
    const stage = host.querySelector('.xlsx-stage') as HTMLElement;
    const W = host.ownerDocument.defaultView as unknown as typeof globalThis;
    const rightPress = (at: { x: number; y: number }) =>
      stage.dispatchEvent(
        new W.MouseEvent('contextmenu', {
          bubbles: true,
          cancelable: true,
          button: 2,
          clientX: at.x,
          clientY: at.y,
        })
      );
    // B2 is in the range that is not the active one.
    rightPress(cellPoint(host, 1, 1));
    expect(viewer.snapshot().selectionRef).toBe('A1:B2 E5');
    rightPress(cellPoint(host, 7, 7));
    expect(viewer.snapshot().selectionRef).toBe('H8');
  });
});

describe('keys on a selection of several ranges', () => {
  test('Enter and Tab go through every range in order, and wrap', async () => {
    const { viewer, scroller } = await mount();
    viewer.selectRanges([
      { anchor: { row: 1, col: 1 }, focus: { row: 2, col: 1 } },
      { anchor: { row: 4, col: 4 } },
    ]);
    expect(viewer.snapshot().activeCell?.ref).toBe('E5');
    key(scroller, 'Enter');
    expect(viewer.snapshot().activeCell?.ref).toBe('B2');
    key(scroller, 'Enter');
    expect(viewer.snapshot().activeCell?.ref).toBe('B3');
    key(scroller, 'Tab');
    expect(viewer.snapshot().activeCell?.ref).toBe('E5');
    key(scroller, 'Enter', { shiftKey: true });
    expect(viewer.snapshot().activeCell?.ref).toBe('B3');
    expect(viewer.snapshot().selectionRef).toBe('B2:B3 E5');
  });

  test('an arrow key collapses the selection to one range', async () => {
    const { viewer, scroller } = await mount();
    viewer.selectRanges([{ anchor: { row: 1, col: 1 } }, { anchor: { row: 4, col: 4 } }]);
    key(scroller, 'ArrowDown');
    expect(viewer.snapshot().selectionRef).toBe('E6');
  });

  test('a committed edit with Enter moves to the next range', async () => {
    const { viewer, scroller } = await mount();
    viewer.selectRanges([{ anchor: { row: 1, col: 1 } }, { anchor: { row: 4, col: 4 } }]);
    key(scroller, '7');
    expect(viewer.commitEdit('down').ok).toBe(true);
    expect(viewer.snapshot().activeCell?.ref).toBe('B2');
    expect(viewer.snapshot().selectionRef).toBe('B2 E5');
  });

  test('cycleSelection skips a hidden cell of the next range', () => {
    const hidden = { row: (i: number) => i === 4, col: () => false };
    const sel = {
      sheetIndex: 0,
      active: { row: 0, col: 0 },
      anchor: { row: 0, col: 0 },
      range: range(0, 0),
      ranges: [range(0, 0), range(4, 1, 5, 1)],
      activeRange: 0,
    };
    const next = cycleSelection(sel, sel.active, false, 1, hidden, []);
    expect(next.active).toEqual({ row: 5, col: 1 });
    expect(next.activeRange).toBe(1);
  });
});

describe('the paint marks every range, clipped to the window', () => {
  test('each range in the window gets the tint and the outline; one outside draws nothing', async () => {
    const { viewer, host } = await mount();
    // The last range holds the active cell, so the view stays at the top left.
    viewer.selectRanges([
      { anchor: { row: 50_000, col: 300 } },
      { anchor: { row: 1, col: 1 }, focus: { row: 2, col: 2 } },
      { anchor: { row: 4, col: 4 } },
    ]);
    flushEditor(viewer);
    expect(host.querySelectorAll('.xlsx-selection-range')).toHaveLength(2);
    expect(host.querySelectorAll('.xlsx-active-cell')).toHaveLength(1);
    const rows = [...host.querySelectorAll('.xlsx-row-heads .xlsx-head-selected')];
    expect(rows.map((h) => (h as HTMLElement).dataset.row)).toEqual(['1', '2', '4']);
  });
});

describe('commands on a selection of several ranges', () => {
  test('a fill applies to every range in one undo step', async () => {
    const { viewer } = await mount();
    viewer.selectRanges([{ anchor: { row: 0, col: 0 } }, { anchor: { row: 4, col: 4 } }]);
    const before = viewer.snapshot().revision;
    expect(viewer.applyFormat({ fill: { type: 'solid', color: 'FFFF00' } }).ok).toBe(true);
    expect(viewer.snapshot().revision).toBe(before + 1);
    const styles = savedStyles(viewer);
    expect(styles.A1).not.toBe('0');
    expect(styles.E5).toBe(styles.A1);
    expect(styles.B1).toBe('0');
    expect(viewer.undo().ok).toBe(true);
    const undone = savedStyles(viewer);
    expect([undone.A1, undone.E5]).toEqual(['0', '0']);
    // The undo selects both ranges again.
    expect(viewer.snapshot().selectionRef).toBe('A1 E5');
    expect(viewer.snapshot().activeCell?.ref).toBe('E5');
  });

  test('Bold toggles every range from the active cell', async () => {
    const { viewer } = await mount();
    viewer.selectRanges([{ anchor: { row: 0, col: 0 } }, { anchor: { row: 4, col: 4 } }]);
    expect(viewer.toggleFormat('bold').ok).toBe(true);
    const styles = savedStyles(viewer);
    expect(styles.E5).toBe(styles.A1);
    expect(styles.A1).not.toBe('0');
  });

  test('Delete clears every range in one undo step', async () => {
    const { viewer, scroller } = await mount();
    viewer.selectRanges([{ anchor: { row: 0, col: 0 } }, { anchor: { row: 4, col: 4 } }]);
    key(scroller, 'Delete');
    viewer.select({ row: 0, col: 0 });
    expect(viewer.snapshot().activeCell?.content).toBe('');
    viewer.select({ row: 4, col: 4 });
    expect(viewer.snapshot().activeCell?.content).toBe('');
    viewer.select({ row: 0, col: 1 });
    expect(viewer.snapshot().activeCell?.content).toBe('2');
    expect(viewer.undo().ok).toBe(true);
    viewer.select({ row: 4, col: 4 });
    expect(viewer.snapshot().activeCell?.content).toBe('5');
    viewer.select({ row: 0, col: 0 });
    expect(viewer.snapshot().activeCell?.content).toBe('1');
  });

  test('the commands that act on one range refuse, and their controls are disabled', async () => {
    const { viewer } = await mount();
    viewer.selectRanges([{ anchor: { row: 0, col: 0 } }, { anchor: { row: 4, col: 4 } }]);
    const revision = viewer.snapshot().revision;
    for (const run of [
      () => viewer.mergeCells(),
      () => viewer.unmergeCells(),
      () => viewer.insertRows(),
      () => viewer.deleteColumns(),
      () => viewer.copy(),
      () => viewer.cut(),
      () => viewer.paste({ text: 'x' }),
    ]) {
      const r = run();
      expect(r.ok).toBe(false);
      if (!r.ok) expect(r.refusal.code).toBe('multiple-ranges');
    }
    expect(viewer.snapshot().revision).toBe(revision);
    const snap = viewer.snapshot();
    for (const id of [
      'cells.insertRows',
      'cells.deleteColumns',
      'number.increaseDecimals',
      'merge.merge',
      'edit.copy',
      'edit.cut',
      'edit.paste',
    ] as const) {
      const state = toolbarCommandState(snap, id, { readClipboard: async () => ({ text: '' }) });
      expect(state.enabled).toBe(false);
      expect(state.disabledReason?.key).toBe('disabledReason.multipleRanges');
    }
    expect(mergeCommandState(snap, 'unmerge').disabledReason?.key).toBe(
      'disabledReason.multipleRanges'
    );
    expect(toolbarCommandState(snap, 'fill.color').enabled).toBe(true);
    expect(toolbarCommandState(snap, 'clear.contents').enabled).toBe(true);
  });

  test('the live region names every range', async () => {
    const { viewer, host } = await mount();
    viewer.selectRanges([
      { anchor: { row: 1, col: 1 }, focus: { row: 2, col: 2 } },
      { anchor: { row: 4, col: 4 } },
    ]);
    flushEditor(viewer);
    const live = host.querySelector('.xlsx-live') as HTMLElement;
    expect(live.textContent).toContain('grid.selectedRanges{"count":"2","ranges":"B2:C3, E5"}');
  });

  test('selectRanges refuses an empty list, too many ranges, and a coordinate that is not finite', async () => {
    const { viewer } = await mount();
    expect(viewer.selectRanges([])).toBe(false);
    const many = Array.from({ length: MAX_SELECTION_RANGES + 1 }, (_, i) => ({
      anchor: { row: i, col: 0 },
    }));
    expect(viewer.selectRanges(many)).toBe(false);
    expect(
      viewer.selectRanges([{ anchor: { row: 0, col: 0 } }, { anchor: { row: NaN, col: 0 } }])
    ).toBe(false);
    expect(viewer.snapshot().selectionRef).toBe('A1');
  });

  test('selectReference selects several space-delimited references, or refuses them all', async () => {
    const { viewer } = await mount();
    expect(viewer.selectReference(' b2:c4  $E$5 ').ok).toBe(true);
    expect(viewer.snapshot().selectionRef).toBe('B2:C4 E5');
    expect(viewer.snapshot().activeCell?.ref).toBe('E5');
    // Whole columns and whole rows read as one reference does, as a heading press selects them.
    expect(viewer.selectReference('a:a C:C').ok).toBe(true);
    expect(viewer.snapshot().selection?.ranges).toHaveLength(2);
    expect(viewer.snapshot().selection?.ranges?.[1]?.start).toEqual({ row: 0, col: 2 });
    expect(viewer.selectReference('B2:C4 E5').ok).toBe(true);
    // One reference that does not read refuses the text, and the selection stays.
    const bad = viewer.selectReference('B2 R1C1');
    expect(bad.ok ? null : bad.refusal.code).toBe('go-to-reference');
    expect(viewer.snapshot().selectionRef).toBe('B2:C4 E5');
    // More references than a selection holds: 65 short ones, inside MAX_GO_TO_LENGTH.
    const capped = Array.from({ length: MAX_SELECTION_RANGES + 1 }, () => 'A1').join(' ');
    const over = viewer.selectReference(capped);
    expect(over.ok ? null : over.refusal.code).toBe('too-many-ranges');
    // One reference that does not read among them refuses the text, not the count.
    const mixed = Array.from({ length: MAX_SELECTION_RANGES + 1 }, (_, i) =>
      i === 2 ? 'R1C1' : 'A1'
    );
    const unread = viewer.selectReference(mixed.join(' '));
    expect(unread.ok ? null : unread.refusal.code).toBe('go-to-reference');
  });

  test('an extension that passes a range equal to another range keeps both', async () => {
    const { viewer, host } = await mount();
    viewer.selectRanges([
      { anchor: { row: 0, col: 0 }, focus: { row: 1, col: 1 } },
      { anchor: { row: 0, col: 0 } },
    ]);
    const grid = host.querySelector('.xlsx-scroller') as HTMLElement;
    const W = host.ownerDocument.defaultView as unknown as typeof globalThis;
    for (const key of ['ArrowRight', 'ArrowDown'])
      grid.dispatchEvent(new W.KeyboardEvent('keydown', { key, shiftKey: true, bubbles: true }));
    expect(viewer.snapshot().selectionRef).toBe('A1:B2 A1:B2');
    // The next step of the same gesture gives the first range back as it was.
    grid.dispatchEvent(
      new W.KeyboardEvent('keydown', { key: 'ArrowDown', shiftKey: true, bubbles: true })
    );
    expect(viewer.snapshot().selectionRef).toBe('A1:B2 A1:B3');
  });

  test('the lowering helpers that act on one range throw for several ranges', () => {
    const sel = {
      sheetIndex: 0,
      active: { row: 0, col: 0 },
      anchor: { row: 0, col: 0 },
      range: range(0, 0, 1, 1),
      ranges: [range(0, 0, 1, 1), range(4, 4)],
      activeRange: 0,
    };
    expect(() => mergeOps('S', sel)).toThrow(RangeError);
    expect(() => unmergeOp('S', sel)).toThrow(RangeError);
    expect(() => structureOp('S', sel, 'insertRows')).toThrow(RangeError);
  });
});

describe('the format limit of several ranges', () => {
  test('counts the union of the ranges, so overlapping ranges count once', () => {
    const box = (top: number, left: number, bottom: number, right: number) => ({
      start: { row: top, col: left },
      end: { row: bottom, col: right },
    });
    const cover = unionCover([box(0, 0, 99, 99), box(0, 0, 98, 99), box(200, 0, 200, 0)]);
    expect(cover.area).toBe(10_001);
    expect([cover.holds(50, 50), cover.holds(150, 0), cover.holds(200, 0)]).toEqual([
      true,
      false,
      true,
    ]);
  });

  test('a format over ranges that add more cells than the limit is refused; one that overlaps is not', async () => {
    const { viewer } = await mount();
    // Two whole columns: about two million new cells together.
    viewer.selectRanges([
      { anchor: { row: 0, col: 0 }, focus: { row: MAX_ROWS - 1, col: 0 } },
      { anchor: { row: 0, col: 2 }, focus: { row: MAX_ROWS - 1, col: 2 } },
    ]);
    expect(toolbarCommandState(viewer.snapshot(), 'font.bold').disabledReason).not.toBeNull();
    // The command refuses the same selection, and writes nothing.
    const refused = viewer.applyFormat({ bold: true });
    expect(refused.ok ? null : refused.refusal.code).toBe('range-too-large');
    // Two ranges of 100 by 100 cells that overlap in all but one row: 10,000 cells, inside it.
    viewer.selectRanges([
      { anchor: { row: 0, col: 0 }, focus: { row: 99, col: 99 } },
      { anchor: { row: 0, col: 0 }, focus: { row: 98, col: 99 } },
    ]);
    expect(toolbarCommandState(viewer.snapshot(), 'font.bold').disabledReason).toBeNull();
  });
});

describe('a protected sheet with several ranges', () => {
  test('a locked cell with content in any range disables Clear contents', async () => {
    const { viewer } = await mount({
      rows: row(10, [c('D10', 1)]),
      beforeMerges: '<sheetProtection sheet="1"/>',
    });
    // The first range holds no content; the second, D9:E11, holds the locked cell D10.
    viewer.selectRanges([
      { anchor: { row: 1, col: 1 }, focus: { row: 2, col: 2 } },
      { anchor: { row: 8, col: 3 }, focus: { row: 10, col: 4 } },
    ]);
    expect(viewer.snapshot().selectionContentLocked).toBe(true);
    expect(toolbarCommandState(viewer.snapshot(), 'clear.contents').enabled).toBe(false);
    viewer.selectRanges([
      { anchor: { row: 1, col: 1 }, focus: { row: 2, col: 2 } },
      { anchor: { row: 20, col: 3 }, focus: { row: 22, col: 4 } },
    ]);
    expect(viewer.snapshot().selectionContentLocked).toBe(false);
  });
});

describe('the clear limit of several ranges', () => {
  test('a cell that two ranges hold counts once, as the store clears it once', async () => {
    // Column A has a column format, so a clear of its empty cells writes `s="0"` cells.
    const { viewer } = await mount({
      cols: '<cols><col min="1" max="1" width="9" style="1"/></cols>',
    });
    const r = (last: number) => ({ anchor: { row: 0, col: 0 }, focus: { row: last, col: 0 } });
    // Two ranges of about 9,000 rows that do not overlap pass the limit together.
    viewer.selectRanges([
      r(9000),
      { anchor: { row: 10000, col: 0 }, focus: { row: 19000, col: 0 } },
    ]);
    expect(toolbarCommandState(viewer.snapshot(), 'clear.formats').disabledReason).not.toBeNull();
    // Two that overlap clear 9,002 cells once, inside the limit.
    viewer.selectRanges([r(9000), r(9001)]);
    expect(viewer.snapshot().selection?.ranges).toHaveLength(2);
    expect(toolbarCommandState(viewer.snapshot(), 'clear.formats').disabledReason).toBeNull();
    expect(viewer.clearFormats().ok).toBe(true);
  });
});

describe('a history selection', () => {
  test('keeps at most MAX_SELECTION_RANGES ranges, and a valid active range', () => {
    const boxes = Array.from({ length: 1000 }, (_, i) => [i, 0, i, 0] as const);
    const { selection } = historyGridSelection(
      {
        sheet: 'S',
        active: { row: 0, col: 0 },
        anchor: { row: 0, col: 0 },
        focus: { row: 0, col: 0 },
        range: [0, 0, 0, 0],
        ranges: boxes,
        activeRange: 999,
      },
      0
    );
    expect(selection.ranges).toHaveLength(MAX_SELECTION_RANGES);
    // Past the kept ranges, the active range is the last kept one, and holds the active cell.
    expect(selection.activeRange).toBe(MAX_SELECTION_RANGES - 1);
    expect(selection.active).toEqual({ row: MAX_SELECTION_RANGES - 1, col: 0 });
    // With no active range given, an active cell that the last range holds stays.
    const kept = historyGridSelection(
      {
        sheet: 'S',
        active: { row: 3, col: 2 },
        anchor: { row: 3, col: 2 },
        focus: { row: 3, col: 2 },
        range: [2, 1, 4, 3],
        ranges: [
          [0, 0, 0, 0],
          [2, 1, 4, 3],
        ],
      },
      0
    ).selection;
    expect(kept.active).toEqual({ row: 3, col: 2 });
  });

  test('the store keeps the same number of ranges as the editor', () => {
    const read = readXlsx(buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] }));
    if (!read.ok) throw new Error('the package did not open');
    const store = new WorkbookDocumentStore(read.pkg);
    const cell = { row: 0, col: 0 };
    const boxes = Array.from({ length: 200 }, (_, i) => [i, 0, i, 0] as const);
    const op = {
      op: 'setCellValue',
      sheet: 'S',
      ref: 'A1',
      value: { type: 'number', value: 2 },
    } as const;
    const selection = {
      sheet: 'S',
      range: [199, 0, 199, 0] as const,
      active: cell,
      anchor: cell,
      focus: cell,
    };
    expect(
      store.transact([op], { selectionBefore: { ...selection, ranges: boxes, activeRange: 199 } })
        .ok
    ).toBe(true);
    const kept = store.selectionForUndo()!;
    expect(kept.ranges).toHaveLength(MAX_SELECTION_RANGES);
    // The active range was past the kept ranges, so the last kept one is active, and `range` is it.
    expect(kept.activeRange).toBe(MAX_SELECTION_RANGES - 1);
    expect(kept.range).toEqual([MAX_SELECTION_RANGES - 1, 0, MAX_SELECTION_RANGES - 1, 0]);
  });
});
