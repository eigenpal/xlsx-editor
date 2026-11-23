// How a sheet rename ends in happy-dom: a press on the grid that blurs a refused rename keeps
// its reason, Escape clears the reason of a refused Enter, a new rename clears the reason of
// an earlier one, a press on the add button commits the rename first, and a renamed tab
// scrolls into view.
// The Vue twin of the React adapter's `sheet-tabs-rename.test.tsx`: Vue renders in the next microtask, so a
// step waits for `nextTick` before it reads the DOM.

import { afterEach, describe, expect, test } from 'bun:test';
import { fireEvent, waitFor } from '@testing-library/react';
import { h, nextTick } from 'vue';
import { buildXlsx, row, c } from '../../../scripts/fixtures/builder.ts';
import { XlsxEditor, type XlsxEditorInstance } from '../src/index.ts';
import { FONTS, change, mount, tabsShow, unmountAll } from './mount.ts';

afterEach(unmountAll);

const three = buildXlsx({
  sheets: [
    { name: 'A', rows: row(1, [c('A1', 1)]) },
    { name: 'B', rows: '' },
    { name: 'C', rows: '' },
  ],
});

async function open(document = three, tabs = 3) {
  let viewer: XlsxEditorInstance | null = null;
  const view = mount(() =>
    h('div', { style: { height: '400px' } }, [
      h(XlsxEditor, {
        document: document,
        fonts: FONTS,
        onReady: (v: XlsxEditorInstance) => (viewer = v),
      }),
    ])
  );
  await tabsShow(view, tabs);
  return { view, viewer: viewer! };
}

type View = Awaited<ReturnType<typeof open>>['view'];
const names = (view: View) => view.getAllByRole('tab').map((t) => t.textContent);
const shownRefusal = (view: View) =>
  view.container.querySelector('.xlsx-refusal')?.textContent ?? null;

/** Start a rename of the first tab and type `text` in its field. */
async function typeName(view: View, text: string): Promise<HTMLInputElement> {
  fireEvent.doubleClick(view.getAllByRole('tab')[0]!);
  await nextTick();
  const field = view.getByLabelText('Sheet name') as HTMLInputElement;
  change(field, text);
  await nextTick();
  return field;
}

describe('a press on the grid during a rename', () => {
  test('a refused name keeps its reason after the press selects a cell', async () => {
    const { view, viewer } = await open();
    const stage = view.container.querySelector('.xlsx-stage') as HTMLElement;
    await typeName(view, 'b');
    const W = stage.ownerDocument.defaultView as unknown as typeof globalThis;
    stage.dispatchEvent(
      new W.PointerEvent('pointerdown', {
        bubbles: true,
        button: 0,
        pointerId: 1,
        clientX: 100,
        clientY: 50,
      })
    );
    await nextTick();
    expect(view.queryByLabelText('Sheet name')).toBeNull();
    expect(names(view)).toEqual(['A', 'B', 'C']);
    // The press selected a cell, and the reason of the refused rename stays.
    expect(viewer.snapshot().selection?.active).not.toEqual({ row: 0, col: 0 });
    expect(viewer.snapshot().refusal?.code).toBe('sheet-name-taken');
    await waitFor(() => expect(shownRefusal(view)).toContain('has the name B.'));
  });
});

describe('Escape after a refused Enter', () => {
  test('ends the rename and clears its reason', async () => {
    const { view, viewer } = await open();
    const field = await typeName(view, 'b');
    fireEvent.keyDown(field, { key: 'Enter' });
    await nextTick();
    expect(viewer.snapshot().refusal?.code).toBe('sheet-name-taken');
    fireEvent.keyDown(field, { key: 'Escape' });
    await nextTick();
    expect(view.queryByLabelText('Sheet name')).toBeNull();
    expect(viewer.snapshot().refusal).toBeNull();
    await waitFor(() => expect(shownRefusal(view)).toBeNull());
  });
});

describe('a new rename after a refused blur', () => {
  test('clears the reason of the earlier rename and keeps another reason', async () => {
    const { view, viewer } = await open();
    fireEvent.blur(await typeName(view, 'b'));
    await nextTick();
    expect(view.queryByLabelText('Sheet name')).toBeNull();
    expect(viewer.snapshot().refusal?.code).toBe('sheet-name-taken');
    fireEvent.doubleClick(view.getAllByRole('tab')[0]!);
    await nextTick();
    expect(view.getByLabelText('Sheet name')).toBeTruthy();
    expect(viewer.snapshot().refusal).toBeNull();
    await waitFor(() => expect(shownRefusal(view)).toBeNull());
    fireEvent.keyDown(view.getByLabelText('Sheet name'), { key: 'Escape' });
    await nextTick();
    // A reason that no rename made stays when a rename starts.
    viewer.moveSheet(0, 9);
    await nextTick();
    expect(viewer.snapshot().refusal?.code).toBe('bad-position');
    fireEvent.doubleClick(view.getAllByRole('tab')[0]!);
    await nextTick();
    expect(view.getByLabelText('Sheet name')).toBeTruthy();
    expect(viewer.snapshot().refusal?.code).toBe('bad-position');
  });
});

describe('a refused Enter', () => {
  test('marks the field invalid and describes it with the reason', async () => {
    const { view } = await open();
    const field = await typeName(view, 'b');
    expect(field.getAttribute('aria-invalid')).toBeNull();
    fireEvent.keyDown(field, { key: 'Enter' });
    await nextTick();
    await waitFor(() => expect(field.getAttribute('aria-invalid')).toBe('true'));
    const id = field.getAttribute('aria-describedby') ?? '';
    const reason = view.container.ownerDocument.getElementById(id);
    expect(reason?.textContent).toContain('has the name B.');
  });
});

describe('the reason of a refused rename in the field', () => {
  // Conditional formatting on a sheet whose name holds a bidi control names sheet `A`.
  const blocked = buildXlsx({
    sheets: [
      { name: 'A', rows: row(1, [c('A1', 1)]) },
      {
        name: 'Other\u202E',
        rows: row(1, [c('A1', 1)]),
        afterMerges:
          '<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1" dxfId="0"><formula>A!A1&gt;0</formula></cfRule></conditionalFormatting>',
      },
    ],
  });

  test('is the text of the status message: the content in words, with no bidi control', async () => {
    const { view, viewer } = await open(blocked, 2);
    const field = await typeName(view, 'X');
    fireEvent.keyDown(field, { key: 'Enter' });
    await nextTick();
    expect(viewer.snapshot().refusal?.code).toBe('structure-blocked');
    await waitFor(() => expect(field.getAttribute('aria-invalid')).toBe('true'));
    const id = field.getAttribute('aria-describedby') ?? '';
    const reason = view.container.ownerDocument.getElementById(id)?.textContent ?? '';
    await waitFor(() => expect(shownRefusal(view)).toBe(reason));
    expect(reason).toContain('conditional formatting');
    expect(reason).not.toContain('/xl/');
    expect(reason).not.toContain('\u202E');
  });
});

describe('an IME composition in the rename field', () => {
  test('Enter and Escape that end a composition neither commit nor cancel the rename', async () => {
    const { view } = await open();
    const field = await typeName(view, 'Neu');
    fireEvent.keyDown(field, { key: 'Enter', keyCode: 229 });
    await nextTick();
    fireEvent.keyDown(field, { key: 'Enter', isComposing: true });
    await nextTick();
    fireEvent.keyDown(field, { key: 'Escape', isComposing: true });
    await nextTick();
    expect(view.queryByLabelText('Sheet name')).not.toBeNull();
    expect(names(view)).toEqual(['A', 'B', 'C']);
    fireEvent.keyDown(field, { key: 'Enter' });
    await nextTick();
    await waitFor(() => expect(names(view)).toEqual(['Neu', 'B', 'C']));
  });
});

describe('undo and redo keys on a focused tab', () => {
  test('undo restores a deleted sheet, and redo deletes it again', async () => {
    const { view, viewer } = await open();
    viewer.deleteSheet(2);
    await nextTick();
    await waitFor(() => expect(names(view)).toEqual(['A', 'B']));
    const tab = view.getAllByRole('tab')[0]!;
    tab.focus();
    expect(fireEvent.keyDown(tab, { key: 'z', ctrlKey: true })).toBe(false);
    await waitFor(() => expect(names(view)).toEqual(['A', 'B', 'C']));
    fireEvent.keyDown(view.getAllByRole('tab')[0]!, { key: 'y', ctrlKey: true });
    await nextTick();
    await waitFor(() => expect(names(view)).toEqual(['A', 'B']));
  });
});

describe('the add button during a rename', () => {
  test('writes the typed name first, adds the sheet, and gives focus to the grid', async () => {
    const { view, viewer } = await open();
    await typeName(view, 'Renamed');
    fireEvent.click(view.getByRole('button', { name: 'Add a sheet' }));
    await nextTick();
    await waitFor(() => expect(names(view)).toEqual(['Renamed', 'Sheet1', 'B', 'C']));
    expect(view.queryByLabelText('Sheet name')).toBeNull();
    const active = view.container.ownerDocument.activeElement;
    expect(active?.classList.contains('xlsx-scroller')).toBe(true);
    expect(viewer.snapshot().canUndo).toBe(true);
  });

  test('a refused name keeps the old name and its reason, and adds no sheet', async () => {
    const { view, viewer } = await open();
    await typeName(view, 'b');
    fireEvent.click(view.getByRole('button', { name: 'Add a sheet' }));
    await nextTick();
    expect(view.queryByLabelText('Sheet name')).toBeNull();
    expect(names(view)).toEqual(['A', 'B', 'C']);
    expect(viewer.snapshot().refusal?.code).toBe('sheet-name-taken');
    const active = view.container.ownerDocument.activeElement;
    expect(active?.classList.contains('xlsx-scroller')).toBe(true);
  });
});

describe('a renamed tab', () => {
  test('scrolls into view in the tab list when the rename changes the sheet list', async () => {
    const { view } = await open();
    const list = view.getByRole('tablist');
    list.style.overflowX = 'auto';
    Object.defineProperty(list, 'scrollWidth', { get: () => 500 });
    Object.defineProperty(list, 'clientWidth', { get: () => 100 });
    const scrolled: (string | null)[] = [];
    let measured: string | null = null;
    const proto = Object.getPrototypeOf(view.getAllByRole('tab')[0]!) as {
      getBoundingClientRect: () => DOMRect;
    };
    const before = proto.getBoundingClientRect;
    proto.getBoundingClientRect = function (this: HTMLElement) {
      if (this.getAttribute('role') !== 'tab') return before.call(this);
      measured = this.textContent;
      return { left: 300, top: 0, width: 60, height: 20, right: 360, bottom: 20 } as DOMRect;
    };
    list.scrollBy = (() => void scrolled.push(measured)) as typeof list.scrollBy;
    try {
      const field = await typeName(view, 'A much longer sheet name');
      scrolled.length = 0;
      fireEvent.keyDown(field, { key: 'Enter' });
      await nextTick();
      await waitFor(() => expect(scrolled).toContain('A much longer sheet name'));
    } finally {
      proto.getBoundingClientRect = before;
    }
  });
});

/** A full press on a tab: mousedown with its click count, then click. */
async function press(tab: HTMLElement, detail = 1) {
  fireEvent.mouseDown(tab, { detail });
  await nextTick();
  fireEvent.click(tab, { detail });
  await nextTick();
}

describe('a press on another tab during a rename', () => {
  test('writes the typed name first, then chooses the tab', async () => {
    const { view, viewer } = await open();
    await typeName(view, 'Renamed');
    await press(view.getAllByRole('tab')[2]!);
    await waitFor(() => expect(names(view)).toEqual(['Renamed', 'B', 'C']));
    expect(view.queryByLabelText('Sheet name')).toBeNull();
    expect(viewer.snapshot().activeSheet).toBe(2);
  });

  test('a refused name keeps its reason, and the press chooses no tab', async () => {
    const { view, viewer } = await open();
    await typeName(view, 'b');
    await press(view.getAllByRole('tab')[2]!);
    expect(view.queryByLabelText('Sheet name')).toBeNull();
    expect(names(view)).toEqual(['A', 'B', 'C']);
    expect(viewer.snapshot().activeSheet).toBe(0);
    expect(viewer.snapshot().refusal?.code).toBe('sheet-name-taken');
    const active = view.container.ownerDocument.activeElement;
    expect(active?.classList.contains('xlsx-scroller')).toBe(true);
  });
});

describe('a double-click on another tab during a rename', () => {
  test('writes the typed name, then starts the rename of that tab', async () => {
    const { view } = await open();
    await typeName(view, 'Renamed');
    const tab = view.getAllByRole('tab')[1]!;
    await press(tab, 1);
    await press(tab, 2);
    fireEvent.doubleClick(tab);
    await nextTick();
    await waitFor(() => expect(names(view)).toEqual(['Renamed', 'B', 'C']));
    const field = view.getByLabelText('Sheet name') as HTMLInputElement;
    expect(field.value).toBe('B');
  });

  test('a refused name starts no other rename', async () => {
    const { view, viewer } = await open();
    await typeName(view, 'b');
    const tab = view.getAllByRole('tab')[1]!;
    await press(tab, 1);
    await press(tab, 2);
    fireEvent.doubleClick(tab);
    await nextTick();
    expect(view.queryByLabelText('Sheet name')).toBeNull();
    expect(names(view)).toEqual(['A', 'B', 'C']);
    expect(viewer.snapshot().refusal?.code).toBe('sheet-name-taken');
  });
});

describe('a right press on a tab during a rename', () => {
  test('Escape on the menu gives focus to the grid, not to the page', async () => {
    const { view } = await open();
    await typeName(view, 'Renamed');
    const tab = view.getAllByRole('tab')[1]!;
    fireEvent.mouseDown(tab, { button: 2, detail: 1 });
    await nextTick();
    fireEvent.contextMenu(tab, { clientX: 40 });
    await nextTick();
    const menu = view.getByRole('menu');
    fireEvent.keyDown(menu, { key: 'Escape' });
    await nextTick();
    expect(view.queryByRole('menu')).toBeNull();
    const active = view.container.ownerDocument.activeElement;
    expect(active).not.toBe(view.container.ownerDocument.body);
    expect(active?.classList.contains('xlsx-scroller')).toBe(true);
    await waitFor(() => expect(names(view)).toEqual(['Renamed', 'B', 'C']));
  });

  test('a menu whose opener is gone at close gives focus to the grid', async () => {
    const { view } = await open();
    await typeName(view, 'Renamed');
    const tab = view.getAllByRole('tab')[1]!;
    // A context menu event alone, with no press before it: the rename field had focus.
    fireEvent.contextMenu(tab, { clientX: 40 });
    await nextTick();
    const menu = view.getByRole('menu');
    fireEvent.keyDown(menu, { key: 'Escape' });
    await nextTick();
    const active = view.container.ownerDocument.activeElement;
    expect(active?.classList.contains('xlsx-scroller')).toBe(true);
  });
});
