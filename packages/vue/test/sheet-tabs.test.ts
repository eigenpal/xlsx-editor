// The sheet tab actions in happy-dom (`bun run test:dom` preloads it): the add button, inline
// rename, and the tab context menu with its keyboard rules.
// The Vue twin of the React adapter's `sheet-tabs.test.tsx`: Vue renders in the next microtask, so a
// step waits for `nextTick` before it reads the DOM.

import { afterEach, describe, expect, test } from 'bun:test';
import { fireEvent, waitFor } from '@testing-library/react';
import { h, nextTick } from 'vue';
import { buildXlsx, c, row } from '../../../scripts/fixtures/builder.ts';
import { XlsxEditor, type XlsxEditorInstance } from '../src/index.ts';
import { FONTS, change, mount, tabsShow, unmountAll } from './mount.ts';

afterEach(unmountAll);

const workbook = buildXlsx({
  sheets: [
    { name: 'First', rows: row(1, [c('A1', 'one')]) },
    { name: 'Second', rows: row(1, [c('A1', 'two')]) },
  ],
});

async function open(bytes = workbook, tabs = 2) {
  let viewer: XlsxEditorInstance | null = null;
  const view = mount(() =>
    h('div', { style: { height: '400px' } }, [
      h(XlsxEditor, {
        document: bytes,
        fonts: FONTS,
        onReady: (v: XlsxEditorInstance) => (viewer = v),
      }),
    ])
  );
  await tabsShow(view, tabs);
  return { view, viewer: viewer! };
}

const names = (view: Awaited<ReturnType<typeof open>>['view']) =>
  view.getAllByRole('tab').map((t) => t.textContent);

describe('SheetTabs actions', () => {
  test('the add button adds a sheet after the shown one and shows it', async () => {
    const { view, viewer } = await open();
    const add = view.getByRole('button', { name: 'Add a sheet' });
    expect(fireEvent.mouseDown(add)).toBe(false);
    fireEvent.click(add);
    await nextTick();
    expect(names(view)).toEqual(['First', 'Sheet1', 'Second']);
    expect(viewer.snapshot().activeSheet).toBe(1);
  });

  test('a position outside the sheet list shows the bad-position reason', async () => {
    const { view, viewer } = await open();
    expect(viewer.moveSheet(0, 9).ok).toBe(false);
    await nextTick();
    await waitFor(() =>
      expect(view.container.querySelector('.xlsx-refusal')?.textContent).toBe(
        'The sheet was not added or moved: the position is outside the sheet list.'
      )
    );
    expect(names(view)).toEqual(['First', 'Second']);
  });

  test('a double-click renames the tab; Enter writes the trimmed name', async () => {
    const { view, viewer } = await open();
    fireEvent.doubleClick(view.getAllByRole('tab')[1]!);
    await nextTick();
    const field = view.getByLabelText('Sheet name') as HTMLInputElement;
    expect(field.value).toBe('Second');
    expect(field.ownerDocument.activeElement).toBe(field);
    change(field, '  Totals ');
    await nextTick();
    fireEvent.keyDown(field, { key: 'Enter' });
    await nextTick();
    expect(names(view)).toEqual(['First', 'Totals']);
    expect(viewer.snapshot().sheets[1]!.name).toBe('Totals');
    expect(view.queryByLabelText('Sheet name')).toBeNull();
  });

  test('Escape cancels a rename, and an unchanged name writes nothing', async () => {
    const { view, viewer } = await open();
    const revision = () =>
      viewer
        .snapshot()
        .sheets.map((s) => s.name)
        .join();
    const before = revision();
    fireEvent.doubleClick(view.getAllByRole('tab')[0]!);
    await nextTick();
    let field = view.getByLabelText('Sheet name') as HTMLInputElement;
    change(field, 'Other');
    await nextTick();
    fireEvent.keyDown(field, { key: 'Escape' });
    await nextTick();
    fireEvent.blur(field);
    await nextTick();
    expect(revision()).toBe(before);
    expect(viewer.undo().ok).toBe(false);

    fireEvent.doubleClick(view.getAllByRole('tab')[0]!);
    await nextTick();
    field = view.getByLabelText('Sheet name') as HTMLInputElement;
    change(field, ' First ');
    await nextTick();
    fireEvent.blur(field);
    await nextTick();
    expect(revision()).toBe(before);
    expect(viewer.undo().ok).toBe(false);
  });

  test('a keydown in the rename field does not reach the tabs', async () => {
    const { view, viewer } = await open();
    fireEvent.doubleClick(view.getAllByRole('tab')[0]!);
    await nextTick();
    const field = view.getByLabelText('Sheet name');
    fireEvent.keyDown(field, { key: 'ArrowRight' });
    await nextTick();
    fireEvent.keyDown(field, { key: 'End' });
    await nextTick();
    expect(viewer.snapshot().activeSheet).toBe(0);
    expect(view.getByLabelText('Sheet name')).toBe(field);
  });

  test('a rename to a taken name shows the refusal and keeps the text and the focus', async () => {
    const { view } = await open();
    fireEvent.doubleClick(view.getAllByRole('tab')[0]!);
    await nextTick();
    const field = view.getByLabelText('Sheet name') as HTMLInputElement;
    change(field, 'second');
    await nextTick();
    fireEvent.keyDown(field, { key: 'Enter' });
    await nextTick();
    await waitFor(() =>
      expect(view.container.querySelector('.xlsx-refusal')?.textContent).toContain(
        'has the name Second.'
      )
    );
    expect(names(view)).toEqual(['First', 'Second']);
    expect(view.getByLabelText('Sheet name')).toBe(field);
    expect(field.value).toBe('second');
    expect(field.ownerDocument.activeElement).toBe(field);
    // A fixed name then goes through.
    change(field, 'Third');
    await nextTick();
    fireEvent.keyDown(field, { key: 'Enter' });
    await nextTick();
    expect(names(view)).toEqual(['Third', 'Second']);
  });

  test('the rename field is outside the tab list, and F2 starts it', async () => {
    const { view } = await open();
    fireEvent.keyDown(view.getAllByRole('tab')[1]!, { key: 'F2' });
    await nextTick();
    const field = view.getByLabelText('Sheet name');
    expect(field.closest('[role="tablist"]')).toBeNull();
    const list = view.getByRole('tablist');
    for (const el of list.querySelectorAll('*'))
      expect(el.getAttribute('role') === 'tab' || el.closest('[role="tab"]') !== null).toBe(true);
  });

  test('the menu opens on a right press and deletes the sheet', async () => {
    const { view, viewer } = await open();
    fireEvent.contextMenu(view.getAllByRole('tab')[1]!);
    await nextTick();
    const menu = view.getByRole('menu', { name: 'Sheet actions' });
    expect(menu.ownerDocument.activeElement?.textContent).toBe('Rename');
    fireEvent.click(view.getByRole('menuitem', { name: 'Delete' }));
    await nextTick();
    expect(view.queryByRole('menu')).toBeNull();
    expect(names(view)).toEqual(['First']);
    expect(viewer.snapshot().sheets.length).toBe(1);
  });

  test('Move right moves the sheet; Move left is disabled on the first tab', async () => {
    const { view, viewer } = await open();
    fireEvent.contextMenu(view.getAllByRole('tab')[0]!);
    await nextTick();
    const left = view.getByRole('menuitem', { name: 'Move left' });
    expect(left.getAttribute('aria-disabled')).toBe('true');
    fireEvent.click(left);
    await nextTick();
    expect(names(view)).toEqual(['First', 'Second']);
    fireEvent.click(view.getByRole('menuitem', { name: 'Move right' }));
    await nextTick();
    expect(names(view)).toEqual(['Second', 'First']);
    expect(viewer.snapshot().sheets.map((s) => s.name)).toEqual(['Second', 'First']);
  });

  test('the spec scenario: A, hidden H, and B; a move of A right gives H, B, A', async () => {
    const bytes = buildXlsx({
      sheets: [
        { name: 'A', rows: '' },
        { name: 'H', rows: '', state: 'hidden' },
        { name: 'B', rows: '' },
      ],
    });
    const { view, viewer } = await open(bytes, 2);
    fireEvent.contextMenu(view.getAllByRole('tab')[0]!);
    await nextTick();
    fireEvent.click(view.getByRole('menuitem', { name: 'Move right' }));
    await nextTick();
    expect(viewer.snapshot().sheets.map((s) => s.name)).toEqual(['H', 'B', 'A']);
    expect(names(view)).toEqual(['B', 'A']);
  });

  test('with hidden sheets, a move passes the next visible tab', async () => {
    const bytes = buildXlsx({
      sheets: [
        { name: 'A', rows: '' },
        { name: 'H1', rows: '', state: 'hidden' },
        { name: 'B', rows: '' },
        { name: 'H2', rows: '', state: 'hidden' },
        { name: 'C', rows: '' },
      ],
    });
    const { view, viewer } = await open(bytes, 3);
    const order = () => viewer.snapshot().sheets.map((s) => s.name);
    fireEvent.contextMenu(view.getAllByRole('tab')[1]!);
    await nextTick();
    fireEvent.click(view.getByRole('menuitem', { name: 'Move right' }));
    await nextTick();
    expect(names(view)).toEqual(['A', 'C', 'B']);
    expect(order()).toEqual(['A', 'H1', 'H2', 'C', 'B']);
    // The last tab cannot move right, and the move left passes C.
    fireEvent.contextMenu(view.getAllByRole('tab')[2]!);
    await nextTick();
    const right = view.getByRole('menuitem', { name: 'Move right' });
    expect(right.getAttribute('aria-disabled')).toBe('true');
    fireEvent.click(view.getByRole('menuitem', { name: 'Move left' }));
    await nextTick();
    expect(names(view)).toEqual(['A', 'B', 'C']);
    fireEvent.contextMenu(view.getAllByRole('tab')[1]!);
    await nextTick();
    fireEvent.click(view.getByRole('menuitem', { name: 'Move left' }));
    await nextTick();
    expect(names(view)).toEqual(['B', 'A', 'C']);
    expect(order()[0]).toBe('B');
  });

  test('the tab with an open menu is expanded, and the menu stays inside the bar', async () => {
    const { view } = await open();
    const bar = view.container.querySelector('.xlsx-sheet-bar') as HTMLElement;
    Object.defineProperty(bar, 'clientWidth', { value: 300, configurable: true });
    const proto = Object.getPrototypeOf(Object.getPrototypeOf(bar)) as object;
    const before = Object.getOwnPropertyDescriptor(proto, 'offsetWidth');
    Object.defineProperty(proto, 'offsetWidth', { get: () => 160, configurable: true });
    try {
      const tab = view.getAllByRole('tab')[1]!;
      expect(tab.getAttribute('aria-expanded')).toBe('false');
      fireEvent.contextMenu(tab, { clientX: 290 });
      await nextTick();
      expect(tab.getAttribute('aria-expanded')).toBe('true');
      expect(view.getByRole('menu').style.left).toBe('140px');
    } finally {
      if (before) Object.defineProperty(proto, 'offsetWidth', before);
    }
  });

  test('Shift+F10 opens the menu, arrows move, and Escape returns focus to the tab', async () => {
    const { view } = await open();
    const tab = view.getAllByRole('tab')[0]!;
    fireEvent.keyDown(tab, { key: 'F10', shiftKey: true });
    await nextTick();
    const menu = view.getByRole('menu');
    const doc = menu.ownerDocument;
    expect(doc.activeElement?.textContent).toBe('Rename');
    fireEvent.keyDown(doc.activeElement!, { key: 'ArrowDown' });
    await nextTick();
    expect(doc.activeElement?.textContent).toBe('Delete');
    fireEvent.keyDown(doc.activeElement!, { key: 'End' });
    await nextTick();
    expect(doc.activeElement?.textContent).toBe('Move right');
    fireEvent.keyDown(doc.activeElement!, { key: 'ArrowDown' });
    await nextTick();
    expect(doc.activeElement?.textContent).toBe('Rename');
    fireEvent.keyDown(doc.activeElement!, { key: 'Escape' });
    await nextTick();
    expect(view.queryByRole('menu')).toBeNull();
    await waitFor(() => expect(doc.activeElement).toBe(view.getAllByRole('tab')[0]!));
  });

  test('the Menu key opens the menu, and Rename starts the rename field', async () => {
    const { view } = await open();
    fireEvent.keyDown(view.getAllByRole('tab')[1]!, { key: 'ContextMenu' });
    await nextTick();
    fireEvent.click(view.getByRole('menuitem', { name: 'Rename' }));
    await nextTick();
    const field = view.getByLabelText('Sheet name') as HTMLInputElement;
    expect(field.value).toBe('Second');
  });
});

describe('SheetTabs in view mode', () => {
  async function openViewing() {
    let viewer: XlsxEditorInstance | null = null;
    const view = mount(() =>
      h('div', { style: { height: '400px' } }, [
        h(XlsxEditor, {
          document: workbook,
          fonts: FONTS,
          mode: 'view',
          onReady: (v: XlsxEditorInstance) => (viewer = v),
        }),
      ])
    );
    await tabsShow(view);
    return { view, viewer: viewer! };
  }
  const REASON = 'The workbook is open for viewing.';

  test('the add button is disabled with the reason of view mode, and adds no sheet', async () => {
    const { view } = await openViewing();
    const add = view.getByRole('button', { name: 'Add a sheet' });
    expect(add.getAttribute('aria-disabled')).toBe('true');
    expect(add.getAttribute('title')).toBe(`Add a sheet: ${REASON}`);
    const reason = add.ownerDocument.getElementById(add.getAttribute('aria-describedby') ?? '');
    expect(reason?.textContent).toBe(REASON);
    fireEvent.click(add);
    await nextTick();
    expect(names(view)).toEqual(['First', 'Second']);
    expect(view.container.querySelector('.xlsx-refusal')).toBeNull();
  });

  test('a double-click and F2 start no rename, and the tab says why', async () => {
    const { view, viewer } = await openViewing();
    const tab = view.getAllByRole('tab')[0]!;
    fireEvent.doubleClick(tab);
    await nextTick();
    expect(view.queryByLabelText('Sheet name')).toBeNull();
    fireEvent.keyDown(tab, { key: 'F2' });
    await nextTick();
    expect(view.queryByLabelText('Sheet name')).toBeNull();
    const why = tab.ownerDocument.getElementById(tab.getAttribute('aria-describedby') ?? '');
    expect(why?.textContent).toBe(REASON);
    // A mouse user sees the reason in the tooltip, as a disabled toolbar control shows it.
    expect(tab.getAttribute('title')).toBe(`First: ${REASON}`);
    // In edit mode the tab has no such description, and its tooltip is its name.
    expect(viewer.setMode('edit')).toBe(true);
    await nextTick();
    await waitFor(() => expect(tab.getAttribute('aria-describedby')).toBeNull());
    expect(tab.getAttribute('title')).toBe('First');
  });

  test('the menu disables Rename, Delete, Move left, and Move right with the reason', async () => {
    const { view } = await openViewing();
    const tab = view.getAllByRole('tab')[0]!;
    fireEvent.contextMenu(tab, { clientX: 10 });
    await nextTick();
    // The sheet menu's rows, not the menu bar's triggers.
    const rows = view.getAllByRole('menuitem').filter((m) => !m.closest('[role="menubar"]'));
    expect(rows.map((r) => r.textContent)).toEqual(['Rename', 'Delete', 'Move left', 'Move right']);
    for (const r of rows) {
      expect(r.getAttribute('aria-disabled')).toBe('true');
      expect(r.getAttribute('title')).toBe(`${r.textContent}: ${REASON}`);
    }
    // Each row is described by the reason of its own command state.
    const described = rows.map((r) => r.getAttribute('aria-describedby') ?? '');
    expect(new Set(described).size).toBe(4);
    for (const id of described)
      expect(tab.ownerDocument.getElementById(id)?.textContent).toBe(REASON);
    fireEvent.click(rows[0]!);
    await nextTick();
    expect(view.queryByLabelText('Sheet name')).toBeNull();
    fireEvent.click(rows[1]!);
    await nextTick();
    fireEvent.click(rows[3]!);
    await nextTick();
    expect(names(view)).toEqual(['First', 'Second']);
  });

  test('a switch back to edit mode enables the add button', async () => {
    const { view, viewer } = await openViewing();
    expect(viewer.setMode('edit')).toBe(true);
    await nextTick();
    const add = view.getByRole('button', { name: 'Add a sheet' });
    await waitFor(() => expect(add.getAttribute('aria-disabled')).toBeNull());
    fireEvent.click(add);
    await nextTick();
    expect(names(view)).toEqual(['First', 'Sheet1', 'Second']);
  });
});

describe('SheetTabs on a workbook that locks its structure', () => {
  const locked = buildXlsx({
    sheets: [
      { name: 'First', rows: row(1, [c('A1', 'one')]) },
      { name: 'Second', rows: row(1, [c('A1', 'two')]) },
    ],
    workbookProtection: 'lockStructure="1"',
  });
  const REASON =
    'The workbook is protected, and adding, renaming, deleting, or moving sheets is not allowed.';

  test('disables the add button and the menu edits, and starts no rename', async () => {
    const { view, viewer } = await open(locked);
    expect(viewer.snapshot().workbookProtection).toEqual({ allowStructure: false });
    const add = view.getByRole('button', { name: 'Add a sheet' });
    expect(add.getAttribute('aria-disabled')).toBe('true');
    expect(add.getAttribute('title')).toBe(`Add a sheet: ${REASON}`);
    const reason = add.ownerDocument.getElementById(add.getAttribute('aria-describedby') ?? '');
    expect(reason?.textContent).toBe(REASON);
    fireEvent.click(add);
    await nextTick();
    const tab = view.getAllByRole('tab')[0]!;
    fireEvent.doubleClick(tab);
    await nextTick();
    expect(view.queryByLabelText('Sheet name')).toBeNull();
    fireEvent.keyDown(tab, { key: 'F2' });
    await nextTick();
    expect(view.queryByLabelText('Sheet name')).toBeNull();
    // The tab describes itself with the reason, so F2 and a double-click are not silent.
    const why = tab.ownerDocument.getElementById(tab.getAttribute('aria-describedby') ?? '');
    expect(why?.textContent).toBe(REASON);
    fireEvent.contextMenu(tab, { clientX: 10 });
    await nextTick();
    // The sheet menu's rows, not the menu bar's triggers.
    const rows = view.getAllByRole('menuitem').filter((m) => !m.closest('[role="menubar"]'));
    for (const r of rows) {
      expect(r.getAttribute('aria-disabled')).toBe('true');
      expect(r.getAttribute('title')).toBe(`${r.textContent}: ${REASON}`);
    }
    for (const r of rows) fireEvent.click(r);
    expect(names(view)).toEqual(['First', 'Second']);
    // The instance refuses a sheet op with the workbook rule, and the strip names it.
    const r = viewer.addSheet();
    expect(r.ok ? null : [r.refusal.code, r.refusal.lock]).toEqual(['locked', 'workbook']);
    await nextTick();
    await waitFor(() =>
      expect(view.container.querySelector('.xlsx-refusal')?.textContent).toBe(
        'The sheets were not changed: the workbook is protected, and adding, renaming, ' +
          'deleting, or moving sheets is not allowed.'
      )
    );
    expect(names(view)).toEqual(['First', 'Second']);
  });
});

describe('the undo key in an open tab menu', () => {
  test('closes the menu and undoes the last step', async () => {
    const { view, viewer } = await open();
    fireEvent.click(view.getByRole('button', { name: 'Add a sheet' }));
    await nextTick();
    expect(names(view)).toEqual(['First', 'Sheet1', 'Second']);
    const tab = view.getAllByRole('tab')[1]!;
    fireEvent.keyDown(tab, { key: 'F10', shiftKey: true });
    await nextTick();
    const menu = view.getByRole('menu');
    fireEvent.keyDown(menu, { key: 'z', ctrlKey: true });
    await nextTick();
    expect(view.queryByRole('menu')).toBeNull();
    expect(names(view)).toEqual(['First', 'Second']);
    expect(viewer.snapshot().canRedo).toBe(true);
    await waitFor(() =>
      expect(view.container.ownerDocument.activeElement?.getAttribute('data-sheet-index')).toBe(
        String(viewer.snapshot().activeSheet)
      )
    );
  });
});
