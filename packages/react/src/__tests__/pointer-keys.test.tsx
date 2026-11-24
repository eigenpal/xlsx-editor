/** @jsxImportSource react */
// The keys of a chrome popup that a pointer press opened, in happy-dom (`bun run test:dom`): a
// navigation key from the grid moves the keyboard into the open menu or list, not the grid
// selection (`listenForPointerPopupKeys`). Also the state of a control that only opens rows (a
// group menu, a submenu) in view mode, the icons of the delete and clear rows, the custom hex
// field, and the Tab return column. The Vue twin is packages/vue/test/pointer-keys.test.ts.

import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react';
import type { EditorMode } from '@xlsx-editor.dev/core/editor';
import { XlsxEditor, type XlsxEditorInstance } from '../index.ts';
import { FONTS, workbook } from './toolbar-stubs.tsx';

afterEach(cleanup);

async function open(mode: EditorMode = 'edit') {
  let viewer: XlsxEditorInstance | null = null;
  const view = render(
    <div style={{ height: 400 }}>
      <XlsxEditor document={workbook()} fonts={FONTS} mode={mode} onReady={(v) => (viewer = v)} />
    </div>
  );
  await waitFor(() => expect(viewer?.snapshot().status).toBe('ready'), { timeout: 5000 });
  const grid = view.container.querySelector<HTMLElement>('.xlsx-scroller')!;
  act(() => viewer!.focus());
  const bar = view.getByRole('menubar');
  return { view, viewer: viewer!, grid, bar };
}

const trigger = (bar: HTMLElement, id: string) =>
  bar.querySelector<HTMLElement>(`[data-menu="${id}"] > [role="menuitem"]`)!;
const label = (el: Element | null) => el?.querySelector('.xlsx-menu-label')?.textContent ?? null;
const key = (el: Element, k: string) => fireEvent.keyDown(el, { key: k });
const active = (el: Element) => el.ownerDocument.activeElement;

describe('a menu opened by a pointer press', () => {
  test('ArrowDown from the grid moves the keyboard into the menu, not the selection', async () => {
    const { viewer, grid, bar } = await open();
    const before = viewer.snapshot().selection;
    fireEvent.click(trigger(bar, 'edit'));
    expect(active(grid)).toBe(grid);
    key(grid, 'ArrowDown');
    expect(viewer.snapshot().selection).toEqual(before);
    expect(label(active(grid))).toBe('Undo');
    // Escape gives the keyboard back to the grid, where it was.
    key(active(grid)!, 'Escape');
    expect(bar.querySelector('[role="menu"]')).toBeNull();
    expect(active(grid)).toBe(grid);
  });

  test('ArrowUp and a letter go to the last row and the row that starts with it', async () => {
    const { grid, bar } = await open();
    fireEvent.click(trigger(bar, 'edit'));
    key(grid, 's');
    expect(label(active(grid))).toBe('Select all');
    key(active(grid)!, 'Escape');
    fireEvent.click(trigger(bar, 'edit'));
    key(grid, 'ArrowUp');
    expect(label(active(grid))).toBe('Select all');
  });

  test('ArrowRight from the grid opens the next menu on its first row', async () => {
    const { grid, bar } = await open();
    fireEvent.click(trigger(bar, 'edit'));
    key(grid, 'ArrowRight');
    expect(trigger(bar, 'view').getAttribute('aria-expanded')).toBe('true');
    expect(label(active(grid))).toBe('Zoom in');
  });

  test('a press on a trigger takes the keyboard from another trigger of the bar', async () => {
    const { bar } = await open();
    const file = trigger(bar, 'file');
    file.focus();
    key(file, 'ArrowDown');
    key(active(bar)!, 'Escape');
    expect(active(bar)).toBe(file);
    fireEvent.click(trigger(bar, 'edit'));
    expect(active(bar)).toBe(trigger(bar, 'edit'));
    key(active(bar)!, 'ArrowDown');
    expect(trigger(bar, 'file').getAttribute('aria-expanded')).toBe('false');
    expect(label(active(bar))).toBe('Undo');
  });

  test('a key with Control, and Tab, stay with the grid', async () => {
    const { grid, bar } = await open();
    fireEvent.click(trigger(bar, 'edit'));
    fireEvent.keyDown(grid, { key: 'ArrowDown', ctrlKey: true });
    key(grid, 'Tab');
    expect(active(grid)).toBe(grid);
  });
});

describe('a toolbar list opened by a pointer press', () => {
  test('a typed digit goes to the font size field, and ArrowDown into the list', async () => {
    const { viewer, view, grid } = await open();
    const size = view.container.querySelector<HTMLElement>('[data-slot="font.size"]')!;
    fireEvent.click(size);
    const typed = fireEvent.keyDown(grid, { key: '1' });
    // The character goes on into the field: its default is not prevented.
    expect(typed).toBe(true);
    expect(active(grid)?.classList.contains('xlsx-toolbar-entry')).toBe(true);
    expect(viewer.snapshot().editing).toBeFalsy();
    fireEvent.click(size);
    fireEvent.click(size);
    const before = viewer.snapshot().selection;
    key(grid, 'ArrowDown');
    expect(viewer.snapshot().selection).toEqual(before);
    expect(active(grid)?.getAttribute('role')).toBe('option');
  });
});

describe('a control that opens rows', () => {
  test('a group menu is disabled with the reason when every row is', async () => {
    const { view } = await open('view');
    const cells = view.container.querySelector<HTMLElement>('[data-group-menu="cells"]')!;
    expect(cells.getAttribute('aria-disabled')).toBe('true');
    expect(cells.getAttribute('title')).toBe(
      'Insert and delete: The workbook is open for viewing.'
    );
    fireEvent.click(cells);
    expect(cells.getAttribute('aria-expanded')).toBe('false');
  });

  test('the Edit > Clear submenu is disabled in view mode, and opens in edit mode', async () => {
    const viewing = await open('view');
    fireEvent.click(trigger(viewing.bar, 'edit'));
    const sub = viewing.bar.querySelector<HTMLElement>('.xlsx-menu-sub > [role="menuitem"]')!;
    expect(sub.getAttribute('aria-disabled')).toBe('true');
    fireEvent.click(sub);
    expect(sub.getAttribute('aria-expanded')).toBe('false');
    cleanup();
    const editing = await open('edit');
    fireEvent.click(trigger(editing.bar, 'edit'));
    const row = editing.bar.querySelector<HTMLElement>('.xlsx-menu-sub > [role="menuitem"]')!;
    expect(row.getAttribute('aria-disabled')).toBeNull();
    fireEvent.click(row);
    expect(row.getAttribute('aria-expanded')).toBe('true');
  });

  test('the delete and clear rows show an icon', async () => {
    const { view } = await open();
    for (const group of ['cells', 'clear']) {
      fireEvent.click(view.container.querySelector(`[data-group-menu="${group}"]`)!);
      const rows = [...view.container.querySelectorAll('.xlsx-toolbar-menu [role="menuitem"]')];
      expect(rows.length).toBeGreaterThan(0);
      for (const row of rows) expect(row.querySelector('svg')).not.toBeNull();
      key(active(view.container)!, 'Escape');
    }
  });
});

describe('the custom hex field', () => {
  test('takes a pasted #RRGGBB whole, and marks text that is no color', async () => {
    const { view } = await open();
    const caret = view.container.querySelector<HTMLElement>(
      '[data-slot="fill.color"] ~ [aria-haspopup]'
    )!;
    fireEvent.click(caret);
    const hex = view.container.querySelector<HTMLInputElement>('.xlsx-toolbar-color-hex')!;
    expect(hex.hasAttribute('maxlength')).toBe(false);
    fireEvent.change(hex, { target: { value: '#00FF00' } });
    expect(hex.value).toBe('00FF00');
    expect(hex.getAttribute('aria-invalid')).toBe('false');
    fireEvent.change(hex, { target: { value: '00F0' } });
    expect(hex.getAttribute('aria-invalid')).toBe('true');
  });
});

describe('the Tab return column', () => {
  test('Enter after typed Tab moves goes down to the column where the run started', async () => {
    const { viewer } = await open();
    act(() => void viewer.select({ row: 24, col: 4 }));
    for (const text of ['a', 'b', 'c']) {
      act(() => void viewer.beginEdit(text));
      act(() => void viewer.commitEdit(text === 'c' ? 'down' : 'right'));
    }
    expect(viewer.snapshot().selection?.active).toEqual({ row: 25, col: 4 });
  });
});
