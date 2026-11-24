// The keys of a chrome popup that a pointer press opened, in happy-dom (`bun run test:dom`). The
// Vue twin of the React adapter's `pointer-keys.test.tsx`: a navigation key from the grid moves
// the keyboard into the open menu or list, a control that only opens rows is disabled in view
// mode, the delete and clear rows show an icon, the custom hex field takes a pasted color, and
// Enter after Tab goes back to the column where the run started.

import { afterEach, describe, expect, test } from 'bun:test';
import { fireEvent, waitFor } from '@testing-library/react';
import { h, nextTick } from 'vue';
import type { EditorMode } from '@xlsx-editor.dev/core/editor';
import { XlsxEditor, type XlsxEditorInstance } from '../src/index.ts';
import { FONTS, mount, unmountAll } from './mount.ts';
import { workbook } from './toolbar-stubs.ts';

afterEach(unmountAll);

async function open(mode: EditorMode = 'edit') {
  let viewer: XlsxEditorInstance | null = null;
  const view = mount(() =>
    h('div', { style: { height: '400px' } }, [
      h(XlsxEditor, {
        document: workbook(),
        fonts: FONTS,
        mode,
        onReady: (v: XlsxEditorInstance) => (viewer = v),
      }),
    ])
  );
  await waitFor(() => expect(viewer?.snapshot().status).toBe('ready'), { timeout: 5000 });
  await nextTick();
  const grid = view.container.querySelector<HTMLElement>('.xlsx-scroller')!;
  viewer!.focus();
  return { view, viewer: viewer!, grid, bar: view.getByRole('menubar') };
}

const trigger = (bar: HTMLElement, id: string) =>
  bar.querySelector<HTMLElement>(`[data-menu="${id}"] > [role="menuitem"]`)!;
const label = (el: Element | null) => el?.querySelector('.xlsx-menu-label')?.textContent ?? null;
const active = (el: Element) => el.ownerDocument.activeElement;

async function click(el: Element) {
  fireEvent.click(el);
  await nextTick();
}

async function key(el: Element, k: string, init: KeyboardEventInit = {}) {
  const result = fireEvent.keyDown(el, { key: k, ...init });
  await nextTick();
  return result;
}

describe('a menu opened by a pointer press', () => {
  test('ArrowDown from the grid moves the keyboard into the menu, not the selection', async () => {
    const { viewer, grid, bar } = await open();
    const before = viewer.snapshot().selection;
    await click(trigger(bar, 'edit'));
    expect(active(grid)).toBe(grid);
    await key(grid, 'ArrowDown');
    expect(viewer.snapshot().selection).toEqual(before);
    expect(label(active(grid))).toBe('Undo');
    // Escape gives the keyboard back to the grid, where it was.
    await key(active(grid)!, 'Escape');
    expect(bar.querySelector('[role="menu"]')).toBeNull();
    expect(active(grid)).toBe(grid);
  });

  test('ArrowUp and a letter go to the last row and the row that starts with it', async () => {
    const { grid, bar } = await open();
    await click(trigger(bar, 'edit'));
    await key(grid, 's');
    expect(label(active(grid))).toBe('Select all');
    await key(active(grid)!, 'Escape');
    await click(trigger(bar, 'edit'));
    await key(grid, 'ArrowUp');
    expect(label(active(grid))).toBe('Select all');
  });

  test('ArrowRight from the grid opens the next menu on its first row', async () => {
    const { grid, bar } = await open();
    await click(trigger(bar, 'edit'));
    await key(grid, 'ArrowRight');
    await nextTick();
    expect(trigger(bar, 'view').getAttribute('aria-expanded')).toBe('true');
    expect(label(active(grid))).toBe('Zoom in');
  });

  test('a press on a trigger takes the keyboard from another trigger of the bar', async () => {
    const { bar } = await open();
    const file = trigger(bar, 'file');
    file.focus();
    await key(file, 'ArrowDown');
    await nextTick();
    await key(active(bar)!, 'Escape');
    expect(active(bar)).toBe(file);
    await click(trigger(bar, 'edit'));
    expect(active(bar)).toBe(trigger(bar, 'edit'));
    await key(active(bar)!, 'ArrowDown');
    expect(trigger(bar, 'file').getAttribute('aria-expanded')).toBe('false');
    expect(label(active(bar))).toBe('Undo');
  });

  test('a key with Control, and Tab, stay with the grid', async () => {
    const { grid, bar } = await open();
    await click(trigger(bar, 'edit'));
    await key(grid, 'ArrowDown', { ctrlKey: true });
    await key(grid, 'Tab');
    expect(active(grid)).toBe(grid);
  });
});

describe('a toolbar list opened by a pointer press', () => {
  test('a typed digit goes to the font size field, and ArrowDown into the list', async () => {
    const { viewer, view, grid } = await open();
    const size = view.container.querySelector<HTMLElement>('[data-slot="font.size"]')!;
    await click(size);
    // The character goes on into the field: its default is not prevented.
    expect(await key(grid, '1')).toBe(true);
    expect(active(grid)?.classList.contains('xlsx-toolbar-entry')).toBe(true);
    expect(viewer.snapshot().editing).toBeFalsy();
    await click(size);
    await click(size);
    const before = viewer.snapshot().selection;
    await key(grid, 'ArrowDown');
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
    await click(cells);
    expect(cells.getAttribute('aria-expanded')).toBe('false');
  });

  test('the Edit > Clear submenu is disabled in view mode, and opens in edit mode', async () => {
    const viewing = await open('view');
    await click(trigger(viewing.bar, 'edit'));
    const sub = viewing.bar.querySelector<HTMLElement>('.xlsx-menu-sub > [role="menuitem"]')!;
    expect(sub.getAttribute('aria-disabled')).toBe('true');
    await click(sub);
    expect(sub.getAttribute('aria-expanded')).toBe('false');
    unmountAll();
    const editing = await open('edit');
    await click(trigger(editing.bar, 'edit'));
    const row = editing.bar.querySelector<HTMLElement>('.xlsx-menu-sub > [role="menuitem"]')!;
    expect(row.getAttribute('aria-disabled')).toBeNull();
    await click(row);
    expect(row.getAttribute('aria-expanded')).toBe('true');
  });

  test('the delete and clear rows show an icon', async () => {
    const { view } = await open();
    for (const group of ['cells', 'clear']) {
      await click(view.container.querySelector(`[data-group-menu="${group}"]`)!);
      const rows = [...view.container.querySelectorAll('.xlsx-toolbar-menu [role="menuitem"]')];
      expect(rows.length).toBeGreaterThan(0);
      for (const row of rows) expect(row.querySelector('svg')).not.toBeNull();
      await key(active(view.container)!, 'Escape');
    }
  });
});

describe('the custom hex field', () => {
  test('takes a pasted #RRGGBB whole, and marks text that is no color', async () => {
    const { view } = await open();
    await click(
      view.container.querySelector<HTMLElement>('[data-slot="fill.color"] ~ [aria-haspopup]')!
    );
    const hex = view.container.querySelector<HTMLInputElement>('.xlsx-toolbar-color-hex')!;
    expect(hex.hasAttribute('maxlength')).toBe(false);
    hex.value = '#00FF00';
    fireEvent.input(hex);
    await nextTick();
    expect(hex.value).toBe('00FF00');
    expect(hex.getAttribute('aria-invalid')).toBe('false');
    hex.value = '00F0';
    fireEvent.input(hex);
    await nextTick();
    expect(hex.getAttribute('aria-invalid')).toBe('true');
  });
});

describe('the Tab return column', () => {
  test('Enter after typed Tab moves goes down to the column where the run started', async () => {
    const { viewer } = await open();
    viewer.select({ row: 24, col: 4 });
    for (const text of ['a', 'b', 'c']) {
      viewer.beginEdit(text);
      viewer.commitEdit(text === 'c' ? 'down' : 'right');
    }
    expect(viewer.snapshot().selection?.active).toEqual({ row: 25, col: 4 });
  });
});
