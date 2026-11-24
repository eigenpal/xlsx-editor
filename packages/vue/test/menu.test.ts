// The menu bar in happy-dom (`bun run test:dom`). The Vue twin of the React adapter's
// `menu.test.tsx`, with the cases that only a Vue host can write: a part's props as a template
// passes them (`slot-id`, a bare `hidden`), and a row's action bound as `onSelect` (`@select`).

import { afterEach, describe, expect, test } from 'bun:test';
import { fireEvent, waitFor } from '@testing-library/react';
import { h, nextTick, type VNodeChild } from 'vue';
import { CHROME_MENUS } from '@xlsx-editor.dev/core/editor';
import {
  StatusMessage,
  XlsxEditor,
  XlsxEditorContent,
  XlsxEditorMenu,
  XlsxEditorRoot,
  XlsxEditorToolbar,
  XlsxEditorToolbarGroup,
  XlsxEditorToolbarSlot,
  XlsxEditorViewport,
  type XlsxEditorInstance,
} from '../src/index.ts';
import { FONTS, mount, unmountAll } from './mount.ts';
import { stubLayout, workbook } from './toolbar-stubs.ts';

afterEach(unmountAll);

/** Mount the editor with `menu` as its menu bar, and wait for the workbook. */
async function open(menu: () => VNodeChild) {
  let viewer: XlsxEditorInstance | null = null;
  const view = mount(() =>
    h(
      XlsxEditorRoot,
      { document: workbook(), fonts: FONTS, onReady: (v: XlsxEditorInstance) => (viewer = v) },
      {
        default: () =>
          h(XlsxEditorViewport, { style: { height: '400px' } }, () => [
            menu(),
            h(XlsxEditorContent),
            h(StatusMessage),
          ]),
      }
    )
  );
  await waitFor(() => expect(viewer?.snapshot().status).toBe('ready'), { timeout: 5000 });
  await nextTick();
  return { view, bar: view.getByRole('menubar'), viewer: viewer! };
}

const triggers = (bar: HTMLElement) =>
  [...bar.querySelectorAll<HTMLElement>(':scope > .xlsx-menu > [role="menuitem"]')].map(
    (b) => b.getAttribute('aria-label') ?? b.textContent
  );

const panel = (bar: HTMLElement) =>
  bar.querySelector<HTMLElement>(':scope > .xlsx-menu > [role="menu"]');

/** The slots of the rows of a panel, in order; a row with no slot as `>label`. */
const rowsOf = (menu: HTMLElement) =>
  [
    ...menu.querySelectorAll<HTMLElement>(
      ':scope > [role^="menuitem"], :scope > [role="none"] > [role="menuitem"]'
    ),
  ].map((r) => r.dataset.slot ?? `>${r.querySelector('.xlsx-menu-label')?.textContent ?? ''}`);

const trigger = (bar: HTMLElement, id: string) =>
  bar.querySelector<HTMLElement>(`[data-menu="${id}"] > [role="menuitem"]`)!;

async function click(el: Element) {
  fireEvent.click(el);
  await nextTick();
}

async function key(el: Element, k: string) {
  fireEvent.keyDown(el, { key: k });
  await nextTick();
  await nextTick();
}

const focused = () => document.activeElement as HTMLElement;

describe('menu bar', () => {
  test('shows the registry menus in order, as one tab stop', async () => {
    const { bar } = await open(() => h(XlsxEditorMenu));
    expect(bar.getAttribute('aria-label')).toBe('Menu bar');
    expect(triggers(bar)).toEqual(['File', 'Edit', 'View', 'Insert', 'Format']);
    const stops = [...bar.querySelectorAll('[role="menuitem"]')].map((b) =>
      b.getAttribute('tabindex')
    );
    expect(stops).toEqual(['0', '-1', '-1', '-1', '-1']);
  });

  test('a press opens a panel with the registry rows, and a second press closes it', async () => {
    const { bar } = await open(() => h(XlsxEditorMenu));
    await click(trigger(bar, 'edit'));
    const edit = panel(bar)!;
    expect(edit.getAttribute('aria-label')).toBe('Edit');
    expect(rowsOf(edit)).toEqual([
      'history.undo',
      'history.redo',
      'edit.cut',
      'edit.copy',
      'edit.paste',
      '>Clear',
      'cells.deleteRows',
      'cells.deleteColumns',
      'edit.selectAll',
    ]);
    await click(trigger(bar, 'edit'));
    expect(panel(bar)).toBeNull();
  });

  test('every registry menu opens, and a disabled row gives its reason', async () => {
    const { bar } = await open(() => h(XlsxEditorMenu));
    for (const menu of CHROME_MENUS) {
      await click(trigger(bar, menu.id));
      expect(panel(bar)).not.toBeNull();
      await click(trigger(bar, menu.id));
    }
    await click(trigger(bar, 'edit'));
    const undo = bar.querySelector<HTMLElement>('[data-slot="history.undo"]')!;
    expect(undo.getAttribute('aria-disabled')).toBe('true');
    expect(undo.getAttribute('title')).toContain('nothing to undo');
    const described = undo.getAttribute('aria-describedby')!;
    expect(document.getElementById(described)?.textContent).toContain('nothing to undo');
  });

  test('File > Save is disabled without onSave, and runs the save with it', async () => {
    let saved: Uint8Array | null = null;
    const { bar } = await open(() =>
      h(XlsxEditorMenu, { onSave: (bytes: Uint8Array) => (saved = bytes) })
    );
    await click(trigger(bar, 'file'));
    const row = bar.querySelector<HTMLElement>('[data-slot="file.save"]')!;
    expect(row.getAttribute('aria-disabled')).toBeNull();
    expect(row.querySelector('.xlsx-menu-shortcut')?.textContent).toMatch(/^(Ctrl\+S|⌘S)$/);
    await click(row);
    expect(panel(bar)).toBeNull();
    await waitFor(() => expect(saved).not.toBeNull());
  });

  test('Control or Command and S saves once with the menu bar and the toolbar both mounted', async () => {
    const saves: Uint8Array[] = [];
    const onSave = (bytes: Uint8Array) => void saves.push(bytes);
    const { view } = await open(() => [
      h(XlsxEditorMenu, { onSave }),
      h(XlsxEditorToolbar, { onSave }),
    ]);
    const grid = view.container.querySelector<HTMLElement>('.xlsx-scroller')!;
    const event = new KeyboardEvent('keydown', {
      key: 's',
      ctrlKey: true,
      bubbles: true,
      cancelable: true,
    });
    grid.dispatchEvent(event);
    expect(event.defaultPrevented).toBe(true);
    await waitFor(() => expect(saves.length).toBe(1));
    await new Promise((r) => setTimeout(r, 50));
    expect(saves.length).toBe(1);
  });

  test('the save key still saves when the host page prevents the default action first', async () => {
    let saved = 0;
    const block = (e: KeyboardEvent) => {
      if (e.key === 's' && e.ctrlKey) e.preventDefault();
    };
    window.addEventListener('keydown', block, true);
    try {
      const { view } = await open(() => h(XlsxEditorMenu, { onSave: () => void (saved += 1) }));
      const grid = view.container.querySelector<HTMLElement>('.xlsx-scroller')!;
      fireEvent.keyDown(grid, { key: 's', ctrlKey: true });
      await waitFor(() => expect(saved).toBe(1));
    } finally {
      window.removeEventListener('keydown', block, true);
    }
  });

  test('File > Open runs onOpen', async () => {
    let opened = 0;
    const { bar, view } = await open(() => h(XlsxEditorMenu, { onOpen: () => (opened += 1) }));
    // The host's Open replaces the picker: the bar adds no file input of its own.
    expect(view.container.querySelector('input[type="file"]')).toBeNull();
    // A keyboard open: the keyboard goes back to the File trigger, not to the page body.
    trigger(bar, 'file').focus();
    await key(trigger(bar, 'file'), 'ArrowDown');
    await click(bar.querySelector<HTMLElement>('[data-slot="file.open"]')!);
    expect(opened).toBe(1);
    expect(focused()).toBe(trigger(bar, 'file'));
  });

  test('a picked file that the editor cannot open keeps the workbook and shows a refusal', async () => {
    const { view, viewer } = await open(() => h(XlsxEditorMenu));
    viewer.beginEdit('kept');
    viewer.commitEdit();
    viewer.select({ row: 0, col: 0 });
    const input = view.container.querySelector<HTMLInputElement>('input[type="file"]')!;
    // The picker lists the workbook and template extensions and their macro-enabled forms.
    expect(input.accept.split(',').slice(0, 4)).toEqual(['.xlsx', '.xltx', '.xlsm', '.xltm']);
    const file = new File(['not a workbook'], 'broken.xlsx');
    Object.defineProperty(input, 'files', { value: [file], configurable: true });
    fireEvent.change(input);
    await waitFor(() => expect(viewer.snapshot().refusal?.code).toBe('open-failed'));
    expect(viewer.snapshot().status).toBe('ready');
    expect(viewer.snapshot().activeCell?.content).toBe('kept');
    await waitFor(() => expect(view.container.textContent).toContain('The file was not opened'));
  });

  test('after a pointer open a row keeps the keyboard in the grid, after a keyboard open on the trigger', async () => {
    let viewer: XlsxEditorInstance | null = null;
    const { bar, viewer: v } = await open(() =>
      h(XlsxEditorMenu, null, () => [
        h(XlsxEditorMenu.Edit, null, () => [
          h(XlsxEditorMenu.Row, { label: 'Note' }),
          h(XlsxEditorMenu.Row, { label: 'Back to the grid', onSelect: () => viewer?.focus() }),
        ]),
      ])
    );
    viewer = v;
    const row = (label: string) =>
      [...panel(bar)!.querySelectorAll<HTMLElement>('[role="menuitem"]')].find(
        (r) => r.textContent === label
      )!;
    // A pointer press: the bar prevents the focus move of the mouse down.
    const press = async (el: HTMLElement) => {
      fireEvent.pointerDown(el);
      if (fireEvent.mouseDown(el)) el.focus();
      fireEvent.pointerUp(el);
      fireEvent.mouseUp(el);
      await click(el);
    };
    v.focus();
    const grid = focused();
    await press(trigger(bar, 'edit'));
    expect(panel(bar)).not.toBeNull();
    expect(focused()).toBe(grid);
    await press(row('Note'));
    expect(panel(bar)).toBeNull();
    expect(focused()).toBe(grid);
    // A keyboard open: F6 to the bar's tab stop (the Edit menu, used last), then ArrowDown.
    await key(grid, 'F6');
    expect(focused()).toBe(trigger(bar, 'edit'));
    await key(focused(), 'ArrowDown');
    expect(bar.contains(focused())).toBe(true);
    await click(row('Note'));
    expect(panel(bar)).toBeNull();
    expect(focused()).toBe(trigger(bar, 'edit'));
    // A keyboard open whose row moves the keyboard: it stays where the row put it.
    trigger(bar, 'edit').focus();
    await key(trigger(bar, 'edit'), 'ArrowDown');
    await click(row('Back to the grid'));
    expect(focused()).toBe(grid);
  });

  test('the packaged File > Open clicks the file input', async () => {
    const { bar, view } = await open(() => h(XlsxEditorMenu));
    await click(trigger(bar, 'file'));
    const input = view.container.querySelector<HTMLInputElement>('input[type="file"]')!;
    let picked = 0;
    input.addEventListener('click', () => (picked += 1));
    await click(bar.querySelector<HTMLElement>('[data-slot="file.open"]')!);
    expect(picked).toBe(1);
  });

  test('a row runs its command', async () => {
    const { bar, viewer } = await open(() => h(XlsxEditorMenu));
    await click(trigger(bar, 'format'));
    await click(bar.querySelector<HTMLElement>('[data-slot="font.bold"]')!);
    expect(viewer.snapshot().activeFormat?.bold).toBe(true);
    expect(panel(bar)).toBeNull();
    await click(trigger(bar, 'format'));
    const bold = bar.querySelector<HTMLElement>('[data-slot="font.bold"]')!;
    expect(bold.getAttribute('role')).toBe('menuitemcheckbox');
    expect(bold.getAttribute('aria-checked')).toBe('true');
  });

  test('the alignment rows are a one-of set: menuitemradio, with aria-checked on the one value', async () => {
    const { bar, viewer } = await open(() =>
      h(XlsxEditorMenu, null, () => [
        h(XlsxEditorMenu.Format, null, () => [
          h(XlsxEditorMenu.Row, { label: 'Compact', selected: true, active: true }),
          h(XlsxEditorMenu.Row, { label: 'Wide', selected: true }),
        ]),
      ])
    );
    const alignment = async () => {
      await click(trigger(bar, 'format'));
      const row = [...panel(bar)!.querySelectorAll<HTMLElement>('[aria-haspopup="menu"]')].find(
        (r) => r.textContent?.includes('Alignment')
      )!;
      await click(row);
      return row.parentElement!;
    };
    await click((await alignment()).querySelector<HTMLElement>('[data-slot="alignment.center"]')!);
    expect(viewer.snapshot().activeFormat?.horizontal).toBe('center');
    const sub = await alignment();
    const horizontal = ['alignment.left', 'alignment.center', 'alignment.right'].map((slot) =>
      sub.querySelector<HTMLElement>(`[data-slot="${slot}"]`)!
    );
    expect(horizontal.map((r) => r.getAttribute('role'))).toEqual([
      'menuitemradio',
      'menuitemradio',
      'menuitemradio',
    ]);
    expect(horizontal.map((r) => r.getAttribute('aria-checked'))).toEqual([
      'false',
      'true',
      'false',
    ]);
    expect(panel(bar)!.querySelector('[data-slot="alignment.wrap"]')?.getAttribute('role')).toBe(
      'menuitemcheckbox'
    );
    const host = [...panel(bar)!.querySelectorAll<HTMLElement>('[role="menuitemradio"]')].filter(
      (r) => !r.dataset.slot
    );
    expect(host.map((r) => r.getAttribute('aria-checked'))).toEqual(['true', 'false']);
  });

  test('Merge cells is disabled on one cell with its reason', async () => {
    const { bar } = await open(() => h(XlsxEditorMenu));
    await click(trigger(bar, 'format'));
    const merge = bar.querySelector<HTMLElement>('[data-slot="merge.merge"]')!;
    expect(merge.getAttribute('aria-disabled')).toBe('true');
    expect(merge.getAttribute('title')).toContain('Select two or more cells');
  });

  test('F6 moves the keyboard from the grid to the menu bar tab stop', async () => {
    const { bar, viewer } = await open(() => h(XlsxEditorMenu));
    viewer.focus();
    // The regions here are the menu bar and the grid: F6 from the grid wraps to the bar.
    await key(focused(), 'F6');
    expect(focused()).toBe(trigger(bar, 'file'));
  });

  test('View > Headings is a checked row that shows the headings, and undo hides them', async () => {
    const { bar, viewer } = await open(() => h(XlsxEditorMenu));
    const row = () => bar.querySelector<HTMLElement>('[data-slot="view.headings"]')!;
    await click(trigger(bar, 'view'));
    expect(row().getAttribute('role')).toBe('menuitemcheckbox');
    expect(row().getAttribute('aria-checked')).toBe('false');
    await click(row());
    expect(viewer.snapshot().headings).toBe(true);
    await click(trigger(bar, 'view'));
    await waitFor(() => expect(row().getAttribute('aria-checked')).toBe('true'));
    viewer.undo();
    await waitFor(() => expect(row().getAttribute('aria-checked')).toBe('false'));
  });

  test('View rows zoom, and Insert > Insert sheet adds a sheet', async () => {
    const { bar, viewer } = await open(() => h(XlsxEditorMenu));
    await click(trigger(bar, 'view'));
    expect(
      bar.querySelector<HTMLElement>('[data-slot="view.zoomReset"]')!.getAttribute('aria-disabled')
    ).toBe('true');
    await click(bar.querySelector<HTMLElement>('[data-slot="view.zoomIn"]')!);
    expect(viewer.snapshot().zoom).toBeGreaterThan(1);
    await click(trigger(bar, 'insert'));
    await click(bar.querySelector<HTMLElement>('[data-slot="sheet.add"]')!);
    expect(viewer.snapshot().sheets.length).toBe(2);
  });
});

describe('menu-only slots in a toolbar group', () => {
  test('Paste in a host toolbar group reads the clipboard of the editor window, and Open stays disabled', async () => {
    const nav = window.navigator as unknown as Record<string, unknown>;
    const was = Object.getOwnPropertyDescriptor(nav, 'clipboard');
    Object.defineProperty(nav, 'clipboard', {
      configurable: true,
      value: { readText: async () => 'from the clipboard' },
    });
    try {
      const { view } = await open(() => [
        h(XlsxEditorMenu),
        h(XlsxEditorToolbar, null, () => [
          h(XlsxEditorToolbarGroup, { id: 'clip', label: 'Clipboard' }, () => [
            h(XlsxEditorToolbarSlot, { slotId: 'edit.paste' }),
            h(XlsxEditorToolbarSlot, { slotId: 'file.open' }),
          ]),
        ]),
      ]);
      const toolbar = view.getByRole('toolbar');
      const paste = toolbar.querySelector<HTMLElement>('[data-slot="edit.paste"]')!;
      expect(paste.getAttribute('aria-disabled')).not.toBe('true');
      const openRow = toolbar.querySelector<HTMLElement>('[data-slot="file.open"]')!;
      expect(openRow.getAttribute('aria-disabled')).toBe('true');
    } finally {
      if (was) Object.defineProperty(nav, 'clipboard', was);
      else delete nav.clipboard;
    }
  });
});

describe('menu bar keys', () => {
  test('arrows move along the bar, and Home and End go to its ends', async () => {
    const { bar } = await open(() => h(XlsxEditorMenu));
    const file = trigger(bar, 'file');
    file.focus();
    await key(file, 'ArrowRight');
    expect(focused()).toBe(trigger(bar, 'edit'));
    expect(trigger(bar, 'edit').getAttribute('tabindex')).toBe('0');
    await key(trigger(bar, 'edit'), 'End');
    expect(focused()).toBe(trigger(bar, 'format'));
    await key(trigger(bar, 'format'), 'ArrowRight');
    expect(focused()).toBe(file);
    await key(file, 'Home');
    expect(focused()).toBe(file);
  });

  test('ArrowDown opens on the first row, ArrowUp on the last, Escape returns to the trigger', async () => {
    const { bar } = await open(() => h(XlsxEditorMenu));
    const edit = trigger(bar, 'edit');
    edit.focus();
    await key(edit, 'ArrowDown');
    expect(focused().dataset.slot).toBe('history.undo');
    await key(focused(), 'End');
    expect(focused().dataset.slot).toBe('edit.selectAll');
    await key(focused(), 'ArrowDown');
    expect(focused().dataset.slot).toBe('history.undo');
    await key(focused(), 'Escape');
    expect(panel(bar)).toBeNull();
    expect(focused()).toBe(edit);
    await key(edit, 'ArrowUp');
    expect(focused().dataset.slot).toBe('edit.selectAll');
  });

  test('a press on the bar keeps the keyboard where it was', async () => {
    const { bar } = await open(() => h(XlsxEditorMenu));
    const press = fireEvent.mouseDown(trigger(bar, 'file'));
    expect(press).toBe(false);
  });

  test('Escape on a closed trigger gives the keyboard to the grid', async () => {
    const { bar, view } = await open(() => h(XlsxEditorMenu));
    const file = trigger(bar, 'file');
    file.focus();
    await key(file, 'Escape');
    expect(view.container.querySelector('.xlsx-scroller')!.contains(document.activeElement)).toBe(
      true
    );
  });

  test('Merge cells over values is disabled with the cell that holds one', async () => {
    const { bar, viewer } = await open(() => h(XlsxEditorMenu));
    viewer.select({ row: 0, col: 0 }, { row: 0, col: 1 });
    await nextTick();
    await click(trigger(bar, 'format'));
    const merge = bar.querySelector<HTMLElement>('[data-slot="merge.merge"]')!;
    expect(merge.getAttribute('aria-disabled')).toBe('true');
    expect(merge.getAttribute('title')).toContain('such as B1');
    await click(merge);
    expect(viewer.snapshot().selectionMerged).toBe(false);
  });

  test('host children stand beside the menubar element, not in it', async () => {
    const { bar, view } = await open(() =>
      h(XlsxEditorMenu, null, { default: () => [h('button', { type: 'button' }, 'Share')] })
    );
    const share = view.getByText('Share');
    expect(bar.contains(share)).toBe(false);
    expect(bar.parentElement).toBe(share.closest('.xlsx-menubar'));
  });

  test('type-ahead, and ArrowRight to the next menu and into a submenu', async () => {
    const { bar } = await open(() => h(XlsxEditorMenu));
    await key(trigger(bar, 'edit'), 'Enter');
    await key(focused(), 'c');
    expect(focused().dataset.slot).toBe('edit.cut');
    await key(focused(), 'ArrowRight');
    expect(panel(bar)?.getAttribute('aria-label')).toBe('View');
    await key(focused(), 'ArrowLeft');
    expect(panel(bar)?.getAttribute('aria-label')).toBe('Edit');
    const clear = panel(bar)!.querySelector<HTMLElement>('[aria-haspopup="menu"]')!;
    clear.focus();
    await key(clear, 'ArrowRight');
    expect(clear.getAttribute('aria-expanded')).toBe('true');
    expect(focused().dataset.slot).toBe('clear.contents');
    await key(focused(), 'ArrowLeft');
    expect(clear.getAttribute('aria-expanded')).toBe('false');
    expect(focused()).toBe(clear);
  });
});

describe('menu bar parts', () => {
  test('a named row or Menu.Item hides or replaces a registry row, as a template passes it', async () => {
    const { bar } = await open(() =>
      h(XlsxEditorMenu, null, () => [
        h(XlsxEditorMenu.Edit, null, () => [
          // A template's bare attribute arrives as an empty string.
          h(XlsxEditorMenu.Paste, { hidden: '' } as never),
          h(XlsxEditorMenu.Item, { 'slot-id': 'edit.copy', label: 'Copy cells' } as never),
          h(XlsxEditorMenu.Item, { slotId: 'clear.formats', hidden: true }),
        ]),
      ])
    );
    await click(trigger(bar, 'edit'));
    const edit = panel(bar)!;
    expect(rowsOf(edit)).toEqual([
      'history.undo',
      'history.redo',
      'edit.cut',
      'edit.copy',
      '>Clear',
      'cells.deleteRows',
      'cells.deleteColumns',
      'edit.selectAll',
    ]);
    expect(edit.querySelector('[data-slot="edit.copy"] .xlsx-menu-label')?.textContent).toBe(
      'Copy cells'
    );
  });

  test('a host row, group, separator, and submenu follow the registry rows', async () => {
    let selected = 0;
    const { bar } = await open(() =>
      h(XlsxEditorMenu, null, () => [
        h(XlsxEditorMenu.File, null, () => [
          h(XlsxEditorMenu.Separator),
          h(XlsxEditorMenu.Group, { label: 'Share' }, () => [
            h(XlsxEditorMenu.Row, {
              label: 'Copy link',
              shortcut: 'L',
              onSelect: () => (selected += 1),
            }),
            h(XlsxEditorMenu.Row, {
              label: 'Publish',
              disabled: true,
              disabledReason: 'Sign in first.',
            }),
          ]),
          h(XlsxEditorMenu.Submenu, { label: 'Recent' }, () => [
            h(XlsxEditorMenu.Row, { label: 'Budget' }),
          ]),
        ]),
      ])
    );
    await click(trigger(bar, 'file'));
    const file = panel(bar)!;
    expect(file.querySelector('[role="separator"]')).not.toBeNull();
    const group = file.querySelector<HTMLElement>('[role="group"]')!;
    expect(group.getAttribute('aria-label')).toBe('Share');
    const publish = [...group.querySelectorAll<HTMLElement>('[role="menuitem"]')][1]!;
    expect(publish.getAttribute('aria-disabled')).toBe('true');
    expect(publish.getAttribute('title')).toBe('Publish: Sign in first.');
    await click(publish);
    expect(panel(bar)).not.toBeNull();
    const link = group.querySelector<HTMLElement>('[role="menuitem"]')!;
    expect(link.querySelector('.xlsx-menu-shortcut')?.textContent).toBe('L');
    await click(link);
    expect(selected).toBe(1);
    expect(panel(bar)).toBeNull();
    await click(trigger(bar, 'file'));
    const recent = panel(bar)!.querySelector<HTMLElement>('[aria-haspopup="menu"]')!;
    await click(recent);
    expect(recent.getAttribute('aria-expanded')).toBe('true');
  });

  test('in development, a composition that does nothing warns, and the bar leaves it out', async () => {
    const warnings: string[] = [];
    const warn = console.warn;
    console.warn = (message: string) => warnings.push(String(message));
    try {
      const { bar } = await open(() =>
        h(XlsxEditorMenu, null, () => [
          h(XlsxEditorMenu.Italic),
          h(XlsxEditorMenu.Menu, { id: 'xlsx-menu-more', label: 'Mine' }),
          h(XlsxEditorMenu.Edit, null, () => [
            h(XlsxEditorMenu.Item, { slotId: 'font.underline', hidden: true }),
          ]),
        ])
      );
      expect(triggers(bar)).toEqual(['File', 'Edit', 'View', 'Insert', 'Format']);
      expect(bar.parentElement!.querySelector('[data-slot="font.italic"]')).toBeNull();
      await click(trigger(bar, 'edit'));
      await waitFor(() => expect(warnings.length).toBe(3));
      expect(warnings.join('\n')).toContain(
        'The XlsxEditorMenu.Item "font.italic" is a direct child of XlsxEditorMenu, not of a menu'
      );
      expect(warnings.join('\n')).toContain('A host menu has the id "xlsx-menu-more"');
      expect(warnings.join('\n')).toContain(
        'A hidden XlsxEditorMenu.Item names "font.underline", but the Edit menu has no row'
      );
    } finally {
      console.warn = warn;
    }
  });

  test('in development, a hidden row in a menu with preset false warns, as it removes nothing', async () => {
    const warnings: string[] = [];
    const warn = console.warn;
    console.warn = (message: string) => warnings.push(String(message));
    try {
      const { bar } = await open(() =>
        h(XlsxEditorMenu, null, () => [
          h(XlsxEditorMenu.View, { preset: false }, () => [
            h(XlsxEditorMenu.ZoomOut, { hidden: true }),
            h(XlsxEditorMenu.Row, { label: 'Full screen' }),
          ]),
        ])
      );
      await click(trigger(bar, 'view'));
      expect(rowsOf(panel(bar)!)).toEqual(['>Full screen']);
      await waitFor(() =>
        expect(warnings.join('\n')).toContain(
          'A hidden XlsxEditorMenu.Item names "view.zoomOut", but the View menu has no row'
        )
      );
    } finally {
      console.warn = warn;
    }
  });

  test('an open panel gets an inline place, and a closed one loses it', async () => {
    const { bar } = await open(() => h(XlsxEditorMenu));
    await click(trigger(bar, 'format'));
    const format = panel(bar)!;
    expect(format.style.left).toMatch(/px$/);
    expect(format.style.top).toMatch(/px$/);
    expect(format.style.maxHeight).toMatch(/px$/);
    const sub = format.querySelector<HTMLElement>('[aria-haspopup="menu"]')!;
    await click(sub);
    const subPanel = sub.parentElement!.querySelector<HTMLElement>('.xlsx-menu-panel-sub')!;
    expect(subPanel.style.left).toMatch(/px$/);
    expect(subPanel.style.top).toMatch(/px$/);
    await click(trigger(bar, 'format'));
    expect(format.style.left).toBe('');
  });

  test('a menu part hides a menu, and Menu.Menu adds a host menu after another', async () => {
    const { bar } = await open(() =>
      h(XlsxEditorMenu, null, () => [
        h(XlsxEditorMenu.View, { hidden: true }),
        h(XlsxEditorMenu.Menu, { id: 'share', label: 'Share', after: 'edit' }, () => [
          h(XlsxEditorMenu.Row, { label: 'Copy link' }),
        ]),
      ])
    );
    expect(triggers(bar)).toEqual(['File', 'Edit', 'Share', 'Insert', 'Format']);
    await click(trigger(bar, 'share'));
    expect(rowsOf(panel(bar)!)).toEqual(['>Copy link']);
  });

  test('preset false renders only the children, and a menu with preset false its rows', async () => {
    const { bar } = await open(() =>
      h(XlsxEditorMenu, { preset: false }, () => [
        h(XlsxEditorMenu.Edit, { preset: false }, () => [
          h(XlsxEditorMenu.Undo),
          h(XlsxEditorMenu.Entry, {
            entry: { kind: 'item', slot: 'edit.copy', shortcut: 'Mod+C' },
          }),
        ]),
      ])
    );
    expect(triggers(bar)).toEqual(['Edit']);
    await click(trigger(bar, 'edit'));
    expect(rowsOf(panel(bar)!)).toEqual(['history.undo', 'edit.copy']);
  });

  test('a menu that does not fit moves into the More menu as a submenu', async () => {
    const restore = stubLayout((el) => (el.dataset.menu ? { width: 60 } : null), 0);
    const width = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientWidth')!;
    Object.defineProperty(HTMLElement.prototype, 'clientWidth', {
      configurable: true,
      get(this: HTMLElement) {
        return this.classList.contains('xlsx-menubar') ? 200 : width.get!.call(this);
      },
    });
    try {
      const { bar } = await open(() => h(XlsxEditorMenu));
      await waitFor(() => expect(bar.querySelector('[data-menu-more]')).not.toBeNull());
      const more = bar.querySelector<HTMLElement>('[data-menu-more] > [role="menuitem"]')!;
      expect(more.getAttribute('aria-label')).toBe('More menus');
      await click(more);
      expect(panel(bar)!.querySelectorAll('[aria-haspopup="menu"]').length).toBeGreaterThan(0);
    } finally {
      Object.defineProperty(HTMLElement.prototype, 'clientWidth', width);
      restore();
    }
  });
});

describe('packaged editor', () => {
  test('shows the menu bar by default, and menu false leaves it out', async () => {
    let viewer: XlsxEditorInstance | null = null;
    const document = workbook();
    const first = mount(() =>
      h(XlsxEditor, {
        document,
        fonts: FONTS,
        onReady: (v: XlsxEditorInstance) => (viewer = v),
      })
    );
    await waitFor(() => expect(viewer?.snapshot().status).toBe('ready'), { timeout: 5000 });
    await nextTick();
    expect(first.queryByRole('menubar')).not.toBeNull();
    const second = mount(() => h(XlsxEditor, { document, fonts: FONTS, menu: false }));
    await nextTick();
    expect(second.queryByRole('menubar')).toBeNull();
  });
});
