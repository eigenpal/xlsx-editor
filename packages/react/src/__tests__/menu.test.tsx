/** @jsxImportSource react */
// The menu bar in happy-dom (`bun run test:dom`): the registry menus and rows, each part
// (`XlsxEditor.Menu.Item`, the named rows, `.Row`, `.Group`, `.Separator`, `.Submenu`, `.Entry`,
// `.Menu`), the keys of the WAI-ARIA menubar pattern, the disabled reasons, hide and replace by
// slot, `preset={false}`, the "More" menu, and the `menu` prop of the packaged editor. The Vue
// twin is packages/vue/test/menu.test.ts.

import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react';
import { type ReactNode } from 'react';
import { CHROME_MENUS } from '@xlsx-editor.dev/core/editor';
import { XlsxEditor, type XlsxEditorInstance } from '../index.ts';
import { FONTS, stubLayout, workbook } from './toolbar-stubs.tsx';

afterEach(cleanup);

/** Mount the editor with `menu` as its menu bar, and wait for the workbook. */
async function open(menu: ReactNode) {
  let viewer: XlsxEditorInstance | null = null;
  const document = workbook();
  const tree = (bar: ReactNode) => (
    <XlsxEditor.Root document={document} fonts={FONTS} onReady={(v) => (viewer = v)}>
      <XlsxEditor.Viewport style={{ height: 400 }}>
        {bar}
        <XlsxEditor.Content />
        <XlsxEditor.StatusMessage />
      </XlsxEditor.Viewport>
    </XlsxEditor.Root>
  );
  const view = render(tree(menu));
  await waitFor(() => expect(viewer?.snapshot().status).toBe('ready'), { timeout: 5000 });
  const rerender = (next: ReactNode) => view.rerender(tree(next));
  return { view, bar: view.getByRole('menubar'), viewer: viewer!, rerender };
}

/** The labels of the bar's triggers, in order. */
const triggers = (bar: HTMLElement) =>
  [...bar.querySelectorAll<HTMLElement>(':scope > .xlsx-menu > [role="menuitem"]')].map(
    (b) => b.getAttribute('aria-label') ?? b.textContent
  );

/** The open panel of the bar. */
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

const key = (el: Element, k: string, init: KeyboardEventInit = {}) =>
  fireEvent.keyDown(el, { key: k, ...init });

describe('menu bar', () => {
  test('shows the registry menus in order, as one tab stop', async () => {
    const { bar } = await open(<XlsxEditor.Menu />);
    expect(bar.getAttribute('aria-label')).toBe('Menu bar');
    expect(triggers(bar)).toEqual(['File', 'Edit', 'View', 'Insert', 'Format']);
    const stops = [...bar.querySelectorAll('[role="menuitem"]')].map((b) =>
      b.getAttribute('tabindex')
    );
    expect(stops).toEqual(['0', '-1', '-1', '-1', '-1']);
  });

  test('a press opens a panel with the registry rows, and a second press closes it', async () => {
    const { bar } = await open(<XlsxEditor.Menu />);
    fireEvent.click(trigger(bar, 'edit'));
    const edit = panel(bar)!;
    expect(edit.getAttribute('aria-label')).toBe('Edit');
    expect(trigger(bar, 'edit').getAttribute('aria-expanded')).toBe('true');
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
    fireEvent.click(trigger(bar, 'edit'));
    expect(panel(bar)).toBeNull();
  });

  test('every registry row is a slot whose state comes from the toolbar command state', async () => {
    const { bar } = await open(<XlsxEditor.Menu />);
    for (const menu of CHROME_MENUS) {
      fireEvent.click(trigger(bar, menu.id));
      expect(panel(bar)).not.toBeNull();
      fireEvent.click(trigger(bar, menu.id));
    }
    // Undo has nothing to undo: disabled, focusable, with the engine's reason.
    fireEvent.click(trigger(bar, 'edit'));
    const undo = bar.querySelector<HTMLElement>('[data-slot="history.undo"]')!;
    expect(undo.getAttribute('aria-disabled')).toBe('true');
    expect(undo.getAttribute('title')).toContain('nothing to undo');
    const described = undo.getAttribute('aria-describedby')!;
    expect(bar.ownerDocument.getElementById(described)?.textContent).toContain('nothing to undo');
  });

  test('File > Save is disabled without onSave, and runs the save with it', async () => {
    const first = await open(<XlsxEditor.Menu />);
    fireEvent.click(trigger(first.bar, 'file'));
    const save = first.bar.querySelector<HTMLElement>('[data-slot="file.save"]')!;
    expect(save.getAttribute('aria-disabled')).toBe('true');
    cleanup();
    let saved: Uint8Array | null = null;
    const { bar } = await open(<XlsxEditor.Menu onSave={(bytes) => (saved = bytes)} />);
    fireEvent.click(trigger(bar, 'file'));
    const row = bar.querySelector<HTMLElement>('[data-slot="file.save"]')!;
    expect(row.getAttribute('aria-disabled')).toBeNull();
    expect(row.querySelector('.xlsx-menu-shortcut')?.textContent).toMatch(/^(Ctrl\+S|⌘S)$/);
    fireEvent.click(row);
    expect(panel(bar)).toBeNull();
    await waitFor(() => expect(saved).not.toBeNull());
  });

  test('Control or Command and S saves once with the menu bar and the toolbar both mounted', async () => {
    const saves: Uint8Array[] = [];
    const onSave = (bytes: Uint8Array) => void saves.push(bytes);
    const { view } = await open(
      <>
        <XlsxEditor.Menu onSave={onSave} />
        <XlsxEditor.Toolbar onSave={onSave} />
      </>
    );
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
      const { view } = await open(<XlsxEditor.Menu onSave={() => void (saved += 1)} />);
      const grid = view.container.querySelector<HTMLElement>('.xlsx-scroller')!;
      key(grid, 's', { ctrlKey: true });
      await waitFor(() => expect(saved).toBe(1));
    } finally {
      window.removeEventListener('keydown', block, true);
    }
  });

  test('File > Open runs onOpen, and is enabled without it through the packaged picker', async () => {
    let opened = 0;
    const { bar, view } = await open(<XlsxEditor.Menu onOpen={() => (opened += 1)} />);
    // The host's Open replaces the picker: the bar adds no file input of its own.
    expect(view.container.querySelector('input[type="file"]')).toBeNull();
    // A keyboard open: the keyboard goes back to the File trigger, not to the page body.
    trigger(bar, 'file').focus();
    key(trigger(bar, 'file'), 'ArrowDown');
    fireEvent.click(bar.querySelector<HTMLElement>('[data-slot="file.open"]')!);
    expect(opened).toBe(1);
    expect(bar.ownerDocument.activeElement).toBe(trigger(bar, 'file'));
    cleanup();
    const packaged = await open(<XlsxEditor.Menu />);
    fireEvent.click(trigger(packaged.bar, 'file'));
    const row = packaged.bar.querySelector<HTMLElement>('[data-slot="file.open"]')!;
    expect(row.getAttribute('aria-disabled')).toBeNull();
    const input = packaged.view.container.querySelector<HTMLInputElement>('input[type="file"]')!;
    // The picker lists the workbook and template extensions and their macro-enabled forms.
    expect(input.accept.split(',').slice(0, 4)).toEqual(['.xlsx', '.xltx', '.xlsm', '.xltm']);
    let picked = 0;
    input.addEventListener('click', () => (picked += 1));
    fireEvent.click(row);
    expect(picked).toBe(1);
  });

  test('a picked file that the editor cannot open keeps the workbook and shows a refusal', async () => {
    const { view, viewer } = await open(<XlsxEditor.Menu />);
    act(() => void viewer.beginEdit('kept'));
    act(() => void viewer.commitEdit());
    act(() => void viewer.select({ row: 0, col: 0 }));
    const input = view.container.querySelector<HTMLInputElement>('input[type="file"]')!;
    const file = new File(['not a workbook'], 'broken.xlsx');
    Object.defineProperty(input, 'files', { value: [file], configurable: true });
    fireEvent.change(input);
    await waitFor(() => expect(viewer.snapshot().refusal?.code).toBe('open-failed'));
    expect(viewer.snapshot().status).toBe('ready');
    expect(viewer.snapshot().activeCell?.content).toBe('kept');
    expect(view.container.textContent).toContain('The file was not opened');
  });

  test('after a pointer open a row keeps the keyboard in the grid, after a keyboard open on the trigger', async () => {
    let viewer: XlsxEditorInstance | null = null;
    const { bar, viewer: v } = await open(
      <XlsxEditor.Menu>
        <XlsxEditor.Menu.Edit>
          <XlsxEditor.Menu.Row label="Note" />
          <XlsxEditor.Menu.Row label="Back to the grid" onSelect={() => viewer?.focus()} />
        </XlsxEditor.Menu.Edit>
      </XlsxEditor.Menu>
    );
    viewer = v;
    const doc = bar.ownerDocument;
    const row = (label: string) =>
      [...panel(bar)!.querySelectorAll<HTMLElement>('[role="menuitem"]')].find(
        (r) => r.textContent === label
      )!;
    // A pointer press: the bar prevents the focus move of the mouse down.
    const press = (el: HTMLElement) => {
      fireEvent.pointerDown(el);
      if (fireEvent.mouseDown(el)) el.focus();
      fireEvent.pointerUp(el);
      fireEvent.mouseUp(el);
      fireEvent.click(el);
    };
    v.focus();
    const grid = doc.activeElement!;
    press(trigger(bar, 'edit'));
    expect(panel(bar)).not.toBeNull();
    expect(doc.activeElement).toBe(grid);
    press(row('Note'));
    expect(panel(bar)).toBeNull();
    expect(doc.activeElement).toBe(grid);
    // A keyboard open: F6 to the bar's tab stop (the Edit menu, used last), then ArrowDown.
    fireEvent.keyDown(grid, { key: 'F6' });
    expect(doc.activeElement).toBe(trigger(bar, 'edit'));
    key(doc.activeElement!, 'ArrowDown');
    expect(bar.contains(doc.activeElement)).toBe(true);
    fireEvent.click(row('Note'));
    expect(panel(bar)).toBeNull();
    expect(doc.activeElement).toBe(trigger(bar, 'edit'));
    // A keyboard open whose row moves the keyboard: it stays where the row put it.
    trigger(bar, 'edit').focus();
    key(trigger(bar, 'edit'), 'ArrowDown');
    fireEvent.click(row('Back to the grid'));
    expect(doc.activeElement).toBe(grid);
  });

  test('a row runs its command and gives the keyboard back to the grid', async () => {
    const { bar, viewer } = await open(<XlsxEditor.Menu />);
    fireEvent.click(trigger(bar, 'format'));
    fireEvent.click(bar.querySelector<HTMLElement>('[data-slot="font.bold"]')!);
    expect(viewer.snapshot().activeFormat?.bold).toBe(true);
    expect(panel(bar)).toBeNull();
    fireEvent.click(trigger(bar, 'format'));
    const bold = bar.querySelector<HTMLElement>('[data-slot="font.bold"]')!;
    expect(bold.getAttribute('role')).toBe('menuitemcheckbox');
    expect(bold.getAttribute('aria-checked')).toBe('true');
  });

  test('the alignment rows are a one-of set: menuitemradio, with aria-checked on the one value', async () => {
    const { bar, viewer } = await open(
      <XlsxEditor.Menu>
        <XlsxEditor.Menu.Format>
          <XlsxEditor.Menu.Row label="Compact" selected active />
          <XlsxEditor.Menu.Row label="Wide" selected />
        </XlsxEditor.Menu.Format>
      </XlsxEditor.Menu>
    );
    const alignment = () => {
      fireEvent.click(trigger(bar, 'format'));
      const row = [...panel(bar)!.querySelectorAll<HTMLElement>('[aria-haspopup="menu"]')].find(
        (r) => r.textContent?.includes('Alignment')
      )!;
      fireEvent.click(row);
      return row.parentElement!;
    };
    fireEvent.click(alignment().querySelector<HTMLElement>('[data-slot="alignment.center"]')!);
    expect(viewer.snapshot().activeFormat?.horizontal).toBe('center');
    const sub = alignment();
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

  test('View rows zoom, and Zoom to 100% is disabled at 100%', async () => {
    const { bar, viewer } = await open(<XlsxEditor.Menu />);
    fireEvent.click(trigger(bar, 'view'));
    const reset = bar.querySelector<HTMLElement>('[data-slot="view.zoomReset"]')!;
    expect(reset.getAttribute('aria-disabled')).toBe('true');
    fireEvent.click(bar.querySelector<HTMLElement>('[data-slot="view.zoomIn"]')!);
    expect(viewer.snapshot().zoom).toBeGreaterThan(1);
  });

  test('View > Headings is a checked row that shows the headings, and undo hides them', async () => {
    const { bar, viewer } = await open(<XlsxEditor.Menu />);
    const row = () => bar.querySelector<HTMLElement>('[data-slot="view.headings"]')!;
    fireEvent.click(trigger(bar, 'view'));
    expect(row().getAttribute('role')).toBe('menuitemcheckbox');
    expect(row().getAttribute('aria-checked')).toBe('false');
    fireEvent.click(row());
    expect(viewer.snapshot().headings).toBe(true);
    fireEvent.click(trigger(bar, 'view'));
    await waitFor(() => expect(row().getAttribute('aria-checked')).toBe('true'));
    act(() => void viewer.undo());
    await waitFor(() => expect(row().getAttribute('aria-checked')).toBe('false'));
  });

  test('Insert > Insert sheet adds a sheet', async () => {
    const { bar, viewer } = await open(<XlsxEditor.Menu />);
    fireEvent.click(trigger(bar, 'insert'));
    fireEvent.click(bar.querySelector<HTMLElement>('[data-slot="sheet.add"]')!);
    expect(viewer.snapshot().sheets.length).toBe(2);
  });

  test('Merge cells is disabled on one cell with its reason', async () => {
    const { bar } = await open(<XlsxEditor.Menu />);
    fireEvent.click(trigger(bar, 'format'));
    const merge = bar.querySelector<HTMLElement>('[data-slot="merge.merge"]')!;
    expect(merge.getAttribute('aria-disabled')).toBe('true');
    expect(merge.getAttribute('title')).toContain('Select two or more cells');
  });
});

describe('menu bar rows', () => {
  test('Merge cells over values is disabled with the cell that holds one', async () => {
    const { bar, viewer } = await open(<XlsxEditor.Menu />);
    act(() => void viewer.select({ row: 0, col: 0 }, { row: 0, col: 1 }));
    fireEvent.click(trigger(bar, 'format'));
    const merge = bar.querySelector<HTMLElement>('[data-slot="merge.merge"]')!;
    expect(merge.getAttribute('aria-disabled')).toBe('true');
    expect(merge.getAttribute('title')).toContain('such as B1');
    fireEvent.click(merge);
    expect(viewer.snapshot().selectionMerged).toBe(false);
  });

  test('host children stand beside the menubar element, not in it', async () => {
    const { bar, view } = await open(
      <XlsxEditor.Menu>
        <button type="button">Share</button>
      </XlsxEditor.Menu>
    );
    const share = view.getByText('Share');
    expect(bar.contains(share)).toBe(false);
    expect(share.closest('.xlsx-menubar')).not.toBeNull();
    expect(bar.parentElement).toBe(share.closest('.xlsx-menubar'));
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
      const { view } = await open(
        <>
          <XlsxEditor.Menu />
          <XlsxEditor.Toolbar>
            <XlsxEditor.Toolbar.Group id="clip" label="Clipboard">
              <XlsxEditor.Toolbar.Slot slotId="edit.paste" />
              <XlsxEditor.Toolbar.Slot slotId="file.open" />
            </XlsxEditor.Toolbar.Group>
          </XlsxEditor.Toolbar>
        </>
      );
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
    const { bar } = await open(<XlsxEditor.Menu />);
    const file = trigger(bar, 'file');
    file.focus();
    key(file, 'ArrowRight');
    expect(bar.ownerDocument.activeElement).toBe(trigger(bar, 'edit'));
    expect(trigger(bar, 'edit').getAttribute('tabindex')).toBe('0');
    key(trigger(bar, 'edit'), 'End');
    expect(bar.ownerDocument.activeElement).toBe(trigger(bar, 'format'));
    key(trigger(bar, 'format'), 'ArrowRight');
    expect(bar.ownerDocument.activeElement).toBe(file);
    key(file, 'ArrowLeft');
    expect(bar.ownerDocument.activeElement).toBe(trigger(bar, 'format'));
    key(trigger(bar, 'format'), 'Home');
    expect(bar.ownerDocument.activeElement).toBe(file);
  });

  test('ArrowDown opens on the first row, ArrowUp on the last, Escape returns to the trigger', async () => {
    const { bar } = await open(<XlsxEditor.Menu />);
    const edit = trigger(bar, 'edit');
    edit.focus();
    key(edit, 'ArrowDown');
    const doc = bar.ownerDocument;
    expect((doc.activeElement as HTMLElement).dataset.slot).toBe('history.undo');
    key(doc.activeElement!, 'ArrowDown');
    expect((doc.activeElement as HTMLElement).dataset.slot).toBe('history.redo');
    key(doc.activeElement!, 'End');
    expect((doc.activeElement as HTMLElement).dataset.slot).toBe('edit.selectAll');
    key(doc.activeElement!, 'ArrowDown');
    expect((doc.activeElement as HTMLElement).dataset.slot).toBe('history.undo');
    key(doc.activeElement!, 'Escape');
    expect(panel(bar)).toBeNull();
    expect(doc.activeElement).toBe(edit);
    key(edit, 'ArrowUp');
    expect((doc.activeElement as HTMLElement).dataset.slot).toBe('edit.selectAll');
  });

  test('Escape on a closed trigger gives the keyboard to the grid', async () => {
    const { bar, view } = await open(<XlsxEditor.Menu />);
    const file = trigger(bar, 'file');
    file.focus();
    key(file, 'Escape');
    expect(view.container.querySelector('.xlsx-scroller')!.contains(document.activeElement)).toBe(
      true
    );
  });

  test('type-ahead moves to the next row that starts with the key', async () => {
    const { bar } = await open(<XlsxEditor.Menu />);
    key(trigger(bar, 'edit'), 'Enter');
    const doc = bar.ownerDocument;
    key(doc.activeElement!, 'c');
    expect((doc.activeElement as HTMLElement).dataset.slot).toBe('edit.cut');
    key(doc.activeElement!, 'c');
    expect((doc.activeElement as HTMLElement).dataset.slot).toBe('edit.copy');
  });

  test('ArrowRight in a panel moves to the next menu, and opens a submenu on its row', async () => {
    const { bar } = await open(<XlsxEditor.Menu />);
    trigger(bar, 'file').focus();
    key(trigger(bar, 'file'), 'ArrowDown');
    const doc = bar.ownerDocument;
    key(doc.activeElement!, 'ArrowRight');
    expect(panel(bar)?.getAttribute('aria-label')).toBe('Edit');
    expect((doc.activeElement as HTMLElement).dataset.slot).toBe('history.undo');
    key(doc.activeElement!, 'ArrowLeft');
    expect(panel(bar)?.getAttribute('aria-label')).toBe('File');
    key(doc.activeElement!, 'ArrowRight');
    // The Clear submenu row: ArrowRight opens it on its first row, ArrowLeft closes it.
    const clear = [...panel(bar)!.querySelectorAll<HTMLElement>('[aria-haspopup="menu"]')][0]!;
    clear.focus();
    key(clear, 'ArrowRight');
    expect(clear.getAttribute('aria-expanded')).toBe('true');
    expect((doc.activeElement as HTMLElement).dataset.slot).toBe('clear.contents');
    key(doc.activeElement!, 'ArrowLeft');
    expect(clear.getAttribute('aria-expanded')).toBe('false');
    expect(doc.activeElement).toBe(clear);
    expect(panel(bar)?.getAttribute('aria-label')).toBe('Edit');
  });

  test('F6 moves the keyboard from the grid to the menu bar tab stop', async () => {
    const { bar, viewer } = await open(<XlsxEditor.Menu />);
    viewer.focus();
    const grid = bar.ownerDocument.activeElement!;
    // The regions here are the menu bar and the grid: F6 from the grid wraps to the bar.
    fireEvent.keyDown(grid, { key: 'F6' });
    expect(bar.contains(bar.ownerDocument.activeElement)).toBe(true);
    expect(bar.ownerDocument.activeElement).toBe(trigger(bar, 'file'));
  });

  test('a press on the bar keeps the keyboard where it was', async () => {
    const { bar } = await open(<XlsxEditor.Menu />);
    const press = fireEvent.mouseDown(trigger(bar, 'file'));
    expect(press).toBe(false);
  });
});

describe('menu bar parts', () => {
  test('a named row or Menu.Item hides or replaces a registry row in place', async () => {
    const { bar } = await open(
      <XlsxEditor.Menu>
        <XlsxEditor.Menu.Edit>
          <XlsxEditor.Menu.Paste hidden />
          <XlsxEditor.Menu.Item slotId="edit.copy" label="Copy cells" />
          <XlsxEditor.Menu.Item slotId="clear.formats" hidden />
        </XlsxEditor.Menu.Edit>
      </XlsxEditor.Menu>
    );
    fireEvent.click(trigger(bar, 'edit'));
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
    const { bar } = await open(
      <XlsxEditor.Menu>
        <XlsxEditor.Menu.File>
          <XlsxEditor.Menu.Separator />
          <XlsxEditor.Menu.Group label="Share">
            <XlsxEditor.Menu.Row label="Copy link" shortcut="L" onSelect={() => (selected += 1)} />
            <XlsxEditor.Menu.Row label="Publish" disabled disabledReason="Sign in first." />
          </XlsxEditor.Menu.Group>
          <XlsxEditor.Menu.Submenu label="Recent">
            <XlsxEditor.Menu.Row label="Budget" />
          </XlsxEditor.Menu.Submenu>
        </XlsxEditor.Menu.File>
      </XlsxEditor.Menu>
    );
    fireEvent.click(trigger(bar, 'file'));
    const file = panel(bar)!;
    expect(file.querySelector('[role="separator"]')).not.toBeNull();
    const group = file.querySelector<HTMLElement>('[role="group"]')!;
    expect(group.getAttribute('aria-label')).toBe('Share');
    const publish = [...group.querySelectorAll<HTMLElement>('[role="menuitem"]')][1]!;
    expect(publish.getAttribute('aria-disabled')).toBe('true');
    expect(publish.getAttribute('title')).toBe('Publish: Sign in first.');
    fireEvent.click(publish);
    expect(panel(bar)).not.toBeNull();
    const link = group.querySelector<HTMLElement>('[role="menuitem"]')!;
    expect(link.querySelector('.xlsx-menu-shortcut')?.textContent).toBe('L');
    fireEvent.click(link);
    expect(selected).toBe(1);
    expect(panel(bar)).toBeNull();
    fireEvent.click(trigger(bar, 'file'));
    const recent = panel(bar)!.querySelector<HTMLElement>('[aria-haspopup="menu"]')!;
    fireEvent.click(recent);
    expect(recent.getAttribute('aria-expanded')).toBe('true');
  });

  test('in development, a composition that does nothing warns, and the bar leaves it out', async () => {
    const warnings: string[] = [];
    const warn = console.warn;
    console.warn = (message: string) => warnings.push(String(message));
    try {
      const { bar } = await open(
        <XlsxEditor.Menu>
          <XlsxEditor.Menu.Bold />
          <XlsxEditor.Menu.Menu id="xlsx-menu-more" label="Mine" />
          <XlsxEditor.Menu.Edit>
            <XlsxEditor.Menu.Item slotId="font.italic" hidden />
          </XlsxEditor.Menu.Edit>
        </XlsxEditor.Menu>
      );
      expect(triggers(bar)).toEqual(['File', 'Edit', 'View', 'Insert', 'Format']);
      expect(bar.parentElement!.querySelector('[data-slot="font.bold"]')).toBeNull();
      await waitFor(() => expect(warnings.length).toBe(3));
      expect(warnings.join('\n')).toContain(
        'The XlsxEditor.Menu.Item "font.bold" is a direct child of XlsxEditor.Menu, not of a menu'
      );
      expect(warnings.join('\n')).toContain('A host menu has the id "xlsx-menu-more"');
      expect(warnings.join('\n')).toContain(
        'A hidden XlsxEditor.Menu.Item names "font.italic", but the Edit menu has no row'
      );
    } finally {
      console.warn = warn;
    }
  });

  test('in development, a hidden row in a menu with preset={false} warns, as it removes nothing', async () => {
    const warnings: string[] = [];
    const warn = console.warn;
    console.warn = (message: string) => warnings.push(String(message));
    try {
      const { bar } = await open(
        <XlsxEditor.Menu>
          <XlsxEditor.Menu.View preset={false}>
            <XlsxEditor.Menu.ZoomIn hidden />
            <XlsxEditor.Menu.Row label="Full screen" />
          </XlsxEditor.Menu.View>
        </XlsxEditor.Menu>
      );
      fireEvent.click(trigger(bar, 'view'));
      expect(rowsOf(panel(bar)!)).toEqual(['>Full screen']);
      await waitFor(() =>
        expect(warnings.join('\n')).toContain(
          'A hidden XlsxEditor.Menu.Item names "view.zoomIn", but the View menu has no row'
        )
      );
    } finally {
      console.warn = warn;
    }
  });

  test('an open panel gets an inline place, and a closed one loses it', async () => {
    const { bar } = await open(<XlsxEditor.Menu />);
    fireEvent.click(trigger(bar, 'format'));
    const format = panel(bar)!;
    expect(format.style.left).toMatch(/px$/);
    expect(format.style.top).toMatch(/px$/);
    expect(format.style.maxHeight).toMatch(/px$/);
    const sub = format.querySelector<HTMLElement>('[aria-haspopup="menu"]')!;
    fireEvent.click(sub);
    const subPanel = sub.parentElement!.querySelector<HTMLElement>('.xlsx-menu-panel-sub')!;
    expect(subPanel.style.left).toMatch(/px$/);
    expect(subPanel.style.top).toMatch(/px$/);
    fireEvent.click(trigger(bar, 'format'));
    expect(format.style.left).toBe('');
  });

  test('a menu part hides a menu, and Menu.Menu adds a host menu after another', async () => {
    const { bar } = await open(
      <XlsxEditor.Menu>
        <XlsxEditor.Menu.View hidden />
        <XlsxEditor.Menu.Menu id="share" label="Share" after="edit">
          <XlsxEditor.Menu.Row label="Copy link" />
        </XlsxEditor.Menu.Menu>
      </XlsxEditor.Menu>
    );
    expect(triggers(bar)).toEqual(['File', 'Edit', 'Share', 'Insert', 'Format']);
    fireEvent.click(trigger(bar, 'share'));
    expect(rowsOf(panel(bar)!)).toEqual(['>Copy link']);
  });

  test('preset={false} renders only the children, and a menu with preset={false} its rows', async () => {
    const { bar } = await open(
      <XlsxEditor.Menu preset={false}>
        <XlsxEditor.Menu.Edit preset={false}>
          <XlsxEditor.Menu.Undo />
          <XlsxEditor.Menu.Entry entry={{ kind: 'item', slot: 'edit.copy', shortcut: 'Mod+C' }} />
        </XlsxEditor.Menu.Edit>
      </XlsxEditor.Menu>
    );
    expect(triggers(bar)).toEqual(['Edit']);
    fireEvent.click(trigger(bar, 'edit'));
    expect(rowsOf(panel(bar)!)).toEqual(['history.undo', 'edit.copy']);
  });

  test('a menu that does not fit moves into the More menu as a submenu', async () => {
    const restore = stubLayout((el) => {
      if (el.dataset.menu) return { width: 60 };
      return null;
    }, 0);
    const width = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientWidth')!;
    Object.defineProperty(HTMLElement.prototype, 'clientWidth', {
      configurable: true,
      get(this: HTMLElement) {
        return this.classList.contains('xlsx-menubar') ? 200 : width.get!.call(this);
      },
    });
    try {
      const { bar } = await open(<XlsxEditor.Menu />);
      await waitFor(() => expect(bar.querySelector('[data-menu-more]')).not.toBeNull());
      expect(triggers(bar).length).toBeLessThan(6);
      const more = bar.querySelector<HTMLElement>('[data-menu-more] > [role="menuitem"]')!;
      expect(more.getAttribute('aria-label')).toBe('More menus');
      fireEvent.click(more);
      const subs = panel(bar)!.querySelectorAll('[aria-haspopup="menu"]');
      expect(subs.length).toBeGreaterThan(0);
    } finally {
      Object.defineProperty(HTMLElement.prototype, 'clientWidth', width);
      restore();
    }
  });
});

describe('packaged editor', () => {
  test('shows the menu bar by default, and menu={false} leaves it out', async () => {
    const document = workbook();
    let viewer: XlsxEditorInstance | null = null;
    const view = render(
      <XlsxEditor document={document} fonts={FONTS} onReady={(v) => (viewer = v)} />
    );
    await waitFor(() => expect(viewer?.snapshot().status).toBe('ready'), { timeout: 5000 });
    expect(view.queryByRole('menubar')).not.toBeNull();
    view.rerender(<XlsxEditor document={document} fonts={FONTS} menu={false} />);
    expect(view.queryByRole('menubar')).toBeNull();
  });
});
