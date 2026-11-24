/** @jsxImportSource react */
// The grid context menu in happy-dom: the default rows, a row that runs its command, Escape,
// hide and replace by slot, `Start`, `preset={false}`, a host row that reads the target, the
// open callback, `disabled`, and the `contextMenu` prop of the packaged editor. The Vue twin is
// packages/vue/test/context-menu.test.ts.

import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react';
import { type ReactNode } from 'react';
import {
  XlsxEditor,
  useContextMenuAnchor,
  useContextMenuTarget,
  type XlsxEditorInstance,
} from '../index.ts';
import { FONTS, workbook } from './toolbar-stubs.tsx';

afterEach(cleanup);

/** Mount the editor with `menu` as its context menu, and wait for the workbook. */
async function open(menu: ReactNode) {
  let viewer: XlsxEditorInstance | null = null;
  const view = render(
    <XlsxEditor.Root document={workbook()} fonts={FONTS} onReady={(v) => (viewer = v)}>
      <XlsxEditor.Viewport style={{ height: 400 }}>
        <XlsxEditor.Content />
        <XlsxEditor.StatusMessage />
        {menu}
      </XlsxEditor.Viewport>
    </XlsxEditor.Root>
  );
  await waitFor(() => expect(viewer?.snapshot().status).toBe('ready'), { timeout: 5000 });
  const scroller = view.container.querySelector<HTMLElement>('.xlsx-scroller')!;
  /** Open the menu with Shift+F10 on the grid; returns whether the key was taken. */
  const shiftF10 = () => {
    scroller.focus();
    let taken = false;
    act(() => {
      taken = !fireEvent.keyDown(scroller, { key: 'F10', shiftKey: true });
    });
    return taken;
  };
  return { view, viewer: viewer!, scroller, shiftF10 };
}

const panel = (root: HTMLElement) =>
  root.querySelector<HTMLElement>('.xlsx-context-menu > [role="menu"]');

/** The slots of the rows, in order; a row with no slot as `>label`, a separator as `|`. */
const rowsOf = (menu: HTMLElement) =>
  [
    ...menu.querySelectorAll<HTMLElement>(
      ':scope > [role^="menuitem"], :scope > [role="separator"]'
    ),
  ].map((r) =>
    r.getAttribute('role') === 'separator'
      ? '|'
      : (r.dataset.slot ?? `>${r.querySelector('.xlsx-menu-label')?.textContent ?? ''}`)
  );

describe('grid context menu', () => {
  test('Shift+F10 opens the cell rows at the active cell, with the first row focused', async () => {
    const { view, viewer, shiftF10 } = await open(<XlsxEditor.ContextMenu />);
    viewer.select({ row: 0, col: 0 });
    expect(shiftF10()).toBe(true);
    const menu = panel(view.container)!;
    expect(menu.getAttribute('aria-label')).toBe('Context menu');
    expect(rowsOf(menu)).toEqual([
      'edit.cut',
      'edit.copy',
      'edit.paste',
      '|',
      'cells.insertRows',
      'cells.insertColumns',
      'cells.deleteRows',
      'cells.deleteColumns',
      '|',
      'clear.contents',
      'clear.formats',
      '|',
      'merge.merge',
      'merge.unmerge',
      '|',
      'number.custom',
    ]);
    expect(document.activeElement?.getAttribute('data-slot')).toBe('edit.cut');
    // A single cell has nothing to merge: the row is disabled with its reason.
    expect(menu.querySelector('[data-slot="merge.merge"]')!.getAttribute('aria-disabled')).toBe(
      'true'
    );
  });

  test('Escape closes the menu and gives the keyboard back to the grid', async () => {
    const { view, viewer, scroller, shiftF10 } = await open(<XlsxEditor.ContextMenu />);
    viewer.select({ row: 0, col: 0 });
    shiftF10();
    act(() => {
      fireEvent.keyDown(document.activeElement!, { key: 'Escape' });
    });
    expect(panel(view.container)).toBeNull();
    expect(document.activeElement).toBe(scroller);
  });

  test('a row runs its command, closes the menu, and gives the keyboard to the grid', async () => {
    const { view, viewer, scroller, shiftF10 } = await open(<XlsxEditor.ContextMenu />);
    viewer.select({ row: 0, col: 0 });
    shiftF10();
    const before = viewer.snapshot().revision;
    act(() => {
      fireEvent.click(panel(view.container)!.querySelector('[data-slot="cells.insertRows"]')!);
    });
    expect(viewer.snapshot().revision).toBeGreaterThan(before);
    expect(panel(view.container)).toBeNull();
    expect(document.activeElement).toBe(scroller);
  });

  test('a hidden named row goes, Start rows go first, and other rows follow', async () => {
    const { view, viewer, shiftF10 } = await open(
      <XlsxEditor.ContextMenu>
        <XlsxEditor.ContextMenu.Start>
          <XlsxEditor.ContextMenu.Row label="First" />
        </XlsxEditor.ContextMenu.Start>
        <XlsxEditor.ContextMenu.Paste hidden />
        <XlsxEditor.ContextMenu.Row label="Last" />
      </XlsxEditor.ContextMenu>
    );
    viewer.select({ row: 0, col: 0 });
    shiftF10();
    const rows = rowsOf(panel(view.container)!);
    expect(rows.slice(0, 4)).toEqual(['>First', '|', 'edit.cut', 'edit.copy']);
    expect(rows).not.toContain('edit.paste');
    expect(rows.at(-1)).toBe('>Last');
  });

  test('preset={false} shows only the children', async () => {
    const { view, viewer, shiftF10 } = await open(
      <XlsxEditor.ContextMenu preset={false}>
        <XlsxEditor.ContextMenu.Copy />
        <XlsxEditor.ContextMenu.Item slotId="merge.merge" label="Join" />
      </XlsxEditor.ContextMenu>
    );
    viewer.select({ row: 0, col: 0 });
    shiftF10();
    const menu = panel(view.container)!;
    expect(rowsOf(menu)).toEqual(['edit.copy', 'merge.merge']);
    expect(menu.querySelector('[data-slot="merge.merge"]')!.textContent).toContain('Join');
  });

  test('a row of your own reads the target and the anchor, and runs after the close', async () => {
    const seen: string[] = [];
    function Describe() {
      const target = useContextMenuTarget();
      const anchor = useContextMenuAnchor();
      const range = target?.range;
      const label = `${target?.kind} ${range?.start.row},${range?.start.col}:${range?.end.row},${range?.end.col} ${anchor?.trigger}`;
      return <XlsxEditor.ContextMenu.Row label={label} onSelect={() => seen.push(label)} />;
    }
    const { view, viewer, shiftF10 } = await open(
      <XlsxEditor.ContextMenu preset={false}>
        <Describe />
      </XlsxEditor.ContextMenu>
    );
    viewer.select({ row: 1, col: 0 }, { row: 2, col: 1 });
    shiftF10();
    const row = panel(view.container)!.querySelector<HTMLElement>('[role="menuitem"]')!;
    expect(row.textContent).toBe('cells 1,0:2,1 keyboard');
    act(() => {
      fireEvent.click(row);
    });
    expect(seen).toEqual(['cells 1,0:2,1 keyboard']);
    expect(panel(view.container)).toBeNull();
  });

  test('onOpenChange reports the open and the close', async () => {
    const changes: boolean[] = [];
    const { view, viewer, shiftF10 } = await open(
      <XlsxEditor.ContextMenu onOpenChange={(o) => changes.push(o)} />
    );
    viewer.select({ row: 0, col: 0 });
    shiftF10();
    act(() => {
      fireEvent.keyDown(panel(view.container)!, { key: 'Escape' });
    });
    expect(changes).toEqual([true, false]);
  });

  test('disabled leaves the key and the browser menu alone', async () => {
    const { view, viewer, shiftF10 } = await open(<XlsxEditor.ContextMenu disabled />);
    viewer.select({ row: 0, col: 0 });
    expect(shiftF10()).toBe(false);
    expect(panel(view.container)).toBeNull();
  });

  test('the packaged editor mounts the menu, and contextMenu={false} leaves it out', async () => {
    const mounted = render(<XlsxEditor document={workbook()} fonts={FONTS} />);
    expect(mounted.container.querySelector('.xlsx-context-menu')).not.toBeNull();
    mounted.unmount();
    const without = render(<XlsxEditor document={workbook()} fonts={FONTS} contextMenu={false} />);
    expect(without.container.querySelector('.xlsx-context-menu')).toBeNull();
  });
});
