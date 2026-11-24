// The grid context menu in happy-dom. The Vue twin of the React adapter's
// `context-menu.test.tsx`, with the cases that only a Vue host can write: a row bound with
// `onSelect` (`@select`), a hidden row as a template passes it, and the `openChange` event.

import { afterEach, describe, expect, test } from 'bun:test';
import { fireEvent, waitFor } from '@testing-library/react';
import { defineComponent, h, nextTick, type VNodeChild } from 'vue';
import {
  StatusMessage,
  XlsxEditor,
  XlsxEditorContent,
  XlsxEditorContextMenu,
  XlsxEditorRoot,
  XlsxEditorViewport,
  useContextMenuAnchor,
  useContextMenuTarget,
  type XlsxEditorInstance,
} from '../src/index.ts';
import { FONTS, mount, unmountAll } from './mount.ts';
import { workbook } from './toolbar-stubs.ts';

afterEach(unmountAll);

/** Mount the editor with `menu` as its context menu, and wait for the workbook. */
async function open(menu: () => VNodeChild) {
  let viewer: XlsxEditorInstance | null = null;
  const view = mount(() =>
    h(
      XlsxEditorRoot,
      { document: workbook(), fonts: FONTS, onReady: (v: XlsxEditorInstance) => (viewer = v) },
      {
        default: () =>
          h(XlsxEditorViewport, { style: { height: '400px' } }, () => [
            h(XlsxEditorContent),
            h(StatusMessage),
            menu(),
          ]),
      }
    )
  );
  await waitFor(() => expect(viewer?.snapshot().status).toBe('ready'), { timeout: 5000 });
  await nextTick();
  const scroller = view.container.querySelector<HTMLElement>('.xlsx-scroller')!;
  const shiftF10 = async () => {
    scroller.focus();
    const taken = !fireEvent.keyDown(scroller, { key: 'F10', shiftKey: true });
    await nextTick();
    await nextTick();
    return taken;
  };
  return { view, viewer: viewer!, scroller, shiftF10 };
}

const panel = (root: HTMLElement) =>
  root.querySelector<HTMLElement>('.xlsx-context-menu > [role="menu"]');

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
  test('Shift+F10 opens the cell rows, with the first row focused', async () => {
    const { view, viewer, shiftF10 } = await open(() => h(XlsxEditorContextMenu));
    viewer.select({ row: 0, col: 0 });
    expect(await shiftF10()).toBe(true);
    const menu = panel(view.container)!;
    expect(menu.getAttribute('aria-label')).toBe('Context menu');
    expect(rowsOf(menu).slice(0, 5)).toEqual([
      'edit.cut',
      'edit.copy',
      'edit.paste',
      '|',
      'cells.insertRows',
    ]);
    expect(rowsOf(menu)).toContain('merge.merge');
    expect(document.activeElement?.getAttribute('data-slot')).toBe('edit.cut');
  });

  test('Escape closes the menu and gives the keyboard back to the grid', async () => {
    const changes: boolean[] = [];
    const { view, viewer, scroller, shiftF10 } = await open(() =>
      h(XlsxEditorContextMenu, { onOpenChange: (o: boolean) => changes.push(o) })
    );
    viewer.select({ row: 0, col: 0 });
    await shiftF10();
    fireEvent.keyDown(document.activeElement!, { key: 'Escape' });
    await nextTick();
    expect(panel(view.container)).toBeNull();
    expect(document.activeElement).toBe(scroller);
    expect(changes).toEqual([true, false]);
  });

  test('a row runs its command and closes the menu', async () => {
    const { view, viewer, shiftF10 } = await open(() => h(XlsxEditorContextMenu));
    viewer.select({ row: 0, col: 0 });
    await shiftF10();
    const before = viewer.snapshot().revision;
    fireEvent.click(panel(view.container)!.querySelector('[data-slot="cells.insertRows"]')!);
    await nextTick();
    expect(viewer.snapshot().revision).toBeGreaterThan(before);
    expect(panel(view.container)).toBeNull();
  });

  test('a hidden row goes, Start rows go first, and other rows follow', async () => {
    const { view, viewer, shiftF10 } = await open(() =>
      h(XlsxEditorContextMenu, null, () => [
        h(XlsxEditorContextMenu.Start, null, () => [
          h(XlsxEditorContextMenu.Row, { label: 'First' }),
        ]),
        h(XlsxEditorContextMenu.Paste, { hidden: '' } as never),
        h(XlsxEditorContextMenu.Row, { label: 'Last' }),
      ])
    );
    viewer.select({ row: 0, col: 0 });
    await shiftF10();
    const rows = rowsOf(panel(view.container)!);
    expect(rows.slice(0, 3)).toEqual(['>First', '|', 'edit.cut']);
    expect(rows).not.toContain('edit.paste');
    expect(rows.at(-1)).toBe('>Last');
  });

  test('a row of your own reads the target and the anchor, and @select runs', async () => {
    const seen: string[] = [];
    const Describe = defineComponent({
      setup() {
        const target = useContextMenuTarget();
        const anchor = useContextMenuAnchor();
        return () => {
          const r = target.value?.range;
          const label = `${target.value?.kind} ${r?.start.row},${r?.start.col}:${r?.end.row},${r?.end.col} ${anchor.value?.trigger}`;
          return h(XlsxEditorContextMenu.Row, { label, onSelect: () => seen.push(label) });
        };
      },
    });
    const { view, viewer, shiftF10 } = await open(() =>
      h(XlsxEditorContextMenu, { preset: false }, () => [h(Describe)])
    );
    viewer.select({ row: 1, col: 0 }, { row: 2, col: 1 });
    await shiftF10();
    const row = panel(view.container)!.querySelector<HTMLElement>('[role="menuitem"]')!;
    expect(row.textContent).toBe('cells 1,0:2,1 keyboard');
    fireEvent.click(row);
    await nextTick();
    expect(seen).toEqual(['cells 1,0:2,1 keyboard']);
  });

  test('disabled leaves the key alone', async () => {
    const { view, viewer, shiftF10 } = await open(() =>
      h(XlsxEditorContextMenu, { disabled: true })
    );
    viewer.select({ row: 0, col: 0 });
    expect(await shiftF10()).toBe(false);
    expect(panel(view.container)).toBeNull();
  });

  test('the packaged editor mounts the menu, and contextMenu: false leaves it out', async () => {
    const mounted = mount(() => h(XlsxEditor, { document: workbook(), fonts: FONTS }));
    await nextTick();
    expect(mounted.container.querySelector('.xlsx-context-menu')).not.toBeNull();
    const without = mount(() =>
      h(XlsxEditor, { document: workbook(), fonts: FONTS, contextMenu: false })
    );
    await nextTick();
    expect(without.container.querySelector('.xlsx-context-menu')).toBeNull();
  });
});
