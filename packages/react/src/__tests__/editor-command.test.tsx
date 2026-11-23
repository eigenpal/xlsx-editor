/** @jsxImportSource react */
// The public command hook and the view mode in happy-dom (`bun run test:dom`): a host control
// reads a slot through `useEditorCommand`, which keeps its result while the slot's state is
// equal; the `mode` prop switches between edit and view without a reload.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { StrictMode } from 'react';
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react';
import { buildXlsx, c, row } from '../../../../scripts/fixtures/builder.ts';
import {
  XlsxEditor,
  useEditorCommand,
  useViewModeWaits,
  type ChromeSlotId,
  type EditorCommandState,
  type EditorMode,
  type FontSource,
  type XlsxEditorInstance,
} from '../index.ts';

const fontsDir = join(
  dirname(
    Bun.resolveSync('@docx-editor.dev/fonts/package.json', join(import.meta.dir, '../../../core'))
  ),
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

afterEach(cleanup);

const workbook = buildXlsx({
  sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 'two')]) }],
});

describe('useEditorCommand', () => {
  async function probe(slot: ChromeSlotId, bytes = workbook) {
    const seen: EditorCommandState[] = [];
    function Probe() {
      seen.push(useEditorCommand(slot));
      return null;
    }
    let viewer: XlsxEditorInstance | null = null;
    const view = render(
      <XlsxEditor.Root document={bytes} fonts={FONTS} onReady={(v) => (viewer = v)}>
        <XlsxEditor.Viewport>
          <XlsxEditor.SheetTabs />
          <Probe />
          <XlsxEditor.StatusMessage />
        </XlsxEditor.Viewport>
      </XlsxEditor.Root>
    );
    await waitFor(() => expect(view.getAllByRole('tab').length).toBe(1), { timeout: 5000 });
    return { seen, viewer: viewer! };
  }

  test('a toggle slot runs its command and shows it applied', async () => {
    const { seen, viewer } = await probe('font.bold');
    expect(seen.at(-1)).toMatchObject({ isEnabled: true, isActive: false, value: null });
    expect(seen.at(-1)!.disabledReason).toBeNull();
    let ran = false;
    act(() => {
      ran = seen.at(-1)!.execute();
    });
    expect(ran).toBe(true);
    expect(viewer.snapshot().activeFormat?.bold).toBe(true);
    expect(seen.at(-1)!.isActive).toBe(true);
  });

  test('the result keeps its identity while its fields are equal', async () => {
    const { seen, viewer } = await probe('font.bold');
    const before = seen.length;
    const last = seen.at(-1);
    act(() => viewer.setZoom(1.5));
    act(() => viewer.select({ row: 0, col: 1 }));
    // Neither change moves the state of bold: no render, and the same object.
    expect(seen.length).toBe(before);
    expect(seen.at(-1)).toBe(last);
  });

  test('a value slot shows its value and takes a picked value', async () => {
    const { seen, viewer } = await probe('font.size');
    expect(seen.at(-1)!.value).toBe('11');
    act(() => {
      expect(seen.at(-1)!.execute('14')).toBe(true);
    });
    expect(viewer.snapshot().activeFormat?.fontSize).toBe(14);
    expect(seen.at(-1)!.value).toBe('14');
    act(() => {
      expect(seen.at(-1)!.execute('not a size')).toBe(false);
    });
  });

  test('a disabled slot gives the translated reason, and execute returns false', async () => {
    const { seen } = await probe('history.undo');
    expect(seen.at(-1)).toMatchObject({
      isEnabled: false,
      disabledReason: 'There is nothing to undo.',
    });
    act(() => {
      expect(seen.at(-1)!.execute()).toBe(false);
    });
  });

  test('a true conditional formatting rule gives a hint, and the hint alone renders again', async () => {
    const rule =
      '<conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>';
    const bytes = buildXlsx({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 2)]), afterMerges: rule }],
      styles: { dxfs: ['<dxf><font><b/></font></dxf>'] },
    });
    const { seen, viewer } = await probe('font.bold', bytes);
    expect(seen.at(-1)).toMatchObject({
      isEnabled: true,
      isActive: false,
      disabledReason: null,
      hint: 'Conditional formatting changes how this cell shows. This control reads and sets the cell format.',
    });
    // B1 is outside the rule: only the hint changes, and the hook gives a new result.
    act(() => viewer.select({ row: 0, col: 1 }));
    expect(seen.at(-1)).toMatchObject({ isEnabled: true, isActive: false, hint: null });
    // In view mode the control is disabled, and a disabled control has no hint.
    act(() => viewer.select({ row: 0, col: 0 }));
    expect(seen.at(-1)!.hint).not.toBeNull();
    act(() => void viewer.setMode('view'));
    expect(seen.at(-1)).toMatchObject({ isEnabled: false, hint: null });
  });

  test('outside a toolbar, save is disabled: the host saves through the ref', async () => {
    const { seen } = await probe('file.save');
    expect(seen.at(-1)!.disabledReason).toBe('Saving is not set up for this editor.');
  });
});

describe('the opening prop', () => {
  test('a press reads the value of the latest render, and the end of the open forgets the refusal', async () => {
    let saves = 0;
    let saveWork = 0;
    let viewer: XlsxEditorInstance | null = null;
    const tree = (opening: boolean) => (
      <div style={{ height: 400 }}>
        <XlsxEditor
          document={workbook}
          fonts={FONTS}
          opening={opening}
          onSave={() => (saves += 1)}
          onReady={(v) => {
            viewer = v;
            // Counts the save work that a press starts.
            const saveInTasks = v.saveInTasks.bind(v);
            v.saveInTasks = () => {
              saveWork += 1;
              return saveInTasks();
            };
          }}
        />
      </div>
    );
    const view = render(tree(true));
    await waitFor(() => expect(view.getAllByRole('tab').length).toBe(1), { timeout: 5000 });
    const save = () => view.getByRole('button', { name: 'Save' });
    const grid = view.container.querySelector('.xlsx-scroller') as HTMLElement;
    const press = () => fireEvent.keyDown(grid, { key: 's', ctrlKey: true });
    expect(save().getAttribute('aria-disabled')).toBe('true');
    expect(save().title).toBe('Save: A workbook is loading.');
    // Control+S in the grid is refused before any save work.
    press();
    expect(viewer!.snapshot().refusal).toMatchObject({ code: 'loading', command: 'save' });
    expect(saveWork).toBe(0);
    // The open ends: Save is enabled, the refusal is gone, and a press saves.
    view.rerender(tree(false));
    expect(save().getAttribute('aria-disabled')).toBeNull();
    expect(viewer!.snapshot().refusal).toBeNull();
    press();
    expect(saveWork).toBe(1);
    await waitFor(() => expect(saves).toBe(1), { timeout: 5000 });
    // Another open starts: the next press reads the new value and starts no save work.
    view.rerender(tree(true));
    press();
    expect(viewer!.snapshot().refusal).toMatchObject({ code: 'loading', command: 'save' });
    expect(saveWork).toBe(1);
    await new Promise((resolve) => setTimeout(resolve, 20));
    expect(saves).toBe(1);
  });

  test('the end of the open keeps a refusal for another reason', async () => {
    let viewer: XlsxEditorInstance | null = null;
    const tree = (opening: boolean) => (
      <div style={{ height: 400 }}>
        <XlsxEditor
          document={workbook}
          fonts={FONTS}
          opening={opening}
          onReady={(v) => (viewer = v)}
        />
      </div>
    );
    const view = render(tree(true));
    await waitFor(() => expect(view.getAllByRole('tab').length).toBe(1), { timeout: 5000 });
    // The workbook has one sheet, so a delete of it is refused.
    act(() => void viewer!.deleteSheet(0));
    const refusal = viewer!.snapshot().refusal;
    expect(refusal).not.toBeNull();
    expect(refusal!.code).not.toBe('loading');
    view.rerender(tree(false));
    expect(viewer!.snapshot().refusal).toEqual(refusal);
  });
});

describe('the mode prop', () => {
  test('view disables the edit controls and refuses typing; edit applies without a reload', async () => {
    let ready = 0;
    let viewer: XlsxEditorInstance | null = null;
    const tree = (mode: EditorMode) => (
      <div style={{ height: 400 }}>
        <XlsxEditor
          document={workbook}
          fonts={FONTS}
          mode={mode}
          onSave={() => undefined}
          onReady={(v) => {
            ready += 1;
            viewer = v;
          }}
        />
      </div>
    );
    const view = render(tree('view'));
    await waitFor(() => expect(view.getAllByRole('tab').length).toBe(1), { timeout: 5000 });
    expect(viewer!.snapshot().mode).toBe('view');
    const bold = view.getByRole('button', { name: 'Bold' });
    expect(bold.getAttribute('aria-disabled')).toBe('true');
    expect(bold.title).toBe('Bold: The workbook is open for viewing.');
    expect(view.getByRole('button', { name: 'Save' }).getAttribute('aria-disabled')).toBeNull();
    const grid = view.container.querySelector('.xlsx-scroller') as HTMLElement;
    expect(grid.getAttribute('aria-readonly')).toBe('true');
    fireEvent.keyDown(grid, { key: 'x' });
    expect(viewer!.snapshot().editing).toBe(false);
    expect(viewer!.snapshot().refusal?.code).toBe('view-mode');

    view.rerender(tree('edit'));
    expect(viewer!.snapshot().mode).toBe('edit');
    expect(viewer!.snapshot().status).toBe('ready');
    expect(ready).toBe(1);
    expect(view.getByRole('button', { name: 'Bold' }).getAttribute('aria-disabled')).toBeNull();
    fireEvent.keyDown(grid, { key: 'x' });
    expect(viewer!.snapshot().editing).toBe(true);
  });

  /** A tree with an optional `mode` prop, and the editor once it is ready. */
  async function mounted(mode?: EditorMode) {
    const ref: { viewer: XlsxEditorInstance | null } = { viewer: null };
    const tree = (m?: EditorMode) => (
      <div style={{ height: 400 }}>
        <XlsxEditor
          document={workbook}
          fonts={FONTS}
          {...(m ? { mode: m } : {})}
          onReady={(v) => (ref.viewer = v)}
        />
      </div>
    );
    const view = render(tree(mode));
    await waitFor(() => expect(view.getAllByRole('tab').length).toBe(1), { timeout: 5000 });
    return { view, tree, viewer: ref.viewer! };
  }

  test('an instance setMode stays while the prop does not change', async () => {
    const { view, tree, viewer } = await mounted();
    act(() => void expect(viewer.setMode('view')).toBe(true));
    expect(viewer.snapshot().mode).toBe('view');
    // Other updates of the editor state do not apply the prop again.
    act(() => void viewer.select({ row: 2, col: 2 }));
    act(() => void viewer.setZoom(1.5));
    view.rerender(tree());
    expect(viewer.snapshot().mode).toBe('view');
    const r = viewer.enterValue('x');
    expect(r.ok ? null : r.refusal.code).toBe('view-mode');
  });

  test('a later prop change wins over an instance call', async () => {
    const { view, tree, viewer } = await mounted('edit');
    act(() => void viewer.setMode('view'));
    view.rerender(tree('view'));
    expect(viewer.snapshot().mode).toBe('view');
    view.rerender(tree('edit'));
    expect(viewer.snapshot().mode).toBe('edit');
  });

  test('a view prop waits for an open edit that the store refuses, with a notice', async () => {
    const { view, tree, viewer } = await mounted('edit');
    act(() => void viewer.beginEdit('=1+', 'enter'));
    view.rerender(tree('view'));
    // The switch is refused: the edit stays open, and the strip says that view mode waits.
    expect(viewer.snapshot().mode).toBe('edit');
    expect(viewer.snapshot().editing).toBe(true);
    await waitFor(() =>
      expect(
        view.getByText('View mode starts when you commit or cancel the cell edit.')
      ).toBeTruthy()
    );
    // A cancel closes the edit, and the waiting switch applies.
    act(() => viewer.cancelEdit());
    expect(viewer.snapshot().mode).toBe('view');
    await waitFor(() =>
      expect(
        view.queryByText('View mode starts when you commit or cancel the cell edit.')
      ).toBeNull()
    );
    // A commit applies a waiting switch too.
    view.rerender(tree('edit'));
    expect(viewer.snapshot().mode).toBe('edit');
    act(() => void viewer.beginEdit('=1+', 'enter'));
    view.rerender(tree('view'));
    expect(viewer.snapshot().mode).toBe('edit');
    const input = view.container.querySelector('.xlsx-cell-editor') as HTMLTextAreaElement;
    act(() => {
      input.value = '7';
      fireEvent.input(input);
    });
    act(() => void expect(viewer.commitEdit().ok).toBe(true));
    expect(viewer.snapshot().mode).toBe('view');
    expect(viewer.snapshot().activeCell?.content).toBe('7');
  });

  test('with the conditional formatting notice, the waiting view mode notice is short', async () => {
    const rule =
      '<conditionalFormatting sqref="A1"><cfRule type="expression" dxfId="0" priority="1"><formula>1</formula></cfRule></conditionalFormatting>';
    const bytes = buildXlsx({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]), afterMerges: rule }],
      styles: { dxfs: ['<dxf/>'] },
    });
    const ref: { viewer: XlsxEditorInstance | null } = { viewer: null };
    const tree = (m: EditorMode) => (
      <XlsxEditor.Root document={bytes} fonts={FONTS} mode={m} onReady={(v) => (ref.viewer = v)}>
        <XlsxEditor.Viewport style={{ height: 400 }}>
          <XlsxEditor.Content />
          <XlsxEditor.StatusMessage />
        </XlsxEditor.Viewport>
      </XlsxEditor.Root>
    );
    const view = render(tree('edit'));
    await waitFor(() => expect(ref.viewer?.snapshot().status).toBe('ready'), { timeout: 5000 });
    expect(ref.viewer!.snapshot().conditionalHidden).toBe(true);
    act(() => void ref.viewer!.beginEdit('=1+', 'enter'));
    view.rerender(tree('view'));
    const brief = view.getByText('View mode waits for the cell edit.');
    expect(brief.getAttribute('aria-hidden')).toBe('true');
    expect((brief.closest('[role="status"]') as HTMLElement).hasAttribute('data-short')).toBe(true);
    // The conditional formatting notice keeps its full text in its live node.
    expect(
      view
        .getByText(/conditional formatting that the grid can leave out/)
        .closest('[role="status"]')
    ).not.toBeNull();
  });

  test('useViewModeWaits follows the wait after each call', async () => {
    const seen: boolean[] = [];
    function Waits() {
      seen.push(useViewModeWaits());
      return null;
    }
    const ref: { viewer: XlsxEditorInstance | null } = { viewer: null };
    const tree = (m: EditorMode) => (
      <XlsxEditor.Root document={workbook} fonts={FONTS} mode={m} onReady={(v) => (ref.viewer = v)}>
        <XlsxEditor.Viewport style={{ height: 400 }}>
          <XlsxEditor.Content />
          <XlsxEditor.StatusMessage />
          <Waits />
        </XlsxEditor.Viewport>
      </XlsxEditor.Root>
    );
    const view = render(tree('edit'));
    await waitFor(() => expect(ref.viewer?.snapshot().status).toBe('ready'), { timeout: 5000 });
    const viewer = ref.viewer!;
    act(() => void viewer.beginEdit('=1+', 'enter'));
    view.rerender(tree('view'));
    expect(seen.at(-1)).toBe(true);
    // A refused instance call keeps the wait.
    act(() => void expect(viewer.setMode('view')).toBe(false));
    expect(seen.at(-1)).toBe(true);
    // An instance call that applies, here with the mode that the editor has, ends it at once.
    act(() => void expect(viewer.setMode('edit')).toBe(true));
    expect(seen.at(-1)).toBe(false);
    act(() => viewer.cancelEdit());
    expect(viewer.snapshot().mode).toBe('edit');
    expect(seen.at(-1)).toBe(false);
    expect(viewer.enterValue('5').ok).toBe(true);
    // A new wait ends when the edit closes, and the switch applies.
    view.rerender(tree('edit'));
    act(() => void viewer.beginEdit('=1+', 'enter'));
    view.rerender(tree('view'));
    expect(seen.at(-1)).toBe(true);
    act(() => viewer.cancelEdit());
    expect(seen.at(-1)).toBe(false);
    expect(viewer.snapshot().mode).toBe('view');
  });

  test('in StrictMode, the second editor starts with no wait, and a waiting view prop applies', async () => {
    const seen: boolean[] = [];
    function Waits() {
      seen.push(useViewModeWaits());
      return null;
    }
    const ref: { viewer: XlsxEditorInstance | null } = { viewer: null };
    const tree = (m: EditorMode) => (
      <StrictMode>
        <XlsxEditor.Root
          document={workbook}
          fonts={FONTS}
          mode={m}
          onReady={(v) => (ref.viewer = v)}
        >
          <XlsxEditor.Viewport style={{ height: 400 }}>
            <XlsxEditor.Content />
            <XlsxEditor.StatusMessage />
            <Waits />
          </XlsxEditor.Viewport>
        </XlsxEditor.Root>
      </StrictMode>
    );
    const view = render(tree('edit'));
    await waitFor(() => expect(ref.viewer?.snapshot().status).toBe('ready'), { timeout: 5000 });
    // The mount runs the root's effects twice: the editor that stays has no wait.
    expect(seen.at(-1)).toBe(false);
    const viewer = ref.viewer!;
    act(() => void viewer.beginEdit('=1+', 'enter'));
    view.rerender(tree('view'));
    expect(seen.at(-1)).toBe(true);
    act(() => viewer.cancelEdit());
    expect(seen.at(-1)).toBe(false);
    expect(viewer.snapshot().mode).toBe('view');
  });
});
