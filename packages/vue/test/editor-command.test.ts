// The public command composable and the view mode in happy-dom, the twin of the React adapter's
// `editor-command.test.tsx`: a host control reads a slot through `useEditorCommand`, whose refs
// change only with the slot's state; the `opening` and `mode` props apply without a reload.
//
// Differences from React: the composable returns refs, so a test reads `.value`; the slot may
// be a ref, so one case changes it. Vue has no StrictMode double mount, so React's StrictMode
// case has no twin.

import { afterEach, describe, expect, test } from 'bun:test';
import { fireEvent, waitFor } from '@testing-library/react';
import { defineComponent, h, nextTick, reactive, shallowRef, watch, type ShallowRef } from 'vue';
import { buildXlsx, c, row } from '../../../scripts/fixtures/builder.ts';
import {
  SheetTabs,
  StatusMessage,
  XlsxEditor,
  XlsxEditorContent,
  XlsxEditorRoot,
  XlsxEditorViewport,
  useEditorCommand,
  useViewModeWaits,
  type ChromeSlotId,
  type EditorMode,
  type UseEditorCommandReturn,
  type XlsxEditorInstance,
} from '../src/index.ts';
import { FONTS, mount, tabsShow, unmountAll } from './mount.ts';

afterEach(unmountAll);

const workbook = buildXlsx({
  sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 'two')]) }],
});

/** The plain state of a command, read from its refs. */
const plain = (command: UseEditorCommandReturn) => ({
  isEnabled: command.isEnabled.value,
  isActive: command.isActive.value,
  disabledReason: command.disabledReason.value,
  hint: command.hint.value,
  value: command.value.value,
});

describe('useEditorCommand', () => {
  async function probe(slot: ChromeSlotId | ShallowRef<ChromeSlotId>, bytes = workbook) {
    let command: UseEditorCommandReturn | null = null;
    let changes = 0;
    const Probe = defineComponent({
      setup() {
        command = useEditorCommand(slot);
        // Each change of any field of the state.
        watch(
          () => plain(command!),
          () => (changes += 1),
          { deep: true }
        );
        return () => null;
      },
    });
    let viewer: XlsxEditorInstance | null = null;
    const view = mount(() =>
      h(
        XlsxEditorRoot,
        { document: bytes, fonts: FONTS, onReady: (v: XlsxEditorInstance) => (viewer = v) },
        {
          default: () =>
            h(XlsxEditorViewport, null, () => [h(SheetTabs), h(Probe), h(StatusMessage)]),
        }
      )
    );
    await tabsShow(view, 1);
    return { command: command!, changes: () => changes, viewer: viewer! };
  }

  test('a toggle slot runs its command and shows it applied', async () => {
    const { command, viewer } = await probe('font.bold');
    expect(plain(command)).toMatchObject({ isEnabled: true, isActive: false, value: null });
    expect(command.disabledReason.value).toBeNull();
    expect(command.execute()).toBe(true);
    expect(viewer.snapshot().activeFormat?.bold).toBe(true);
    expect(command.isActive.value).toBe(true);
  });

  test('the refs keep their values while the slot state is equal', async () => {
    const { command, changes, viewer } = await probe('font.bold');
    const before = changes();
    const reason = command.disabledReason.value;
    viewer.setZoom(1.5);
    viewer.select({ row: 0, col: 1 });
    await nextTick();
    // Neither change moves the state of bold.
    expect(changes()).toBe(before);
    expect(command.disabledReason.value).toBe(reason);
  });

  test('a value slot shows its value and takes a picked value', async () => {
    const { command, viewer } = await probe('font.size');
    expect(command.value.value).toBe('11');
    expect(command.execute('14')).toBe(true);
    expect(viewer.snapshot().activeFormat?.fontSize).toBe(14);
    expect(command.value.value).toBe('14');
    expect(command.execute('not a size')).toBe(false);
  });

  test('a disabled slot gives the translated reason, and execute returns false', async () => {
    const { command } = await probe('history.undo');
    expect(plain(command)).toMatchObject({
      isEnabled: false,
      disabledReason: 'There is nothing to undo.',
    });
    expect(command.execute()).toBe(false);
  });

  test('a slot ref: the state follows the slot', async () => {
    const slot = shallowRef<ChromeSlotId>('history.undo');
    const { command } = await probe(slot);
    expect(command.isEnabled.value).toBe(false);
    slot.value = 'font.bold';
    await nextTick();
    expect(plain(command)).toMatchObject({ isEnabled: true, disabledReason: null });
    expect(command.execute()).toBe(true);
    expect(command.isActive.value).toBe(true);
  });

  test('a true conditional formatting rule gives a hint, and the hint alone changes', async () => {
    const rule =
      '<conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>';
    const bytes = buildXlsx({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 2)]), afterMerges: rule }],
      styles: { dxfs: ['<dxf><font><b/></font></dxf>'] },
    });
    const { command, viewer } = await probe('font.bold', bytes);
    expect(plain(command)).toMatchObject({
      isEnabled: true,
      isActive: false,
      disabledReason: null,
      hint: 'Conditional formatting changes how this cell shows. This control reads and sets the cell format.',
    });
    // B1 is outside the rule: only the hint changes.
    viewer.select({ row: 0, col: 1 });
    expect(plain(command)).toMatchObject({ isEnabled: true, isActive: false, hint: null });
    // In view mode the control is disabled, and a disabled control has no hint.
    viewer.select({ row: 0, col: 0 });
    expect(command.hint.value).not.toBeNull();
    viewer.setMode('view');
    expect(plain(command)).toMatchObject({ isEnabled: false, hint: null });
  });

  test('outside a toolbar, save is disabled: the host saves through the ref', async () => {
    const { command } = await probe('file.save');
    expect(command.disabledReason.value).toBe('Saving is not set up for this editor.');
  });

  test('outside a root, the command is disabled and execute returns false', () => {
    let command: UseEditorCommandReturn | null = null;
    mount(() =>
      h(
        defineComponent({
          setup() {
            command = useEditorCommand('font.bold');
            return () => null;
          },
        })
      )
    );
    expect(command!.execute()).toBe(false);
  });
});

describe('the opening prop', () => {
  test('a press reads the value of the latest render, and the end of the open forgets the refusal', async () => {
    let saves = 0;
    let saveWork = 0;
    let viewer: XlsxEditorInstance | null = null;
    const state = reactive({ opening: true });
    const view = mount(() =>
      h('div', { style: { height: '400px' } }, [
        h(XlsxEditor, {
          document: workbook,
          fonts: FONTS,
          opening: state.opening,
          onSave: () => (saves += 1),
          onReady: (v: XlsxEditorInstance) => {
            viewer = v;
            // Counts the save work that a press starts.
            const saveInTasks = v.saveInTasks.bind(v);
            v.saveInTasks = () => {
              saveWork += 1;
              return saveInTasks();
            };
          },
        }),
      ])
    );
    await tabsShow(view, 1);
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
    state.opening = false;
    await nextTick();
    expect(save().getAttribute('aria-disabled')).toBeNull();
    expect(viewer!.snapshot().refusal).toBeNull();
    press();
    expect(saveWork).toBe(1);
    await waitFor(() => expect(saves).toBe(1), { timeout: 5000 });
    // Another open starts: the next press reads the new value and starts no save work.
    state.opening = true;
    await nextTick();
    press();
    expect(viewer!.snapshot().refusal).toMatchObject({ code: 'loading', command: 'save' });
    expect(saveWork).toBe(1);
    await new Promise((resolve) => setTimeout(resolve, 20));
    expect(saves).toBe(1);
  });

  test('the end of the open keeps a refusal for another reason', async () => {
    let viewer: XlsxEditorInstance | null = null;
    const state = reactive({ opening: true });
    const view = mount(() =>
      h('div', { style: { height: '400px' } }, [
        h(XlsxEditor, {
          document: workbook,
          fonts: FONTS,
          opening: state.opening,
          onReady: (v: XlsxEditorInstance) => (viewer = v),
        }),
      ])
    );
    await tabsShow(view, 1);
    // The workbook has one sheet, so a delete of it is refused.
    viewer!.deleteSheet(0);
    const refusal = viewer!.snapshot().refusal;
    expect(refusal).not.toBeNull();
    expect(refusal!.code).not.toBe('loading');
    state.opening = false;
    await nextTick();
    expect(viewer!.snapshot().refusal).toEqual(refusal);
  });
});

describe('the mode prop', () => {
  test('view disables the edit controls and refuses typing; edit applies without a reload', async () => {
    let ready = 0;
    let viewer: XlsxEditorInstance | null = null;
    const state = reactive<{ mode: EditorMode }>({ mode: 'view' });
    const view = mount(() =>
      h('div', { style: { height: '400px' } }, [
        h(XlsxEditor, {
          document: workbook,
          fonts: FONTS,
          mode: state.mode,
          onSave: () => undefined,
          onReady: (v: XlsxEditorInstance) => {
            ready += 1;
            viewer = v;
          },
        }),
      ])
    );
    await tabsShow(view, 1);
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

    state.mode = 'edit';
    await nextTick();
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
    const state = reactive<{ mode?: EditorMode }>({ mode });
    const view = mount(() =>
      h('div', { style: { height: '400px' } }, [
        h(XlsxEditor, {
          document: workbook,
          fonts: FONTS,
          ...(state.mode ? { mode: state.mode } : {}),
          onReady: (v: XlsxEditorInstance) => (ref.viewer = v),
        }),
      ])
    );
    await tabsShow(view, 1);
    const set = async (m?: EditorMode) => {
      state.mode = m;
      await nextTick();
    };
    return { view, set, viewer: ref.viewer! };
  }

  test('an instance setMode stays while the prop does not change', async () => {
    const { set, viewer } = await mounted();
    expect(viewer.setMode('view')).toBe(true);
    expect(viewer.snapshot().mode).toBe('view');
    // Other updates of the editor state do not apply the prop again.
    viewer.select({ row: 2, col: 2 });
    viewer.setZoom(1.5);
    await set(undefined);
    expect(viewer.snapshot().mode).toBe('view');
    const r = viewer.enterValue('x');
    expect(r.ok ? null : r.refusal.code).toBe('view-mode');
  });

  test('a later prop change wins over an instance call', async () => {
    const { set, viewer } = await mounted('edit');
    viewer.setMode('view');
    await set('view');
    expect(viewer.snapshot().mode).toBe('view');
    await set('edit');
    expect(viewer.snapshot().mode).toBe('edit');
  });

  test('a view prop waits for an open edit that the store refuses, with a notice', async () => {
    const { view, set, viewer } = await mounted('edit');
    viewer.beginEdit('=1+', 'enter');
    await set('view');
    // The switch is refused: the edit stays open, and the strip says that view mode waits.
    expect(viewer.snapshot().mode).toBe('edit');
    expect(viewer.snapshot().editing).toBe(true);
    const waits = 'View mode starts when you commit or cancel the cell edit.';
    await waitFor(() => expect(view.getByText(waits)).toBeTruthy());
    // A cancel closes the edit, and the waiting switch applies.
    viewer.cancelEdit();
    expect(viewer.snapshot().mode).toBe('view');
    await waitFor(() => expect(view.queryByText(waits)).toBeNull());
    // A commit applies a waiting switch too.
    await set('edit');
    expect(viewer.snapshot().mode).toBe('edit');
    viewer.beginEdit('=1+', 'enter');
    await set('view');
    expect(viewer.snapshot().mode).toBe('edit');
    const input = view.container.querySelector('.xlsx-cell-editor') as HTMLTextAreaElement;
    input.value = '7';
    fireEvent.input(input);
    expect(viewer.commitEdit().ok).toBe(true);
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
    const state = reactive<{ mode: EditorMode }>({ mode: 'edit' });
    const view = mount(() =>
      h(
        XlsxEditorRoot,
        {
          document: bytes,
          fonts: FONTS,
          mode: state.mode,
          onReady: (v: XlsxEditorInstance) => (ref.viewer = v),
        },
        {
          default: () =>
            h(XlsxEditorViewport, { style: { height: '400px' } }, () => [
              h(XlsxEditorContent),
              h(StatusMessage),
            ]),
        }
      )
    );
    await waitFor(() => expect(ref.viewer?.snapshot().status).toBe('ready'), { timeout: 5000 });
    expect(ref.viewer!.snapshot().conditionalHidden).toBe(true);
    ref.viewer!.beginEdit('=1+', 'enter');
    state.mode = 'view';
    await nextTick();
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
    let waits: { value: boolean } | null = null;
    const Waits = defineComponent({
      setup() {
        waits = useViewModeWaits();
        return () => null;
      },
    });
    const ref: { viewer: XlsxEditorInstance | null } = { viewer: null };
    const state = reactive<{ mode: EditorMode }>({ mode: 'edit' });
    mount(() =>
      h(
        XlsxEditorRoot,
        {
          document: workbook,
          fonts: FONTS,
          mode: state.mode,
          onReady: (v: XlsxEditorInstance) => (ref.viewer = v),
        },
        {
          default: () =>
            h(XlsxEditorViewport, { style: { height: '400px' } }, () => [
              h(XlsxEditorContent),
              h(StatusMessage),
              h(Waits),
            ]),
        }
      )
    );
    await waitFor(() => expect(ref.viewer?.snapshot().status).toBe('ready'), { timeout: 5000 });
    const viewer = ref.viewer!;
    const set = async (m: EditorMode) => {
      state.mode = m;
      await nextTick();
    };
    viewer.beginEdit('=1+', 'enter');
    await set('view');
    expect(waits!.value).toBe(true);
    // A refused instance call keeps the wait.
    expect(viewer.setMode('view')).toBe(false);
    expect(waits!.value).toBe(true);
    // An instance call that applies, here with the mode that the editor has, ends it at once.
    expect(viewer.setMode('edit')).toBe(true);
    expect(waits!.value).toBe(false);
    viewer.cancelEdit();
    expect(viewer.snapshot().mode).toBe('edit');
    expect(waits!.value).toBe(false);
    expect(viewer.enterValue('5').ok).toBe(true);
    // A new wait ends when the edit closes, and the switch applies.
    await set('edit');
    viewer.beginEdit('=1+', 'enter');
    await set('view');
    expect(waits!.value).toBe(true);
    viewer.cancelEdit();
    expect(waits!.value).toBe(false);
    expect(viewer.snapshot().mode).toBe('view');
  });
});
