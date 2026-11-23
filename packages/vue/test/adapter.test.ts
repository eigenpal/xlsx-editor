// The Vue adapter in happy-dom, the twin of the React adapter's `adapter.test.tsx`: the root
// opens a workbook, the tabs show the visible sheets and switch the grid, the formula bar
// follows the selection, the zoom control zooms, chrome presses keep the keyboard on the grid,
// and the props, emits, and template ref act as React's props, callbacks, and ref do.
//
// Differences from React: the stylesheet rules of the status strip and of the scope class are
// core's and are checked once, in the React test. A render count reads Vue's `updated` hook,
// where React's test reads a `Profiler`.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { fireEvent, waitFor } from '@testing-library/react';
import {
  Teleport,
  defineComponent,
  h,
  nextTick,
  onMounted,
  provide,
  reactive,
  shallowRef,
  type ShallowRef,
} from 'vue';
import { buildXlsx, c, row } from '../../../scripts/fixtures/builder.ts';
import {
  SheetTabs,
  StatusMessage,
  XlsxEditor,
  XlsxEditorContent,
  XlsxEditorRoot,
  XlsxEditorViewport,
  ZoomControl,
  useEditorColorMode,
  useEditorState,
  useScopeClassName,
  xlsxEditorKey,
  type EditorSnapshot,
  type FontSource,
  type NumberLocale,
  type XlsxEditorInstance,
  type XlsxEditorRef,
} from '../src/index.ts';
import { FONTS, field, mount, tabsShow, unmountAll, workbook, type Mounted } from './mount.ts';

afterEach(unmountAll);

async function open(): Promise<{ view: Mounted; viewer: XlsxEditorInstance }> {
  let viewer: XlsxEditorInstance | null = null;
  const view = mount(() =>
    h('div', { style: { height: '400px' } }, [
      h(XlsxEditor, {
        document: workbook,
        fonts: FONTS,
        onReady: (v: XlsxEditorInstance) => (viewer = v),
      }),
    ])
  );
  await tabsShow(view);
  return { view, viewer: viewer! };
}

describe('XlsxEditor', () => {
  test('shows one tab for each visible sheet, in workbook order', async () => {
    const { view } = await open();
    expect(view.getAllByRole('tab').map((t) => t.textContent)).toEqual(['First', 'Second']);
    expect(view.getByRole('region', { name: 'Workbook editor' })).toBeTruthy();
    expect(view.getByRole('tablist', { name: 'Sheets' })).toBeTruthy();
  });

  test('opens a workbook from an ArrayBuffer', async () => {
    const buffer = workbook.buffer.slice(
      workbook.byteOffset,
      workbook.byteOffset + workbook.byteLength
    ) as ArrayBuffer;
    const view = mount(() => h(XlsxEditor, { document: buffer, fonts: FONTS }));
    await tabsShow(view);
  });

  test('a tab switches the sheet; arrow keys move between tabs', async () => {
    const { view, viewer } = await open();
    const [first, second] = view.getAllByRole('tab');
    expect(first!.getAttribute('aria-selected')).toBe('true');
    fireEvent.click(second!);
    expect(viewer.snapshot().activeSheet).toBe(2);
    await nextTick();
    fireEvent.keyDown(view.getAllByRole('tab')[1]!, { key: 'ArrowLeft' });
    expect(viewer.snapshot().activeSheet).toBe(0);
  });

  test('the formula bar shows the active reference and the formula text', async () => {
    const { view, viewer } = await open();
    viewer.select({ row: 4, col: 0 });
    await nextTick();
    expect(field(view, 'Cell reference').value).toBe('A5');
    expect(field(view, 'Cell content').value).toBe('=A2*2');
  });

  test('the formula bar follows the text of an open edit', async () => {
    const { view, viewer } = await open();
    const grid = view.container.querySelector('.xlsx-scroller') as HTMLElement;
    viewer.select({ row: 0, col: 0 });
    await nextTick();
    expect(field(view, 'Cell content').value).toBe('one');
    fireEvent.keyDown(grid, { key: '7' });
    await nextTick();
    const input = view.getByLabelText('Cell editor') as HTMLTextAreaElement;
    expect(field(view, 'Cell content').value).toBe('7');
    input.value = '72';
    fireEvent.input(input);
    await nextTick();
    expect(field(view, 'Cell content').value).toBe('72');
    fireEvent.keyDown(input, { key: 'Escape' });
    await nextTick();
    expect(field(view, 'Cell content').value).toBe('one');
  });

  test('the formula bar is not a live region: the grid announces each move once', async () => {
    const { view } = await open();
    for (const name of ['Cell reference', 'Cell content']) {
      expect(field(view, name).closest('[aria-live]:not([aria-live="off"])')).toBeNull();
      expect(field(view, name).getAttribute('role')).toBeNull();
    }
  });

  test('a value edit in a workbook with a formula shows the stale notice', async () => {
    const { view, viewer } = await open();
    const notice = () => view.container.querySelector('.xlsx-notice');
    // The status node is mounted before the notice, so its text is announced when it arrives.
    expect(notice()?.getAttribute('role')).toBe('status');
    expect(notice()?.textContent).toBe('');
    viewer.enterValue('7');
    await waitFor(() =>
      expect(notice()?.textContent).toContain('Formula results can be out of date')
    );
    expect(notice()?.getAttribute('role')).toBe('status');
    expect(viewer.snapshot().stale).toBe(true);
  });

  test('the status strip keeps its place above the grid, and a refusal shows with a short notice', async () => {
    const { view, viewer } = await open();
    const grid = view.container.querySelector('.xlsx-canvas') as HTMLElement;
    const strip = view.container.querySelector('.xlsx-status-strip') as HTMLElement;
    expect(strip.textContent).toBe('');
    expect(grid.previousElementSibling).toBe(strip);
    viewer.enterValue('7');
    viewer.beginEdit('=SUM(', 'enter');
    viewer.commitEdit();
    await waitFor(() => expect(strip.querySelector('.xlsx-refusal')).not.toBeNull());
    const notice = strip.querySelector('.xlsx-notice') as HTMLElement;
    expect(notice.hasAttribute('data-short')).toBe(true);
    expect(notice.querySelector('.xlsx-notice-brief')!.textContent).toBe(
      'Formula results can be out of date.'
    );
    expect(notice.querySelector('.xlsx-notice-brief')!.getAttribute('aria-hidden')).toBe('true');
    expect(notice.querySelector('.xlsx-notice-full')!.textContent).toContain(
      'the editor shows the stored results'
    );
    expect(notice.title).toContain('the editor shows the stored results');
    expect(view.container.querySelector('.xlsx-status-strip')).toBe(strip);
    expect(grid.previousElementSibling).toBe(strip);
  });

  test('F6 and Shift+F6 move the keyboard among the toolbar, the formula bar, the grid, and the footer', async () => {
    const { view } = await open();
    const nameBox = view.getByLabelText('Cell reference');
    const grid = view.container.querySelector('.xlsx-scroller') as HTMLElement;
    const toolbar = view.getByRole('toolbar');
    const footer = view.container.querySelector('.xlsx-footer') as HTMLElement;
    grid.focus();
    const f6 = (shiftKey = false) =>
      fireEvent.keyDown(view.container.ownerDocument.activeElement!, { key: 'F6', shiftKey });
    f6();
    expect(footer.contains(document.activeElement)).toBe(true);
    // The footer takes the keyboard on the selected sheet tab.
    expect(document.activeElement?.getAttribute('aria-selected')).toBe('true');
    f6();
    expect(toolbar.contains(document.activeElement)).toBe(true);
    // The formula bar takes the keyboard on the name box.
    f6();
    expect(document.activeElement).toBe(nameBox);
    f6();
    expect(document.activeElement).toBe(grid);
    f6(true);
    expect(document.activeElement).toBe(nameBox);
    f6(true);
    expect(toolbar.contains(document.activeElement)).toBe(true);
    f6(true);
    expect(footer.contains(document.activeElement)).toBe(true);
  });

  test('F6 in the cell editor commits the edit and moves to the next region', async () => {
    const { view, viewer } = await open();
    const footer = view.container.querySelector('.xlsx-footer') as HTMLElement;
    viewer.select({ row: 1, col: 1 });
    viewer.beginEdit('7');
    await nextTick();
    const input = view.getByLabelText('Cell content') as HTMLTextAreaElement;
    const editor = view.container.querySelector('.xlsx-scroller textarea, .xlsx-scroller input');
    expect(editor).not.toBeNull();
    (editor as HTMLElement).focus();
    fireEvent.keyDown(editor!, { key: 'F6' });
    expect(viewer.snapshot().editing).toBe(false);
    expect(footer.contains(document.activeElement)).toBe(true);
    viewer.select({ row: 1, col: 1 });
    await nextTick();
    expect(input.value).toBe('7');
  });

  test('the zoom control zooms and resets', async () => {
    const { view, viewer } = await open();
    expect(view.getByRole('group', { name: 'Zoom level' })).toBeTruthy();
    fireEvent.click(view.getByRole('button', { name: 'Zoom in' }));
    expect(viewer.snapshot().zoom).toBe(1.1);
    await nextTick();
    view.getByRole('button', { name: 'Reset the zoom to 100%. The zoom is 110%.' });
    fireEvent.click(view.getByRole('button', { name: /Reset the zoom/ }));
    expect(viewer.snapshot().zoom).toBe(1);
  });

  test('a zoom button at its limit keeps the focus and does nothing', async () => {
    const { view, viewer } = await open();
    viewer.setZoom(4);
    await nextTick();
    const zoomIn = view.getByRole('button', { name: 'Zoom in' });
    zoomIn.focus();
    expect(zoomIn.getAttribute('aria-disabled')).toBe('true');
    expect(zoomIn.hasAttribute('disabled')).toBe(false);
    fireEvent.click(zoomIn);
    expect(viewer.snapshot().zoom).toBe(4);
    expect(document.activeElement).toBe(zoomIn);
    fireEvent.click(view.getByRole('button', { name: 'Zoom out' }));
    await nextTick();
    expect(zoomIn.getAttribute('aria-disabled')).toBe('false');
  });

  test('pressing chrome does not take the focus from the grid', async () => {
    const { view } = await open();
    const press = fireEvent.mouseDown(view.getAllByRole('tab')[0]!);
    expect(press).toBe(false);
  });

  test('an edit from the grid commits through the engine, and a refusal shows once', async () => {
    const { view, viewer } = await open();
    const grid = view.container.querySelector('.xlsx-scroller') as HTMLElement;
    viewer.select({ row: 1, col: 1 });
    fireEvent.keyDown(grid, { key: '5' });
    const input = view.getByLabelText('Cell editor') as HTMLTextAreaElement;
    fireEvent.keyDown(input, { key: 'Enter' });
    viewer.select({ row: 1, col: 1 });
    await nextTick();
    expect(field(view, 'Cell content').value).toBe('5');
    fireEvent.keyDown(grid, { key: '=' });
    input.value = '=SUM(';
    fireEvent.keyDown(input, { key: 'Enter' });
    const refusal =
      'The formula was not applied: it cannot be read. Check the quotes, parentheses, and references.';
    await waitFor(() =>
      expect(view.container.querySelector('.xlsx-refusal')?.textContent).toBe(refusal)
    );
    // One announcement: the engine's live region. The shown refusal has no live role.
    expect(view.queryByRole('alert')).toBeNull();
    expect(view.container.querySelector('.xlsx-live')?.textContent).toBe(refusal);
    fireEvent.keyDown(input, { key: 'Escape' });
    expect(viewer.snapshot().refusal).toBeNull();
  });

  test('the template ref is small: it saves, reads the state, and hands out the editor', async () => {
    for (const component of [XlsxEditor, XlsxEditorRoot]) {
      const handle = shallowRef<XlsxEditorRef | null>(null);
      let viewer: XlsxEditorInstance | null = null;
      const view = mount(() =>
        h(
          component,
          {
            ref: handle,
            document: workbook,
            fonts: FONTS,
            onReady: (v: XlsxEditorInstance) => (viewer = v),
          },
          { default: () => h(XlsxEditorViewport, null, () => [h(SheetTabs)]) }
        )
      );
      await tabsShow(view);
      viewer!.enterValue('edited');
      expect(handle.value!.save()!.byteLength).toBeGreaterThan(0);
      expect(handle.value!.getEditor()).toBe(viewer);
      expect(handle.value!.snapshot()).toBe(viewer!.snapshot());
      expect((await handle.value!.saveInTasks())!.bytes.byteLength).toBeGreaterThan(0);
      expect(
        Object.keys(handle.value!)
          .filter((k) => !k.startsWith('$') && !k.startsWith('_'))
          .sort()
      ).toEqual(['focus', 'getEditor', 'load', 'save', 'saveInTasks', 'snapshot']);
      expect(handle.value!.getEditor()!.undo().ok).toBe(true);
      expect(viewer!.snapshot().canUndo).toBe(false);
      const loaded = await handle.value!.load(
        buildXlsx({ sheets: [{ name: 'Only', rows: row(1, [c('A1', 1)]) }] })
      );
      expect(loaded?.sheets.map((s) => s.name)).toEqual(['Only']);
      view.unmount();
    }
  });

  test('before the editor mounts, the state and the handle read the mode prop', () => {
    const modes: string[] = [];
    const handleModes: string[] = [];
    const handle = shallowRef<XlsxEditorRef | null>(null);
    const Probe = defineComponent({
      setup() {
        const mode = useEditorState((s) => s.mode);
        modes.push(mode.value);
        // A child mounts before its root creates the editor.
        onMounted(() => {
          handleModes.push(handle.value!.snapshot().mode);
          expect(handle.value!.getEditor()).toBeNull();
        });
        return () => null;
      },
    });
    mount(() =>
      h(XlsxEditorRoot, { ref: handle, fonts: FONTS, mode: 'view' }, { default: () => h(Probe) })
    );
    expect(modes[0]).toBe('view');
    expect(handleModes).toEqual(['view']);
    expect(handle.value!.getEditor()!.snapshot().mode).toBe('view');
  });

  test('in development, a layout with the grid and no StatusMessage gets one warning', async () => {
    const warnings: string[] = [];
    const warn = console.warn;
    console.warn = (message: string) => warnings.push(message);
    try {
      const bare = mount(() =>
        h(XlsxEditorRoot, { fonts: FONTS }, { default: () => h(XlsxEditorContent) })
      );
      await waitFor(() => expect(warnings.length).toBe(1));
      expect(warnings[0]).toContain('StatusMessage');
      bare.unmount();
      warnings.length = 0;
      mount(() =>
        h(
          XlsxEditorRoot,
          { fonts: FONTS },
          { default: () => [h(XlsxEditorContent), h(StatusMessage)] }
        )
      );
      await nextTick();
      expect(warnings).toEqual([]);
    } finally {
      console.warn = warn;
    }
  });

  test('useEditorState renders again only when its slice changes', async () => {
    let renders = 0;
    let seen = -2;
    const pick = (s: EditorSnapshot) => s.activeSheet;
    const Probe = defineComponent({
      setup() {
        const sheet = useEditorState(pick);
        return () => {
          seen = sheet.value;
          renders += 1;
          return null;
        };
      },
    });
    let viewer: XlsxEditorInstance | null = null;
    const view = mount(() =>
      h(
        XlsxEditorRoot,
        { document: workbook, fonts: FONTS, onReady: (v: XlsxEditorInstance) => (viewer = v) },
        { default: () => h(XlsxEditorViewport, null, () => [h(SheetTabs), h(Probe)]) }
      )
    );
    await tabsShow(view);
    expect(seen).toBe(0);
    const before = renders;
    viewer!.setZoom(1.5);
    viewer!.select({ row: 3, col: 3 });
    await nextTick();
    expect(renders).toBe(before);
    viewer!.setActiveSheet(2);
    await nextTick();
    expect(seen).toBe(2);
    expect(renders).toBe(before + 1);
  });

  test('useEditorState reads no stale snapshot and survives key repeat with no warning', async () => {
    const seen: number[] = [];
    const pick = (s: EditorSnapshot) => s.selection?.active.row ?? -1;
    let state: ShallowRef<number> | null = null;
    const Probe = defineComponent({
      setup() {
        const active = useEditorState(pick);
        state = active as ShallowRef<number>;
        return () => {
          seen.push(active.value);
          return null;
        };
      },
    });
    let viewer: XlsxEditorInstance | null = null;
    const view = mount(() =>
      h(
        XlsxEditorRoot,
        { document: workbook, fonts: FONTS, onReady: (v: XlsxEditorInstance) => (viewer = v) },
        {
          default: () =>
            h(XlsxEditorViewport, null, () => [h(XlsxEditorContent), h(SheetTabs), h(Probe)]),
        }
      )
    );
    await tabsShow(view);
    const rows: number[] = [];
    const reads: number[] = [];
    const off = viewer!.subscribe(() => {
      rows.push(viewer!.snapshot().selection!.active.row);
      // The composable's ref is current inside the editor's own notification.
      reads.push(state!.value);
    });
    const grid = view.container.querySelector('.xlsx-scroller') as HTMLElement;
    const messages: unknown[] = [];
    const error = console.error;
    const warn = console.warn;
    console.error = (...args: unknown[]) => void messages.push(args);
    console.warn = (...args: unknown[]) => void messages.push(args);
    try {
      for (let i = 0; i < 500; i += 1) fireEvent.keyDown(grid, { key: 'ArrowDown' });
      await nextTick();
    } finally {
      console.error = error;
      console.warn = warn;
      off();
    }
    expect(messages).toEqual([]);
    expect(rows.length).toBeGreaterThanOrEqual(500);
    expect(rows.every((r, i) => i === 0 || r >= rows[i - 1]!)).toBe(true);
    expect(reads).toEqual(rows);
    expect(seen.at(-1)).toBe(viewer!.snapshot().selection!.active.row);
    expect(seen.at(-1)).toBe(500);
  });

  test('the chrome parts render only when their own slice changes', async () => {
    // Every component's own render after its first, by component name.
    const updates: Record<string, number> = {};
    let viewer: XlsxEditorInstance | null = null;
    const view = mount(
      () =>
        h(
          XlsxEditorRoot,
          { document: workbook, fonts: FONTS, onReady: (v: XlsxEditorInstance) => (viewer = v) },
          {
            default: () =>
              h(XlsxEditorViewport, null, () => [h(ZoomControl), h(SheetTabs), h(StatusMessage)]),
          }
        ),
      (app) =>
        app.mixin({
          updated(this: { $options: { name?: string } }) {
            const name = this.$options.name;
            if (name) updates[name] = (updates[name] ?? 0) + 1;
          },
        })
    );
    await tabsShow(view);
    await nextTick();
    const parts = () =>
      Object.fromEntries(
        Object.entries(updates).filter(([name]) =>
          ['ZoomControl', 'SheetTabs', 'StatusMessage'].includes(name)
        )
      );
    const before = parts();
    viewer!.select({ row: 3, col: 3 });
    await nextTick();
    expect(parts()).toEqual(before);
    viewer!.setZoom(1.5);
    await nextTick();
    expect(parts()).toEqual({ ...before, ZoomControl: (before.ZoomControl ?? 0) + 1 });
  });

  test('useEditorState reads the new slice when its deps change', async () => {
    const seen: unknown[] = [];
    const pickers: Record<string, (s: EditorSnapshot) => unknown> = {
      status: (s) => s.status,
      sheet: (s) => s.activeSheet,
    };
    const state = reactive({ pick: 'status' });
    const Probe = defineComponent({
      setup() {
        const slice = useEditorState(
          (s) => pickers[state.pick]!(s),
          Object.is,
          () => state.pick
        );
        return () => {
          seen.push(slice.value);
          return null;
        };
      },
    });
    const view = mount(() =>
      h(
        XlsxEditorRoot,
        { document: workbook, fonts: FONTS },
        { default: () => h(XlsxEditorViewport, null, () => [h(SheetTabs), h(Probe)]) }
      )
    );
    await tabsShow(view);
    expect(seen.at(-1)).toBe('ready');
    state.pick = 'sheet';
    await nextTick();
    expect(seen.at(-1)).toBe(0);
  });

  test('the arrow keys pass over a chart sheet tab', async () => {
    const { viewer } = await open();
    const sheets = [
      { index: 0, name: 'One', visible: true, worksheet: true },
      { index: 1, name: 'Chart', visible: true, worksheet: false },
      { index: 2, name: 'Two', visible: true, worksheet: true },
    ];
    let active = 0;
    const listeners = new Set<() => void>();
    let snap = { ...viewer.snapshot(), sheets, activeSheet: active };
    let refuse = false;
    const fake = {
      ...viewer,
      snapshot: () => snap,
      subscribe: (f: () => void) => (listeners.add(f), () => void listeners.delete(f)),
      setActiveSheet: (i: number) => {
        if (!sheets[i]!.worksheet || refuse) return false;
        active = i;
        snap = { ...snap, activeSheet: active };
        for (const f of listeners) f();
        return true;
      },
    } as unknown as XlsxEditorInstance;
    const Provider = defineComponent({
      setup(_, { slots }) {
        provide(xlsxEditorKey, shallowRef(fake));
        return () => slots.default?.();
      },
    });
    const view = mount(() =>
      h(Provider, null, { default: () => h('div', { 'data-probe': 'tabs' }, [h(SheetTabs)]) })
    );
    const tabs = () =>
      view.container.querySelectorAll<HTMLButtonElement>('[data-probe="tabs"] [role="tab"]');
    expect(tabs().length).toBe(3);
    fireEvent.keyDown(tabs()[0]!, { key: 'ArrowRight' });
    expect(active).toBe(2);
    await nextTick();
    fireEvent.keyDown(tabs()[2]!, { key: 'ArrowLeft' });
    expect(active).toBe(0);
    await nextTick();
    // A refused sheet change keeps the focus on the selected tab.
    tabs()[0]!.focus();
    refuse = true;
    fireEvent.keyDown(tabs()[0]!, { key: 'ArrowRight' });
    await nextTick();
    expect(active).toBe(0);
    expect(document.activeElement).toBe(tabs()[0]!);
  });

  test('a file that is not a workbook shows an alert', async () => {
    const view = mount(() => h(XlsxEditor, { document: new Uint8Array([1, 2, 3]), fonts: FONTS }));
    await waitFor(() => expect(view.getByRole('alert').textContent).toContain('not a workbook'), {
      timeout: 5000,
    });
  });

  test('a font load failure emits fontError once and shows the fonts alert', async () => {
    const errors: string[] = [];
    const view = mount(() =>
      h(XlsxEditor, {
        document: workbook,
        fonts: () => Promise.reject(new Error('offline')),
        onFontError: (f: { code: string }) => errors.push(f.code),
      })
    );
    await waitFor(() => expect(view.getByRole('alert').textContent).toContain('fonts'), {
      timeout: 5000,
    });
    expect(errors).toEqual(['fonts']);
  });

  test('XlsxEditor sees `.once` listeners, and `@save.once` leaves Save disabled', async () => {
    const saved: Uint8Array[] = [];
    const started: Promise<unknown>[] = [];
    const errors: string[] = [];
    let viewer: XlsxEditorInstance | null = null;
    // A template compiles `.once` to `on<Event>Once`.
    const view = mount(() =>
      h(XlsxEditor, {
        document: workbook,
        fonts: FONTS,
        onReady: (v: XlsxEditorInstance) => (viewer = v),
        onSave: (b: Uint8Array) => saved.push(b),
        onSaveStartOnce: (s: Promise<unknown>) => started.push(s),
      })
    );
    await waitFor(() => expect(viewer?.snapshot().status).toBe('ready'), { timeout: 5000 });
    const save = view.getByRole('button', { name: 'Save' });
    expect(save.getAttribute('aria-disabled')).toBeNull();
    fireEvent.click(save);
    await waitFor(() => expect(saved.length).toBe(1));
    fireEvent.click(save);
    await waitFor(() => expect(saved.length).toBe(2));
    // The `.once` listener ran for the first press only.
    expect(started.length).toBe(1);
    const once = mount(() =>
      h(XlsxEditor, { document: workbook, fonts: FONTS, onSaveOnce: () => undefined })
    );
    await waitFor(() => expect(once.queryAllByRole('tab').length).toBeGreaterThan(0), {
      timeout: 5000,
    });
    expect(once.getByRole('button', { name: 'Save' }).getAttribute('aria-disabled')).toBe('true');
    const failed = mount(() =>
      h(XlsxEditor, {
        document: workbook,
        fonts: () => Promise.reject(new Error('offline')),
        onFontErrorOnce: (f: { code: string }) => errors.push(f.code),
      })
    );
    await waitFor(() => expect(failed.getByRole('alert').textContent).toContain('fonts'), {
      timeout: 5000,
    });
    expect(errors).toEqual(['fonts']);
  });

  test('a new fonts or numberLocale value applies in place, and the workbook stays', async () => {
    const bytes = buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1.5)]) }] });
    let viewer: XlsxEditorInstance | null = null;
    const calls: string[] = [];
    const state = reactive<{ fonts: FontSource[]; numberLocale?: NumberLocale }>({
      fonts: FONTS,
    });
    const view = mount(() =>
      h('div', { style: { height: '400px' } }, [
        h(XlsxEditor, {
          document: bytes,
          fonts: state.fonts,
          ...(state.numberLocale ? { numberLocale: state.numberLocale } : {}),
          onReady: (v: XlsxEditorInstance) => {
            viewer = v;
            const setFonts = v.setFonts.bind(v);
            v.setFonts = (f) => (calls.push('setFonts'), setFonts(f));
          },
        }),
      ])
    );
    await waitFor(() => expect(viewer!.snapshot().status).toBe('ready'), { timeout: 5000 });
    viewer!.select({ row: 0, col: 0 });
    await nextTick();
    expect(field(view, 'Cell content').value).toBe('1.5');
    // The same faces again, in a new list, load nothing.
    state.fonts = [...FONTS];
    await nextTick();
    expect(calls).toEqual([]);
    state.fonts = FONTS.map((f) => ({ ...f, bytes: new Uint8Array(f.bytes) }));
    state.numberLocale = { decimal: ',', group: '.' };
    await nextTick();
    expect(calls).toEqual(['setFonts']);
    await nextTick();
    expect(field(view, 'Cell content').value).toBe('1,5');
    expect(viewer!.snapshot().status).toBe('ready');
    expect(viewer!.snapshot().selectionRef).toBe('A1');
  });

  test('ready fires after the first open settles, and a select from it sticks', async () => {
    const bytes = buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B2', 2)]) }] });
    const seen: string[] = [];
    let viewer: XlsxEditorInstance | null = null;
    const state = reactive<{ document?: Uint8Array }>({ document: bytes });
    const ui = () =>
      h('div', { style: { height: '300px' } }, [
        h(XlsxEditor, {
          ...(state.document ? { document: state.document } : {}),
          fonts: FONTS,
          onReady: (v: XlsxEditorInstance) => {
            viewer = v;
            seen.push(v.snapshot().status);
            v.select({ row: 1, col: 1 });
          },
        }),
      ]);
    const view = mount(ui);
    await waitFor(() => expect(seen).toEqual(['ready']), { timeout: 5000 });
    expect(viewer!.snapshot().selectionRef).toBe('B2');
    // A change of `document` to undefined closes the workbook; `ready` does not fire again.
    state.document = undefined;
    await nextTick();
    expect(viewer!.snapshot().status).toBe('empty');
    expect(viewer!.snapshot().sheets).toEqual([]);
    expect(viewer!.select({ row: 0, col: 0 })).toBe(false);
    expect(seen).toEqual(['ready']);
    view.unmount();
    // Without a document, `ready` fires right after the attach.
    mount(ui);
    await waitFor(() => expect(seen).toEqual(['ready', 'empty']));
  });

  test('an inline numberLocale with a state-setting change listener does not loop', async () => {
    const bytes = buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1.5)]) }] });
    let renders = 0;
    let viewer: XlsxEditorInstance | null = null;
    const changed: { last: EditorSnapshot | null } = { last: null };
    const Host = defineComponent({
      setup() {
        const snap = shallowRef<EditorSnapshot | null>(null);
        return () => {
          renders += 1;
          void snap.value;
          return h('div', { style: { height: '300px' } }, [
            h(XlsxEditor, {
              document: bytes,
              fonts: [...FONTS],
              numberLocale: { decimal: ',', group: '.', builtins: { 2: '0,00' } },
              onReady: (v: XlsxEditorInstance) => (viewer = v),
              onChange: (s: EditorSnapshot) => {
                changed.last = s;
                snap.value = s;
              },
            }),
          ]);
        };
      },
    });
    const messages: unknown[] = [];
    const error = console.error;
    const warn = console.warn;
    console.error = (...a: unknown[]) => void messages.push(a[0]);
    console.warn = (...a: unknown[]) => void messages.push(a[0]);
    try {
      mount(() => h(Host));
      await waitFor(() => expect(viewer?.snapshot().status).toBe('ready'), { timeout: 5000 });
      const settled = renders;
      await new Promise((r) => setTimeout(r, 200));
      expect(renders - settled).toBeLessThan(3);
      expect(messages).toEqual([]);
      // `change` delivers the snapshot of the ready workbook, and again after an edit.
      expect(changed.last?.status).toBe('ready');
      expect(changed.last?.sheets.length).toBe(1);
      expect(changed.last?.canUndo).toBe(false);
      viewer!.toggleFormat('bold');
      await nextTick();
      expect(changed.last?.canUndo).toBe(true);
    } finally {
      console.error = error;
      console.warn = warn;
    }
  });

  test('useScopeClassName follows colorMode, and a teleported host popup resolves the tokens', async () => {
    // happy-dom computes a custom property that a rule declares on the element itself, so the
    // test reads `--doc-bg` from the computed style of the popup and of the viewport.
    const sheet = document.createElement('style');
    sheet.textContent = readFileSync(
      join(import.meta.dir, '../../core/src/styles/editor.css'),
      'utf8'
    );
    document.head.append(sheet);
    try {
      const inside: string[] = [];
      const Inside = defineComponent({
        setup() {
          // The color mode makes this part render again after a change, so the test reads the
          // scope of that render.
          const scope = useScopeClassName();
          const { colorMode } = useEditorColorMode();
          return () => {
            inside.push(`${colorMode.value}:${scope.value}`);
            return null;
          };
        },
      });
      const HostPopup = defineComponent({
        setup() {
          const scope = useScopeClassName();
          return () =>
            h(Teleport, { to: 'body' }, [
              h('div', { class: `${scope.value}acme-popup`, 'data-popup': '' }),
            ]);
        },
      });
      const state = reactive<{ colorMode: 'dark' | 'light' }>({ colorMode: 'light' });
      const view = mount(() =>
        h(
          XlsxEditorRoot,
          { fonts: FONTS, colorMode: state.colorMode },
          {
            default: () => [
              h(XlsxEditorViewport, null, { default: () => h(Inside) }),
              h(HostPopup),
            ],
          }
        )
      );
      const popup = () => document.body.querySelector('[data-popup]') as HTMLElement;
      const viewport = () => view.container.querySelector('.xlsx-editor') as HTMLElement;
      const token = (el: HTMLElement) => getComputedStyle(el).getPropertyValue('--doc-bg');
      expect(inside.at(-1)).toBe('light:');
      expect(popup().parentElement).toBe(document.body);
      expect(popup().className).toBe('xlsx-editor acme-popup');
      expect(viewport().className).toBe('xlsx-editor xlsx-editor-viewport');
      expect(token(popup())).not.toBe('');
      expect(token(popup())).toBe(token(viewport()));
      const lightBg = token(popup());

      state.colorMode = 'dark';
      await nextTick();
      expect(inside.at(-1)).toBe('dark:');
      expect(popup().className).toBe('xlsx-editor dark acme-popup');
      expect(token(popup())).toBe(token(viewport()));
      expect(token(popup())).not.toBe(lightBg);
      expect(getComputedStyle(popup()).getPropertyValue('color-scheme')).toBe('dark');
    } finally {
      sheet.remove();
    }
  });

  test('outside a root, useScopeClassName is the light scope', () => {
    const seen: string[] = [];
    mount(() =>
      h(
        defineComponent({
          setup() {
            seen.push(useScopeClassName().value);
            return () => null;
          },
        })
      )
    );
    expect(seen).toEqual(['xlsx-editor ']);
  });

  test('colorMode puts the dark class on the viewport, and system follows the OS', async () => {
    const viewport = (view: Mounted) => view.container.querySelector('.xlsx-editor') as HTMLElement;
    const light = mount(() => h(XlsxEditor, { fonts: FONTS }));
    expect(viewport(light).classList.contains('dark')).toBe(false);
    light.unmount();
    const state = reactive<{ colorMode: 'dark' | 'light' | 'system' }>({ colorMode: 'dark' });
    const dark = mount(() => h(XlsxEditor, { fonts: FONTS, colorMode: state.colorMode }));
    expect(viewport(dark).classList.contains('dark')).toBe(true);
    state.colorMode = 'light';
    await nextTick();
    expect(viewport(dark).classList.contains('dark')).toBe(false);
    dark.unmount();
    const original = window.matchMedia;
    window.matchMedia = ((query: string) => ({
      matches: query === '(prefers-color-scheme: dark)',
      media: query,
      addEventListener: () => {},
      removeEventListener: () => {},
    })) as unknown as typeof window.matchMedia;
    try {
      const seen: string[] = [];
      const Probe = defineComponent({
        setup() {
          const { colorMode } = useEditorColorMode();
          return () => {
            seen.push(colorMode.value);
            return null;
          };
        },
      });
      state.colorMode = 'system';
      const system = mount(() =>
        h(
          XlsxEditorRoot,
          { fonts: FONTS, colorMode: state.colorMode },
          { default: () => h(XlsxEditorViewport, null, () => [h(Probe)]) }
        )
      );
      // The operating system is read after mount, so a server render starts light.
      await nextTick();
      expect(viewport(system).classList.contains('dark')).toBe(true);
      expect(seen.at(-1)).toBe('dark');
      state.colorMode = 'light';
      await nextTick();
      expect(seen.at(-1)).toBe('light');
    } finally {
      window.matchMedia = original;
    }
  });

  test('a locale catalog replaces strings and falls back to English', async () => {
    const view = mount(() =>
      h(XlsxEditor, { fonts: FONTS, locale: { status: { empty: 'Leer', loading: null } } })
    );
    await waitFor(() => expect(view.getByRole('status').textContent).toBe('Leer'));
  });

  test('the host class, style, and attributes land on the viewport', async () => {
    const view = mount(() =>
      h(XlsxEditor, {
        fonts: FONTS,
        class: 'host-class',
        style: { height: '321px' },
        id: 'sheet-editor',
        'data-probe': 'yes',
      })
    );
    const region = view.container.querySelector('.xlsx-editor') as HTMLElement;
    expect(region.classList.contains('host-class')).toBe(true);
    expect(region.style.height).toBe('321px');
    expect(region.id).toBe('sheet-editor');
    expect(region.getAttribute('data-probe')).toBe('yes');
    expect(view.container.querySelectorAll('.host-class').length).toBe(1);
  });

  test('a viewport with label false renders no region; a label names it', () => {
    const none = mount(() =>
      h(XlsxEditorRoot, null, { default: () => h(XlsxEditorViewport, { label: false }) })
    );
    expect(none.container.querySelector('.xlsx-editor')!.hasAttribute('role')).toBe(false);
    const named = mount(() =>
      h(XlsxEditorRoot, null, { default: () => h(XlsxEditorViewport, { label: 'Budget' }) })
    );
    expect(named.getByRole('region', { name: 'Budget' })).toBeTruthy();
  });

  test('unmount destroys the editor', async () => {
    const { view, viewer } = await open();
    let calls = 0;
    viewer.subscribe(() => (calls += 1));
    view.unmount();
    // A destroyed editor has no listeners and no surface.
    viewer.setZoom(2);
    expect(calls).toBe(0);
    expect(document.querySelector('.xlsx-scroller')).toBeNull();
  });
});
