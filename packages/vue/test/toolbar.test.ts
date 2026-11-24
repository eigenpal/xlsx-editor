// The editing toolbar in happy-dom (`bun run test:dom`): it renders the chrome registry, a press
// runs the engine command, a disabled control shows the engine's reason, every popup closes
// through the one close rule, and chrome presses keep the keyboard on the grid. The Vue twin of
// the React adapter's `toolbar.test.tsx`. It leaves out the tests of the core functions and of
// the core stylesheet, which the React file covers once for both adapters.

import { afterEach, describe, expect, test } from 'bun:test';
import { createEvent, fireEvent, waitFor } from '@testing-library/react';
import { h, nextTick } from 'vue';
import { buildXlsx, c, row } from '../../../scripts/fixtures/builder.ts';
import {
  XlsxEditor,
  XlsxEditorContent,
  XlsxEditorRoot,
  StatusMessage,
  XlsxEditorToolbar,
  XlsxEditorViewport,
  type XlsxEditorInstance,
} from '../src/index.ts';
import { CHROME_ICON_PATHS, MAX_ROWS } from '@xlsx-editor.dev/core/editor';
import { FONTS, change, mount, unmountAll } from './mount.ts';
import { type Clip, clipBox, stubLayout, stubScrollHeight, workbook } from './toolbar-stubs.ts';

afterEach(unmountAll);

const CF =
  '<conditionalFormatting sqref="A1"><cfRule type="expression" dxfId="0" priority="1"><formula>1</formula></cfRule></conditionalFormatting>';

async function open(bytes = workbook(), onSave?: (b: Uint8Array) => void) {
  let viewer: XlsxEditorInstance | null = null;
  const view = mount(() =>
    h('div', { style: { height: '400px' } }, [
      h(XlsxEditor, {
        document: bytes,
        fonts: FONTS,
        onReady: (v: XlsxEditorInstance) => (viewer = v),
        ...(onSave ? { onSave } : {}),
      }),
    ])
  );
  await waitFor(() => expect(view.getAllByRole('tab').length).toBe(1), { timeout: 5000 });
  const grid = view.container.querySelector('.xlsx-scroller') as HTMLElement;
  return { view, viewer: viewer!, grid };
}

describe('XlsxEditor.Toolbar', () => {
  /**
   * Open the number format list in the toolbar row, in a window of `windowHeight` pixels, with
   * a trigger from 60 to 88, a list of 354 pixels, and items of `item` pixels. With
   * `pastLeft`, the editor starts at 20, the trigger at 100, and the list runs past the editor's
   * left edge, from -50 to 150. Gives the popup.
   */
  async function openRowList(
    windowHeight: number,
    item: number,
    clip?: Clip,
    scale = 1,
    pastLeft = false
  ) {
    const height = Object.getOwnPropertyDescriptor(window, 'innerHeight');
    Object.defineProperty(window, 'innerHeight', { configurable: true, value: windowHeight });
    const x = (left: number, right: number) => (pastLeft ? { left, right } : {});
    const restore = stubLayout((el) =>
      (scale !== 1 || pastLeft) && el.classList.contains('xlsx-editor')
        ? { width: 400 * scale, height: 400 * scale, ...x(20, 20 + 400 * scale) }
        : clip !== undefined && el.hasAttribute('data-test-clip')
          ? clipBox(clip)
          : el.classList.contains('xlsx-toolbar-dropdown')
            ? { top: 60, bottom: 88, height: 28, ...x(100, 128) }
            : el.hasAttribute('data-popup')
              ? { top: 92, bottom: 446, height: 354, ...x(-50, 150) }
              : el.getAttribute('role') === 'option'
                ? { height: item }
                : null
    );
    const unstub = stubScrollHeight((el) => (el.hasAttribute('data-popup') ? 354 : null));
    try {
      const { view } = await open();
      // A host container that clips its overflow, around the editor.
      const host = view.container.firstElementChild as HTMLElement;
      if (clip !== undefined) {
        host.style.overflow = 'hidden';
        host.setAttribute('data-test-clip', '');
      }
      // A host that scales the editor: its box shows at `scale` times its layout size.
      const editor = view.container.querySelector<HTMLElement>('.xlsx-editor')!;
      for (const size of ['offsetWidth', 'offsetHeight'])
        Object.defineProperty(editor, size, { configurable: true, value: 400 });
      fireEvent.click(view.getByRole('button', { name: /^Number format/ }));
      await nextTick();
      return view
        .getByRole('listbox', { name: 'Number format' })
        .closest<HTMLElement>('[data-popup]')!;
    } finally {
      unstub();
      restore();
      if (height) Object.defineProperty(window, 'innerHeight', height);
      else delete (window as unknown as Record<string, unknown>).innerHeight;
    }
  }

  test('a list in the toolbar row is capped to the room below it, so the page never scrolls', async () => {
    const popup = await openRowList(340, 24);
    // 340 less the trigger's bottom (88), less the gap of 4 twice: 244 pixels below it.
    expect(popup.dataset.side).toBeUndefined();
    expect(popup.style.maxHeight).toBe('244px');
    expect(popup.hasAttribute('data-capped')).toBe(true);
  });

  test('in a host container that clips its overflow, a list is capped to the container', async () => {
    // A container of 200 pixels in a window of 900: 200 less 88, less the gap of 4 twice.
    const popup = await openRowList(900, 24, 200);
    expect([popup.dataset.side, popup.style.maxHeight]).toEqual([undefined, '104px']);
  });

  test('inside a host that scales the editor by 2, a list is capped in its own pixels', async () => {
    // 340 less 88, less the gap of 4 window pixels twice at scale 2: 236 window pixels, or 118
    // of the list's pixels.
    const popup = await openRowList(340, 24, undefined, 2);
    expect([popup.dataset.side, popup.style.maxHeight]).toEqual([undefined, '118px']);
  });

  test('inside a host that scales the editor by 2, one item is measured in its own pixels', async () => {
    // 194 less 88, less the gap of 4 window pixels twice at scale 2: 90 window pixels, or 45 of
    // the list's pixels. An item of 60 window pixels is 30 of the list's pixels, so it shows.
    const popup = await openRowList(194, 60, undefined, 2);
    expect([popup.dataset.side, popup.style.maxHeight]).toEqual([undefined, '45px']);
  });

  test('inside a host that scales the editor by 2, a list is held inside its left edge', async () => {
    // The editor's left edge is 80 window pixels before the trigger: 40 of the list's pixels.
    const popup = await openRowList(900, 24, undefined, 2, true);
    expect([popup.style.left, popup.style.right]).toEqual(['-40px', 'auto']);
  });

  test('inside a host that scales the editor by 2, a list opens above only when the room shows one item in its own pixels', async () => {
    // A window of 100 pixels and a gap of 8 window pixels: 4 below the trigger, and 52 above,
    // less the gap: 44 window pixels, or 22 of the list's pixels. An item of 60 window pixels is
    // 30 of the list's pixels and does not fit above, so the list stays below, and it is not
    // capped.
    const popup = await openRowList(100, 60, undefined, 2);
    expect([popup.dataset.side, popup.style.maxHeight]).toEqual([undefined, '']);
  });

  test('in a host container that clips above the trigger, the room above ends at the container', async () => {
    // A container from 30 to 110: 18 pixels below the trigger, and 26 above it. The list opens
    // above, capped to 26 less the gap: 22 pixels.
    const popup = await openRowList(900, 12, [30, 110]);
    expect([popup.dataset.side, popup.style.maxHeight]).toEqual(['top', '22px']);
  });

  test('a list opens above its trigger only when the room there shows one item', async () => {
    // A window of 100 pixels: 4 below the trigger, and 60 less the gap of 4 twice (52) above.
    const above = await openRowList(100, 24);
    expect([above.dataset.side, above.style.maxHeight]).toEqual(['top', '52px']);
    unmountAll();
    // An item of 60 pixels does not fit above: the list stays below, and it is not capped.
    const below = await openRowList(100, 60);
    expect([below.dataset.side, below.style.maxHeight]).toEqual([undefined, '']);
  });

  test('Control or Command and S saves from the editor, and leaves other keys and places alone', async () => {
    const saved: Uint8Array[] = [];
    const { view, viewer, grid } = await open(workbook(), (b) => saved.push(b));
    // In the grid: the save runs, and the browser's own save does not.
    let ev = createEvent.keyDown(grid, { key: 's', ctrlKey: true });
    fireEvent(grid, ev);
    expect(ev.defaultPrevented).toBe(true);
    await waitFor(() => expect(saved.length).toBe(1));
    ev = createEvent.keyDown(grid, { key: 'S', metaKey: true });
    fireEvent(grid, ev);
    await waitFor(() => expect(saved.length).toBe(2));
    // In the cell editor: the open edit commits first, as the control does.
    fireEvent.keyDown(grid, { key: '7' });
    await nextTick();
    const input = view.getByLabelText('Cell editor');
    ev = createEvent.keyDown(input, { key: 's', ctrlKey: true });
    fireEvent(input, ev);
    expect(ev.defaultPrevented).toBe(true);
    await waitFor(() => expect(saved.length).toBe(3));
    expect(viewer.snapshot().editing).toBe(false);
    // While an input method composes, the save takes the refusal path of the control.
    fireEvent.keyDown(grid, { key: '8' });
    await nextTick();
    const input2 = view.getByLabelText('Cell editor');
    fireEvent.compositionStart(input2);
    await nextTick();
    fireEvent.keyDown(input2, { key: 's', ctrlKey: true });
    await nextTick();
    await waitFor(() => expect(saved.length).toBe(3));
    expect(viewer.snapshot().refusal?.code).toBe('composing');
    fireEvent.compositionEnd(input2);
    await nextTick();
    viewer.cancelEdit();
    await nextTick();
    // Shift or Alt with it, and a key outside the editor, are not the shortcut.
    fireEvent.keyDown(grid, { key: 's', ctrlKey: true, shiftKey: true });
    await nextTick();
    const outside = document.createElement('input');
    document.body.append(outside);
    ev = createEvent.keyDown(outside, { key: 's', ctrlKey: true });
    fireEvent(outside, ev);
    expect(ev.defaultPrevented).toBe(false);
    outside.remove();
    await waitFor(() => expect(saved.length).toBe(3));
    // In an open dialog, the key stays with the dialog.
    fireEvent.click(view.getByRole('button', { name: /^Number format/ }));
    await nextTick();
    fireEvent.click(view.getByRole('button', { name: 'More number formats…' }));
    await nextTick();
    const field = view.getByRole('textbox', { name: 'Format code' });
    ev = createEvent.keyDown(field, { key: 's', ctrlKey: true });
    fireEvent(field, ev);
    expect(ev.defaultPrevented).toBe(true);
    // A save delivers its bytes later, so the check waits for a delivery that must not come.
    await new Promise((resolve) => setTimeout(resolve, 300));
    await nextTick();
    expect(saved.length).toBe(3);
    expect(view.getByRole('dialog')).toBeTruthy();
  });

  test('XlsxEditor tells onSaveStart when a press of Save starts a save', async () => {
    const started: Promise<unknown>[] = [];
    const saved: Uint8Array[] = [];
    let viewer: XlsxEditorInstance | null = null;
    const view = mount(() =>
      h(XlsxEditor, {
        document: workbook(),
        fonts: FONTS,
        onReady: (v: XlsxEditorInstance) => (viewer = v),
        onSave: (b: Uint8Array) => saved.push(b),
        onSaveStart: (s: Promise<unknown>) => started.push(s),
      })
    );
    await waitFor(() => expect(viewer?.snapshot().status).toBe('ready'), { timeout: 5000 });
    fireEvent.click(view.getByRole('button', { name: 'Save' }));
    await nextTick();
    expect(started.length).toBe(1);
    await waitFor(() => expect(saved.length).toBe(1));
  });

  test('Control and S saves with the save control hidden', async () => {
    const saved: Uint8Array[] = [];
    let viewer: XlsxEditorInstance | null = null;
    const view = mount(() =>
      h(
        XlsxEditorRoot,
        { document: workbook(), fonts: FONTS, onReady: (v: XlsxEditorInstance) => (viewer = v) },
        {
          default: () =>
            h(XlsxEditorViewport, { style: { height: '400px' } }, () => [
              h(XlsxEditorToolbar, {
                hidden: ['file.save'],
                onSave: (b: Uint8Array) => saved.push(b),
              }),
              h(XlsxEditorContent),
              h(StatusMessage),
            ]),
        }
      )
    );
    await waitFor(() => expect(viewer?.snapshot().status).toBe('ready'), { timeout: 5000 });
    const grid = view.container.querySelector('.xlsx-scroller') as HTMLElement;
    const ev = createEvent.keyDown(grid, { key: 's', ctrlKey: true });
    fireEvent(grid, ev);
    expect(ev.defaultPrevented).toBe(true);
    await waitFor(() => expect(saved.length).toBe(1));
  });

  test('a refused save says that the workbook was not saved, and a dropped edit stays shown after the load', async () => {
    const saved: Uint8Array[] = [];
    const { view, viewer, grid } = await open(workbook(), (b) => saved.push(b));
    fireEvent.keyDown(grid, { key: '=' });
    await nextTick();
    const input = view.getByLabelText('Cell editor') as HTMLTextAreaElement;
    input.value = '=SUM(1';
    fireEvent.input(input);
    await nextTick();
    await nextTick();
    fireEvent.keyDown(input, { key: 's', ctrlKey: true });
    await nextTick();
    expect(saved.length).toBe(0);
    const refusal = viewer.snapshot().refusal;
    expect(refusal?.command).toBe('save');
    expect(refusal?.code).toBe('unreadable-formula');
    const live = view.container.querySelector('.xlsx-live')!;
    const edit =
      'The workbook was not saved: the formula of the cell edit cannot be read. Check the quotes, parentheses, and references.';
    expect(live.textContent).toBe(edit);
    await waitFor(() =>
      expect(view.container.querySelector('.xlsx-refusal')?.textContent).toBe(edit)
    );
    // The load drops the open edit. A save during the load is refused: both messages show under
    // the loading text, and the new workbook clears them.
    let loading: Promise<unknown> = Promise.resolve();
    loading = viewer.load(workbook());
    // No wait here: the save must come while the load runs.
    fireEvent.keyDown(grid, { key: 's', ctrlKey: true });
    expect(saved.length).toBe(0);
    expect(viewer.snapshot().refusal).toEqual({ code: 'loading', command: 'save' });
    expect(viewer.snapshot().editDropped).toBe(true);
    const busy = 'The workbook was not saved: a workbook is loading.';
    expect(live.textContent).toBe(busy);
    const shown = () =>
      [...view.container.querySelectorAll('.xlsx-message-refusal')].map((n) => n.textContent);
    await waitFor(() =>
      expect(shown()).toEqual(['The open cell edit was dropped: its text was not stored.', busy])
    );
    await loading;
    await nextTick();
    // In the ready workbook, the strip still shows that the typed text was dropped, until the
    // next selection.
    const dropped = 'The open cell edit was dropped: its text was not stored.';
    expect(viewer.snapshot().refusal).toEqual({ code: 'edit-dropped' });
    expect(view.container.querySelector('.xlsx-message-refusal')).toBeNull();
    await waitFor(() =>
      expect(view.container.querySelector('.xlsx-refusal')?.textContent).toBe(dropped)
    );
    viewer.select({ row: 1, col: 1 });
    await nextTick();
    expect(viewer.snapshot().refusal).toBeNull();
    expect(viewer.snapshot().editDropped).toBe(false);
    await waitFor(() => expect(view.container.querySelector('.xlsx-refusal')).toBeNull());
  });

  test('without onSave, Control and S is left to the browser', async () => {
    const { grid } = await open();
    const ev = createEvent.keyDown(grid, { key: 's', ctrlKey: true });
    fireEvent(grid, ev);
    expect(ev.defaultPrevented).toBe(false);
  });

  test('renders every group of the registry, with names from the catalog', async () => {
    const { view } = await open();
    const bar = view.getByRole('toolbar', { name: 'Formatting controls' });
    const groups = [...bar.querySelectorAll('[role="group"]')].map((g) =>
      g.getAttribute('aria-label')
    );
    expect(groups).toEqual([
      'History',
      'Font',
      'Fill',
      'Borders',
      'Alignment',
      'Number format',
      'Insert and delete',
      'Clear',
      'File',
    ]);
    expect(view.getByRole('button', { name: /^Select font family: \S/ })).toBeTruthy();
  });

  test('a press on bold toggles the format of the selection, and shows it pressed', async () => {
    const { view, viewer } = await open();
    const bold = view.getByRole('button', { name: 'Bold' });
    expect(bold.getAttribute('aria-pressed')).toBe('false');
    fireEvent.click(bold);
    await nextTick();
    expect(viewer.snapshot().activeFormat?.bold).toBe(true);
    expect(view.getByRole('button', { name: 'Bold' }).getAttribute('aria-pressed')).toBe('true');
    fireEvent.click(view.getByRole('button', { name: 'Bold' }));
    await nextTick();
    expect(viewer.snapshot().activeFormat?.bold).toBe(false);
  });

  test('a protected sheet disables bold with the locked reason, and typing opens no editor', async () => {
    const protectedSheet = buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 1)]),
          beforeMerges: '<sheetProtection sheet="1"/>',
        },
      ],
    });
    const { view, viewer } = await open(protectedSheet);
    // The sheet shows that it is protected, with a status that a screen reader announces.
    const notice = view.getByText(
      'This sheet is protected. Locked cells and locked actions cannot be edited.'
    );
    expect(notice.closest('[role="status"]')).not.toBeNull();
    const bold = view.getByRole('button', { name: 'Bold' });
    expect(bold.getAttribute('aria-disabled')).toBe('true');
    expect(bold.title).toBe('Bold: The sheet is protected, and formatting cells is not allowed.');
    fireEvent.click(bold);
    await nextTick();
    expect(viewer.snapshot().activeFormat?.bold).toBe(false);
    const entered: { r?: ReturnType<typeof viewer.enterValue> } = {};
    entered.r = viewer.enterValue('5');
    await nextTick();
    const r = entered.r!;
    expect(r.ok ? null : r.refusal.code).toBe('locked');
    expect(view.getAllByText(/The sheet is protected, and/).length).toBeGreaterThan(0);
    // With a refusal shown, the protection notice shows its short text, visually only: the
    // text of the live node does not change on a refusal or a selection, so it is not
    // announced again.
    expect(view.getByText('This sheet is protected.').getAttribute('aria-hidden')).toBe('true');
    const live = notice.closest('[role="status"]') as HTMLElement;
    expect(live.hasAttribute('data-short')).toBe(true);
    const text = live.textContent;
    viewer.select({ row: 2, col: 2 });
    await nextTick();
    expect(live.hasAttribute('data-short')).toBe(false);
    expect(live.textContent).toBe(text);
    viewer.enterValue('6');
    await nextTick();
    expect(live.textContent).toBe(text);
  });

  test('a disabled undo stays focusable, describes its reason, and enables after an edit', async () => {
    const { view, viewer } = await open();
    const undo = view.getByRole('button', { name: 'Undo' }) as HTMLButtonElement;
    expect(undo.disabled).toBe(false);
    expect(undo.getAttribute('aria-disabled')).toBe('true');
    expect(undo.title).toBe('Undo: There is nothing to undo.');
    const described = document.getElementById(undo.getAttribute('aria-describedby') ?? '');
    expect(described?.textContent).toBe('There is nothing to undo.');
    // A press on a disabled control does nothing.
    fireEvent.click(undo);
    await nextTick();
    expect(viewer.snapshot().canRedo).toBe(false);
    fireEvent.click(view.getByRole('button', { name: 'Italic' }));
    await nextTick();
    const enabled = view.getByRole('button', { name: 'Undo' });
    expect(enabled.getAttribute('aria-disabled')).toBeNull();
    expect(enabled.getAttribute('aria-describedby')).toBeNull();
  });

  test('a control that becomes disabled under focus keeps the focus', async () => {
    const { view } = await open();
    fireEvent.click(view.getByRole('button', { name: 'Italic' }));
    await nextTick();
    const undo = view.getByRole('button', { name: 'Undo' });
    undo.focus();
    fireEvent.click(undo);
    await nextTick();
    expect(undo.getAttribute('aria-disabled')).toBe('true');
    expect(document.activeElement).toBe(undo);
    // The arrow keys still move from it, and land on disabled controls too.
    fireEvent.keyDown(undo, { key: 'ArrowRight' });
    await nextTick();
    expect(document.activeElement).toBe(view.getByRole('button', { name: 'Redo' }));
  });

  test('Escape from the grid closes a dropdown, and the keyboard stays in the grid', async () => {
    const { view, viewer, grid } = await open();
    viewer.select({ row: 0, col: 1 });
    await nextTick();
    viewer.focus();
    await nextTick();
    expect(document.activeElement).toBe(grid);
    const trigger = view.getByRole('button', { name: /^Select font family/ });
    expect(fireEvent.mouseDown(trigger)).toBe(false);
    fireEvent.click(trigger);
    await nextTick();
    expect(view.getByRole('listbox', { name: 'Select font family' })).toBeTruthy();
    const escape = fireEvent.keyDown(grid, { key: 'Escape' });
    expect(escape).toBe(false);
    await nextTick();
    expect(view.queryByRole('listbox')).toBeNull();
    expect(document.activeElement).toBe(grid);
    expect(viewer.snapshot().selectionRef).toBe('B1');
  });

  test('a dropdown opened from the keyboard closes on Escape back to its trigger', async () => {
    const { view } = await open();
    const trigger = view.getByRole('button', { name: /^Number format/ });
    trigger.focus();
    fireEvent.keyDown(trigger, { key: 'ArrowDown' });
    await nextTick();
    const list = view.getByRole('listbox', { name: 'Number format' });
    await waitFor(() => expect(list.contains(document.activeElement)).toBe(true));
    expect(document.activeElement?.getAttribute('aria-selected')).toBe('true');
    fireEvent.keyDown(document.activeElement!, { key: 'Escape' });
    await nextTick();
    expect(view.queryByRole('listbox')).toBeNull();
    expect(document.activeElement).toBe(view.getByRole('button', { name: /^Number format/ }));
  });

  test('a pick applies the value and returns the keyboard to the grid', async () => {
    const { view, viewer, grid } = await open();
    fireEvent.click(view.getByRole('button', { name: /^Font size/ }));
    await nextTick();
    fireEvent.click(view.getByRole('option', { name: '14' }));
    await nextTick();
    expect(viewer.snapshot().activeFormat?.fontSize).toBe(14);
    expect(view.queryByRole('listbox')).toBeNull();
    expect(document.activeElement).toBe(grid);
    fireEvent.click(view.getByRole('button', { name: 'Fill color options' }));
    await nextTick();
    const yellow = view.getByRole('button', { name: 'Yellow' });
    expect(yellow.getAttribute('aria-pressed')).toBe('false');
    fireEvent.click(yellow);
    await nextTick();
    expect(viewer.snapshot().activeFormat?.fillColor).toBe('FFFF00');
  });

  test('only the selected option has the check mark, and each popup trigger has a caret icon', async () => {
    const { view, viewer } = await open();
    const trigger = view.getByRole('button', { name: /^Font size/ });
    fireEvent.click(trigger);
    await nextTick();
    fireEvent.click(view.getByRole('option', { name: '14' }));
    await nextTick();
    expect(viewer.snapshot().activeFormat?.fontSize).toBe(14);
    fireEvent.click(trigger);
    await nextTick();
    const checked = [
      ...view.getByRole('listbox').querySelectorAll<HTMLElement>('[role="option"]'),
    ].filter((o) => o.querySelector('svg.xlsx-toolbar-check'));
    expect(checked.map((o) => [o.textContent, o.getAttribute('aria-selected')])).toEqual([
      ['14', 'true'],
    ]);
    // The mark is hidden from assistive technology: the option's name stays the value.
    expect(view.getByRole('option', { name: '14' })).toBe(checked[0]);
    expect(checked[0]!.querySelector('svg.xlsx-toolbar-check')?.getAttribute('aria-hidden')).toBe(
      'true'
    );
    // A field takes `expand_more`, and a color trigger and a group menu take `arrow_drop_down`.
    const caret = (name: RegExp) =>
      view
        .getByRole('button', { name })
        .querySelector('svg.xlsx-toolbar-caret path')
        ?.getAttribute('d');
    expect(caret(/^Font size/)).toBe(CHROME_ICON_PATHS.expand_more[0]);
    expect(caret(/^Fill color options/)).toBe(CHROME_ICON_PATHS.arrow_drop_down[0]);
    const menu = view.container.querySelector('[data-group-menu] svg.xlsx-toolbar-caret path');
    expect(menu?.getAttribute('d')).toBe(CHROME_ICON_PATHS.arrow_drop_down[0]);
    expect(view.container.querySelector('span.xlsx-toolbar-caret')).toBeNull();
    for (const icon of view.container.querySelectorAll('svg.xlsx-toolbar-caret'))
      expect(icon.getAttribute('aria-hidden')).toBe('true');
  });

  test('a press outside an open popup closes it', async () => {
    const { view, grid } = await open();
    fireEvent.click(view.getByRole('button', { name: 'Font color options' }));
    await nextTick();
    expect(view.getByRole('dialog', { name: 'Font color' })).toBeTruthy();
    fireEvent.pointerDown(grid);
    await nextTick();
    expect(view.queryByRole('dialog')).toBeNull();
  });

  test('a structural command that design D4 refuses is disabled with its reason', async () => {
    const { view } = await open(workbook(CF));
    fireEvent.click(view.getByRole('button', { name: 'Insert and delete' }));
    await nextTick();
    const item = view.getByRole('menuitem', { name: /Delete rows/ });
    expect(item.getAttribute('aria-disabled')).toBe('true');
    // The reason names the content in words, never the part path.
    expect(item.title).toContain('conditional formatting on sheet');
    expect(item.title).not.toContain('/xl/');
    expect(item.textContent).toContain('can refer to cells that would move');
  });

  test('a rule that the editor does not show gives a notice for the sheet', async () => {
    const full =
      'This sheet has conditional formatting that the grid can leave out. Where it does, a cell can show its own format or the format of another rule. A save keeps every rule.';
    const { view } = await open(workbook(CF));
    // The status strip keeps its notices mounted; only the conditional formatting one has text.
    const shown = [...view.container.querySelectorAll('.xlsx-notice')].filter((n) => n.textContent);
    expect(shown.map((n) => [n.getAttribute('role'), n.getAttribute('title')])).toEqual([
      ['status', full],
    ]);
    expect(shown[0]!.querySelector('.xlsx-notice-full')!.textContent).toBe(full);
    expect(shown[0]!.hasAttribute('data-short')).toBe(false);
  });

  test('with the protection notice, the conditional formatting notice shows its short text', async () => {
    const bytes = buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 1)]),
          beforeMerges: '<sheetProtection sheet="1"/>',
          afterMerges: CF,
        },
      ],
    });
    const { view } = await open(bytes);
    const brief = view.getByText('The grid can leave out some conditional formatting.');
    expect(brief.getAttribute('aria-hidden')).toBe('true');
    const live = brief.closest('[role="status"]') as HTMLElement;
    expect(live.hasAttribute('data-short')).toBe(true);
    expect(live.querySelector('.xlsx-notice-full')!.textContent).toContain(
      'conditional formatting that the grid can leave out'
    );
  });

  test('a true rule at the active cell explains the bold control in its tooltip', async () => {
    const rule =
      '<conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>';
    const bytes = buildXlsx({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]), afterMerges: rule }],
      styles: { dxfs: ['<dxf><font><b/></font></dxf>'] },
    });
    const { view } = await open(bytes);
    const bold = view.getByRole('button', { name: 'Bold' });
    expect(bold.getAttribute('aria-disabled')).toBeNull();
    const hint =
      'Conditional formatting changes how this cell shows. This control reads and sets the cell format.';
    // The tooltip keeps the name of the control, and the hint is the description.
    expect(bold.title).toBe(`Bold\n${hint}`);
    const described = view.container.querySelector(
      `#${CSS.escape(bold.getAttribute('aria-describedby')!)}`
    );
    expect(described?.textContent).toBe(hint);
    // Italic shows as the cell format shows it: no hint.
    const italic = view.getByRole('button', { name: 'Italic' });
    expect(italic.title).toBe('Italic');
    expect(italic.getAttribute('aria-describedby')).toBeNull();
    expect(
      [...view.container.querySelectorAll('.xlsx-notice')].filter((n) => n.textContent)
    ).toEqual([]);
  });

  test('a fill rule gives the hint to the fill control, and not to bold', async () => {
    const rule =
      '<conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>';
    const bytes = buildXlsx({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]), afterMerges: rule }],
      styles: {
        dxfs: ['<dxf><fill><patternFill><bgColor rgb="FFFFFF00"/></patternFill></fill></dxf>'],
      },
    });
    const { view } = await open(bytes);
    const fill = view.getByRole('button', { name: 'Fill color' });
    expect(fill.title).toBe(
      'Fill color\nConditional formatting changes how this cell shows. This control reads and sets the cell format.'
    );
    expect(fill.getAttribute('aria-describedby')).toBeTruthy();
    expect(view.getByRole('button', { name: 'Bold' }).title).toBe('Bold');
  });

  test('a menu item runs its command and closes the menu', async () => {
    const { view, viewer } = await open();
    fireEvent.click(view.getByRole('button', { name: 'Clear' }));
    await nextTick();
    fireEvent.click(view.getByRole('menuitem', { name: /Clear contents/ }));
    await nextTick();
    expect(view.queryByRole('menu')).toBeNull();
    expect(viewer.snapshot().activeCell?.content).toBe('');
    fireEvent.click(view.getByRole('button', { name: 'Insert and delete' }));
    await nextTick();
    fireEvent.click(view.getByRole('menuitem', { name: /Insert rows/ }));
    await nextTick();
    expect(viewer.snapshot().canUndo).toBe(true);
  });

  test('save hands the bytes to the host, and is disabled without a receiver', async () => {
    const first = await open();
    const off = first.view.getByRole('button', { name: 'Save' }) as HTMLButtonElement;
    expect(off.getAttribute('aria-disabled')).toBe('true');
    expect(off.title).toBe('Save: Saving is not set up for this editor.');
    unmountAll();
    const saved: Uint8Array[] = [];
    const { view } = await open(workbook(), (b) => saved.push(b));
    fireEvent.click(view.getByRole('button', { name: 'Save' }));
    await nextTick();
    await waitFor(() => expect(saved.length).toBe(1));
    expect(saved[0]!.byteLength).toBeGreaterThan(0);
  });

  test('a press on a disabled control keeps the focus on the grid', async () => {
    const { view } = await open();
    const bar = view.getByRole('toolbar');
    expect(fireEvent.mouseDown(bar)).toBe(false);
  });

  test('the arrow keys move between the controls of the bar', async () => {
    const { view } = await open();
    const bold = view.getByRole('button', { name: 'Bold' });
    bold.focus();
    fireEvent.keyDown(bold, { key: 'ArrowRight' });
    await nextTick();
    expect(document.activeElement).toBe(view.getByRole('button', { name: 'Italic' }));
    fireEvent.keyDown(document.activeElement!, { key: 'ArrowLeft' });
    await nextTick();
    expect(document.activeElement).toBe(bold);
  });
});

describe('the toolbar keyboard', () => {
  test('Enter and Space open a list with the keyboard on the selected item', async () => {
    const { view } = await open();
    for (const key of ['Enter', ' ']) {
      const trigger = view.getByRole('button', { name: /^Font size/ });
      trigger.focus();
      fireEvent.keyDown(trigger, { key });
      await nextTick();
      const list = view.getByRole('listbox', { name: 'Font size' });
      await waitFor(() => expect(list.contains(document.activeElement)).toBe(true));
      expect(document.activeElement?.getAttribute('aria-selected')).toBe('true');
      fireEvent.keyDown(document.activeElement!, { key: 'Escape' });
      await nextTick();
      expect(view.queryByRole('listbox')).toBeNull();
    }
  });

  test('Down on an open popup moves the keyboard into it', async () => {
    const { view } = await open();
    const trigger = view.getByRole('button', { name: 'Fill color options' });
    fireEvent.click(trigger);
    await nextTick();
    trigger.focus();
    fireEvent.keyDown(trigger, { key: 'ArrowDown' });
    await nextTick();
    const palette = view.getByRole('dialog', { name: 'Fill color' });
    // The cell has no fill: the keyboard starts on "No fill", and Down moves to the grid.
    expect(document.activeElement).toBe(view.getByRole('button', { name: 'No fill' }));
    fireEvent.keyDown(document.activeElement!, { key: 'ArrowDown' });
    await nextTick();
    expect(document.activeElement).toBe(view.getByRole('button', { name: 'Background 1' }));
    expect(palette.contains(document.activeElement)).toBe(true);
  });

  test('Enter opens a group menu with the keyboard on its first item', async () => {
    const { view } = await open();
    const trigger = view.getByRole('button', { name: 'Clear' });
    trigger.focus();
    fireEvent.keyDown(trigger, { key: 'Enter' });
    await nextTick();
    await waitFor(() =>
      expect(document.activeElement).toBe(view.getByRole('menuitem', { name: /Clear contents/ }))
    );
  });

  test('a disabled trigger opens nothing', async () => {
    const { view, viewer } = await open();
    viewer.select({ row: 0, col: 0 }, { row: MAX_ROWS - 1, col: 3 });
    await nextTick();
    const trigger = view.getByRole('button', { name: /^Font size/ });
    expect(trigger.getAttribute('aria-disabled')).toBe('true');
    fireEvent.keyDown(trigger, { key: 'Enter' });
    await nextTick();
    fireEvent.click(trigger);
    await nextTick();
    expect(view.queryByRole('listbox')).toBeNull();
  });

  test('a typed font size applies, and a size out of range is marked invalid with a message', async () => {
    const { view, viewer, grid } = await open();
    fireEvent.click(view.getByRole('button', { name: /^Font size/ }));
    await nextTick();
    const entry = view.getByRole('textbox', { name: /^Type a font size/ }) as HTMLInputElement;
    change(entry, '500');
    await nextTick();
    fireEvent.keyDown(entry, { key: 'Enter' });
    await nextTick();
    expect(entry.getAttribute('aria-invalid')).toBe('true');
    expect(viewer.snapshot().activeFormat?.fontSize).not.toBe(500);
    // The entry shows why, and names the message as its description.
    const message = view.getByRole('alert');
    expect(message.textContent).toBe('The size was not applied: type a number from 1 to 409.');
    expect(entry.getAttribute('aria-describedby')).toBe(message.id);
    change(entry, '13,5');
    await nextTick();
    fireEvent.keyDown(entry, { key: 'Enter' });
    await nextTick();
    expect(viewer.snapshot().activeFormat?.fontSize).toBe(13.5);
    expect(view.queryByRole('listbox')).toBeNull();
    expect(document.activeElement).toBe(grid);
  });

  test('a popup near the bottom of the window opens above, inside the editor', async () => {
    const { view } = await open();
    const height = Object.getOwnPropertyDescriptor(window, 'innerHeight');
    Object.defineProperty(window, 'innerHeight', { configurable: true, value: 600 });
    const restore = stubLayout((el) =>
      el.hasAttribute('data-popup')
        ? { left: -20, right: 100, top: 494, bottom: 894, width: 120, height: 400 }
        : el.classList.contains('xlsx-toolbar-dropdown')
          ? { left: 0, right: 136, top: 460, bottom: 490, height: 30 }
          : el.classList.contains('xlsx-editor')
            ? { left: 0, right: 1000, top: 0, bottom: 1000 }
            : null
    );
    try {
      fireEvent.click(view.getByRole('button', { name: /^Number format/ }));
      await nextTick();
      const popup = view.getByRole('listbox').closest('[data-popup]') as HTMLElement;
      expect(popup.dataset.side).toBe('top');
      expect(popup.style.left).toBe('0px');
    } finally {
      restore();
      if (height) Object.defineProperty(window, 'innerHeight', height);
    }
  });
});

describe('hidden slots', () => {
  test('the toolbar renders without the hidden slots', async () => {
    const view = mount(() =>
      h(
        XlsxEditorRoot,
        { document: workbook(), fonts: FONTS },
        {
          default: () =>
            h(XlsxEditorViewport, null, () => [
              h(XlsxEditorToolbar, { hidden: ['file.save', 'history.undo', 'history.redo'] }),
              h(XlsxEditorContent),
              h(StatusMessage),
            ]),
        }
      )
    );
    const bar = view.getByRole('toolbar');
    expect(bar.querySelector('[data-group="file"]')).toBeNull();
    expect(bar.querySelector('[data-group="history"]')).toBeNull();
    expect(bar.querySelector('[data-slot="font.bold"]')).toBeTruthy();
  });
});
