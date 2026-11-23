// The number format dialog in happy-dom: the last item of the number format list
// opens it; the code field starts at the active cell's code; the preview follows the typed
// code through the engine; a code that the store refuses is marked with its reason and applies
// nothing; Apply writes the code in one undo step and returns the keyboard to the grid; the
// workbook's codes render as text and fill the field; and the slot alone is an icon button.
// The Vue twin of the React adapter's `number-format-dialog.test.tsx`: Vue renders in the next microtask, so a
// step waits for `nextTick` before it reads the DOM.

import { afterEach, describe, expect, test } from 'bun:test';
import { fireEvent, waitFor } from '@testing-library/react';
import { h, nextTick } from 'vue';
import { buildXlsx, c, row } from '../../../scripts/fixtures/builder.ts';
import {
  SheetTabs,
  XlsxEditorContent,
  XlsxEditorRoot,
  XlsxEditorToolbar,
  XlsxEditorViewport,
  type XlsxEditorInstance,
} from '../src/index.ts';
import { FONTS, change, mount, tabsShow, unmountAll } from './mount.ts';

afterEach(unmountAll);

const workbook = (numFmts: readonly string[] = []) =>
  buildXlsx({
    sheets: [{ name: 'S', rows: row(1, [c('A1', 1234.5), c('B1', 'two')]) }],
    styles: { numFmts },
  });

async function open(bytes = workbook(), hidden?: readonly ('number.format' | 'number.custom')[]) {
  let viewer: XlsxEditorInstance | null = null;
  const view = mount(() =>
    h(
      XlsxEditorRoot,
      { document: bytes, fonts: FONTS, onReady: (v: XlsxEditorInstance) => (viewer = v) },
      {
        default: () =>
          h(XlsxEditorViewport, null, () => [
            h(XlsxEditorToolbar, hidden ? { hidden } : {}),
            h(XlsxEditorContent),
            h(SheetTabs),
          ]),
      }
    )
  );
  await tabsShow(view, 1);
  const grid = view.container.querySelector('.xlsx-scroller') as HTMLElement;
  return { view, viewer: viewer!, grid };
}

/** Open the dialog through the last item of the number format list. */
async function openDialog(view: Awaited<ReturnType<typeof open>>['view']) {
  fireEvent.click(view.getByRole('button', { name: /^Number format/ }));
  await nextTick();
  fireEvent.click(view.getByRole('button', { name: 'More number formats…' }));
  await nextTick();
  const dialog = view.getByRole('dialog', { name: 'Number format' });
  const field = view.getByRole('textbox', { name: 'Format code' }) as HTMLInputElement;
  // Vue skips a listener for an event whose time stamp is not later than the time it attached
  // the listener. A key in the same millisecond as the mount reaches only some listeners of the
  // dialog, so wait for the next millisecond, as PPTX Editor's `dialog-frame.test.ts` does.
  const mounted = Date.now();
  while (Date.now() <= mounted) await new Promise((r) => setTimeout(r, 1));
  return { dialog, field };
}

/** Type `text` into the code field, and let Vue render it. */
async function type(field: HTMLInputElement, text: string): Promise<void> {
  change(field, text);
  await nextTick();
}

describe('the number format dialog', () => {
  test('opens from the list with the active code, previews, and applies in one undo step', async () => {
    const { view, viewer, grid } = await open();
    const { dialog, field } = await openDialog(view);
    expect(view.queryByRole('listbox')).toBeNull();
    expect(dialog.getAttribute('aria-modal')).toBe('true');
    expect(dialog.hasAttribute('data-xlsx-dialog')).toBe(true);
    expect(field.value).toBe('General');
    expect(document.activeElement).toBe(field);
    expect([field.selectionStart, field.selectionEnd]).toEqual([0, 'General'.length]);
    await type(field, '#,##0.00 "kg"');
    expect(dialog.querySelector('output')?.textContent).toBe('1,234.50 kg');
    fireEvent.click(view.getByRole('button', { name: 'Apply' }));
    await nextTick();
    expect(view.queryByRole('dialog')).toBeNull();
    expect(viewer.snapshot().activeFormat?.numberFormat).toMatchObject({
      id: 164,
      code: '#,##0.00 "kg"',
    });
    expect(document.activeElement).toBe(grid);
    expect(viewer.undo().ok).toBe(true);
    expect(viewer.snapshot().activeFormat?.numberFormat.id).toBe(0);
  });

  test('Enter in the field applies', async () => {
    const { view, viewer } = await open();
    const { field } = await openDialog(view);
    await type(field, '0.0%');
    fireEvent.keyDown(field, { key: 'Enter' });
    await nextTick();
    expect(view.queryByRole('dialog')).toBeNull();
    expect(viewer.snapshot().activeFormat?.numberFormat).toMatchObject({ id: 164, code: '0.0%' });
  });

  test('a code that the store refuses shows its reason, is marked, and applies nothing', async () => {
    const { view, viewer } = await open();
    const { dialog, field } = await openDialog(view);
    await type(field, '"open');
    const reason =
      'The format code cannot be read. Check the sections, the quotes, and the square brackets.';
    expect(dialog.querySelector('output')?.textContent).toBe(reason);
    // The preview shows the reason, so the field is marked before Apply.
    expect(field.getAttribute('aria-invalid')).toBe('true');
    fireEvent.click(view.getByRole('button', { name: 'Apply' }));
    await nextTick();
    expect(view.getByRole('alert').textContent).toBe(reason);
    expect(dialog.querySelector<HTMLElement>('.xlsx-dialog__preview')?.hidden).toBe(true);
    expect(field.getAttribute('aria-invalid')).toBe('true');
    expect(field.getAttribute('aria-describedby')).toContain(view.getByRole('alert').id);
    expect(document.activeElement).toBe(field);
    expect(view.getByRole('dialog')).toBeTruthy();
    expect(viewer.snapshot().canUndo).toBe(false);
    await type(field, '[Red');
    expect(field.getAttribute('aria-invalid')).toBe('true');
    await type(field, '0.0');
    expect(field.getAttribute('aria-invalid')).toBe('false');
    await type(field, '"again');
    expect(field.getAttribute('aria-invalid')).toBe('true');
    // An empty field is marked only after an Apply.
    await type(field, '');
    expect(field.getAttribute('aria-invalid')).toBe('false');
    fireEvent.click(view.getByRole('button', { name: 'Apply' }));
    await nextTick();
    expect(view.getByRole('alert').textContent).toBe('Type a format code.');
  });

  test('the workbook codes are listed as text, and a pick fills the field', async () => {
    const hostile = '"<img src=x onerror=alert(1)>"0';
    const { view } = await open(
      workbook([
        '<numFmt numFmtId="165" formatCode="0.000"/>',
        `<numFmt numFmtId="166" formatCode="${hostile.replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}"/>`,
      ])
    );
    const { dialog, field } = await openDialog(view);
    const list = view.getByRole('listbox', {
      name: 'Formats in this workbook',
    }) as HTMLSelectElement;
    expect([...list.querySelectorAll('option')].map((o) => o.textContent)).toEqual([
      '0.000',
      hostile,
    ]);
    expect(dialog.querySelector('img')).toBeNull();
    change(list, '0.000');
    await nextTick();
    expect(field.value).toBe('0.000');
    expect(dialog.querySelector('output')?.textContent).toBe('1234.500');
  });

  test('no workbook code is selected until the field holds one, and the first one picks', async () => {
    const { view } = await open(
      workbook([
        '<numFmt numFmtId="165" formatCode="0.000"/>',
        '<numFmt numFmtId="166" formatCode="0.0000"/>',
      ])
    );
    const { field } = await openDialog(view);
    const list = view.getByRole('listbox', {
      name: 'Formats in this workbook',
    }) as HTMLSelectElement;
    expect(list.selectedIndex).toBe(-1);
    change(list, '0.000');
    await nextTick();
    expect(field.value).toBe('0.000');
    await type(field, '0.00000');
    expect(list.selectedIndex).toBe(-1);
    await type(field, '0.0000');
    expect(list.value).toBe('0.0000');
  });

  test('Apply of the unchanged code writes nothing, so a cell keeps an id with no code', async () => {
    const bytes = buildXlsx({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 5, { s: 1 })]) }],
      styles: {
        cellXfs: ['<xf numFmtId="44" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>'],
      },
    });
    const { view, viewer, grid } = await open(bytes);
    const { field } = await openDialog(view);
    expect(field.value).toBe('General');
    fireEvent.click(view.getByRole('button', { name: 'Apply' }));
    await nextTick();
    expect(view.queryByRole('dialog')).toBeNull();
    expect(viewer.snapshot().canUndo).toBe(false);
    expect(viewer.snapshot().activeFormat?.numberFormat.id).toBe(44);
    expect(document.activeElement).toBe(grid);
  });

  test('Apply of the unchanged code on a merged cell writes nothing', async () => {
    const bytes = buildXlsx({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 5)]), merges: ['A1:B2'] }],
    });
    const { view, viewer } = await open(bytes);
    expect(viewer.snapshot().selectionRef).toBe('A1:B2');
    const { field } = await openDialog(view);
    fireEvent.click(view.getByRole('button', { name: 'Apply' }));
    await nextTick();
    expect(view.queryByRole('dialog')).toBeNull();
    expect(field.value).toBe('General');
    expect(viewer.snapshot().canUndo).toBe(false);
  });

  test('a double-click on a workbook code applies it', async () => {
    const { view, viewer } = await open(workbook(['<numFmt numFmtId="165" formatCode="0.000"/>']));
    await openDialog(view);
    const list = view.getByRole('listbox', {
      name: 'Formats in this workbook',
    }) as HTMLSelectElement;
    change(list, '0.000');
    await nextTick();
    fireEvent.doubleClick(list.querySelector('option')!);
    await nextTick();
    expect(view.queryByRole('dialog')).toBeNull();
    expect(viewer.snapshot().activeFormat?.numberFormat).toMatchObject({ id: 165, code: '0.000' });
  });

  test('Enter in the list of workbook codes applies the picked code', async () => {
    const { view, viewer } = await open(workbook(['<numFmt numFmtId="165" formatCode="0.000"/>']));
    await openDialog(view);
    const list = view.getByRole('listbox', {
      name: 'Formats in this workbook',
    }) as HTMLSelectElement;
    change(list, '0.000');
    await nextTick();
    fireEvent.keyDown(list, { key: 'Enter' });
    await nextTick();
    expect(view.queryByRole('dialog')).toBeNull();
    expect(viewer.snapshot().activeFormat?.numberFormat).toMatchObject({ id: 165, code: '0.000' });
  });

  test('Apply of the unchanged code sets it on every cell of a range', async () => {
    const bytes = buildXlsx({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 1, { s: 1 })]) + row(2, [c('A2', 2)]) }],
      styles: {
        numFmts: ['<numFmt numFmtId="164" formatCode="0.000"/>'],
        cellXfs: ['<xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>'],
      },
    });
    const { view, viewer } = await open(bytes);
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    await nextTick();
    const { field } = await openDialog(view);
    expect(field.value).toBe('0.000');
    fireEvent.click(view.getByRole('button', { name: 'Apply' }));
    await nextTick();
    viewer.select({ row: 1, col: 0 });
    await nextTick();
    expect(viewer.snapshot().activeFormat?.numberFormat).toMatchObject({ id: 164, code: '0.000' });
  });

  test('an empty cell shows the sample note under the preview', async () => {
    const { view, viewer } = await open();
    viewer.select({ row: 5, col: 5 });
    await nextTick();
    const { field } = await openDialog(view);
    const note = view.getByText('The active cell is empty, so the preview shows a sample number.');
    expect(field.getAttribute('aria-describedby')).toContain(note.id);
  });

  test('a workbook with no custom codes says so', async () => {
    const { view } = await open();
    await openDialog(view);
    expect(view.getByText('This workbook has no custom formats.')).toBeTruthy();
    expect(view.getByText('Type a format code, such as #,##0.00 or 0.0%.')).toBeTruthy();
  });

  test('Escape closes it, and the keyboard goes back to the list trigger that opened it', async () => {
    const { view, viewer } = await open();
    const trigger = view.getByRole('button', { name: /^Number format/ });
    trigger.focus();
    fireEvent.keyDown(trigger, { key: 'ArrowDown' });
    await nextTick();
    const more = view.getByRole('button', { name: 'More number formats…' });
    await waitFor(() => expect(document.activeElement?.getAttribute('aria-selected')).toBe('true'));
    more.focus();
    fireEvent.click(more);
    await nextTick();
    const { field } = {
      field: view.getByRole('textbox', { name: 'Format code' }) as HTMLInputElement,
    };
    fireEvent.keyDown(field, { key: 'Escape' });
    await nextTick();
    expect(view.queryByRole('dialog')).toBeNull();
    expect(document.activeElement).toBe(view.getByRole('button', { name: /^Number format/ }));
    expect(viewer.snapshot().canUndo).toBe(false);
  });

  test('a refused start from the keyboard opens nothing, and the keyboard goes to the grid', async () => {
    const { view, viewer } = await open();
    viewer.beginEdit('=SUM((');
    await nextTick();
    const trigger = view.getByRole('button', { name: /^Number format/ });
    trigger.focus();
    fireEvent.keyDown(trigger, { key: 'ArrowDown' });
    await nextTick();
    const more = view.getByRole('button', { name: 'More number formats…' });
    await waitFor(() => expect(document.activeElement?.getAttribute('aria-selected')).toBe('true'));
    more.focus();
    fireEvent.click(more);
    await nextTick();
    expect(view.queryByRole('dialog')).toBeNull();
    expect(viewer.snapshot().editing).toBe(true);
    expect(document.activeElement).not.toBe(document.body);
    expect(view.container.contains(document.activeElement)).toBe(true);
  });

  test('a selection that a host moves closes it', async () => {
    const { view, viewer } = await open();
    await openDialog(view);
    viewer.select({ row: 4, col: 3 });
    await nextTick();
    expect(view.queryByRole('dialog')).toBeNull();
    expect(viewer.snapshot().canUndo).toBe(false);
  });

  test('a change to view mode closes it', async () => {
    const { view, viewer } = await open();
    await openDialog(view);
    viewer.setMode('view');
    await nextTick();
    expect(view.queryByRole('dialog')).toBeNull();
  });

  test('a load of another workbook closes it and applies nothing', async () => {
    const { view, viewer } = await open();
    const { field } = await openDialog(view);
    await type(field, '0.0%');
    await viewer.load(workbook(['<numFmt numFmtId="165" formatCode="0.000"/>']));
    await nextTick();
    await waitFor(() => expect(view.queryByRole('dialog')).toBeNull());
    expect(viewer.snapshot().canUndo).toBe(false);
    expect(viewer.snapshot().activeFormat?.numberFormat.id).toBe(0);
  });

  test('Tab and Shift+Tab wrap inside the dialog', async () => {
    const { view } = await open();
    const { field } = await openDialog(view);
    const apply = view.getByRole('button', { name: 'Apply' });
    apply.focus();
    fireEvent.keyDown(apply, { key: 'Tab' });
    await nextTick();
    expect(document.activeElement).toBe(field);
    fireEvent.keyDown(field, { key: 'Tab', shiftKey: true });
    await nextTick();
    expect(document.activeElement).toBe(apply);
  });

  test('a press on the backdrop closes it, and a drag out of the field does not', async () => {
    const { view, viewer } = await open();
    const { dialog, field } = await openDialog(view);
    // The panel's box is empty in happy-dom, so any point on the dialog element is outside it.
    fireEvent.mouseDown(field, { clientX: 5, clientY: 5 });
    await nextTick();
    fireEvent.click(dialog, { clientX: 500, clientY: 500 });
    await nextTick();
    expect(view.getByRole('dialog')).toBeTruthy();
    fireEvent.mouseDown(dialog, { clientX: 500, clientY: 500 });
    await nextTick();
    fireEvent.click(dialog, { clientX: 500, clientY: 500 });
    await nextTick();
    expect(view.queryByRole('dialog')).toBeNull();
    expect(viewer.snapshot().canUndo).toBe(false);
  });

  test('a code that a host adds while the dialog is open shows in the list', async () => {
    const { view, viewer } = await open();
    await openDialog(view);
    expect(view.queryByRole('listbox')).toBeNull();
    viewer.applyFormat({ numberFormat: { code: '0.000' } });
    await nextTick();
    const list = view.getByRole('listbox', {
      name: 'Formats in this workbook',
    }) as HTMLSelectElement;
    expect([...list.querySelectorAll('option')].map((o) => o.textContent)).toContain('0.000');
  });

  test('with no listed code selected, an arrow key starts at an end of the list', async () => {
    const { view } = await open(
      workbook([
        '<numFmt numFmtId="164" formatCode="0.0"/>',
        '<numFmt numFmtId="165" formatCode="0.000"/>',
      ])
    );
    const { field } = await openDialog(view);
    const list = view.getByRole('listbox', {
      name: 'Formats in this workbook',
    }) as HTMLSelectElement;
    await type(field, '#');
    expect(list.selectedIndex).toBe(-1);
    fireEvent.keyDown(list, { key: 'ArrowUp' });
    await nextTick();
    expect(field.value).toBe('0.000');
    await type(field, '#');
    fireEvent.keyDown(list, { key: 'ArrowDown' });
    await nextTick();
    expect(field.value).toBe('0.0');
  });

  test('a code that hides the value gets a note, and a blank code is refused', async () => {
    const { view } = await open();
    const { field } = await openDialog(view);
    await type(field, ';;;');
    expect(view.getByText('The code shows no text for this value.')).toBeTruthy();
    await type(field, '   ');
    expect(view.queryByText('The code shows no text for this value.')).toBeNull();
    expect(view.getByText('Type a format code.')).toBeTruthy();
  });

  test('keys in the dialog do not reach the grid', async () => {
    const { view, viewer } = await open();
    const { field } = await openDialog(view);
    fireEvent.keyDown(field, { key: 'b', ctrlKey: true });
    await nextTick();
    expect(viewer.snapshot().activeFormat?.bold).toBe(false);
  });

  test('with the slot hidden, the number format list has no item for the dialog', async () => {
    const { view } = await open(workbook(), ['number.custom']);
    fireEvent.click(view.getByRole('button', { name: /^Number format/ }));
    await nextTick();
    expect(view.queryByRole('button', { name: 'More number formats…' })).toBeNull();
  });

  test('without the number format list, the slot is an icon button that opens it', async () => {
    const { view } = await open(workbook(), ['number.format']);
    const button = view.getByRole('button', { name: 'More number formats…' });
    expect(button.getAttribute('aria-haspopup')).toBe('dialog');
    expect(button.querySelector('svg')).toBeTruthy();
    fireEvent.click(button);
    await nextTick();
    expect(view.getByRole('dialog', { name: 'Number format' })).toBeTruthy();
  });
});
