// The editable formula bar in happy-dom, the twin of the React adapter's
// `formula-bar.test.tsx`: the content field edits the one open cell edit (the grid's cell editor
// shows the same text), Enter, Tab, and Escape act as in the cell editor, an input method holds
// Enter, a refused commit keeps the text and the keyboard, and the name box selects a typed
// reference or shows why it cannot.
//
// Differences from React: a Vue field listens to `input`, so `typeIn` sets the value and fires
// `input` where React's test fires `change`; each step awaits the render that Vue runs after it.

import { afterEach, describe, expect, test } from 'bun:test';
import { fireEvent, waitFor } from '@testing-library/react';
import { h, nextTick, reactive } from 'vue';
import { buildXlsx, c, row } from '../../../scripts/fixtures/builder.ts';
import {
  FormulaBar,
  MAX_CONTENT_FIELD_CHARS,
  StatusMessage,
  XlsxEditor,
  XlsxEditorContent,
  XlsxEditorRoot,
  type XlsxEditorInstance,
} from '../src/index.ts';
import { FONTS, mount, tabsShow, unmountAll } from './mount.ts';

afterEach(unmountAll);

const workbook = buildXlsx({
  sheets: [
    { name: 'First', rows: row(1, [c('A1', 'one'), c('B1', 'a\nb')]) },
    { name: 'Second', rows: row(1, [c('A1', 'two')]) },
  ],
  definedNames:
    '<definedName name="Rates">Second!$B$2</definedName>' +
    '<definedName name="Total">SUM(First!$A$1)</definedName>',
});

async function open(mode: 'edit' | 'view' = 'edit', bytes: Uint8Array = workbook) {
  let viewer: XlsxEditorInstance | null = null;
  const view = mount(() =>
    h('div', { style: { height: '400px' } }, [
      h(XlsxEditor, {
        document: bytes,
        fonts: FONTS,
        mode,
        onReady: (v: XlsxEditorInstance) => (viewer = v),
      }),
    ])
  );
  await tabsShow(view);
  const content = view.getByLabelText('Cell content') as HTMLTextAreaElement;
  const nameBox = view.getByLabelText('Cell reference') as HTMLInputElement;
  const cellEditor = view.container.querySelector('.xlsx-cell-editor') as HTMLTextAreaElement;
  const grid = view.container.querySelector('.xlsx-scroller') as HTMLElement;
  return { view, viewer: viewer!, content, nameBox, cellEditor, grid };
}

/** Type `text` into the field as one input, and let Vue render it. */
async function typeIn(field: HTMLInputElement | HTMLTextAreaElement, text: string): Promise<void> {
  field.value = text;
  fireEvent.input(field);
  await nextTick();
  await nextTick();
}

const strip = (view: { container: Element } | Element) =>
  ('container' in view ? view.container : view).querySelector('.xlsx-refusal')?.textContent ?? '';

const longText = (length: number) =>
  buildXlsx({
    sheets: [
      { name: 'First', rows: row(1, [c('A1', 'a'.repeat(length))]) },
      { name: 'Second', rows: row(1, [c('A1', 'two')]) },
    ],
  });

describe('the content field with a very long text', () => {
  test('holds only the start of the text, read-only, and says so', async () => {
    const full = MAX_CONTENT_FIELD_CHARS + 5;
    const { view, viewer, content } = await open('edit', longText(full));
    expect(content.value.length).toBe(MAX_CONTENT_FIELD_CHARS);
    expect(content.readOnly).toBe(true);
    const note = view.container.querySelector('.xlsx-cell-content-lines');
    expect(note?.textContent).toContain('Press F2 in the grid');
    expect(note?.textContent).toContain(
      'The edit applies only when the text is not too long for one cell.'
    );
    expect(content.getAttribute('aria-describedby')).toBe(note?.id ?? '');
    // In a narrow bar the note gives way and ends in an ellipsis, so its title holds all of it.
    expect(note?.classList.contains('xlsx-cell-content-note')).toBe(true);
    expect(note?.getAttribute('title')).toBe(note?.textContent ?? '');
    expect(viewer.snapshot().activeCell?.content.length).toBe(full);
  });

  test('an open edit keeps the complete text through the keys of the field', async () => {
    const full = MAX_CONTENT_FIELD_CHARS + 5;
    const { viewer, content } = await open('edit', longText(full));
    viewer.beginEdit();
    await nextTick();
    expect(content.value.length).toBe(MAX_CONTENT_FIELD_CHARS);
    content.focus();
    fireEvent.keyDown(content, { key: 'Enter', altKey: true });
    await nextTick();
    await typeIn(content, 'short');
    expect(viewer.snapshot().editText?.length).toBe(full);
    fireEvent.compositionStart(content);
    await nextTick();
    fireEvent.blur(content);
    await nextTick();
    expect(viewer.snapshot().editText?.length).toBe(full);
    fireEvent.keyDown(content, { key: 'Enter' });
    await nextTick();
    viewer.select({ row: 0, col: 0 });
    await nextTick();
    expect(viewer.snapshot().activeCell?.content.length).toBe(full);
  });

  test('a text of the largest length stays complete and editable', async () => {
    const { view, content } = await open('edit', longText(MAX_CONTENT_FIELD_CHARS));
    expect(content.value.length).toBe(MAX_CONTENT_FIELD_CHARS);
    expect(content.readOnly).toBe(false);
    expect(view.container.querySelector('.xlsx-cell-content-lines')).toBeNull();
  });

  test('a commit past the limit of a cell value names the limit', async () => {
    const { view, viewer, content } = await open();
    content.focus();
    await typeIn(content, 'a'.repeat(32_768));
    fireEvent.keyDown(content, { key: 'Enter' });
    await nextTick();
    expect(viewer.snapshot().editing).toBe(true);
    expect(strip(view)).toContain('the text is too long for one cell');
  });

  test('the limit counts UTF-16 code units, also for a text outside the Basic Multilingual Plane', async () => {
    // Each character outside the plane takes two code units: 20,000 characters are 40,000 units.
    const outside = '\u{10400}';
    const { view, viewer, content } = await open();
    content.focus();
    await typeIn(content, outside.repeat(20_000));
    fireEvent.keyDown(content, { key: 'Enter' });
    await nextTick();
    expect(viewer.snapshot().editing).toBe(true);
    expect(strip(view)).toBe(
      'The edit was not applied: the text is too long for one cell. Shorten the text, then try again.'
    );
    // A save commits the open edit first, so it is refused for the same text.
    viewer.save();
    await nextTick();
    expect(viewer.snapshot().refusal).toMatchObject({ code: 'text-too-long', command: 'save' });
    expect(strip(view)).toBe(
      'The workbook was not saved: the text of the cell edit is too long for one cell. Shorten the text, then save.'
    );
    // 16,383 such characters are 32,766 code units, inside the limit: the edit applies.
    await typeIn(content, outside.repeat(16_383));
    fireEvent.keyDown(content, { key: 'Enter' });
    await nextTick();
    expect(viewer.snapshot().editing).toBe(false);
    viewer.select({ row: 0, col: 0 });
    await nextTick();
    expect(viewer.snapshot().activeCell?.content.length).toBe(32_766);
  });
});

describe('the content field', () => {
  test('is a text field with the active cell content, and keeps its line breaks', async () => {
    const { viewer, content } = await open();
    expect(content.tagName).toBe('TEXTAREA');
    expect(content.value).toBe('one');
    viewer.select({ row: 0, col: 1 });
    await nextTick();
    expect(content.value).toBe('a\nb');
  });

  test('typing opens the cell edit, and the grid shows the same text live', async () => {
    const { viewer, content, cellEditor } = await open();
    content.focus();
    await typeIn(content, 'one more');
    expect(viewer.snapshot().editing).toBe(true);
    expect(viewer.snapshot().editText).toBe('one more');
    expect(cellEditor.value).toBe('one more');
    expect(cellEditor.hidden).toBe(false);
    expect(document.activeElement).toBe(content);
    // A key in the cell editor shows in the bar: one edit, one text.
    cellEditor.value = 'one more!';
    fireEvent.input(cellEditor);
    await nextTick();
    expect(content.value).toBe('one more!');
  });

  test('Enter commits and moves down; Tab commits and moves right', async () => {
    const { viewer, content, grid } = await open();
    content.focus();
    await typeIn(content, '=1+2');
    fireEvent.keyDown(content, { key: 'Enter' });
    await nextTick();
    expect(viewer.snapshot().selectionRef).toBe('A2');
    expect(document.activeElement).toBe(grid);
    viewer.select({ row: 0, col: 0 });
    await nextTick();
    expect(content.value).toBe('=1+2');
    await typeIn(content, '5');
    fireEvent.keyDown(content, { key: 'Tab' });
    await nextTick();
    expect(viewer.snapshot().selectionRef).toBe('B1');
    fireEvent.keyDown(content, { key: 'Enter', shiftKey: true });
    await nextTick();
    fireEvent.keyDown(content, { key: 'Tab', shiftKey: true });
    await nextTick();
    expect(viewer.snapshot().selectionRef).toBe('A1');
    expect(content.value).toBe('5');
  });

  test('Escape cancels the edit and gives the grid the keyboard', async () => {
    const { viewer, content, grid } = await open();
    content.focus();
    await typeIn(content, 'changed');
    fireEvent.keyDown(content, { key: 'Escape' });
    await nextTick();
    expect(viewer.snapshot().editing).toBe(false);
    expect(content.value).toBe('one');
    expect(document.activeElement).toBe(grid);
    // With no open edit, Escape gives the grid the keyboard too.
    content.focus();
    fireEvent.keyDown(content, { key: 'Escape' });
    await nextTick();
    expect(document.activeElement).toBe(grid);
  });

  test('blur ends a composition that had no end event', async () => {
    const { viewer, content } = await open();
    content.focus();
    fireEvent.compositionStart(content);
    await nextTick();
    await typeIn(content, 'か');
    expect(viewer.snapshot().composing).toBe(true);
    fireEvent.blur(content);
    await nextTick();
    expect(viewer.snapshot().composing).toBe(false);
    expect(viewer.commitEdit().ok).toBe(true);
  });

  test('an unmount ends a composition that had no end event', async () => {
    const { view, viewer, content } = await open();
    content.focus();
    fireEvent.compositionStart(content);
    await nextTick();
    await typeIn(content, 'か');
    view.unmount();
    expect(viewer.snapshot().composing).toBe(false);
  });

  test('a refused F6 commit marks the field invalid with its reason', async () => {
    const { viewer, content } = await open();
    content.focus();
    await typeIn(content, '=1+');
    fireEvent.keyDown(content, { key: 'F6' });
    await nextTick();
    expect(viewer.snapshot().editing).toBe(true);
    expect(content.getAttribute('aria-invalid')).toBe('true');
  });

  test('a composing input method holds Enter, and its text commits after it ends', async () => {
    const { viewer, content } = await open();
    content.focus();
    fireEvent.compositionStart(content);
    await nextTick();
    await typeIn(content, 'か');
    expect(viewer.snapshot().composing).toBe(true);
    fireEvent.keyDown(content, { key: 'Enter', keyCode: 229 });
    await nextTick();
    fireEvent.keyDown(content, { key: 'Enter' });
    await nextTick();
    expect(viewer.snapshot().editing).toBe(true);
    content.value = '漢';
    fireEvent.compositionEnd(content);
    await nextTick();
    expect(viewer.snapshot().composing).toBe(false);
    expect(viewer.snapshot().editText).toBe('漢');
    fireEvent.keyDown(content, { key: 'Enter' });
    await nextTick();
    expect(viewer.snapshot().editing).toBe(false);
    viewer.select({ row: 0, col: 0 });
    await nextTick();
    expect(content.value).toBe('漢');
  });

  test('a refused commit keeps the text, the field, and shows the reason', async () => {
    const { view, viewer, content } = await open();
    content.focus();
    await typeIn(content, '=1+');
    fireEvent.keyDown(content, { key: 'Enter' });
    await nextTick();
    expect(viewer.snapshot().editing).toBe(true);
    expect(content.value).toBe('=1+');
    expect(content.getAttribute('aria-invalid')).toBe('true');
    expect(document.activeElement).toBe(content);
    expect(strip(view)).not.toBe('');
  });

  test('the invalid field names its reason, and is valid again once its text changes', async () => {
    const { content } = await open();
    content.focus();
    await typeIn(content, '=1+');
    fireEvent.keyDown(content, { key: 'Enter' });
    await nextTick();
    const ids = content.getAttribute('aria-describedby')!.split(' ');
    const reason = document.getElementById(ids[0]!)!.textContent;
    expect(reason).toBe(strip(document.body));
    expect(reason).toContain('formula');
    await typeIn(content, '=1+2');
    expect(content.getAttribute('aria-invalid')).toBeNull();
    expect(content.getAttribute('aria-describedby')!.split(' ')).toHaveLength(1);
  });

  test('Alt+Enter adds a line break', async () => {
    const { viewer, content } = await open();
    content.focus();
    content.setSelectionRange(3, 3);
    fireEvent.keyDown(content, { key: 'Enter', altKey: true });
    await nextTick();
    expect(viewer.snapshot().editText).toBe('one\n');
    expect(content.value).toBe('one\n');
    expect(content.selectionStart).toBe(4);
    // A text of more than one line shows its line count.
    expect(document.querySelector('.xlsx-cell-content-lines')?.textContent).toBe('2 lines');
  });

  test('the field scrolls to the whole line of the caret', async () => {
    const { content } = await open();
    content.style.lineHeight = '24px';
    content.focus();
    content.setSelectionRange(3, 3);
    fireEvent.keyDown(content, { key: 'Enter', altKey: true });
    await nextTick();
    expect(content.scrollTop).toBe(24);
    await typeIn(content, 'one\ntwo\nthree');
    content.setSelectionRange(5, 5);
    fireEvent.select(content);
    await nextTick();
    expect(content.scrollTop).toBe(24);
    content.setSelectionRange(0, 0);
    fireEvent.select(content);
    await nextTick();
    expect(content.scrollTop).toBe(0);
  });

  test('a forward selection shows the line of its moving end', async () => {
    const { content } = await open();
    content.style.lineHeight = '24px';
    content.focus();
    await typeIn(content, 'one\ntwo\nthree');
    content.setSelectionRange(1, 6, 'forward');
    fireEvent.select(content);
    await nextTick();
    expect(content.scrollTop).toBe(24);
  });

  test('a refusal from before the edit does not mark the new edit invalid', async () => {
    const { viewer, content } = await open();
    viewer.renameSheet(0, 'Second');
    await nextTick();
    expect(viewer.snapshot().refusal).not.toBeNull();
    content.focus();
    await typeIn(content, 'x');
    expect(viewer.snapshot().editing).toBe(true);
    expect(content.getAttribute('aria-invalid')).toBeNull();
  });

  test('a bar that mounts during an edit is invalid only for a refusal of its text', async () => {
    let viewer: XlsxEditorInstance | null = null;
    const state = reactive({ bar: false });
    const layout = async (bar: boolean) => {
      state.bar = bar;
      await nextTick();
    };
    const view = mount(() =>
      h('div', { style: { height: '400px' } }, [
        h(
          XlsxEditorRoot,
          { document: workbook, fonts: FONTS, onReady: (v: XlsxEditorInstance) => (viewer = v) },
          {
            default: () => [
              state.bar ? h(FormulaBar) : null,
              h(XlsxEditorContent),
              h(StatusMessage),
            ],
          }
        ),
      ])
    );
    await waitFor(() => expect(viewer?.snapshot().status).toBe('ready'), { timeout: 5000 });
    // A refused rename, then an edit: the edit opens without the earlier reason.
    viewer!.renameSheet(0, '');
    await nextTick();
    viewer!.beginEdit('x');
    await nextTick();
    expect(viewer!.snapshot().refusal).toBeNull();
    await layout(true);
    const content = view.getByLabelText('Cell content') as HTMLTextAreaElement;
    expect(content.getAttribute('aria-invalid')).toBeNull();
    // A refused commit, then a new bar: the bar shows the refused text as invalid.
    viewer!.setEditText('=1+');
    await nextTick();
    viewer!.commitEdit();
    await nextTick();
    await layout(false);
    await layout(true);
    const again = view.getByLabelText('Cell content') as HTMLTextAreaElement;
    expect(again.getAttribute('aria-invalid')).toBe('true');
  });

  test('Alt+Enter inserts the break as typed text when the document can', async () => {
    // A document with the insert command: it types the text, as a browser does.
    const doc = document as unknown as { execCommand?: unknown };
    const setValue = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!;
    doc.execCommand = (_: string, __: boolean, text: string) => {
      const el = document.activeElement as HTMLTextAreaElement;
      const at = el.selectionStart;
      setValue.call(el, `${el.value.slice(0, at)}${text}${el.value.slice(el.selectionEnd)}`);
      el.setSelectionRange(at + text.length, at + text.length);
      el.dispatchEvent(new Event('input', { bubbles: true }));
      return true;
    };
    try {
      const { viewer, content } = await open();
      content.focus();
      content.setSelectionRange(1, 1);
      fireEvent.keyDown(content, { key: 'Enter', altKey: true });
      await nextTick();
      expect(viewer.snapshot().editText).toBe('o\nne');
      expect(content.value).toBe('o\nne');
      // A refused edit gives the field back its text.
      unmountAll();
      const locked = buildXlsx({
        sheets: [
          {
            name: 'First',
            rows: row(1, [c('A1', 'one')]),
            beforeMerges: '<sheetProtection sheet="1"/>',
          },
          { name: 'Second', rows: row(1, [c('A1', 'two')]) },
        ],
      });
      const shut = await open('edit', locked);
      shut.content.focus();
      shut.content.setSelectionRange(1, 1);
      fireEvent.keyDown(shut.content, { key: 'Enter', altKey: true });
      await nextTick();
      expect(shut.viewer.snapshot().refusal?.code).toBe('locked');
      expect(shut.content.value).toBe('one');
    } finally {
      delete doc.execCommand;
    }
  });

  test('a refused Alt+Enter leaves the field as it was', async () => {
    const locked = buildXlsx({
      sheets: [
        {
          name: 'First',
          rows: row(1, [c('A1', 'one')]),
          beforeMerges: '<sheetProtection sheet="1"/>',
        },
        { name: 'Second', rows: row(1, [c('A1', 'two')]) },
      ],
    });
    const { viewer, content } = await open('edit', locked);
    content.focus();
    content.setSelectionRange(3, 3);
    fireEvent.keyDown(content, { key: 'Enter', altKey: true });
    await nextTick();
    expect(viewer.snapshot().editing).toBe(false);
    expect(viewer.snapshot().refusal?.code).toBe('locked');
    expect(content.value).toBe('one');
  });

  test('history keys with no open edit step the workbook history', async () => {
    const { viewer, content, nameBox } = await open();
    viewer.enterValue('changed');
    await nextTick();
    content.focus();
    fireEvent.keyDown(content, { key: 'z', ctrlKey: true });
    await nextTick();
    expect(content.value).toBe('one');
    nameBox.focus();
    fireEvent.keyDown(nameBox, { key: 'y', ctrlKey: true });
    await nextTick();
    expect(content.value).toBe('changed');
    // With a typed reference, the keys belong to the name box.
    await typeIn(nameBox, 'B');
    fireEvent.keyDown(nameBox, { key: 'z', ctrlKey: true });
    await nextTick();
    expect(content.value).toBe('changed');
  });

  test('history keys in the name box leave an open cell edit as it is', async () => {
    const { viewer, content, nameBox } = await open();
    content.focus();
    await typeIn(content, 'oneXYZ');
    nameBox.focus();
    fireEvent.keyDown(nameBox, { key: 'z', ctrlKey: true });
    await nextTick();
    expect([viewer.snapshot().editing, viewer.snapshot().editText]).toEqual([true, 'oneXYZ']);
  });

  test('F6 visits the formula bar, and commits its open edit first', async () => {
    const { viewer, content, nameBox, grid } = await open();
    grid.focus();
    fireEvent.keyDown(grid, { key: 'F6', shiftKey: true });
    await nextTick();
    expect(document.activeElement).toBe(nameBox);
    content.focus();
    await typeIn(content, 'moved on');
    fireEvent.keyDown(content, { key: 'F6' });
    await nextTick();
    expect(viewer.snapshot().editing).toBe(false);
    expect(viewer.snapshot().activeCell?.content).toBe('moved on');
    expect(document.activeElement).toBe(grid);
    // A refused commit keeps the keyboard in the bar.
    content.focus();
    await typeIn(content, '=1+');
    fireEvent.keyDown(content, { key: 'F6' });
    await nextTick();
    expect(viewer.snapshot().editing).toBe(true);
    expect(document.activeElement).toBe(content);
  });

  test('a press on the line count puts the caret at the end of the shown line', async () => {
    const { viewer, content, grid } = await open();
    viewer.select({ row: 0, col: 1 });
    await nextTick();
    grid.focus();
    const mark = document.querySelector('.xlsx-cell-content-lines') as HTMLElement;
    expect(mark.textContent).toBe('2 lines');
    fireEvent.mouseDown(mark);
    await nextTick();
    expect(document.activeElement).toBe(content);
    expect([content.selectionStart, content.selectionEnd]).toEqual([1, 1]);
    expect(viewer.snapshot().selectionRef).toBe('B1');
  });

  test('a press on the line count of a field scrolled to line 2 puts the caret at its end', async () => {
    const { viewer, content, grid } = await open();
    viewer.select({ row: 0, col: 1 });
    await nextTick();
    grid.focus();
    // The test document has no layout: the field shows line 2 through its line height and scroll.
    content.style.lineHeight = '20px';
    Object.defineProperty(content, 'scrollTop', {
      configurable: true,
      get: () => 20,
      set: () => {},
    });
    fireEvent.mouseDown(document.querySelector('.xlsx-cell-content-lines') as HTMLElement);
    await nextTick();
    expect(document.activeElement).toBe(content);
    expect([content.selectionStart, content.selectionEnd]).toEqual([3, 3]);
  });

  test('Shift and a press on the line count extend the text selection', async () => {
    const { viewer, content } = await open();
    viewer.select({ row: 0, col: 1 });
    await nextTick();
    content.focus();
    content.setSelectionRange(0, 0);
    fireEvent.mouseDown(document.querySelector('.xlsx-cell-content-lines') as HTMLElement, {
      shiftKey: true,
    });
    expect([content.selectionStart, content.selectionEnd]).toEqual([0, 1]);
    expect(content.selectionDirection).toBe('forward');
    // From a caret after the end of the shown line, the selection runs backward to it.
    content.setSelectionRange(3, 3);
    fireEvent.mouseDown(document.querySelector('.xlsx-cell-content-lines') as HTMLElement, {
      shiftKey: true,
    });
    expect([content.selectionStart, content.selectionEnd]).toEqual([1, 3]);
    expect(content.selectionDirection).toBe('backward');
    // A backward selection keeps its end as the anchor.
    content.setSelectionRange(0, 3, 'backward');
    fireEvent.mouseDown(document.querySelector('.xlsx-cell-content-lines') as HTMLElement, {
      shiftKey: true,
    });
    expect([content.selectionStart, content.selectionEnd]).toEqual([1, 3]);
    expect(content.selectionDirection).toBe('backward');
  });

  test('Shift and a press on the line count put the caret when the field has no keyboard', async () => {
    const { viewer, content, grid } = await open();
    viewer.select({ row: 0, col: 1 });
    await nextTick();
    content.setSelectionRange(0, 0);
    grid.focus();
    fireEvent.mouseDown(document.querySelector('.xlsx-cell-content-lines') as HTMLElement, {
      shiftKey: true,
    });
    expect(document.activeElement).toBe(content);
    expect([content.selectionStart, content.selectionEnd]).toEqual([1, 1]);
  });

  test('a press on the field keeps the selection', async () => {
    const { viewer, content } = await open();
    viewer.select({ row: 0, col: 1 });
    await nextTick();
    const pressed = fireEvent.mouseDown(content);
    expect(pressed).toBe(true);
    expect(viewer.snapshot().selectionRef).toBe('B1');
  });

  test('is read-only in view mode', async () => {
    const { content } = await open('view');
    expect(content.readOnly).toBe(true);
  });
});

describe('the name box', () => {
  test('Enter selects a typed cell, range, or sheet reference', async () => {
    const { viewer, nameBox, grid } = await open();
    nameBox.focus();
    await typeIn(nameBox, 'b2:c5');
    fireEvent.keyDown(nameBox, { key: 'Enter' });
    await nextTick();
    expect(viewer.snapshot().selectionRef).toBe('B2:C5');
    expect(nameBox.value).toBe('B2:C5');
    expect(document.activeElement).toBe(grid);
    nameBox.focus();
    await typeIn(nameBox, 'Second!A3');
    fireEvent.keyDown(nameBox, { key: 'Enter' });
    await nextTick();
    expect(viewer.snapshot().activeSheet).toBe(1);
    expect(viewer.snapshot().selectionRef).toBe('A3');
    nameBox.focus();
    await typeIn(nameBox, 'rates');
    fireEvent.keyDown(nameBox, { key: 'Enter' });
    await nextTick();
    expect([viewer.snapshot().activeSheet, viewer.snapshot().selectionRef]).toEqual([1, 'B2']);
  });

  test('a refused reference keeps the text, marks the field, and shows why', async () => {
    const { view, viewer, nameBox } = await open();
    nameBox.focus();
    await typeIn(nameBox, 'Total');
    fireEvent.keyDown(nameBox, { key: 'Enter' });
    await nextTick();
    expect(nameBox.value).toBe('Total');
    expect(nameBox.getAttribute('aria-invalid')).toBe('true');
    expect(document.activeElement).toBe(nameBox);
    expect(strip(view)).toContain('defined name Total is not one range');
    const ids = nameBox.getAttribute('aria-describedby')!.split(' ');
    expect(document.getElementById(ids[0]!)!.textContent).toBe(strip(view));
    await typeIn(nameBox, 'nowhere!');
    expect(nameBox.getAttribute('aria-invalid')).toBeNull();
    fireEvent.keyDown(nameBox, { key: 'Enter' });
    await nextTick();
    expect(strip(view)).toContain('type a cell, a range, whole columns or rows');
    // Escape drops the typed text and the reason, and the grid takes the keyboard.
    fireEvent.keyDown(nameBox, { key: 'Escape' });
    await nextTick();
    expect(nameBox.value).toBe('A1');
    expect(viewer.snapshot().refusal).toBeNull();
    expect(nameBox.getAttribute('aria-invalid')).toBeNull();
  });

  test('Enter in an empty name box selects nothing and is no error', async () => {
    const { viewer, nameBox, grid } = await open();
    nameBox.focus();
    await typeIn(nameBox, '  ');
    fireEvent.keyDown(nameBox, { key: 'Enter' });
    await nextTick();
    expect(nameBox.value).toBe('A1');
    expect(nameBox.getAttribute('aria-invalid')).toBeNull();
    expect(viewer.snapshot().refusal).toBeNull();
    expect(document.activeElement).toBe(grid);
  });

  test('Escape drops the reason of a refused reference, also after more typing', async () => {
    const { view, viewer, nameBox } = await open();
    nameBox.focus();
    await typeIn(nameBox, 'foo bar');
    fireEvent.keyDown(nameBox, { key: 'Enter' });
    await nextTick();
    expect(viewer.snapshot().refusal?.code).toBe('go-to-reference');
    await typeIn(nameBox, 'foo barx');
    fireEvent.keyDown(nameBox, { key: 'Escape' });
    await nextTick();
    expect(viewer.snapshot().refusal).toBeNull();
    expect(strip(view)).toBe('');
  });

  test('blur drops the typed reference and its reason', async () => {
    const { view, viewer, nameBox } = await open();
    nameBox.focus();
    await typeIn(nameBox, 'nope!A1');
    fireEvent.keyDown(nameBox, { key: 'Enter' });
    await nextTick();
    expect(viewer.snapshot().refusal?.code).toBe('go-to-sheet');
    fireEvent.blur(nameBox);
    await nextTick();
    expect(nameBox.value).toBe('A1');
    expect(nameBox.getAttribute('aria-invalid')).toBeNull();
    expect(viewer.snapshot().refusal).toBeNull();
    expect(strip(view)).toBe('');
  });

  test('a refused commit of the open edit is not a refused reference', async () => {
    const { viewer, content, nameBox } = await open();
    content.focus();
    await typeIn(content, '=1+');
    fireEvent.keyDown(content, { key: 'Enter' });
    await nextTick();
    nameBox.focus();
    await typeIn(nameBox, 'B2');
    fireEvent.keyDown(nameBox, { key: 'Enter' });
    await nextTick();
    expect(viewer.snapshot().refusal?.code).toBe('formula-syntax');
    expect(nameBox.getAttribute('aria-invalid')).toBeNull();
    expect(document.activeElement).toBe(content);
    expect(content.getAttribute('aria-invalid')).toBe('true');
  });

  test('a composing input method holds Enter', async () => {
    const { viewer, nameBox } = await open();
    nameBox.focus();
    await typeIn(nameBox, 'C3');
    fireEvent.keyDown(nameBox, { key: 'Enter', keyCode: 229 });
    await nextTick();
    expect(viewer.snapshot().selectionRef).toBe('A1');
  });
});
