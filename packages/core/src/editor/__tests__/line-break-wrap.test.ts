// Text that the cell editor commits with a line break (Alt+Enter) turns on `@wrapText` of the
// cell (`x:alignment`, `CT_CellAlignment`), in the same transaction as the value, so each line
// shows and one undo restores both. A value that a command writes keeps the cell format. The
// engine is driven directly, in happy-dom.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { createXlsxEditor, type XlsxEditorInstance } from '../index.ts';
import { flushEditor } from './flush-editor.ts';
import type { FontSource } from '../../layout/index.ts';

const fontsDir = join(
  dirname(Bun.resolveSync('@docx-editor.dev/fonts/package.json', import.meta.dir)),
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

let editors: XlsxEditorInstance[] = [];
afterEach(() => {
  for (const e of editors) e.destroy();
  editors = [];
});

const WRAPPED =
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf>';

async function mount() {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.appendChild(host);
  const editor = createXlsxEditor({ fonts: FONTS });
  editors.push(editor);
  editor.attach(host);
  await editor.load(
    buildXlsx({
      styles: { cellXfs: [WRAPPED] },
      sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 'old', { s: 1 })]) }],
    })
  );
  flushEditor(editor);
  const input = host.querySelector('.xlsx-cell-editor') as HTMLTextAreaElement;
  return { editor, win, input };
}

const wraps = (e: XlsxEditorInstance) => e.snapshot().activeFormat?.wrapText;

describe('a line break that the cell editor commits', () => {
  test('turns on wrap text, and one undo restores the value and the format', async () => {
    const { editor } = await mount();
    editor.select({ row: 1, col: 0 });
    expect(wraps(editor)).toBe(false);
    expect(editor.setEditText('cellA\ncellB').ok).toBe(true);
    expect(editor.commitEdit().ok).toBe(true);
    editor.select({ row: 1, col: 0 });
    expect(editor.snapshot().activeCell?.content).toBe('cellA\ncellB');
    expect(wraps(editor)).toBe(true);
    expect(editor.undo().ok).toBe(true);
    editor.select({ row: 1, col: 0 });
    expect(editor.snapshot().activeCell?.content ?? '').toBe('');
    expect(wraps(editor)).toBe(false);
  });

  test('Alt+Enter in the cell editor adds the line break, and Enter commits it with wrap text', async () => {
    const { editor, win, input } = await mount();
    editor.select({ row: 2, col: 0 });
    expect(editor.beginEdit('cellA').ok).toBe(true);
    const view = win as unknown as typeof globalThis;
    const key = (init: KeyboardEventInit) =>
      input.dispatchEvent(
        new view.KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init })
      );
    key({ key: 'Enter', altKey: true });
    expect(editor.snapshot().editText).toBe('cellA\n');
    editor.setEditText('cellA\ncellB');
    key({ key: 'Enter' });
    editor.select({ row: 2, col: 0 });
    expect(editor.snapshot().activeCell?.content).toBe('cellA\ncellB');
    expect(wraps(editor)).toBe(true);
  });

  test('text with no line break keeps the cell format', async () => {
    const { editor } = await mount();
    editor.select({ row: 1, col: 0 });
    editor.setEditText('one line');
    expect(editor.commitEdit().ok).toBe(true);
    editor.select({ row: 1, col: 0 });
    expect(wraps(editor)).toBe(false);
  });

  test('a cell that wraps already takes the value only', async () => {
    const { editor } = await mount();
    editor.select({ row: 0, col: 1 });
    expect(wraps(editor)).toBe(true);
    editor.setEditText('a\nb');
    expect(editor.commitEdit().ok).toBe(true);
    // One undo restores the old text, and the format stays.
    expect(editor.undo().ok).toBe(true);
    editor.select({ row: 0, col: 1 });
    expect(editor.snapshot().activeCell?.content).toBe('old');
    expect(wraps(editor)).toBe(true);
  });
});

test('a value with a line break that a command writes keeps the cell format', async () => {
  const { editor } = await mount();
  editor.select({ row: 1, col: 0 });
  expect(editor.enterValue('cellA\ncellB').ok).toBe(true);
  expect(editor.snapshot().activeCell?.content).toBe('cellA\ncellB');
  expect(wraps(editor)).toBe(false);
});
