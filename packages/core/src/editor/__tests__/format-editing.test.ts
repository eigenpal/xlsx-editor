// Format editing from the grid (task 7.4) in a happy-dom window of its own: the format state of
// the active cell, the Mod+B, Mod+I, and Mod+U toggles that read it, the format commands of the
// viewer, their refusals, and one undo step for each command.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { createXlsxEditor, runToolbarCommand, type XlsxEditorInstance } from '../index.ts';
import { flushEditor } from './flush-editor.ts';
import type { FontSource } from '../../layout/index.ts';
import { readXlsx, styleLists, workbookIndex } from '../../store/index.ts';

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

let viewers: XlsxEditorInstance[] = [];
afterEach(() => {
  for (const v of viewers) v.destroy();
  viewers = [];
});

const sheet = () =>
  buildXlsx({
    sheets: [
      {
        name: 'S',
        rows: row(1, [c('A1', 1), c('B1', 2, { s: 1 })]) + row(2, [c('A2', 3)]),
      },
    ],
    styles: {
      fonts: ['<font><b/><sz val="11"/><name val="Calibri"/></font>'],
      cellXfs: ['<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'],
    },
  });

async function mount(bytes: Uint8Array) {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.appendChild(host);
  const viewer = createXlsxEditor({ fonts: FONTS, translate: (k) => k });
  viewers.push(viewer);
  viewer.attach(host);
  const scroller = host.querySelector('.xlsx-scroller') as HTMLElement;
  Object.defineProperty(scroller, 'clientWidth', { get: () => 800 });
  Object.defineProperty(scroller, 'clientHeight', { get: () => 400 });
  await viewer.load(bytes);
  flushEditor(viewer);
  const live = host.querySelector('.xlsx-live') as HTMLElement;
  return { viewer, scroller, live };
}

const key = (el: Element, k: string, init: Record<string, unknown> = {}) =>
  el.dispatchEvent(
    new (el.ownerDocument!.defaultView as unknown as typeof globalThis).KeyboardEvent('keydown', {
      key: k,
      bubbles: true,
      cancelable: true,
      ...init,
    })
  );

const format = (v: XlsxEditorInstance) => v.snapshot().activeFormat!;

describe('format state and toggles', () => {
  test('the snapshot shows the format of the active cell', async () => {
    const { viewer } = await mount(sheet());
    viewer.select({ row: 0, col: 1 });
    expect(format(viewer)).toMatchObject({
      bold: true,
      italic: false,
      underline: false,
      fontName: 'Calibri',
      fontSize: 11,
      fillColor: null,
      horizontal: 'general',
      numberFormat: { id: 0, code: 'General' },
    });
    viewer.select({ row: 0, col: 0 });
    expect(format(viewer).bold).toBe(false);
  });

  test('Mod+B toggles from the active cell: the state the snapshot shows', async () => {
    const { viewer, scroller } = await mount(sheet());
    // The active cell B1 is bold, so the toggle removes bold from A1 and B1.
    viewer.select({ row: 0, col: 1 }, { row: 0, col: 0 });
    key(scroller, 'b', { ctrlKey: true });
    expect(format(viewer).bold).toBe(false);
    viewer.select({ row: 0, col: 0 });
    expect(format(viewer).bold).toBe(false);
    key(scroller, 'B', { metaKey: true });
    expect(format(viewer).bold).toBe(true);
    key(scroller, 'i', { ctrlKey: true });
    key(scroller, 'u', { ctrlKey: true });
    expect(format(viewer)).toMatchObject({ bold: true, italic: true, underline: true });
    // Each toggle is one undo step.
    key(scroller, 'z', { ctrlKey: true });
    expect(format(viewer)).toMatchObject({ italic: true, underline: false });
  });

  test('an empty cell of a bold row or column shows bold, and Mod+B removes it', async () => {
    const { viewer, scroller } = await mount(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            cols: '<cols><col min="3" max="3" width="9" style="1" customWidth="1"/></cols>',
            rows: '<row r="2" s="1" customFormat="1"><c r="A2"><v>1</v></c></row>',
          },
        ],
        styles: {
          fonts: ['<font><b/><sz val="11"/><name val="Calibri"/></font>'],
          cellXfs: ['<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'],
        },
      })
    );
    // B2: no cell, in the bold row. C1: no cell, in the bold column. A2: a cell of its own.
    for (const [at, bold] of [
      [{ row: 1, col: 1 }, true],
      [{ row: 0, col: 2 }, true],
      [{ row: 1, col: 0 }, false],
      [{ row: 0, col: 1 }, false],
    ] as const) {
      viewer.select(at);
      expect(format(viewer).bold).toBe(bold);
    }
    for (const at of [
      { row: 1, col: 1 },
      { row: 0, col: 2 },
    ]) {
      viewer.select(at);
      key(scroller, 'b', { ctrlKey: true });
      expect(format(viewer).bold).toBe(false);
    }
  });

  test('a toggle key with Shift or Alt does nothing', async () => {
    const { viewer, scroller } = await mount(sheet());
    key(scroller, 'b', { ctrlKey: true, altKey: true });
    key(scroller, 'B', { ctrlKey: true, shiftKey: true });
    expect(viewer.snapshot().canUndo).toBe(false);
  });
});

describe('format commands', () => {
  test('applyFormat sets every property of the patch on the selection in one step', async () => {
    const { viewer } = await mount(sheet());
    viewer.select({ row: 0, col: 0 }, { row: 2, col: 2 });
    const r = viewer.applyFormat({
      fontName: 'Courier New',
      fontSize: 14,
      fontColor: '1F4E79',
      fill: { type: 'solid', color: 'FFFF00' },
      borders: { edges: 'outline' },
      horizontal: 'center',
      vertical: 'top',
      wrapText: true,
      numberFormat: { code: '0.00' },
    });
    expect(r.ok).toBe(true);
    expect(format(viewer)).toMatchObject({
      fontName: 'Courier New',
      fontSize: 14,
      fontColor: '1F4E79',
      fillColor: 'FFFF00',
      horizontal: 'center',
      vertical: 'top',
      wrapText: true,
      // A code that equals a built-in format is written as the built-in id.
      numberFormat: { id: 2, code: '0.00' },
    });
    const pkg = readXlsx(viewer.save()!);
    if (!pkg.ok) throw new Error('reopen');
    expect(styleLists(pkg.pkg, workbookIndex(pkg.pkg)!).numFmts.size).toBe(0);
    expect(viewer.undo().ok).toBe(true);
    expect(format(viewer).fontName).toBe('Carlito');
    expect(viewer.snapshot().canUndo).toBe(false);
  });

  test('an unreadable number format code is refused before the store, with no undo step', async () => {
    const { viewer, live } = await mount(sheet());
    const r = viewer.applyFormat({ numberFormat: { code: '0.00"' } });
    expect(r.ok ? null : r.refusal.code).toBe('unreadable-number-format');
    expect(viewer.snapshot().refusal?.code).toBe('unreadable-number-format');
    expect(live.textContent).toBe('editing.refused.unreadable-number-format');
    expect(viewer.snapshot().canUndo).toBe(false);
    expect(viewer.applyFormat({ numberFormat: { code: '[Red]0.0;(0.0)' } }).ok).toBe(true);
    expect(viewer.snapshot().refusal).toBeNull();
    expect(format(viewer).numberFormat).toEqual({ id: 164, code: '[Red]0.0;(0.0)', hasCode: true });
  });

  test('a whole column is refused with range-too-large, and a whole row formats', async () => {
    const { viewer } = await mount(sheet());
    viewer.select({ row: 0, col: 0 }, { row: 1_048_575, col: 0 });
    const r = viewer.applyFormat({ bold: true });
    expect(r.ok ? null : r.refusal.code).toBe('range-too-large');
    viewer.select({ row: 4, col: 0 }, { row: 4, col: 16_383 });
    expect(viewer.applyFormat({ italic: true }).ok).toBe(true);
  });

  test('clearFormats keeps the value; clearAll removes both in one step', async () => {
    const { viewer } = await mount(sheet());
    viewer.select({ row: 0, col: 1 });
    expect(viewer.clearFormats().ok).toBe(true);
    expect(format(viewer).bold).toBe(false);
    expect(viewer.snapshot().activeCell?.content).toBe('2');
    viewer.undo();
    expect(viewer.clearAll().ok).toBe(true);
    expect(viewer.snapshot().activeCell?.content).toBe('');
    expect(format(viewer).bold).toBe(false);
    viewer.undo();
    expect(viewer.snapshot().activeCell?.content).toBe('2');
    expect(format(viewer).bold).toBe(true);
    const saved = readXlsx(viewer.save()!);
    expect(saved.ok).toBe(true);
  });

  test('a format or clear command commits the open edit first, or is refused with it', async () => {
    const { viewer, scroller } = await mount(sheet());
    const input = () => scroller.ownerDocument.querySelector('textarea') as HTMLTextAreaElement;
    key(scroller, '7');
    input().value = '7';
    expect(viewer.applyFormat({ italic: true }).ok).toBe(true);
    expect(viewer.snapshot().editing).toBe(false);
    expect(viewer.snapshot().activeCell?.content).toBe('7');
    expect(format(viewer).italic).toBe(true);
    key(scroller, '8');
    input().value = '8';
    expect(viewer.clearAll().ok).toBe(true);
    expect(viewer.snapshot().activeCell?.content).toBe('');
    // Undo takes back the clear, then the committed edit.
    viewer.undo();
    expect(viewer.snapshot().activeCell?.content).toBe('8');
    key(scroller, '=');
    input().value = '=SUM(';
    for (const r of [viewer.clearAll(), viewer.clearFormats(), viewer.toggleFormat('bold')])
      expect(r.ok ? null : r.refusal.code).toBe('unreadable-formula');
    expect(viewer.snapshot().editing).toBe(true);
    expect(input().value).toBe('=SUM(');
  });

  test('a format command with no workbook is refused', () => {
    const viewer = createXlsxEditor({ fonts: FONTS });
    viewers.push(viewer);
    const r = viewer.applyFormat({ bold: true });
    expect(r.ok ? null : r.refusal.code).toBe('not-editable');
    expect(viewer.toggleFormat('bold').ok).toBe(false);
  });

  test('the font color takes `auto`, the automatic color, and the fill color does not', async () => {
    const { viewer } = await mount(sheet());
    expect(runToolbarCommand(viewer, 'font.color', 'FF0000').ok).toBe(true);
    expect(viewer.snapshot().activeFormat?.fontColor).toBe('FF0000');
    expect(runToolbarCommand(viewer, 'font.color', 'auto').ok).toBe(true);
    // The automatic font color paints black (`cascade/color.ts`).
    expect(viewer.snapshot().activeFormat?.fontColor).toBe('000000');
    expect(runToolbarCommand(viewer, 'fill.color', 'auto').ok).toBe(false);
  });
});
