// The engine side of the formula bar, in a happy-dom window of its own: `setEditText` edits
// the one open cell edit (the cell editor shows the same text), `commitEdit(move)` commits and
// moves as Enter and Tab do, a composing input method holds every commit, and
// `selectReference` selects what the name box names or refuses it with a typed reason.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { createXlsxEditor, MAX_COLS, MAX_ROWS, type XlsxEditorInstance } from '../index.ts';
import { readGoTo } from '../go-to.ts';
import { flushEditor } from './flush-editor.ts';
import type { FontSource } from '../../layout/index.ts';
import { readXlsx, workbookIndex } from '../../store/index.ts';

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

const workbook = buildXlsx({
  sheets: [
    { name: 'Data', rows: row(1, [c('A1', 1), c('B1', 'a\nb')]) + row(2, [c('A2', 2)]) },
    { name: 'Hidden', state: 'hidden', rows: row(1, [c('A1', 'secret')]) },
    { name: 'My sheet', rows: row(1, [c('A1', 'two')]) },
  ],
  definedNames:
    '<definedName name="Rates">Data!$A$1:$A$2</definedName>' +
    '<definedName name="Total">SUM(Data!$A$1:$A$2)</definedName>' +
    '<definedName name="Secret">Hidden!$A$1</definedName>' +
    '<definedName name="Spot">\'My sheet\'!$B$3</definedName>' +
    '<definedName name="Spot" localSheetId="0">Data!$C$4</definedName>' +
    '<definedName name="Only" localSheetId="2">\'My sheet\'!$A$1</definedName>',
});

async function mount(bytes: Uint8Array = workbook) {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.appendChild(host);
  const bar = doc.createElement('textarea');
  doc.body.appendChild(bar);
  const viewer = createXlsxEditor({ fonts: FONTS });
  viewers.push(viewer);
  viewer.attach(host);
  const scroller = host.querySelector('.xlsx-scroller') as HTMLElement;
  Object.defineProperty(scroller, 'clientWidth', { get: () => 800 });
  Object.defineProperty(scroller, 'clientHeight', { get: () => 400 });
  await viewer.load(bytes);
  flushEditor(viewer);
  const input = host.querySelector('.xlsx-cell-editor') as HTMLTextAreaElement;
  const live = host.querySelector('.xlsx-live') as HTMLElement;
  return { viewer, doc, host, bar, input, scroller, live };
}

const active = (v: XlsxEditorInstance) => v.snapshot().selection?.active;

describe('setEditText', () => {
  test('opens the edit without the keyboard, and the cell editor shows the same text', async () => {
    const { viewer, doc, bar, input } = await mount();
    bar.focus();
    expect(viewer.setEditText('12').ok).toBe(true);
    const s = viewer.snapshot();
    expect([s.editing, s.editText, s.editChanged, s.composing]).toEqual([true, '12', true, false]);
    expect(input.value).toBe('12');
    expect(input.hidden).toBe(false);
    expect(doc.activeElement).toBe(bar);
    // A later key replaces the text of the same edit: there is one text, never a copy.
    viewer.setEditText('123');
    expect([viewer.snapshot().editText, input.value]).toEqual(['123', '123']);
  });

  test('a key in the cell editor continues the edit that the bar opened', async () => {
    const { viewer, input } = await mount();
    viewer.setEditText('4');
    input.value = '45';
    input.dispatchEvent(
      new (input.ownerDocument.defaultView as unknown as typeof globalThis).Event('input')
    );
    expect(viewer.snapshot().editText).toBe('45');
    expect(viewer.commitEdit().ok).toBe(true);
    expect(viewer.snapshot().activeCell?.content).toBe('45');
  });

  test('keeps a line break of the cell text', async () => {
    const { viewer } = await mount();
    viewer.select({ row: 0, col: 1 });
    expect(viewer.snapshot().activeCell?.content).toBe('a\nb');
    viewer.setEditText('a\nbc');
    expect(viewer.commitEdit('down').ok).toBe(true);
    viewer.select({ row: 0, col: 1 });
    expect(viewer.snapshot().activeCell?.content).toBe('a\nbc');
  });

  test('is refused as beginEdit is: in view mode, and nothing opens', async () => {
    const { viewer } = await mount();
    viewer.setMode('view');
    const r = viewer.setEditText('1');
    expect(r.ok).toBe(false);
    expect(!r.ok && r.refusal.code).toBe('view-mode');
    expect(viewer.snapshot().editing).toBe(false);
  });

  test('an edit that opens clears a refusal from before it', async () => {
    const { viewer } = await mount();
    expect(viewer.renameSheet(0, '').ok).toBe(false);
    expect(viewer.setEditText('2').ok).toBe(true);
    expect(viewer.snapshot().refusal).toBeNull();
  });

  test('an edit that opens keeps the reason of a dropped edit', async () => {
    const { viewer } = await mount();
    viewer.setEditText('changed');
    await viewer.load(workbook);
    flushEditor(viewer);
    expect(viewer.snapshot().refusal?.code).toBe('edit-dropped');
    expect(viewer.setEditText('2').ok).toBe(true);
    expect(viewer.snapshot().refusal?.code).toBe('edit-dropped');
    expect(viewer.snapshot().editDropped).toBe(true);
  });

  test('a text that is not a string throws, and opens no edit', async () => {
    const { viewer } = await mount();
    expect(() => viewer.setEditText(null as never)).toThrow(TypeError);
    expect(viewer.snapshot().editing).toBe(false);
  });

  test('a composing input method holds every commit until it ends', async () => {
    const { viewer } = await mount();
    viewer.setEditText('か', { composing: true });
    expect(viewer.snapshot().composing).toBe(true);
    const held = viewer.commitEdit('down');
    expect(!held.ok && held.refusal.code).toBe('composing');
    expect(viewer.select({ row: 3, col: 3 })).toBe(false);
    expect(viewer.snapshot().editing).toBe(true);
    viewer.setEditText('漢', { composing: false });
    expect(viewer.snapshot().composing).toBe(false);
    expect(viewer.commitEdit('down').ok).toBe(true);
    expect(active(viewer)).toEqual({ row: 1, col: 0 });
    viewer.select({ row: 0, col: 0 });
    expect(viewer.snapshot().activeCell?.content).toBe('漢');
  });
});

describe('the formula bar and the cell editor share one edit', () => {
  test('a key in the cell editor ends a composition that the bar reported', async () => {
    const { viewer, input } = await mount();
    viewer.setEditText('ab', { composing: true });
    input.value = 'abc';
    input.dispatchEvent(
      new (input.ownerDocument.defaultView as unknown as typeof globalThis).Event('input')
    );
    expect(viewer.snapshot().composing).toBe(false);
    expect(viewer.commitEdit().ok).toBe(true);
  });

  test('a refused commit keeps the keyboard in the bar after an edit that F2 opened', async () => {
    const { viewer, doc, bar, input } = await mount();
    viewer.beginEdit();
    input.value = '=1+';
    input.dispatchEvent(
      new (input.ownerDocument.defaultView as unknown as typeof globalThis).Event('input')
    );
    bar.focus();
    const r = viewer.commitEdit('down');
    expect(!r.ok && r.refusal.code).toBe('formula-syntax');
    expect(doc.activeElement).toBe(bar);
  });

  test('a refused commit from the cell editor keeps the keyboard in the cell editor', async () => {
    const { viewer, doc, input } = await mount();
    viewer.beginEdit('=1+');
    expect(viewer.commitEdit().ok).toBe(false);
    expect(doc.activeElement).toBe(input);
  });
});

describe('commitEdit with a move', () => {
  test('Enter and Tab moves: down, up, right, and left, and the grid takes the keyboard', async () => {
    const { viewer, doc, bar, scroller } = await mount();
    viewer.select({ row: 1, col: 1 });
    bar.focus();
    viewer.setEditText('7');
    expect(viewer.commitEdit('down').ok).toBe(true);
    expect(active(viewer)).toEqual({ row: 2, col: 1 });
    expect(doc.activeElement).toBe(scroller);
    viewer.setEditText('8');
    expect(viewer.commitEdit('right').ok).toBe(true);
    expect(active(viewer)).toEqual({ row: 2, col: 2 });
    viewer.commitEdit('up');
    viewer.commitEdit('left');
    // With no open edit, the move alone applies.
    expect(active(viewer)).toEqual({ row: 1, col: 1 });
    viewer.select({ row: 1, col: 1 });
    expect(viewer.snapshot().activeCell?.content).toBe('7');
    viewer.select({ row: 2, col: 1 });
    expect(viewer.snapshot().activeCell?.content).toBe('8');
  });

  test('moves inside a multi-cell selection and keeps it', async () => {
    const { viewer } = await mount();
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 1 });
    viewer.setEditText('5');
    viewer.commitEdit('right');
    expect(active(viewer)).toEqual({ row: 0, col: 1 });
    expect(viewer.snapshot().selectionRef).toBe('A1:B2');
  });

  test('a refused commit keeps the edit, its text, and the keyboard in the bar', async () => {
    const { viewer, doc, bar, live } = await mount();
    bar.focus();
    viewer.setEditText('=1+');
    const r = viewer.commitEdit('down');
    expect(!r.ok && r.refusal.code).toBe('formula-syntax');
    const s = viewer.snapshot();
    expect([s.editing, s.editText, s.refusal?.code]).toEqual([true, '=1+', 'formula-syntax']);
    expect(active(viewer)).toEqual({ row: 0, col: 0 });
    expect(doc.activeElement).toBe(bar);
    expect(live.textContent).toContain('formula');
    // Cancel restores the cell, and the grid takes the keyboard.
    viewer.cancelEdit();
    expect(viewer.snapshot().activeCell?.content).toBe('1');
    expect(doc.activeElement?.classList.contains('xlsx-scroller')).toBe(true);
  });

  test('a move that is not an edit move throws, and changes nothing', async () => {
    const { viewer } = await mount();
    viewer.setEditText('9');
    for (const move of ['bogus', '__proto__', 'toString']) {
      expect(() => viewer.commitEdit(move as never)).toThrow(TypeError);
    }
    expect([viewer.snapshot().editing, viewer.snapshot().editText]).toEqual([true, '9']);
    expect(active(viewer)).toEqual({ row: 0, col: 0 });
  });

  test('a press on the grid commits a bar edit, or keeps a refused one in the bar', async () => {
    const { viewer, doc, host, bar } = await mount();
    const W = doc.defaultView as unknown as typeof globalThis;
    const stage = host.querySelector('.xlsx-stage') as HTMLElement;
    const press = () =>
      stage.dispatchEvent(
        new W.PointerEvent('pointerdown', { bubbles: true, button: 0, clientX: 200, clientY: 100 })
      );
    bar.focus();
    viewer.setEditText('=SUM(A1');
    press();
    const s = viewer.snapshot();
    expect([s.editing, s.editText, s.refusal?.code]).toEqual([
      true,
      '=SUM(A1',
      'unreadable-formula',
    ]);
    expect(active(viewer)).toEqual({ row: 0, col: 0 });
    expect(doc.activeElement).toBe(bar);
    // A formula that the engine can store commits on the press; it inserts no reference.
    viewer.setEditText('=SUM(A2)');
    press();
    expect(viewer.snapshot().editing).toBe(false);
    viewer.select({ row: 0, col: 0 });
    expect(viewer.snapshot().activeCell?.content).toBe('=SUM(A2)');
  });

  test('one commit is one undo step', async () => {
    const { viewer } = await mount();
    viewer.setEditText('9');
    viewer.commitEdit('down');
    expect(viewer.undo().ok).toBe(true);
    viewer.select({ row: 0, col: 0 });
    expect(viewer.snapshot().activeCell?.content).toBe('1');
  });
});

describe('selectReference', () => {
  test('selects a cell, a range, whole columns, and whole rows, in any case', async () => {
    const { viewer } = await mount();
    expect(viewer.selectReference('b2').ok).toBe(true);
    expect(viewer.snapshot().selectionRef).toBe('B2');
    expect(viewer.selectReference(' C5:b2 ').ok).toBe(true);
    expect(viewer.snapshot().selectionRef).toBe('B2:C5');
    expect(active(viewer)).toEqual({ row: 1, col: 1 });
    viewer.selectReference('$D$4');
    expect(viewer.snapshot().selectionRef).toBe('D4');
    viewer.selectReference('C:D');
    expect(viewer.snapshot().selection?.range).toEqual({
      start: { row: 0, col: 2 },
      end: { row: MAX_ROWS - 1, col: 3 },
    });
    viewer.selectReference('3:3');
    expect(viewer.snapshot().selection?.range).toEqual({
      start: { row: 2, col: 0 },
      end: { row: 2, col: MAX_COLS - 1 },
    });
  });

  test('a defined name goes to its range; a name of the shown sheet comes first', async () => {
    const { viewer } = await mount();
    expect(viewer.selectReference('rates').ok).toBe(true);
    expect(viewer.snapshot().selectionRef).toBe('A1:A2');
    // On Data, the name that Data scopes wins; on another sheet, the workbook name.
    viewer.selectReference('Spot');
    expect([viewer.snapshot().activeSheet, viewer.snapshot().selectionRef]).toEqual([0, 'C4']);
    viewer.setActiveSheet(2);
    viewer.selectReference('Spot');
    expect([viewer.snapshot().activeSheet, viewer.snapshot().selectionRef]).toEqual([2, 'B3']);
  });

  test('the view scrolls to the active cell, not to the far corner of the range', async () => {
    const { viewer, scroller } = await mount();
    viewer.selectReference('C:E');
    flushEditor(viewer);
    expect(scroller.scrollTop).toBe(0);
    viewer.selectReference('B2:Z500');
    flushEditor(viewer);
    expect([scroller.scrollTop, scroller.scrollLeft]).toEqual([0, 0]);
    expect(active(viewer)).toEqual({ row: 1, col: 1 });
  });

  test('whole columns and rows select the band and do not grow through merged cells', async () => {
    const merged = buildXlsx({
      sheets: [{ name: 'M', rows: row(1, [c('A1', 'title')]), merges: ['A1:C1', 'B5:D5'] }],
    });
    const { viewer } = await mount(merged);
    viewer.selectReference('C:E');
    expect(viewer.snapshot().selection?.range).toEqual({
      start: { row: 0, col: 2 },
      end: { row: MAX_ROWS - 1, col: 4 },
    });
    expect(active(viewer)).toEqual({ row: 0, col: 2 });
    viewer.selectReference('4:5');
    expect(viewer.snapshot().selection?.range).toEqual({
      start: { row: 3, col: 0 },
      end: { row: 4, col: MAX_COLS - 1 },
    });
    // A band commits an open edit first, and a refused commit refuses it.
    viewer.setEditText('=1+');
    const r = viewer.selectReference('B:B');
    expect(!r.ok && r.refusal.code).toBe('formula-syntax');
    expect(viewer.snapshot().selection?.range.start).toEqual({ row: 3, col: 0 });
  });

  test('a sheet prefix reads the name that the sheet scopes', async () => {
    const { viewer } = await mount();
    viewer.setActiveSheet(2);
    expect(viewer.selectReference('data!Spot').ok).toBe(true);
    expect([viewer.snapshot().activeSheet, viewer.snapshot().selectionRef]).toEqual([0, 'C4']);
    // A name that only another sheet scopes is a defined name, not unknown text.
    const r = viewer.selectReference("'My sheet'!Spot");
    expect(!r.ok && r.refusal.code).toBe('go-to-name');
    // The reason names the text as typed, with its sheet.
    expect(!r.ok && r.refusal.subject).toBe("'My sheet'!Spot");
    const other = viewer.selectReference('Only');
    expect(!other.ok && other.refusal.code).toBe('go-to-name');
  });

  test('is refused with composing while an input method composes, whatever the text', async () => {
    const { viewer } = await mount();
    viewer.setEditText('か', { composing: true });
    for (const text of ['B2', 'nothing at all']) {
      const r = viewer.selectReference(text);
      expect(!r.ok && r.refusal.code).toBe('composing');
    }
  });

  test('a sheet name shows that sheet', async () => {
    const { viewer } = await mount();
    expect(viewer.selectReference("'my SHEET'!A1:B2").ok).toBe(true);
    expect(viewer.snapshot().activeSheet).toBe(2);
    expect(viewer.snapshot().selectionRef).toBe('A1:B2');
  });

  test('refuses what it cannot select, with a typed reason, and changes nothing', async () => {
    const { viewer, live } = await mount();
    viewer.select({ row: 1, col: 0 });
    for (const [text, code, subject] of [
      ['Total', 'go-to-name', 'Total'],
      ['Secret', 'go-to-name', 'Secret'],
      ['Unknown', 'go-to-reference', undefined],
      ['Hidden!A1', 'go-to-sheet', 'Hidden'],
      ['Hidden!Secret', 'go-to-sheet', 'Hidden'],
      ['Nope!A1', 'go-to-sheet', 'Nope'],
      ['A0', 'go-to-reference', undefined],
      ['XFE1', 'go-to-reference', undefined],
      ['A1 B2', 'go-to-reference', undefined],
      ['=A1', 'go-to-reference', undefined],
      ['', 'go-to-reference', undefined],
      ['[1]Data!A1', 'go-to-reference', undefined],
      ['Data:Data!A1', 'go-to-reference', undefined],
      ['R1C1', 'go-to-reference', undefined],
    ] as const) {
      const r = viewer.selectReference(text);
      expect([text, !r.ok && r.refusal.code, !r.ok && r.refusal.subject]).toEqual([
        text,
        code,
        subject,
      ]);
      expect(viewer.snapshot().refusal?.code).toBe(code);
      expect([viewer.snapshot().activeSheet, viewer.snapshot().selectionRef]).toEqual([0, 'A2']);
    }
    expect(live.textContent).toContain('Nothing was selected');
  });

  test('commits an open edit first; a refused commit refuses the selection', async () => {
    const { viewer } = await mount();
    viewer.setEditText('3');
    expect(viewer.selectReference('C3').ok).toBe(true);
    viewer.select({ row: 0, col: 0 });
    expect(viewer.snapshot().activeCell?.content).toBe('3');
    viewer.setEditText('=1+');
    const r = viewer.selectReference('C3');
    expect(!r.ok && r.refusal.code).toBe('formula-syntax');
    expect(viewer.snapshot().selectionRef).toBe('A1');
  });

  test('selects in view mode, where no cell can be edited', async () => {
    const { viewer } = await mount();
    viewer.setMode('view');
    expect(viewer.selectReference('B2:C3').ok).toBe(true);
    expect(viewer.snapshot().selectionRef).toBe('B2:C3');
  });

  test('is refused with not-editable when no workbook is open', () => {
    const viewer = createXlsxEditor({ fonts: FONTS });
    viewers.push(viewer);
    const r = viewer.selectReference('A1');
    expect(!r.ok && r.refusal.code).toBe('not-editable');
  });

  test('commitEdit with a move is refused with not-editable when no workbook is open', () => {
    const viewer = createXlsxEditor({ fonts: FONTS });
    viewers.push(viewer);
    const r = viewer.commitEdit('down');
    expect(!r.ok && r.refusal.code).toBe('not-editable');
  });
});

describe('readGoTo', () => {
  const read = readXlsx(workbook);
  if (!read.ok) throw new Error('the workbook did not read');
  const { pkg } = read;
  const index = workbookIndex(pkg)!;
  const sheets = [
    { index: 0, name: 'Data', visible: true, worksheet: true },
    { index: 1, name: 'Chart', visible: true, worksheet: false },
  ];
  test('a chartsheet is not a worksheet to go to', () => {
    expect(readGoTo('Chart!A1', pkg, index, sheets, 0)).toEqual({
      code: 'go-to-sheet',
      subject: 'Chart',
    });
  });
  test('a defined name that names a formula is told apart from text that is no reference', () => {
    expect(readGoTo('total', pkg, index, sheets, 0)).toEqual({
      code: 'go-to-name',
      subject: 'total',
    });
    expect(readGoTo('Other', pkg, index, sheets, 0)).toMatchObject({ code: 'go-to-reference' });
  });
  test('a hidden sheet whose name differs only in case does not hide a visible one', () => {
    const twins = [
      { index: 0, name: 'data', visible: false, worksheet: true },
      { index: 1, name: 'Data', visible: true, worksheet: true },
    ];
    expect(readGoTo('DATA!B2', pkg, index, twins, 1)).toMatchObject({ sheet: 1 });
  });
  test('a long text is refused before it is read', () => {
    expect(readGoTo('A'.repeat(300), pkg, index, sheets, 0)).toMatchObject({
      code: 'go-to-reference',
    });
    // Spaces count: the cap holds before the text is trimmed.
    expect(readGoTo(`${' '.repeat(300)}B2`, pkg, index, sheets, 0)).toMatchObject({
      code: 'go-to-reference',
    });
  });
});
