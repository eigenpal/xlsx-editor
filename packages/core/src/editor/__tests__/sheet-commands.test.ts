// The sheet commands of the editor (design D1 and D4 of the change `workbook-structure-editing`)
// in a happy-dom window of their own: add shows the new sheet with the default name of the
// catalog, rename changes the tabs and the formulas, delete shows a neighbor, move keeps the
// moved sheet shown with its selection, a refusal shows through the refusal channel, and one
// undo step puts the sheets back. An undo and a redo of a sheet op show and name the right
// sheet.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { createXlsxEditor, type XlsxEditorInstance } from '../index.ts';
import {
  parseXmlPart,
  readXlsx,
  saveXlsx,
  serializeXmlPart,
  workbookIndex,
  xmlPartOf,
} from '../../store/index.ts';
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

let viewers: XlsxEditorInstance[] = [];
afterEach(() => {
  for (const v of viewers) v.destroy();
  viewers = [];
});

const translate = (key: string, params?: Readonly<Record<string, string | number>>) => {
  if (key === 'sheetTabs.newSheetName') return `Sheet${params?.n}`;
  // A sheet step names the change and the shown sheet.
  if (params?.name !== undefined && params.sheet !== undefined)
    return `${key} [${params.name}] {${params.sheet}}`;
  if (params?.name !== undefined) return `${key} [${params.name}]`;
  return params?.subject !== undefined ? `${key} [${params.subject}]` : key;
};

const workbook = (sheets = ['Sheet1', 'Sheet2']) =>
  buildXlsx({
    sheets: sheets.map((name, i) => ({
      name,
      rows: i === 1 ? row(1, [c('A1', 5, { f: `${sheets[0]}!B2` })]) : row(2, [c('B2', 5)]),
    })),
  });

async function mount(bytes: Uint8Array) {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.appendChild(host);
  const viewer = createXlsxEditor({ fonts: FONTS, translate });
  viewers.push(viewer);
  viewer.attach(host);
  const scroller = host.querySelector('.xlsx-scroller') as HTMLElement;
  Object.defineProperty(scroller, 'clientWidth', { get: () => 800 });
  Object.defineProperty(scroller, 'clientHeight', { get: () => 400 });
  await viewer.load(bytes);
  flushEditor(viewer);
  const live = host.querySelector('.xlsx-live') as HTMLElement;
  return { viewer, live, host };
}

const tabs = (v: XlsxEditorInstance) => v.snapshot().sheets.map((s) => s.name);
const shown = (v: XlsxEditorInstance) => v.snapshot().sheets[v.snapshot().activeSheet]?.name;
const content = (v: XlsxEditorInstance, sheet: number, r: number, col: number) => {
  v.setActiveSheet(sheet);
  v.select({ row: r, col });
  return v.snapshot().activeCell?.content;
};
const code = (r: { ok: boolean; refusal?: { code: string } }) => (r.ok ? null : r.refusal?.code);

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** Worksheets A and B, then the chartsheet Chart1. `lock` protects the workbook structure. */
function chartsheetWorkbook(lock = false): Uint8Array {
  const read = readXlsx(
    buildXlsx({
      sheets: [
        { name: 'A', rows: row(1, [c('A1', 1)]) },
        { name: 'B', rows: '' },
      ],
      extraParts: [
        [
          'xl/chartsheets/sheet9.xml',
          'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml',
          new TextEncoder().encode(
            '<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"/></sheetViews></chartsheet>'
          ),
          `${REL}/chartsheet`,
        ],
      ],
    })
  );
  if (!read.ok) throw new Error('the fixture does not read');
  const main = workbookIndex(read.pkg)!.mainPart;
  let text = serializeXmlPart(xmlPartOf(read.pkg, main)!).replace(
    '</sheets>',
    '<sheet name="Chart1" sheetId="9" r:id="rIdX1"/></sheets>'
  );
  if (lock) {
    const at = text.includes('<bookViews') ? '<bookViews' : '<sheets>';
    text = text.replace(at, `<workbookProtection lockStructure="1"/>${at}`);
  }
  const parsed = parseXmlPart(main, text);
  if (!parsed.ok) throw new Error('the patched workbook part does not parse');
  const parts = new Map(read.pkg.parts);
  parts.set(main, { kind: 'xml', partName: main, xml: parsed.part });
  return saveXlsx({ ...read.pkg, parts });
}

describe('sheet commands of the editor', () => {
  test('addSheet adds the smallest free default name after the shown sheet, and shows it', async () => {
    const { viewer } = await mount(workbook(['Sheet1', 'Sheet3']));
    expect(viewer.addSheet().ok).toBe(true);
    expect(tabs(viewer)).toEqual(['Sheet1', 'Sheet2', 'Sheet3']);
    expect(shown(viewer)).toBe('Sheet2');
    expect(viewer.snapshot().activeCell?.ref).toBe('A1');
    expect(viewer.addSheet('Data').ok).toBe(true);
    expect(tabs(viewer)).toEqual(['Sheet1', 'Sheet2', 'Data', 'Sheet3']);
    expect(viewer.undo().ok).toBe(true);
    expect(tabs(viewer)).toEqual(['Sheet1', 'Sheet2', 'Sheet3']);
    expect(viewer.snapshot().sheets[viewer.snapshot().activeSheet]?.visible).toBe(true);
    const saved = viewer.save();
    expect(saved).not.toBeNull();
  });

  test('a taken or invalid name is refused through the refusal channel', async () => {
    const { viewer, live } = await mount(workbook());
    expect(code(viewer.addSheet('sheet1'))).toBe('sheet-name-taken');
    // The subject is the existing sheet as its tab shows it, not the typed name.
    expect(viewer.snapshot().refusal).toMatchObject({
      code: 'sheet-name-taken',
      subject: 'Sheet1',
    });
    expect(live.textContent).toBe('editing.refused.sheet-name-taken [Sheet1]');
    expect(code(viewer.renameSheet(0, 'a/b'))).toBe('sheet-name-invalid');
    expect(code(viewer.renameSheet(9, 'X'))).toBe('unknown-sheet');
    expect(viewer.snapshot().canUndo).toBe(false);
  });

  test('renameSheet changes the tab and every formula that names the sheet', async () => {
    const { viewer } = await mount(workbook());
    expect(viewer.renameSheet(0, 'Data 1').ok).toBe(true);
    expect(tabs(viewer)).toEqual(['Data 1', 'Sheet2']);
    expect(content(viewer, 1, 0, 0)).toBe("='Data 1'!B2");
    expect(viewer.undo().ok).toBe(true);
    expect(content(viewer, 1, 0, 0)).toBe('=Sheet1!B2');
  });

  test('deleteSheet of the shown sheet shows a neighbor, and references become #REF!', async () => {
    const { viewer } = await mount(workbook());
    expect(shown(viewer)).toBe('Sheet1');
    expect(viewer.deleteSheet(0).ok).toBe(true);
    expect(tabs(viewer)).toEqual(['Sheet2']);
    expect(shown(viewer)).toBe('Sheet2');
    expect(content(viewer, 0, 0, 0)).toBe('=#REF!');
    // The undo step keeps the selection of the sheet that the delete showed: the undo shows
    // the restored sheet again.
    expect(viewer.undo().ok).toBe(true);
    expect(tabs(viewer)).toEqual(['Sheet1', 'Sheet2']);
    expect(shown(viewer)).toBe('Sheet1');
  });

  test('deleteSheet announces the deleted sheet and that undo restores it', async () => {
    const { viewer, live } = await mount(workbook());
    expect(viewer.deleteSheet(1).ok).toBe(true);
    flushEditor(viewer);
    expect(live.textContent).toBe('grid.sheetDeleted [Sheet2]');
  });

  test('addSheet and moveSheet announce the added or moved sheet', async () => {
    const { viewer, live } = await mount(workbook());
    expect(viewer.addSheet('New').ok).toBe(true);
    flushEditor(viewer);
    expect(live.textContent?.trim()).toBe('grid.sheetAdded [New]');
    expect(viewer.moveSheet(0, 2).ok).toBe(true);
    flushEditor(viewer);
    expect(live.textContent?.trim()).toBe('grid.sheetMoved [Sheet1]');
  });

  test('moveSheet to the same position announces nothing', async () => {
    const { viewer, live } = await mount(workbook());
    live.textContent = '';
    expect(viewer.moveSheet(1, 1).ok).toBe(true);
    flushEditor(viewer);
    expect(live.textContent?.trim()).toBe('');
    expect(viewer.snapshot().canUndo).toBe(false);
  });

  test('a repeated delete or cut announcement changes the live region again', async () => {
    const { viewer, live } = await mount(workbook(['Sheet1', 'Sheet2', 'Sheet3']));
    expect(viewer.deleteSheet(1).ok).toBe(true);
    flushEditor(viewer);
    const first = live.textContent;
    expect(viewer.undo().ok).toBe(true);
    flushEditor(viewer);
    live.textContent = first;
    expect(viewer.deleteSheet(1).ok).toBe(true);
    flushEditor(viewer);
    expect(live.textContent).not.toBe(first);
    expect(live.textContent?.trim()).toBe('grid.sheetDeleted [Sheet2]');
    viewer.select({ row: 1, col: 1 });
    expect(viewer.cut().ok).toBe(true);
    const cut = live.textContent;
    expect(cut?.trim()).toBe('grid.cut');
    expect(viewer.cut().ok).toBe(true);
    expect(live.textContent).not.toBe(cut);
  });

  test('deleteSheet of the last visible sheet is refused with last-sheet', async () => {
    const { viewer, live } = await mount(buildXlsx({ sheets: [{ name: 'Only', rows: '' }] }));
    expect(code(viewer.deleteSheet(0))).toBe('last-sheet');
    expect(live.textContent).toBe('editing.refused.last-sheet [Only]');
  });

  test('moveSheet keeps the moved sheet shown, with its selection', async () => {
    const { viewer } = await mount(workbook(['Sheet1', 'Sheet2', 'Sheet3']));
    viewer.select({ row: 2, col: 1 });
    expect(viewer.moveSheet(0, 2).ok).toBe(true);
    expect(tabs(viewer)).toEqual(['Sheet2', 'Sheet3', 'Sheet1']);
    expect(viewer.snapshot().activeSheet).toBe(2);
    expect(viewer.snapshot().selection?.sheetIndex).toBe(2);
    expect(viewer.snapshot().activeCell?.ref).toBe('B3');
    const undoable = viewer.snapshot().canUndo;
    expect(viewer.moveSheet(2, 2).ok).toBe(true);
    expect(viewer.snapshot().canUndo).toBe(undoable);
    expect(viewer.undo().ok).toBe(true);
    expect(tabs(viewer)).toEqual(['Sheet1', 'Sheet2', 'Sheet3']);
    expect(viewer.snapshot().activeSheet).toBe(0);
    expect(viewer.snapshot().activeCell?.ref).toBe('B3');
  });

  test('moveSheet of a chartsheet is refused with structure-blocked and the sheet name', async () => {
    const { viewer, live } = await mount(chartsheetWorkbook());
    expect(tabs(viewer)).toEqual(['A', 'B', 'Chart1']);
    const r = viewer.moveSheet(2, 0);
    expect(r.ok ? null : { code: r.refusal.code, subject: r.refusal.subject }).toEqual({
      code: 'structure-blocked',
      subject: 'Chart1',
    });
    expect(r.ok ? null : r.refusal.about).toEqual({ kind: 'not-worksheet', name: 'Chart1' });
    expect(live.textContent).toBe('editing.notWorksheetRefused.named [Chart1]');
    expect(tabs(viewer)).toEqual(['A', 'B', 'Chart1']);
    expect(viewer.snapshot().canUndo).toBe(false);
    // A worksheet still moves past the chartsheet.
    expect(viewer.moveSheet(0, 2).ok).toBe(true);
    expect(tabs(viewer)).toEqual(['B', 'Chart1', 'A']);
  });

  test('moveSheet of a chartsheet gives the refusal of a delete in each editor state', async () => {
    // View mode, an input method that composes, and the workbook protection refuse a move of a
    // chartsheet as they refuse its delete, before the kind of the sheet.
    const both = (v: XlsxEditorInstance) => [code(v.moveSheet(2, 0)), code(v.deleteSheet(2))];
    const viewing = await mount(chartsheetWorkbook());
    expect(viewing.viewer.setMode('view')).toBe(true);
    expect(both(viewing.viewer)).toEqual(['view-mode', 'view-mode']);

    const locked = await mount(chartsheetWorkbook(true));
    const lockOf = (r: ReturnType<XlsxEditorInstance['moveSheet']>) =>
      r.ok ? null : [r.refusal.code, r.refusal.lock];
    expect(lockOf(locked.viewer.moveSheet(2, 0))).toEqual(['locked', 'workbook']);
    expect(lockOf(locked.viewer.deleteSheet(2))).toEqual(['locked', 'workbook']);

    const composing = await mount(chartsheetWorkbook());
    const { viewer, host } = composing;
    viewer.select({ row: 0, col: 0 });
    const scroller = host.querySelector('.xlsx-scroller') as HTMLElement;
    const input = host.querySelector('.xlsx-cell-editor') as HTMLTextAreaElement;
    const W = scroller.ownerDocument!.defaultView as unknown as typeof globalThis;
    scroller.dispatchEvent(new W.KeyboardEvent('keydown', { key: '7', bubbles: true }));
    input.dispatchEvent(new W.Event('compositionstart'));
    expect(viewer.snapshot().composing).toBe(true);
    expect(both(viewer)).toEqual(['composing', 'composing']);
    expect(tabs(viewer)).toEqual(['A', 'B', 'Chart1']);
  });

  test('deleteSheet of a chartsheet says that the editor deletes only a worksheet', async () => {
    const { viewer, live } = await mount(chartsheetWorkbook());
    const r = viewer.deleteSheet(2);
    expect(r.ok ? null : [r.refusal.code, r.refusal.about]).toEqual([
      'structure-blocked',
      { kind: 'not-worksheet', name: 'Chart1' },
    ]);
    expect(live.textContent).toBe('editing.notWorksheetRefused.named [Chart1]');
  });

  test('a tab whose name finds another sheet, or none, is refused with sheet-name-ambiguous', async () => {
    // A file can give two sheets one name. The tab names its sheet by position, so a command on
    // the second tab never acts on the first sheet.
    const { viewer } = await mount(
      buildXlsx({
        sheets: [
          { name: 'Data', rows: row(1, [c('A1', 1)]) },
          { name: 'Data', rows: row(1, [c('A1', 2)]) },
          { name: 'Other', rows: '' },
        ],
      })
    );
    // The subject is the twin sheet, as the store names it for a rename of a twin.
    for (const r of [viewer.renameSheet(1, 'X'), viewer.deleteSheet(1), viewer.moveSheet(1, 2)])
      expect(r.ok ? null : { code: r.refusal.code, subject: r.refusal.subject }).toEqual({
        code: 'sheet-name-ambiguous',
        subject: 'Data',
      });
    expect(tabs(viewer)).toEqual(['Data', 'Data', 'Other']);
    expect(viewer.snapshot().canUndo).toBe(false);
  });
});

describe('undo and redo of a sheet op', () => {
  test('an undo and a redo of a rename name the sheet; the redo selects on the new name', async () => {
    const { viewer, live } = await mount(workbook());
    viewer.select({ row: 1, col: 1 });
    expect(viewer.renameSheet(0, 'Data').ok).toBe(true);
    expect(viewer.undo().ok).toBe(true);
    flushEditor(viewer);
    expect(shown(viewer)).toBe('Sheet1');
    expect(live.textContent).toBe('grid.sheetRenameUndone [Sheet1] {Sheet1}');
    viewer.select({ row: 4, col: 4 });
    expect(viewer.redo().ok).toBe(true);
    flushEditor(viewer);
    expect(shown(viewer)).toBe('Data');
    expect(viewer.snapshot().selection?.active).toEqual({ row: 1, col: 1 });
    expect(live.textContent).toBe('grid.sheetRenameRedone [Data] {Data}');
  });

  test('an undo and a redo of a rename of a sheet that does not show name the new name', async () => {
    const { viewer, live } = await mount(workbook());
    expect(viewer.renameSheet(1, 'Data').ok).toBe(true);
    expect(viewer.undo().ok).toBe(true);
    flushEditor(viewer);
    expect(live.textContent).toBe('grid.sheetRenameUndone [Sheet2] {Sheet1}');
    expect(viewer.redo().ok).toBe(true);
    flushEditor(viewer);
    expect(live.textContent).toBe('grid.sheetRenameRedone [Data] {Sheet1}');
  });

  test('an undo and a redo of a move name the moved sheet', async () => {
    const { viewer, live } = await mount(workbook());
    expect(viewer.moveSheet(1, 0).ok).toBe(true);
    expect(tabs(viewer)).toEqual(['Sheet2', 'Sheet1']);
    expect(viewer.undo().ok).toBe(true);
    flushEditor(viewer);
    expect(tabs(viewer)).toEqual(['Sheet1', 'Sheet2']);
    expect(live.textContent).toBe('grid.sheetMoveUndone [Sheet2] {Sheet1}');
    expect(viewer.redo().ok).toBe(true);
    flushEditor(viewer);
    expect(live.textContent).toBe('grid.sheetMoveRedone [Sheet2] {Sheet1}');
  });

  test('an undo and a redo of a delete name the deleted sheet and the shown sheet', async () => {
    const { viewer, live } = await mount(workbook());
    viewer.setActiveSheet(1);
    expect(viewer.deleteSheet(1).ok).toBe(true);
    expect(shown(viewer)).toBe('Sheet1');
    expect(viewer.undo().ok).toBe(true);
    flushEditor(viewer);
    expect(shown(viewer)).toBe('Sheet2');
    expect(live.textContent).toBe('grid.sheetDeleteUndone [Sheet2] {Sheet2}');
    expect(viewer.redo().ok).toBe(true);
    flushEditor(viewer);
    expect(tabs(viewer)).toEqual(['Sheet1']);
    expect(shown(viewer)).toBe('Sheet1');
    expect(live.textContent).toBe('grid.sheetDeleteRedone [Sheet2] {Sheet1}');
  });

  test('an undo of an add names the removed sheet, and a redo shows the added sheet again', async () => {
    const { viewer, live } = await mount(workbook());
    expect(viewer.addSheet().ok).toBe(true);
    expect(shown(viewer)).toBe('Sheet3');
    expect(viewer.undo().ok).toBe(true);
    flushEditor(viewer);
    expect(shown(viewer)).toBe('Sheet1');
    expect(live.textContent).toBe('grid.sheetAddUndone [Sheet3] {Sheet1}');
    expect(viewer.redo().ok).toBe(true);
    flushEditor(viewer);
    expect(tabs(viewer)).toEqual(['Sheet1', 'Sheet3', 'Sheet2']);
    expect(shown(viewer)).toBe('Sheet3');
    expect(viewer.snapshot().selection?.active).toEqual({ row: 0, col: 0 });
    expect(live.textContent).toBe('grid.sheetAddRedone [Sheet3] {Sheet3}');
  });
});
