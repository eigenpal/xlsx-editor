// Conditional formatting and the toolbar (change `workbook-conditional-formatting`, design D4):
// the format controls read and set the cell format, a true rule gives a hint to each control
// whose state it changes, and the snapshot flags a sheet whose rules the grid can leave out.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import {
  createXlsxEditor,
  toolbarCommandState,
  type ChromeSlotId,
  type XlsxEditorInstance,
} from '../index.ts';
import { flushEditor } from './flush-editor.ts';
import type { FontSource } from '../../layout/index.ts';
import { readXlsx, xmlPartOf } from '../../store/package/package.ts';
import { sheetByName, workbookIndex } from '../../store/workbook/workbook-index.ts';
import { sheetIndex } from '../../store/workbook/sheet-index.ts';
import { cellFormat, conditionalFormat, EN_US } from '../../cascade/index.ts';
import { cellAnnouncement } from '../cell-announcement.ts';

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

const CF =
  '<conditionalFormatting sqref="A1"><cfRule type="expression" dxfId="0" priority="1"><formula>1</formula></cfRule></conditionalFormatting>';

const documents = new WeakMap<XlsxEditorInstance, Document>();
/** The document a test viewer is mounted in. */
const document_of = (viewer: XlsxEditorInstance): Document => documents.get(viewer)!;

async function open(bytes: Uint8Array) {
  const win = new Window();
  const host = (win.document as unknown as Document).createElement('div');
  (win.document as unknown as Document).body.appendChild(host);
  const viewer = createXlsxEditor({ fonts: FONTS });
  viewers.push(viewer);
  viewer.attach(host);
  documents.set(viewer, win.document as unknown as Document);
  await viewer.load(bytes);
  flushEditor(viewer);
  return viewer;
}

describe('conditional formatting and the toolbar', () => {
  test('the bold state shows the cell format, not the format of a true rule', async () => {
    const rule =
      '<conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>';
    const bytes = buildXlsx({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]), afterMerges: rule }],
      styles: { dxfs: ['<dxf><font><b/></font></dxf>'] },
    });
    // The rule is true: the overlay that layout applies makes A1 bold.
    const read = readXlsx(bytes);
    if (!read.ok) throw new Error('read');
    const index = workbookIndex(read.pkg)!;
    const root = sheetIndex(xmlPartOf(read.pkg, sheetByName(index, 'S')!.partName)!).root;
    const base = cellFormat(read.pkg, index, undefined).format;
    const value = { kind: 'number', value: 1 } as const;
    expect(conditionalFormat(read.pkg, index, root, 0, 0, value, base).font.bold).toBe(true);
    const viewer = await open(bytes);
    // The grid paints A1 bold; the press sets bold on the cell format, which is off.
    expect(toolbarCommandState(viewer.snapshot(), 'font.bold').active).toBe(false);
  });
});

describe('conditional formatting signals', () => {
  const cellIs = (operator: string, operand: string) =>
    `<conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="0" priority="1" operator="${operator}"><formula>${operand}</formula></cfRule></conditionalFormatting>`;
  const book = (afterMerges: string, dxf = '<dxf><font><b/></font></dxf>', ext?: string) =>
    buildXlsx({
      sheets: [
        { name: 'S', rows: row(1, [c('A1', 1), c('B1', 2)]), afterMerges, ...(ext ? { ext } : {}) },
      ],
      styles: { dxfs: [dxf] },
    });

  test('a true rule gives a hint to each control whose state it changes', async () => {
    const viewer = await open(book(cellIs('greaterThan', '0')));
    expect(viewer.snapshot().conditionalFormat?.bold).toBe(true);
    expect(viewer.snapshot().activeFormat?.bold).toBe(false);
    const hint = (id: ChromeSlotId) => toolbarCommandState(viewer.snapshot(), id).hint;
    expect(hint('font.bold')).toEqual({ key: 'hint.conditionalFormat' });
    // The rule sets bold only: other controls, and controls that are not format controls, get
    // no hint.
    for (const id of ['font.italic', 'fill.color', 'alignment.left', 'history.undo'] as const)
      expect({ id, hint: hint(id) }).toEqual({ id, hint: undefined });
    // B1 is outside the rule's range.
    viewer.select({ row: 0, col: 1 }, { row: 0, col: 1 });
    expect(viewer.snapshot().conditionalFormat).toBeNull();
    expect(hint('font.bold')).toBeUndefined();
  });

  test('a formula cell without a cached value gives no hint, as layout paints no rule', async () => {
    const bytes = buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: '<row r="1"><c r="A1"><f>B1</f></c></row>',
          afterMerges: cellIs('equal', '0'),
        },
      ],
      styles: { dxfs: ['<dxf><font><b/></font></dxf>'] },
    });
    const viewer = await open(bytes);
    // An empty value compares as 0, so without the guard the rule would be true here.
    expect(viewer.snapshot().conditionalFormat).toBeNull();
    expect(toolbarCommandState(viewer.snapshot(), 'font.bold').hint).toBeUndefined();
  });

  test('a rule that sets a state that the cell format already has gives no hint', async () => {
    const viewer = await open(book(cellIs('greaterThan', '0')));
    expect(viewer.snapshot().conditionalFormat?.bold).toBe(true);
    // Bold on the cell format: the rule changes the format object, not the control state.
    expect(viewer.toggleFormat('bold').ok).toBe(true);
    expect(viewer.snapshot().activeFormat?.bold).toBe(true);
    expect(viewer.snapshot().conditionalFormat).toBeNull();
    expect(toolbarCommandState(viewer.snapshot(), 'font.bold').hint).toBeUndefined();
  });

  test('a fill rule gives the hint to the fill control only', async () => {
    const fill = '<dxf><fill><patternFill><bgColor rgb="FFFFFF00"/></patternFill></fill></dxf>';
    const viewer = await open(book(cellIs('greaterThan', '0'), fill));
    expect(viewer.snapshot().conditionalFormat?.fillColor).toBe('FFFF00');
    expect(toolbarCommandState(viewer.snapshot(), 'fill.color').hint?.key).toBe(
      'hint.conditionalFormat'
    );
    expect(toolbarCommandState(viewer.snapshot(), 'font.bold').hint).toBeUndefined();
  });

  test('a cell of a merged range reads the rules of the top-left cell', async () => {
    const bytes = buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 1)]),
          merges: ['A1:B1'],
          afterMerges: cellIs('greaterThan', '0'),
        },
      ],
      styles: { dxfs: ['<dxf><font><b/></font></dxf>'] },
    });
    const viewer = await open(bytes);
    // A press on B1 makes A1, the top-left cell of the range, the active cell.
    viewer.select({ row: 0, col: 1 }, { row: 0, col: 1 });
    expect(viewer.snapshot().activeCell?.ref).toBe('A1');
    expect(viewer.snapshot().conditionalFormat?.bold).toBe(true);
    expect(toolbarCommandState(viewer.snapshot(), 'font.bold').hint?.key).toBe(
      'hint.conditionalFormat'
    );
  });

  test('a false rule, or a true rule that changes nothing, gives no hint', async () => {
    const off = await open(book(cellIs('lessThan', '0')));
    expect(off.snapshot().conditionalFormat).toBeNull();
    const same = await open(book(cellIs('greaterThan', '0'), '<dxf/>'));
    expect(same.snapshot().conditionalFormat).toBeNull();
    expect(toolbarCommandState(same.snapshot(), 'font.bold').hint).toBeUndefined();
  });

  test('more shown rules than one cell walks give the notice', async () => {
    const many = (n: number) =>
      Array.from(
        { length: n },
        (_, i) =>
          `<conditionalFormatting sqref="A${i + 1}"><cfRule type="cellIs" dxfId="0" priority="${i + 1}" operator="equal"><formula>0</formula></cfRule></conditionalFormatting>`
      ).join('');
    expect((await open(book(many(64)))).snapshot().conditionalHidden).toBe(false);
    expect((await open(book(many(65)))).snapshot().conditionalHidden).toBe(true);
  });

  test('a layout past its check budget gives the notice', async () => {
    const ranges = Array.from({ length: 256 }, () => 'A1:A1048576 XFD1:XFD1048576').join(' ');
    const rules = [1, 2]
      .map(
        (p) =>
          `<conditionalFormatting sqref="${ranges}"><cfRule type="cellIs" dxfId="0" priority="${p}" operator="equal"><formula>0</formula></cfRule></conditionalFormatting>`
      )
      .join('');
    const bytes = buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: Array.from(
            { length: 2000 },
            (_, i) => `<row r="${i + 1}" ht="1" customHeight="1"/>`
          ).join(''),
          formatPr: 'defaultRowHeight="1" customHeight="1"',
          cols: '<cols><col min="1" max="16384" width="0.2" customWidth="1"/></cols>',
          afterMerges: rules,
        },
        { name: 'T', rows: row(1, [c('A1', 1)]) },
      ],
      styles: { dxfs: ['<dxf><font><b/></font></dxf>'] },
    });
    const viewer = await open(bytes);
    expect(viewer.snapshot().conditionalHidden).toBe(false);
    // A1 is empty, and `equal 0` is true there: before the cut, bold gets the hint.
    expect(viewer.snapshot().conditionalFormat?.bold).toBe(true);
    expect(toolbarCommandState(viewer.snapshot(), 'font.bold').hint).toBeDefined();
    // happy-dom has no layout: give the grid a size, and lay it out again.
    const scroller = document_of(viewer).querySelector('.xlsx-scroller')!;
    let height = 1000;
    let width = 1500;
    Object.defineProperty(scroller, 'clientWidth', { get: () => width });
    Object.defineProperty(scroller, 'clientHeight', { get: () => height });
    viewer.select({ row: 0, col: 0 }, { row: 0, col: 0 });
    flushEditor(viewer);
    expect(viewer.snapshot().conditionalHidden).toBe(true);
    // Past the budget, the grid can show a cell format where a rule is true: no hint.
    expect(viewer.snapshot().activeCell?.ref).toBe('A1');
    expect(viewer.snapshot().conditionalFormat).toBeNull();
    expect(toolbarCommandState(viewer.snapshot(), 'font.bold').hint).toBeUndefined();
    // A smaller window stays inside the budget, and the notice stays: it never toggles with
    // the window.
    height = 20;
    width = 20;
    viewer.select({ row: 0, col: 2 }, { row: 0, col: 2 });
    flushEditor(viewer);
    expect(viewer.snapshot().conditionalHidden).toBe(true);
    // Another sheet clears the flag, and a return to the sheet starts without it: the small
    // window stays inside the budget.
    expect(viewer.setActiveSheet(1)).toBe(true);
    flushEditor(viewer);
    expect(viewer.snapshot().conditionalHidden).toBe(false);
    expect(viewer.setActiveSheet(0)).toBe(true);
    flushEditor(viewer);
    expect(viewer.snapshot().conditionalHidden).toBe(false);
    // A large window cuts again. Another workbook then opens at the same sheet index without
    // the flag, and its true rule gives the hint.
    height = 1000;
    width = 1500;
    viewer.select({ row: 0, col: 1 }, { row: 0, col: 1 });
    flushEditor(viewer);
    expect(viewer.snapshot().conditionalHidden).toBe(true);
    await viewer.load(book(cellIs('greaterThan', '0')));
    flushEditor(viewer);
    expect(viewer.snapshot().activeSheet).toBe(0);
    expect(viewer.snapshot().conditionalHidden).toBe(false);
    expect(viewer.snapshot().conditionalFormat?.bold).toBe(true);
  });

  test('the budget flag stays with its sheet when a sheet op moves or deletes it', async () => {
    const ranges = Array.from({ length: 256 }, () => 'A1:A1048576 XFD1:XFD1048576').join(' ');
    const rules = [1, 2]
      .map(
        (p) =>
          `<conditionalFormatting sqref="${ranges}"><cfRule type="cellIs" dxfId="0" priority="${p}" operator="equal"><formula>0</formula></cfRule></conditionalFormatting>`
      )
      .join('');
    const bytes = buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: Array.from(
            { length: 2000 },
            (_, i) => `<row r="${i + 1}" ht="1" customHeight="1"/>`
          ).join(''),
          formatPr: 'defaultRowHeight="1" customHeight="1"',
          cols: '<cols><col min="1" max="16384" width="0.2" customWidth="1"/></cols>',
          afterMerges: rules,
        },
        { name: 'T', rows: row(1, [c('A1', 0)]), afterMerges: cellIs('equal', '0') },
      ],
      styles: { dxfs: ['<dxf><font><b/></font></dxf>'] },
    });
    const viewer = await open(bytes);
    const scroller = document_of(viewer).querySelector('.xlsx-scroller')!;
    let size = 1000;
    Object.defineProperty(scroller, 'clientWidth', { get: () => size * 1.5 });
    Object.defineProperty(scroller, 'clientHeight', { get: () => size });
    viewer.select({ row: 0, col: 0 }, { row: 0, col: 0 });
    flushEditor(viewer);
    expect(viewer.snapshot().conditionalHidden).toBe(true);
    expect(viewer.setMode('edit')).toBe(true);
    // A move keeps the flag on its sheet at the new index, also when the window gets small.
    size = 20;
    expect(viewer.moveSheet(0, 1).ok).toBe(true);
    flushEditor(viewer);
    expect(viewer.snapshot().activeSheet).toBe(1);
    expect(viewer.snapshot().conditionalHidden).toBe(true);
    expect(viewer.moveSheet(1, 0).ok).toBe(true);
    flushEditor(viewer);
    // A delete of the shown sheet shows T at the same index. The flag of S does not pass to T:
    // T has no notice, and its true rule gives the hint.
    expect(viewer.deleteSheet(0).ok).toBe(true);
    flushEditor(viewer);
    expect(viewer.snapshot().activeSheet).toBe(0);
    expect(viewer.snapshot().activeCell?.ref).toBe('A1');
    expect(viewer.snapshot().conditionalHidden).toBe(false);
    expect(viewer.snapshot().conditionalFormat?.bold).toBe(true);
    expect(toolbarCommandState(viewer.snapshot(), 'font.bold').hint).toBeDefined();
  });

  test('a copy and the cell announcement give the text that the grid shows under a rule', async () => {
    const percent = '<dxf><numFmt numFmtId="200" formatCode="0%"/></dxf>';
    const viewer = await open(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            rows: row(1, [c('A1', 0.5), c('B1', -0.5), c('C1', '')]),
            afterMerges: `<conditionalFormatting sqref="A1:C1"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>`,
          },
        ],
        styles: { dxfs: [percent] },
      })
    );
    // A1 is true and shows 50%; B1 is false and keeps its cell format; C1 is empty.
    viewer.select({ row: 0, col: 0 }, { row: 0, col: 2 });
    const copy = viewer.copy();
    expect(copy.ok && copy.content.text).toBe('50%\t-0.5\t');
    const read = readXlsx(viewer.save()!);
    if (!read.ok) throw new Error('read');
    const index = workbookIndex(read.pkg)!;
    const partName = sheetByName(index, 'S')!.partName;
    const t = (key: string, args?: Record<string, string>) => `${key} ${args?.content ?? ''}`;
    const said = (col: number, rules = true) =>
      cellAnnouncement(read.pkg, index, partName, { row: 0, col }, EN_US, t, rules);
    expect(said(0)).toBe('grid.activeCell 50%');
    expect(said(1)).toBe('grid.activeCell -0.5');
    // After a budget cut, the cell format gives the text.
    expect(said(0, false)).toBe('grid.activeCell 0.5');
  });

  test('after a budget cut, a copy and the announcement give the text of the cell format', async () => {
    const ranges = Array.from({ length: 255 }, () => 'A1:A1048576 XFD1:XFD1048576').join(' ');
    const rules = [1, 2]
      .map(
        (p) =>
          `<conditionalFormatting sqref="${ranges}"><cfRule type="cellIs" dxfId="0" priority="${p}" operator="equal"><formula>0</formula></cfRule></conditionalFormatting>`
      )
      .join('');
    const percent = `<conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="1" priority="3" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>`;
    const bytes = buildXlsx({
      sheets: [
        {
          name: 'S',
          rows:
            row(1, [c('A1', 0.5)]) +
            Array.from(
              { length: 2000 },
              (_, i) => `<row r="${i + 2}" ht="1" customHeight="1"/>`
            ).join(''),
          formatPr: 'defaultRowHeight="1" customHeight="1"',
          cols: '<cols><col min="1" max="16384" width="0.2" customWidth="1"/></cols>',
          afterMerges: rules + percent,
        },
      ],
      styles: {
        dxfs: [
          '<dxf><font><b/></font></dxf>',
          '<dxf><numFmt numFmtId="200" formatCode="0%"/></dxf>',
        ],
      },
    });
    const viewer = await open(bytes);
    const live = () => document_of(viewer).querySelector('.xlsx-live')?.textContent ?? '';
    const copyText = () => {
      const r = viewer.copy();
      return r.ok ? r.content.text : r.refusal.code;
    };
    // Before a cut, A1 shows 50%.
    viewer.select({ row: 0, col: 1 }, { row: 0, col: 1 });
    viewer.select({ row: 0, col: 0 }, { row: 0, col: 0 });
    flushEditor(viewer);
    expect(viewer.snapshot().conditionalHidden).toBe(false);
    expect(live()).toContain('50%');
    expect(copyText()).toBe('50%');
    const scroller = document_of(viewer).querySelector('.xlsx-scroller')!;
    Object.defineProperty(scroller, 'clientWidth', { get: () => 1500 });
    Object.defineProperty(scroller, 'clientHeight', { get: () => 1000 });
    viewer.select({ row: 0, col: 1 }, { row: 0, col: 1 });
    flushEditor(viewer);
    expect(viewer.snapshot().conditionalHidden).toBe(true);
    // After the cut, the grid can show the cell format: the copy and the announcement follow it.
    viewer.select({ row: 0, col: 0 }, { row: 0, col: 0 });
    flushEditor(viewer);
    expect(live()).toContain('0.5');
    expect(live()).not.toContain('50%');
    expect(copyText()).toBe('0.5');
  });

  test('a copy past its check budget is refused, and the same range gives the same result', async () => {
    const letters = (i: number): string =>
      i < 26 ? String.fromCharCode(65 + i) : letters(Math.floor(i / 26) - 1) + letters(i % 26);
    const ranges = Array.from({ length: 256 }, () => 'A1:DX200 XFD1:XFD1048576').join(' ');
    const rules = [1, 2]
      .map(
        (p) =>
          `<conditionalFormatting sqref="${ranges}"><cfRule type="cellIs" dxfId="0" priority="${p}" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>`
      )
      .join('');
    const rows = Array.from({ length: 128 }, (_, r) =>
      row(
        r + 1,
        Array.from({ length: 128 }, (_, k) => c(`${letters(k)}${r + 1}`, 0.5))
      )
    ).join('');
    const viewer = await open(
      buildXlsx({
        sheets: [{ name: 'S', rows, afterMerges: rules }],
        styles: { dxfs: ['<dxf><numFmt numFmtId="200" formatCode="0%"/></dxf>'] },
      })
    );
    expect(viewer.snapshot().conditionalHidden).toBe(false);
    // One cell copies inside the budget, under the rule.
    viewer.select({ row: 0, col: 0 }, { row: 0, col: 0 });
    const one = viewer.copy();
    expect(one.ok && one.content.text).toBe('50%');
    // The whole range runs out of the budget, and the copy charges every row in full, so the
    // result depends only on the range: each copy and each cut of it is refused, never mixed.
    viewer.select({ row: 0, col: 0 }, { row: 127, col: 127 });
    const code = (r: { ok: boolean; refusal?: { code: string } }) =>
      r.ok ? null : r.refusal?.code;
    expect(code(viewer.copy())).toBe('copy-conditional-limit');
    expect(code(viewer.copy())).toBe('copy-conditional-limit');
    expect(viewer.setMode('edit')).toBe(true);
    expect(code(viewer.cut())).toBe('copy-conditional-limit');
    flushEditor(viewer);
    // The status strip shows the refusal. The grid did not change, so the sheet keeps its rules:
    // no notice, the hint stays, and a smaller copy again gives the shown text.
    expect(viewer.snapshot().refusal?.code).toBe('copy-conditional-limit');
    expect(viewer.snapshot().conditionalHidden).toBe(false);
    viewer.select({ row: 0, col: 0 }, { row: 0, col: 0 });
    flushEditor(viewer);
    expect(viewer.snapshot().conditionalFormat).not.toBeNull();
    const again = viewer.copy();
    expect(again.ok && again.content.text).toBe('50%');
  });

  test('a copy charges each row in full, so a second copy of a range is refused too', async () => {
    // Each rule has 512 ranges, and only `A1:A3950` meets the rows of the copy: the cells are
    // cheap, and the rows cost 3,950 × 1,024 checks, more than the budget.
    const ranges = ['A1:A3950', ...Array.from({ length: 511 }, () => 'XFD5000:XFD5000')].join(' ');
    const rules = [1, 2]
      .map(
        (p) =>
          `<conditionalFormatting sqref="${ranges}"><cfRule type="cellIs" dxfId="0" priority="${p}" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>`
      )
      .join('');
    const rows = Array.from({ length: 3950 }, (_, r) => row(r + 1, [c(`A${r + 1}`, 0.5)])).join('');
    const viewer = await open(
      buildXlsx({
        sheets: [{ name: 'S', rows, afterMerges: rules }],
        styles: { dxfs: ['<dxf><numFmt numFmtId="200" formatCode="0%"/></dxf>'] },
      })
    );
    expect(viewer.snapshot().conditionalHidden).toBe(false);
    viewer.select({ row: 0, col: 0 }, { row: 3949, col: 0 });
    const code = (r: { ok: boolean; refusal?: { code: string } }) =>
      r.ok ? null : r.refusal?.code;
    // The first copy fills the row cache. The second copy still pays for each row once.
    expect(code(viewer.copy())).toBe('copy-conditional-limit');
    expect(code(viewer.copy())).toBe('copy-conditional-limit');
  });

  test('a copy charges each row once, so a wide range inside the row budget copies', async () => {
    const letters = (i: number): string =>
      i < 26 ? String.fromCharCode(65 + i) : letters(Math.floor(i / 26) - 1) + letters(i % 26);
    // Each rule has 512 ranges: 1,024 pairs. 100 rows cost 102,400 checks; a charge for each of
    // the 4,000 cells would cost 4,096,000, more than the budget.
    const ranges = ['A1:AN100', ...Array.from({ length: 511 }, () => 'XFD5000:XFD5000')].join(' ');
    const rules = [1, 2]
      .map(
        (p) =>
          `<conditionalFormatting sqref="${ranges}"><cfRule type="cellIs" dxfId="0" priority="${p}" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>`
      )
      .join('');
    const rows = Array.from({ length: 100 }, (_, r) =>
      row(
        r + 1,
        Array.from({ length: 40 }, (_, k) => c(`${letters(k)}${r + 1}`, 0.5))
      )
    ).join('');
    const offRule = Array.from({ length: 100 }, (_, r) =>
      row(
        r + 1,
        Array.from({ length: 40 }, (_, k) => c(`${letters(k + 100)}${r + 1}`, 0.5))
      )
    );
    const viewer = await open(
      buildXlsx({
        sheets: [
          { name: 'S', rows, afterMerges: rules },
          { name: 'T', rows: offRule.join(''), afterMerges: rules },
        ],
        styles: { dxfs: ['<dxf><numFmt numFmtId="200" formatCode="0%"/></dxf>'] },
      })
    );
    viewer.select({ row: 0, col: 0 }, { row: 99, col: 39 });
    const under = viewer.copy();
    if (!under.ok) throw new Error(under.refusal.code);
    const texts = under.content.text.split(/\t|\r\n/);
    expect(texts.length).toBe(4000);
    expect(texts.every((text) => text === '50%')).toBe(true);
    // Sheet T has the same rules, and no rule meets the copied cells: it copies too.
    expect(viewer.setActiveSheet(1)).toBe(true);
    flushEditor(viewer);
    viewer.select({ row: 0, col: 100 }, { row: 99, col: 139 });
    const off = viewer.copy();
    if (!off.ok) throw new Error(off.refusal.code);
    expect(off.content.text.split(/\t|\r\n/).every((text) => text === '0.5')).toBe(true);
  });

  test('a sheet with a rule that the editor does not show gets the notice', async () => {
    const shown = await open(book(cellIs('greaterThan', '0')));
    expect(shown.snapshot().conditionalHidden).toBe(false);
    const expression = await open(book(CF));
    expect(expression.snapshot().conditionalHidden).toBe(true);
    const reference = await open(book(cellIs('greaterThan', '$B$1')));
    expect(reference.snapshot().conditionalHidden).toBe(true);
    const ext =
      '<ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings/></ext>';
    const extended = await open(book('', undefined, ext));
    expect(extended.snapshot().conditionalHidden).toBe(true);
  });
});
