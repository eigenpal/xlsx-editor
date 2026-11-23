// The stale notice of a sheet (design D3) counts the formula cells that the stale flag counts:
// also a formula cell that the sheet index skips, so a value edit elsewhere shows the notice on
// that sheet and the save keeps the flag. A wrapped formula in a shown cell also gets the stale
// mark in the grid. The grid does not show a wrapped row or an `x:c` outside a row, so for
// those the notice is the only mark.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { createXlsxEditor, type XlsxEditorInstance } from '../index.ts';
import { flushEditor } from './flush-editor.ts';
import type { FontSource } from '../../layout/index.ts';
import { readXlsx, sheetByName, workbookIndex } from '../../store/index.ts';
import { EN_US } from '../../cascade/index.ts';
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
  return { viewer, host };
}

const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
// The prefix that `@Requires` names is declared, as the reader requires.
const X14 = 'urn:example:x14';

/**
 * The forms of a formula cell that the sheet index skips: an `x:c` directly in `x:sheetData`
 * (the typed tree keeps it generic), a row inside `mc:AlternateContent`, and an `x:f` inside
 * `mc:AlternateContent` in a cell.
 */
const FORMS: readonly (readonly [string, string])[] = [
  ['a generic cell', '<c r="A3"><f>1+1</f><v>2</v></c>'],
  [
    'a wrapped row',
    `<mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}"><mc:Choice Requires="x14">` +
      '<row r="3"><c r="A3"><f>1+1</f><v>2</v></c></row></mc:Choice></mc:AlternateContent>',
  ],
  [
    'a wrapped formula in a cell',
    `<row r="3"><c r="A3"><mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}"><mc:Choice Requires="x14">` +
      '<f>1+1</f></mc:Choice></mc:AlternateContent><v>2</v></c></row>',
  ],
];

describe('the stale notice and the formula scan', () => {
  for (const [name, rows] of FORMS)
    test(`${name}: the sheet shows the notice after an edit on another sheet`, async () => {
      const { viewer } = await mount(
        buildXlsx({
          sheets: [
            { name: 'S', rows: row(1, [c('A1', 1)]) },
            { name: 'T', rows: row(1, [c('A1', 1)]) + rows },
          ],
        })
      );
      expect(viewer.setActiveSheet(1)).toBe(true);
      expect(viewer.snapshot().stale).toBe(false);
      expect(viewer.setActiveSheet(0)).toBe(true);
      viewer.select({ row: 0, col: 0 });
      expect(viewer.enterValue('5').ok).toBe(true);
      // Sheet `S` has no formula cell, so it shows no notice.
      expect(viewer.snapshot().stale).toBe(false);
      expect(viewer.setActiveSheet(1)).toBe(true);
      expect(viewer.snapshot().stale).toBe(true);
      const saved = readXlsx(viewer.save()!);
      expect(saved.ok && workbookIndex(saved.pkg)!.fullCalcOnLoad).toBe(true);
    });

  test('a wrapped formula in a shown cell gets the stale mark in the grid', async () => {
    const { viewer, host } = await mount(
      buildXlsx({
        sheets: [{ name: 'T', rows: row(1, [c('A1', 1)]) + FORMS[2]![1] }],
        calcPr: 'calcId="191029" fullCalcOnLoad="1"',
      })
    );
    expect(viewer.snapshot().stale).toBe(true);
    flushEditor(viewer);
    expect(host.querySelector('[data-cell="A3"]')?.getAttribute('data-stale')).toBe('1');
    expect(host.querySelector('[data-cell="A1"]')?.hasAttribute('data-stale')).toBe(false);
    // The formula bar shows the formula, as the mark says.
    viewer.select({ row: 2, col: 0 });
    expect(viewer.snapshot().activeCell?.content).toBe('=1+1');
  });

  test('a cached value in the formula branch shows, and a workbook that is not stale marks nothing', async () => {
    const both =
      `<row r="3"><c r="A3"><mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}"><mc:Choice Requires="x14">` +
      '<f>1+1</f><v>2</v></mc:Choice><mc:Fallback><f>1+1</f><v>2</v></mc:Fallback>' +
      '</mc:AlternateContent></c></row>';
    const { viewer, host } = await mount(
      buildXlsx({ sheets: [{ name: 'T', rows: row(1, [c('A1', 1)]) + both }] })
    );
    flushEditor(viewer);
    const cell = host.querySelector('[data-cell="A3"]');
    expect(cell?.hasAttribute('data-stale')).toBe(false);
    expect(cell?.textContent).toContain('2');
    viewer.select({ row: 2, col: 0 });
    expect(viewer.snapshot().activeCell?.content).toBe('=1+1');
  });

  test('a cached value only in the Fallback shows with the formula of a Choice', async () => {
    const split =
      `<row r="3"><c r="A3"><mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}"><mc:Choice Requires="x14">` +
      '<f>3+4</f></mc:Choice><mc:Fallback><v>7</v></mc:Fallback></mc:AlternateContent></c></row>';
    const { viewer, host } = await mount(
      buildXlsx({ sheets: [{ name: 'T', rows: row(1, [c('A1', 1)]) + split }] })
    );
    flushEditor(viewer);
    const cell = host.querySelector('[data-cell="A3"]');
    expect(cell?.hasAttribute('data-stale')).toBe(false);
    expect(cell?.textContent).toContain('7');
    viewer.select({ row: 2, col: 0 });
    expect(viewer.snapshot().activeCell?.content).toBe('=3+4');
  });

  test('a value next to another formula is not the result of the shown formula', async () => {
    const pair =
      `<row r="3"><c r="A3"><mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}"><mc:Choice Requires="x14">` +
      '<f>B1*2</f><v>5</v></mc:Choice><mc:Fallback><f>OLD()</f></mc:Fallback>' +
      '</mc:AlternateContent></c></row>';
    const { viewer, host } = await mount(
      buildXlsx({ sheets: [{ name: 'T', rows: row(1, [c('A1', 1)]) + pair }] })
    );
    flushEditor(viewer);
    // The Fallback formula shows, and the result of the Choice formula does not. The Fallback
    // formula has no cached value, so the cell is stale.
    const cell = host.querySelector('[data-cell="A3"]');
    expect(cell?.textContent ?? '').not.toContain('5');
    expect(cell?.getAttribute('data-stale')).toBe('1');
    viewer.select({ row: 2, col: 0 });
    expect(viewer.snapshot().activeCell?.content).toBe('=OLD()');
  });

  test('a direct formula with no text shows its value', async () => {
    const empty = '<row r="3"><c r="A3"><f></f><v>3</v></c></row>';
    const { viewer } = await mount(
      buildXlsx({ sheets: [{ name: 'T', rows: row(1, [c('A1', 1)]) + empty }] })
    );
    viewer.select({ row: 2, col: 0 });
    expect(viewer.snapshot().activeCell?.content).toBe('3');
  });

  test('a wrapped shared formula child shows the formula of its direct master', async () => {
    const member =
      `<row r="3"><c r="A3"><mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}"><mc:Choice Requires="x14">` +
      '<f t="shared" si="0"/></mc:Choice></mc:AlternateContent><v>9</v></c></row>';
    const master =
      '<row r="2"><c r="A2"><f t="shared" ref="A2:A3" si="0">A1*2</f><v>2</v></c></row>';
    const { viewer } = await mount(
      buildXlsx({ sheets: [{ name: 'T', rows: row(1, [c('A1', 1)]) + master + member }] })
    );
    viewer.select({ row: 2, col: 0 });
    // The child has no text of its own: the formula bar shows the master's formula moved to the
    // child, as the grid marks the child a formula and shows its cached value.
    expect(viewer.snapshot().activeCell?.content).toBe('=A2*2');
    viewer.select({ row: 1, col: 0 });
    expect(viewer.snapshot().activeCell?.content).toBe('=A1*2');
  });

  test('the live region says the formula and the stale words for a wrapped formula', () => {
    const read = readXlsx(
      buildXlsx({
        sheets: [{ name: 'T', rows: row(1, [c('A1', 1)]) + FORMS[2]![1] }],
        calcPr: 'calcId="191029" fullCalcOnLoad="1"',
      })
    );
    if (!read.ok) throw new Error('read');
    const index = workbookIndex(read.pkg)!;
    const t = (key: string, args?: Record<string, string>) =>
      `${key}[${args?.formula ?? ''}|${args?.text ?? ''}]`;
    const said = cellAnnouncement(
      read.pkg,
      index,
      sheetByName(index, 'T')!.partName,
      { row: 2, col: 0 },
      EN_US,
      t
    );
    expect(said).toBe('grid.staleResult[|grid.activeFormula[=1+1|]]');
  });

  test('a workbook opened stale shows the notice on such a sheet', async () => {
    const { viewer } = await mount(
      buildXlsx({
        sheets: [{ name: 'T', rows: row(1, [c('A1', 1)]) + FORMS[0]![1] }],
        calcPr: 'calcId="191029" fullCalcOnLoad="1"',
      })
    );
    expect(viewer.snapshot().stale).toBe(true);
  });
});
