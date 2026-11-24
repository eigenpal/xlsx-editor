// The sheet commands of a sheet that the editor added, in a workbook with a chart. The chart
// refuses a rename and a delete of each sheet of the file (`workbookPartBlocker`). A sheet that
// the add made takes both, unless the chart names it (`c:f`, ECMA-376 Part 1, 21.2.2.65). The
// sheet menu state (`sheetCommandState`) and the store agree on each sheet, and after a save
// and a reopen the added sheet is a sheet of the file.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { createXlsxEditor, sheetCommandState, type XlsxEditorInstance } from '../index.ts';
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

const translate = (key: string, params?: Readonly<Record<string, string | number>>) =>
  key === 'sheetTabs.newSheetName' ? `Sheet${params?.n}` : key;

/** Sheet1 and Sheet2, and a chart part whose series reads `ref`. */
const chartWorkbook = (ref: string) =>
  buildXlsx({
    sheets: [
      { name: 'Sheet1', rows: row(1, [c('A1', 1)]) },
      { name: 'Sheet2', rows: row(1, [c('A1', 2)]) },
    ],
    extraParts: [
      [
        'xl/charts/chart1.xml',
        'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
        new TextEncoder().encode(
          `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numRef><c:f>${ref}</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`
        ),
      ],
    ],
  });

let viewers: XlsxEditorInstance[] = [];
afterEach(() => {
  for (const v of viewers) v.destroy();
  viewers = [];
});

async function mount(bytes: Uint8Array) {
  const doc = new Window().document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.appendChild(host);
  const viewer = createXlsxEditor({ fonts: FONTS, translate });
  viewers.push(viewer);
  viewer.attach(host);
  await viewer.load(bytes);
  flushEditor(viewer);
  return viewer;
}

/** Whether the sheet menu enables Rename and Delete of the sheet at `index`. */
const menu = (v: XlsxEditorInstance, index: number) => [
  sheetCommandState(v.snapshot(), 'rename', index).enabled,
  sheetCommandState(v.snapshot(), 'delete', index).enabled,
];
const tabs = (v: XlsxEditorInstance) => v.snapshot().sheets.map((s) => s.name);
const code = (r: { ok: boolean; refusal?: { code: string } }) => (r.ok ? 'ok' : r.refusal?.code);

describe('a sheet that the editor added to a workbook with a chart', () => {
  // The add puts the new sheet after the shown sheet, Sheet1.
  test('takes Rename and Delete from the menu and the store; the sheets of the file do not', async () => {
    const v = await mount(chartWorkbook('Sheet1!$A$1'));
    expect(menu(v, 0)).toEqual([false, false]);
    expect(code(v.addSheet())).toBe('ok');
    expect(tabs(v)).toEqual(['Sheet1', 'Sheet3', 'Sheet2']);
    expect(v.snapshot().sheets[1]!.editRefusal).toBeNull();
    expect(v.snapshot().sheets[0]!.editRefusal?.code).toBe('structure-blocked');
    expect(menu(v, 1)).toEqual([true, true]);
    expect(code(v.renameSheet(1, 'Notes'))).toBe('ok');
    expect(menu(v, 1)).toEqual([true, true]);
    expect(code(v.enterValue('5'))).toBe('ok');
    expect(menu(v, 1)).toEqual([true, true]);
    expect(code(v.renameSheet(2, 'Data'))).toBe('structure-blocked');
    expect(menu(v, 2)).toEqual([false, false]);
    expect(code(v.deleteSheet(1))).toBe('ok');
    expect(tabs(v)).toEqual(['Sheet1', 'Sheet2']);
    // Undo gives the sheet back, and it stays an added sheet.
    expect(v.undo().ok).toBe(true);
    expect(tabs(v)).toEqual(['Sheet1', 'Notes', 'Sheet2']);
    expect(menu(v, 1)).toEqual([true, true]);
  });

  test('a chart formula that names it disables Rename and Delete, as the store refuses them', async () => {
    const v = await mount(chartWorkbook('Sheet3!$A$1'));
    expect(code(v.addSheet())).toBe('ok');
    expect(v.snapshot().sheets[1]!.editRefusal).toMatchObject({
      code: 'structure-blocked',
      about: { kind: 'chart' },
    });
    expect(menu(v, 1)).toEqual([false, false]);
    expect(sheetCommandState(v.snapshot(), 'rename', 1).disabledReason?.key).toBe(
      'disabledReason.sheetEditBlocked'
    );
    expect(code(v.renameSheet(1, 'Notes'))).toBe('structure-blocked');
    expect(code(v.deleteSheet(1))).toBe('structure-blocked');
    // A cell edit keeps the reason of each sheet (the snapshot does not scan again).
    expect(code(v.enterValue('5'))).toBe('ok');
    expect(menu(v, 1)).toEqual([false, false]);
  });

  test('is a sheet of the file after a save and a reopen', async () => {
    const v = await mount(chartWorkbook('Sheet1!$A$1'));
    expect(code(v.addSheet())).toBe('ok');
    const saved = v.save();
    if (!saved) throw new Error('no save');
    const reopened = await mount(saved);
    expect(tabs(reopened)).toEqual(['Sheet1', 'Sheet3', 'Sheet2']);
    expect(menu(reopened, 1)).toEqual([false, false]);
    expect(code(reopened.renameSheet(1, 'Notes'))).toBe('structure-blocked');
  });
});
