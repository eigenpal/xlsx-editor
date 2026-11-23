// One test for each scenario of the copy and paste requirements in the change
// `workbook-structure-editing`, from "Copy and paste in the workbook" to "Edits that keep a copy".
// Each test name holds the title of its scenario.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import {
  buildXlsx,
  c,
  row,
  type FixtureWorkbook,
} from '../../../../../scripts/fixtures/builder.ts';
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

let viewers: XlsxEditorInstance[] = [];
afterEach(() => {
  for (const v of viewers) v.destroy();
  viewers = [];
});

async function mount(wb: FixtureWorkbook) {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.appendChild(host);
  const viewer = createXlsxEditor({ fonts: FONTS, translate: (key) => key });
  viewers.push(viewer);
  viewer.attach(host);
  const scroller = host.querySelector('.xlsx-scroller') as HTMLElement;
  Object.defineProperty(scroller, 'clientWidth', { get: () => 800 });
  Object.defineProperty(scroller, 'clientHeight', { get: () => 400 });
  await viewer.load(buildXlsx(wb));
  flushEditor(viewer);
  return viewer;
}

function copyAt(v: XlsxEditorInstance, r: number, col: number) {
  v.select({ row: r, col });
  const out = v.copy();
  if (!out.ok) throw new Error(out.refusal.code);
  return out.content.token;
}

/** Paste at the cell, and give the refusal code, or null when the paste applies. */
function pasteAt(v: XlsxEditorInstance, token: string, r: number, col: number) {
  v.select({ row: r, col });
  const out = v.paste({ token, text: 'plain' });
  return out.ok ? null : out.refusal.code;
}

function content(v: XlsxEditorInstance, r: number, col: number) {
  v.select({ row: r, col });
  return v.snapshot().activeCell?.content ?? '';
}

/** Paste on the first sheet at `D20`: the paste is refused, is refused again, and writes nothing. */
function expectExpired(v: XlsxEditorInstance, token: string) {
  expect(v.setActiveSheet(0)).toBe(true);
  expect(pasteAt(v, token, 19, 3)).toBe('copy-expired');
  expect(pasteAt(v, token, 19, 3)).toBe('copy-expired');
  expect(content(v, 19, 3)).toBe('');
}

const ABSOLUTE: FixtureWorkbook = {
  sheets: [
    { name: 'S', rows: row(5, [c('A5', 7), c('B5', 7, { f: '$A$5' })]) },
    { name: 'T', rows: row(1, [c('A1', 1)]) },
  ],
};
const NAMED: FixtureWorkbook = {
  sheets: [
    { name: 'T', rows: row(1, [c('A1', 1, { f: 'S!A1' })]) },
    { name: 'S', rows: row(1, [c('A1', 1), c('B1', 2)]) },
  ],
};
const SPAN: FixtureWorkbook = {
  sheets: [
    { name: 'S', rows: row(1, [c('A1', 3, { f: 'SUM(Jan:Mar!A1)' })]) },
    { name: 'Jan', rows: row(1, [c('A1', 1)]) },
    { name: 'Feb', rows: row(1, [c('A1', 1)]) },
    { name: 'Mar', rows: row(1, [c('A1', 1)]) },
    { name: 'Other', rows: row(1, [c('A1', 2)]) },
  ],
};
const SCOPED: FixtureWorkbook = {
  sheets: [
    { name: 'S', rows: row(1, [c('A1', 1, { f: 'Rate' }), c('B1', 1)]) },
    { name: 'T', rows: row(1, [c('A1', 2)]) },
  ],
  definedNames: '<definedName name="Rate" localSheetId="0">S!$B$1</definedName>',
};

const insertRow = (v: XlsxEditorInstance) => {
  v.select({ row: 0, col: 0 });
  expect(v.insertRows().ok).toBe(true);
};

describe('the copy and paste requirements', () => {
  test('scenario "A relative formula"', async () => {
    const v = await mount({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 1, { f: 'A1' })]) }],
    });
    v.select({ row: 0, col: 1 });
    expect(v.toggleFormat('bold').ok).toBe(true);
    const token = copyAt(v, 0, 1);
    expect(pasteAt(v, token, 2, 1)).toBeNull();
    expect(content(v, 2, 1)).toBe('=A3');
    expect(v.snapshot().activeFormat).toMatchObject({ bold: true });
  });

  test('scenario "A reference that leaves the sheet"', async () => {
    const v = await mount({
      sheets: [{ name: 'S', rows: row(2, [c('B2', 1, { f: 'A1' })]) }],
    });
    const token = copyAt(v, 1, 1);
    expect(pasteAt(v, token, 1, 0)).toBe('paste-reference-outside');
    expect(content(v, 1, 0)).toBe('');
  });

  test('scenario "A sheet rename after the copy"', async () => {
    const v = await mount(NAMED);
    const token = copyAt(v, 0, 0);
    expect(v.renameSheet(1, 'Renamed').ok).toBe(true);
    expectExpired(v, token);
  });

  test('scenario "A row insert after the copy"', async () => {
    const v = await mount(ABSOLUTE);
    const token = copyAt(v, 4, 1);
    insertRow(v);
    expectExpired(v, token);
  });

  test('scenario "A column delete after the copy"', async () => {
    const v = await mount(ABSOLUTE);
    const token = copyAt(v, 4, 1);
    v.select({ row: 0, col: 0 });
    expect(v.deleteColumns().ok).toBe(true);
    expectExpired(v, token);
  });

  test('scenario "A row insert on another sheet after the copy"', async () => {
    const v = await mount(ABSOLUTE);
    const token = copyAt(v, 4, 1);
    expect(v.setActiveSheet(1)).toBe(true);
    insertRow(v);
    expectExpired(v, token);
  });

  test('scenario "An undo of a row insert after the copy"', async () => {
    const v = await mount(ABSOLUTE);
    const token = copyAt(v, 4, 1);
    insertRow(v);
    expect(v.undo().ok).toBe(true);
    expectExpired(v, token);
  });

  test('scenario "A redo of a row insert after the copy"', async () => {
    const v = await mount(ABSOLUTE);
    insertRow(v);
    expect(v.undo().ok).toBe(true);
    const token = copyAt(v, 4, 1);
    expect(v.redo().ok).toBe(true);
    expectExpired(v, token);
  });

  test('scenario "A row insert after a copy of values"', async () => {
    const v = await mount(ABSOLUTE);
    const token = copyAt(v, 4, 0);
    insertRow(v);
    expect(pasteAt(v, token, 19, 3)).toBeNull();
    expect(content(v, 19, 3)).toBe('7');
  });

  test('scenario "A cell edit on a named sheet after the copy"', async () => {
    const v = await mount(NAMED);
    const token = copyAt(v, 0, 0);
    expect(v.setActiveSheet(1)).toBe(true);
    v.select({ row: 0, col: 1 });
    expect(v.enterValue('5').ok).toBe(true);
    expect(v.setActiveSheet(0)).toBe(true);
    expect(pasteAt(v, token, 19, 3)).toBeNull();
    expect(content(v, 19, 3)).toBe('=S!D20');
  });

  test('scenario "A change of a copied cell format after the copy"', async () => {
    const v = await mount({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 2)]) }],
    });
    v.select({ row: 0, col: 0 });
    expect(v.toggleFormat('bold').ok).toBe(true);
    const token = copyAt(v, 0, 0);
    expect(v.undo().ok).toBe(true);
    v.select({ row: 0, col: 1 });
    expect(v.toggleFormat('italic').ok).toBe(true);
    expectExpired(v, token);
  });

  test('scenario "A rename of the source sheet after the copy of a scoped name"', async () => {
    const v = await mount(SCOPED);
    const token = copyAt(v, 0, 0);
    expect(v.renameSheet(0, 'Renamed').ok).toBe(true);
    expect(v.setActiveSheet(1)).toBe(true);
    expect(pasteAt(v, token, 2, 0)).toBe('copy-expired');
    expect(content(v, 2, 0)).toBe('');
  });

  test('scenario "A paste on the source sheet after the copy of a scoped name"', async () => {
    const v = await mount(SCOPED);
    const token = copyAt(v, 0, 0);
    expect(v.renameSheet(0, 'Renamed').ok).toBe(true);
    expect(pasteAt(v, token, 2, 0)).toBeNull();
    expect(content(v, 2, 0)).toBe('=Rate');
  });

  test('scenario "A move of a sheet out of a span after the copy"', async () => {
    const v = await mount(SPAN);
    const token = copyAt(v, 0, 0);
    const moved = v.moveSheet(2, 4);
    expect(moved.ok ? null : moved.refusal.code).toBe('sheet-range-blocked');
    expect(v.snapshot().sheets.map((s) => s.name)).toEqual(['S', 'Jan', 'Feb', 'Mar', 'Other']);
    expect(pasteAt(v, token, 19, 3)).toBeNull();
    expect(content(v, 19, 3)).toBe('=SUM(Jan:Mar!D20)');
  });

  test('scenario "A move of a sheet outside a span after the copy"', async () => {
    const v = await mount(SPAN);
    const token = copyAt(v, 0, 0);
    expect(v.moveSheet(4, 0).ok).toBe(true);
    expect(v.setActiveSheet(1)).toBe(true);
    expect(pasteAt(v, token, 19, 3)).toBe('copy-expired');
    expect(content(v, 19, 3)).toBe('');
  });

  test('scenario "An add of a sheet after the copy of a span"', async () => {
    const v = await mount(SPAN);
    const token = copyAt(v, 0, 0);
    expect(v.addSheet().ok).toBe(true);
    expectExpired(v, token);
  });

  test('scenario "A rename of a sheet inside a span after the copy"', async () => {
    const v = await mount(SPAN);
    const token = copyAt(v, 0, 0);
    expect(v.renameSheet(2, 'Renamed').ok).toBe(true);
    expectExpired(v, token);
  });

  test('scenario "A delete of a sheet inside a span after the copy"', async () => {
    const v = await mount(SPAN);
    const token = copyAt(v, 0, 0);
    expect(v.deleteSheet(2).ok).toBe(true);
    expectExpired(v, token);
  });

  test('scenario "An undo of a sheet delete after the copy of a span"', async () => {
    const v = await mount(SPAN);
    expect(v.deleteSheet(2).ok).toBe(true);
    const token = copyAt(v, 0, 0);
    expect(v.undo().ok).toBe(true);
    expectExpired(v, token);
  });

  test('scenario "A redo of a sheet delete after the copy of a span"', async () => {
    const v = await mount(SPAN);
    expect(v.deleteSheet(2).ok).toBe(true);
    expect(v.undo().ok).toBe(true);
    expect(v.setActiveSheet(0)).toBe(true);
    const token = copyAt(v, 0, 0);
    expect(v.redo().ok).toBe(true);
    expectExpired(v, token);
  });

  test('scenario "A delete and an add of the named sheet after the copy"', async () => {
    const v = await mount({
      sheets: [
        { name: 'S', rows: row(1, [c('A1', 1, { f: 'Data!A1' })]) },
        { name: 'Data', rows: row(1, [c('A1', 1)]) },
      ],
    });
    const token = copyAt(v, 0, 0);
    expect(v.deleteSheet(1).ok).toBe(true);
    expect(v.addSheet().ok).toBe(true);
    expect(v.renameSheet(1, 'Data').ok).toBe(true);
    expectExpired(v, token);
  });

  test('scenario "An edit that changes nothing after the copy"', async () => {
    const v = await mount(ABSOLUTE);
    const token = copyAt(v, 4, 1);
    v.select({ row: 0, col: 5 });
    expect(v.deleteColumns().ok).toBe(true);
    expect(pasteAt(v, token, 19, 3)).toBeNull();
    expect(content(v, 19, 3)).toBe('=$A$5');
  });

  test('scenario "A sheet range that reads two ways"', async () => {
    const v = await mount({
      sheets: [{ name: 'S', rows: row(1, [c('C1', 1, { f: 'SUM(A1:S!B2)' })]) }],
    });
    const token = copyAt(v, 0, 2);
    expect(pasteAt(v, token, 3, 3)).toBe('paste-unreadable-formula');
    expect(content(v, 3, 3)).toBe('');
  });

  test('scenario "A same-sheet cut of a sheet range that reads two ways"', async () => {
    const v = await mount({
      sheets: [{ name: 'S', rows: row(1, [c('C1', 1, { f: 'SUM(A1:S!B2)' })]) }],
    });
    v.select({ row: 0, col: 2 });
    const cut = v.cut();
    if (!cut.ok) throw new Error(cut.refusal.code);
    expect(pasteAt(v, cut.content.token, 3, 3)).toBe('paste-unreadable-formula');
    expect(content(v, 0, 2)).toBe('=SUM(A1:S!B2)');
  });

  test('scenario "A copied formula that fails the check of typed input"', async () => {
    const v = await mount({
      sheets: [{ name: 'S', rows: row(1, ['<c r="A1"><f>1+</f><v>0</v></c>']) }],
    });
    const token = copyAt(v, 0, 0);
    v.select({ row: 2, col: 0 });
    const pasted = v.paste({ token });
    expect(pasted.ok ? null : [pasted.refusal.code, pasted.refusal.subject]).toEqual([
      'paste-unreadable-formula',
      'S!A1',
    ]);
  });

  test('scenario "Undo a paste"', async () => {
    const v = await mount({
      sheets: [
        { name: 'S', rows: row(1, [c('A1', 1), c('C1', 8)]) + row(2, [c('A2', 2), c('C2', 9)]) },
      ],
    });
    v.select({ row: 0, col: 2 });
    expect(v.toggleFormat('bold').ok).toBe(true);
    v.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    const copied = v.copy();
    if (!copied.ok) throw new Error(copied.refusal.code);
    v.select({ row: 0, col: 2 });
    expect(v.paste({ token: copied.content.token }).ok).toBe(true);
    expect(content(v, 0, 2)).toBe('1');
    expect(v.undo().ok).toBe(true);
    expect(content(v, 0, 2)).toBe('8');
    expect(v.snapshot().activeFormat).toMatchObject({ bold: true });
    expect(content(v, 1, 2)).toBe('9');
  });

  test('scenario "A rename of a sheet that no copied formula names"', async () => {
    const v = await mount({
      sheets: [
        { name: 'S', rows: row(1, [c('A1', 1), c('B1', 2, { f: 'A1+1' })]) },
        { name: 'T', rows: row(1, [c('A1', 1)]) },
      ],
    });
    const token = copyAt(v, 0, 1);
    expect(v.renameSheet(1, 'Renamed').ok).toBe(true);
    expect(pasteAt(v, token, 1, 1)).toBeNull();
    expect(content(v, 1, 1)).toBe('=A2+1');
  });
});
