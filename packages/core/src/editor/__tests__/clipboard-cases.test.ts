// More clipboard cases of design D2 of the change `workbook-structure-editing`: a cut formula
// with a name that its sheet or the target sheet scopes, the spec example of a cut to another
// sheet, a formula that cannot be read on paste, a copy of a data table formula, and a copy
// whose sheets changed before the paste. Also approximation `cut-cell-attachments`: a cut and
// a copy leave hyperlinks, data validation, and conditional formatting with the source cells.

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
import {
  attr,
  child,
  children,
  readXlsx,
  sheetByName,
  workbookIndex,
  xmlPartOf,
} from '../../store/index.ts';
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
  const viewer = createXlsxEditor({ fonts: FONTS });
  viewers.push(viewer);
  viewer.attach(host);
  await viewer.load(buildXlsx(wb));
  flushEditor(viewer);
  return viewer;
}

function content(v: XlsxEditorInstance, r: number, col: number) {
  v.select({ row: r, col });
  return v.snapshot().activeCell?.content;
}

const code = (r: { ok: boolean; refusal?: { code: string } }) => (r.ok ? null : r.refusal?.code);

function cut(v: XlsxEditorInstance) {
  const r = v.cut();
  if (!r.ok) throw new Error(r.refusal.code);
  return r.content;
}

describe('a cut formula on another sheet', () => {
  test('a name that the source sheet scopes gets the source sheet prefix', async () => {
    const v = await mount({
      sheets: [
        { name: 'Src', rows: row(1, [c('A1', 1), c('B1', 1, { f: 'myName+A1+other' })]) },
        { name: 'T', rows: row(1, [c('A1', 2)]) },
      ],
      definedNames:
        '<definedName name="MYNAME" localSheetId="0">Src!$A$1</definedName>' +
        '<definedName name="myName" localSheetId="1">T!$A$1</definedName>' +
        '<definedName name="other">Src!$A$1</definedName>',
    });
    v.select({ row: 0, col: 1 });
    const data = cut(v);
    expect(v.setActiveSheet(1)).toBe(true);
    v.select({ row: 3, col: 3 });
    expect(v.paste({ token: data.token }).ok).toBe(true);
    // The name of the source sheet keeps its binding; a workbook name keeps its text.
    expect(content(v, 3, 3)).toBe('=Src!myName+Src!A1+other');
    const reopened = readXlsx(v.save()!);
    expect(reopened.ok).toBe(true);
  });
});

describe('a formula that cannot be read on paste', () => {
  test('a copy with an unreadable formula refuses the paste with paste-unreadable-formula', async () => {
    // An unclosed string: the copy keeps the text, and the move on paste cannot read it.
    const v = await mount({
      sheets: [{ name: 'S', rows: row(1, ['<c r="A1"><f>"open</f><v>0</v></c>']) }],
    });
    v.select({ row: 0, col: 0 });
    const r = v.copy();
    if (!r.ok) throw new Error(r.refusal.code);
    v.select({ row: 2, col: 0 });
    const pasted = v.paste({ token: r.content.token });
    expect(code(pasted)).toBe('paste-unreadable-formula');
    expect(pasted.ok ? null : pasted.refusal.subject).toBe('S!A1');
  });
});

describe('a copy of a data table formula', () => {
  test('is refused with copy-array-formula', async () => {
    const v = await mount({
      sheets: [
        {
          name: 'S',
          rows: row(1, [
            '<c r="A1"><f t="dataTable" ref="A1:A2" dt2D="0" dtr="0" r1="B1"/><v>2</v></c>',
            c('B1', 2),
          ]),
        },
      ],
    });
    v.select({ row: 0, col: 0 }, { row: 0, col: 1 });
    const r = v.copy();
    expect(code(r)).toBe('copy-array-formula');
    expect(r.ok ? null : r.refusal.subject).toBe('S!A1:A2');
  });

  test('a range without the master cell is refused too, with the whole range', async () => {
    const v = await mount({
      sheets: [
        {
          name: 'S',
          rows:
            row(1, ['<c r="B1"><f t="array" ref="B1:C3">A1:B3*2</f><v>2</v></c>']) +
            row(2, ['<c r="B2"><v>4</v></c>']),
        },
      ],
    });
    v.select({ row: 1, col: 1 }, { row: 3, col: 3 });
    const r = v.copy();
    expect(code(r)).toBe('copy-array-formula');
    expect(r.ok ? null : r.refusal.subject).toBe('S!B1:C3');
    v.select({ row: 4, col: 0 }, { row: 5, col: 1 });
    expect(v.copy().ok).toBe(true);
  });
});

describe('the spec example of a cut to another sheet', () => {
  test("S1!B1 with =A1 pastes on S2 as ='S1'!A1, and a name that S1 scopes gets 'S1'!", async () => {
    const v = await mount({
      sheets: [
        {
          name: 'S1',
          rows: row(1, [c('A1', 1), c('B1', 1, { f: 'A1' }), c('C1', 1, { f: 'nm' })]),
        },
        { name: 'S2', rows: row(1, [c('A1', 2)]) },
      ],
      definedNames: '<definedName name="nm" localSheetId="0">S1!$A$1</definedName>',
    });
    v.select({ row: 0, col: 1 }, { row: 0, col: 2 });
    const data = cut(v);
    expect(v.setActiveSheet(1)).toBe(true);
    v.select({ row: 2, col: 1 });
    expect(v.paste({ token: data.token }).ok).toBe(true);
    expect(content(v, 2, 1)).toBe("='S1'!A1");
    expect(content(v, 2, 2)).toBe("='S1'!nm");
  });
});

describe('a cut formula with a name that only the target sheet scopes', () => {
  test('refuses the paste with paste-name-scope, and nothing changes', async () => {
    const v = await mount({
      sheets: [
        { name: 'Src', rows: row(1, [c('A1', 1), c('B1', 1, { f: 'nm+1' })]) },
        { name: 'T', rows: row(1, [c('A1', 2)]) },
      ],
      definedNames:
        '<definedName name="nm">Src!$A$1</definedName>' +
        '<definedName name="nm" localSheetId="1">T!$A$1</definedName>',
    });
    v.select({ row: 0, col: 1 });
    const data = cut(v);
    expect(v.setActiveSheet(1)).toBe(true);
    v.select({ row: 3, col: 3 });
    const r = v.paste({ token: data.token });
    expect(code(r)).toBe('paste-name-scope');
    expect(r.ok ? null : r.refusal.subject).toBe('Src!B1');
    expect(content(v, 3, 3)).toBe('');
    expect(v.setActiveSheet(0)).toBe(true);
    expect(content(v, 0, 1)).toBe('=nm+1');
  });
});

describe('a copy whose sheets changed before the paste', () => {
  const book = () =>
    mount({
      sheets: [
        { name: 'S', rows: row(1, [c('A1', 1)]) },
        { name: 'T', rows: row(1, [c('A1', 1, { f: 'S!A1' }), c('B1', 5)]) },
      ],
    });

  function copyT(v: XlsxEditorInstance, col: number) {
    expect(v.setActiveSheet(1)).toBe(true);
    v.select({ row: 0, col });
    const r = v.copy();
    if (!r.ok) throw new Error(r.refusal.code);
    return r.content;
  }

  test('scenario "A sheet rename after the copy", and a delete, refuse the paste with copy-expired', async () => {
    for (const change of [
      (v: XlsxEditorInstance) => v.renameSheet(0, 'New'),
      (v: XlsxEditorInstance) => v.deleteSheet(0),
    ]) {
      const v = await book();
      const data = copyT(v, 0);
      expect(change(v).ok).toBe(true);
      expect(v.setActiveSheet(0)).toBe(true);
      v.select({ row: 2, col: 0 });
      const r = v.paste({ token: data.token, text: data.text });
      expect(code(r)).toBe('copy-expired');
      expect(r.ok ? null : r.refusal.subject).toBe('T!A1');
      expect(content(v, 2, 0)).toBe('');
      // The same token stays refused with the same code, and never takes the plain lane.
      expect(code(v.paste({ token: data.token, text: data.text }))).toBe('copy-expired');
      expect(content(v, 2, 0)).toBe('');
    }
  });

  test('a new sheet, or a copy of values only, keeps the copy', async () => {
    const v = await book();
    const formula = copyT(v, 0);
    expect(v.addSheet('Extra').ok).toBe(true);
    expect(v.setActiveSheet(1)).toBe(true);
    v.select({ row: 2, col: 0 });
    expect(v.paste({ token: formula.token }).ok).toBe(true);
    expect(content(v, 2, 0)).toBe('=S!A3');
    const value = copyT(v, 1);
    expect(v.renameSheet(0, 'New').ok).toBe(true);
    expect(v.setActiveSheet(1)).toBe(true);
    v.select({ row: 3, col: 1 });
    expect(v.paste({ token: value.token }).ok).toBe(true);
    expect(content(v, 3, 1)).toBe('5');
  });
});

describe('a copied formula on another sheet', () => {
  const scoped = () =>
    mount({
      sheets: [
        { name: 'S1', rows: row(1, [c('A1', 1, { f: 'Rate*2+B1' }), c('B1', 3)]) },
        { name: 'S2', rows: '' },
      ],
      definedNames:
        '<definedName name="Rate" localSheetId="0">S1!$B$1</definedName>' +
        '<definedName name="Rate">99</definedName>',
    });

  test('a name that the source sheet scopes gets the source sheet, and references move', async () => {
    const v = await scoped();
    v.select({ row: 0, col: 0 });
    const r = v.copy();
    if (!r.ok) throw new Error(r.refusal.code);
    expect(v.setActiveSheet(1)).toBe(true);
    v.select({ row: 2, col: 0 });
    expect(v.paste({ token: r.content.token }).ok).toBe(true);
    expect(content(v, 2, 0)).toBe("='S1'!Rate*2+B3");
    expect(readXlsx(v.save()!).ok).toBe(true);
  });

  test('a name that only the target sheet scopes refuses the paste with paste-name-scope', async () => {
    const v = await mount({
      sheets: [
        { name: 'S1', rows: row(1, [c('A1', 1, { f: 'nm+1' })]) },
        { name: 'S2', rows: '' },
      ],
      definedNames:
        '<definedName name="nm">S1!$A$1</definedName>' +
        '<definedName name="nm" localSheetId="1">S2!$A$1</definedName>',
    });
    v.select({ row: 0, col: 0 });
    const r = v.copy();
    if (!r.ok) throw new Error(r.refusal.code);
    expect(v.setActiveSheet(1)).toBe(true);
    v.select({ row: 2, col: 0 });
    const pasted = v.paste({ token: r.content.token });
    expect(code(pasted)).toBe('paste-name-scope');
    expect(pasted.ok ? null : pasted.refusal.subject).toBe('S1!A1');
    expect(content(v, 2, 0)).toBe('');
    // On the source sheet the name keeps its binding, so the paste is applied.
    expect(v.setActiveSheet(0)).toBe(true);
    v.select({ row: 2, col: 0 });
    expect(v.paste({ token: r.content.token }).ok).toBe(true);
    expect(content(v, 2, 0)).toBe('=nm+1');
  });

  test('a function name that the source sheet scopes refuses the paste on another sheet', async () => {
    const v = await mount({
      sheets: [
        { name: 'S1', rows: row(1, [c('A1', 1, { f: 'myFn(2)' })]) },
        { name: 'S2', rows: '' },
      ],
      definedNames: '<definedName name="myFn" localSheetId="0">LAMBDA(x,x*2)</definedName>',
    });
    v.select({ row: 0, col: 0 });
    const r = v.copy();
    if (!r.ok) throw new Error(r.refusal.code);
    expect(v.setActiveSheet(1)).toBe(true);
    v.select({ row: 2, col: 0 });
    // The store refuses `S1!myFn(2)` (a name and then a group: `missing-operator`), so the
    // call cannot keep its binding on another sheet: the reason says to paste on the source.
    const pasted = v.paste({ token: r.content.token });
    expect(code(pasted)).toBe('paste-name-scope');
    expect(pasted.ok ? null : pasted.refusal.subject).toBe('S1!A1');
    expect(content(v, 2, 0)).toBe('');
    // On the source sheet the call is not a predefined function, so the store refuses it as
    // typed input, and the paste names the function.
    expect(v.setActiveSheet(0)).toBe(true);
    v.select({ row: 2, col: 0 });
    const same = v.paste({ token: r.content.token });
    expect(same.ok ? null : [same.refusal.code, same.refusal.subject]).toEqual([
      'unsupported-function',
      'MYFN',
    ]);
    expect(content(v, 2, 0)).toBe('');
  });

  test('a copy with a scoped name depends on its source sheet for a paste on another sheet', async () => {
    const v = await scoped();
    v.select({ row: 0, col: 0 });
    const copied = v.copy();
    if (!copied.ok) throw new Error(copied.refusal.code);
    expect(v.renameSheet(0, 'Renamed').ok).toBe(true);
    expect(v.setActiveSheet(1)).toBe(true);
    v.select({ row: 2, col: 0 });
    expect(code(v.paste({ token: copied.content.token }))).toBe('copy-expired');
  });

  test('a function name that only the target sheet scopes refuses the paste', async () => {
    const v = await mount({
      sheets: [
        { name: 'S1', rows: row(1, [c('A1', 1, { f: 'myFn(2)' })]) },
        { name: 'S2', rows: '' },
      ],
      definedNames:
        '<definedName name="myFn">LAMBDA(x,x*2)</definedName>' +
        '<definedName name="myFn" localSheetId="1">LAMBDA(x,x*3)</definedName>',
    });
    v.select({ row: 0, col: 0 });
    const r = v.copy();
    if (!r.ok) throw new Error(r.refusal.code);
    expect(v.setActiveSheet(1)).toBe(true);
    v.select({ row: 2, col: 0 });
    expect(code(v.paste({ token: r.content.token }))).toBe('paste-name-scope');
  });

  test('a delete and a new sheet with the scoping name refuses the paste with copy-expired', async () => {
    const v = await scoped();
    v.select({ row: 0, col: 0 });
    const r = v.copy();
    if (!r.ok) throw new Error(r.refusal.code);
    expect(v.deleteSheet(0).ok).toBe(true);
    expect(v.addSheet('S1').ok).toBe(true);
    v.select({ row: 4, col: 0 });
    expect(code(v.paste({ token: r.content.token }))).toBe('copy-expired');
  });
});

describe('a copy whose named sheet was deleted and added again', () => {
  const book = () =>
    mount({
      sheets: [
        { name: 'S', rows: row(1, [c('A1', 0, { f: 'Data!A1' })]) },
        { name: 'Data', rows: row(1, [c('A1', 5)]) },
        { name: 'Z', rows: '' },
      ],
    });

  test('refuses the paste with copy-expired', async () => {
    const v = await book();
    v.select({ row: 0, col: 0 });
    const r = v.copy();
    if (!r.ok) throw new Error(r.refusal.code);
    expect(v.deleteSheet(1).ok).toBe(true);
    expect(v.addSheet('Data').ok).toBe(true);
    expect(v.setActiveSheet(0)).toBe(true);
    v.select({ row: 3, col: 0 });
    expect(code(v.paste({ token: r.content.token }))).toBe('copy-expired');
    expect(content(v, 3, 0)).toBe('');
  });

  test('an undo of the delete restores the sheet, and the copy stays expired', async () => {
    const v = await book();
    v.select({ row: 0, col: 0 });
    const r = v.copy();
    if (!r.ok) throw new Error(r.refusal.code);
    expect(v.deleteSheet(1).ok).toBe(true);
    expect(v.undo().ok).toBe(true);
    expect(v.setActiveSheet(0)).toBe(true);
    v.select({ row: 3, col: 0 });
    const pasted = v.paste({ token: r.content.token });
    expect(pasted.ok ? null : pasted.refusal.code).toBe('copy-expired');
    expect(content(v, 3, 0)).toBe('');
  });

  test('a move of the named sheet ends the copy', async () => {
    const v = await book();
    v.select({ row: 0, col: 0 });
    const r = v.copy();
    if (!r.ok) throw new Error(r.refusal.code);
    expect(v.moveSheet(1, 2).ok).toBe(true);
    expect(v.setActiveSheet(0)).toBe(true);
    v.select({ row: 3, col: 0 });
    const pasted = v.paste({ token: r.content.token });
    expect(pasted.ok ? null : pasted.refusal.code).toBe('copy-expired');
  });
});

describe('a copied formula with an external sheet range', () => {
  test('a rename of a local sheet with the same name keeps the copy and the text', async () => {
    const v = await mount({
      sheets: [
        { name: 'Jan', rows: row(1, [c('A1', 0, { f: 'SUM([1]Jan:Mar!A1)' })]) },
        { name: 'Mar', rows: '' },
      ],
      // `[1]` is the first external reference: the store reads a pasted formula against them.
      externalReferences: '<externalReference r:id="rIdX1"/>',
      extraParts: [
        [
          '/xl/externalLinks/externalLink1.xml',
          'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml',
          new TextEncoder().encode(
            '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>'
          ),
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink',
        ],
      ],
    });
    v.select({ row: 0, col: 0 });
    const r = v.copy();
    if (!r.ok) throw new Error(r.refusal.code);
    expect(v.renameSheet(0, 'Q1').ok).toBe(true);
    v.select({ row: 2, col: 0 });
    expect(v.paste({ token: r.content.token }).ok).toBe(true);
    expect(content(v, 2, 0)).toBe('=SUM([1]Jan:Mar!A3)');
    expect(content(v, 0, 0)).toBe('=SUM([1]Jan:Mar!A1)');
  });
});

describe('a pasted formula that the store refuses as typed input', () => {
  test('a copy of 1+ refuses with paste-unreadable-formula and the source cell', async () => {
    const v = await mount({
      sheets: [{ name: 'S', rows: row(1, ['<c r="A1"><f>1+</f><v>0</v></c>']) }],
    });
    v.select({ row: 0, col: 0 });
    const r = v.copy();
    if (!r.ok) throw new Error(r.refusal.code);
    v.select({ row: 2, col: 0 });
    const pasted = v.paste({ token: r.content.token });
    expect(pasted.ok ? null : [pasted.refusal.code, pasted.refusal.subject]).toEqual([
      'paste-unreadable-formula',
      'S!A1',
    ]);
    expect(content(v, 2, 0)).toBe('');
  });

  test('a same-sheet cut of an open quote refuses with paste-unreadable-formula', async () => {
    const v = await mount({
      sheets: [{ name: 'S', rows: row(1, ['<c r="A1"><f>"open</f><v>0</v></c>']) }],
    });
    v.select({ row: 0, col: 0 });
    const r = v.cut();
    if (!r.ok) throw new Error(r.refusal.code);
    v.select({ row: 0, col: 2 });
    const pasted = v.paste({ token: r.content.token });
    expect(pasted.ok ? null : [pasted.refusal.code, pasted.refusal.subject]).toEqual([
      'paste-unreadable-formula',
      'S!A1',
    ]);
  });

  test('a same-sheet cut of an unquoted sheet range refuses, as typed input does', async () => {
    // `A1:S!B2` reads as the sheet range `A1:S`, and the check of typed input refuses it, so
    // the store cannot write the text at the target either.
    const v = await mount({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('C1', 2, { f: 'SUM(A1:S!B2)' })]) }],
    });
    v.select({ row: 0, col: 2 });
    const r = v.cut();
    if (!r.ok) throw new Error(r.refusal.code);
    v.select({ row: 3, col: 3 });
    const pasted = v.paste({ token: r.content.token });
    expect(pasted.ok ? null : [pasted.refusal.code, pasted.refusal.subject]).toEqual([
      'paste-unreadable-formula',
      'S!C1',
    ]);
    expect(content(v, 0, 2)).toBe('=SUM(A1:S!B2)');
  });
});

describe('approximation cut-cell-attachments', () => {
  test('a cut leaves the hyperlink, data validation, and conditional formatting of its cells', async () => {
    const v = await mount({
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 'link'), c('B1', 2)]),
          afterMerges:
            '<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1" dxfId="0"><formula>TRUE</formula></cfRule></conditionalFormatting>' +
            '<dataValidations count="1"><dataValidation type="whole" sqref="A1"><formula1>1</formula1></dataValidation></dataValidations>' +
            '<hyperlinks><hyperlink ref="A1" location="S!B1" display="link"/></hyperlinks>',
        },
      ],
    });
    v.select({ row: 0, col: 0 });
    const data = cut(v);
    v.select({ row: 0, col: 2 });
    expect(v.paste({ token: data.token }).ok).toBe(true);
    expect(content(v, 0, 2)).toBe('link');
    expect(content(v, 0, 0)).toBe('');
    const r = readXlsx(v.save()!);
    if (!r.ok) throw new Error('reopen');
    const root = xmlPartOf(r.pkg, sheetByName(workbookIndex(r.pkg)!, 'S')!.partName)!.root;
    const refs = (list: string, item: string, name: string) =>
      children(child(root, list), item).map((el) => attr(el, name));
    expect(refs('x:hyperlinks', 'x:hyperlink', 'ref')).toEqual(['A1']);
    expect(refs('x:dataValidations', 'x:dataValidation', 'sqref')).toEqual(['A1']);
    expect(children(root, 'x:conditionalFormatting').map((el) => attr(el, 'sqref'))).toEqual([
      'A1',
    ]);
  });
});

describe('a copied cell format index past the cell formats', () => {
  test('copies format 0, the format that the cell shows, so the paste is applied', async () => {
    // A1 states `s="9"`, row 2 states `s="9"` with `customFormat`, and column C states
    // `style="9"`: the workbook has fewer cell formats, so each cell shows format 0.
    const v = await mount({
      sheets: [
        {
          name: 'S',
          cols: '<cols><col min="3" max="3" width="9" style="9"/></cols>',
          rows: row(1, [c('A1', 1, { s: 9 })]) + row(2, [], 's="9" customFormat="1"'),
        },
      ],
    });
    v.select({ row: 0, col: 0 }, { row: 1, col: 2 });
    const r = v.copy();
    if (!r.ok) throw new Error(r.refusal.code);
    v.select({ row: 4, col: 0 });
    expect(code(v.paste({ token: r.content.token }))).toBeNull();
    expect(content(v, 4, 0)).toBe('1');
    const saved = readXlsx(v.save()!);
    if (!saved.ok) throw new Error('reopen');
    const root = xmlPartOf(saved.pkg, sheetByName(workbookIndex(saved.pkg)!, 'S')!.partName)!.root;
    const pasted = children(child(root, 'x:sheetData'), 'x:row')
      .filter((el) => Number(attr(el, 'r')) >= 5)
      .flatMap((el) => children(el, 'x:c'));
    for (const cell of pasted) expect(attr(cell, 's') ?? '0').toBe('0');
  });
});

describe('the texts of the clipboard refusals', () => {
  const en = JSON.parse(readFileSync(`${import.meta.dir}/../../../../i18n/en.json`, 'utf8'));
  const refused = (code: string) => String(en.editing.refused[code]);

  test('each gives advice, and the quote advice holds only for a range of sheets', () => {
    expect(refused('paste-unreadable-formula')).toContain(
      'Type the formula in the target cell. If it holds a range of sheets, quote the sheet names.'
    );
    expect(refused('copy-unreadable-formula')).toContain(
      'Fix the formula, or copy a range without it.'
    );
    expect(refused('range-too-large')).toContain('Select or paste a smaller range');
  });
});
