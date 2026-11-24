// A macro-enabled workbook or template opens, edits, and saves like a plain one. Its main part
// has a content type that ECMA-376 Part 1, 12.3.23 does not list, with the same root element
// (`x:workbook`) and the same `officeDocument` relationship. The macro part stays as bytes that
// are never run, and a save keeps it and the main-part content type (Part 2, 7.2.3).

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { createXlsxEditor, refusalText, type XlsxEditorInstance } from '../index.ts';
import en from '../../../../i18n/en.json';
import { flushEditor } from './flush-editor.ts';
import { runToolbarCommand, toolbarCommandState } from '../toolbar-commands.ts';
import type { FontSource } from '../../layout/index.ts';
import { readXlsx } from '../../store/index.ts';

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

const SML = 'application/vnd.openxmlformats-officedocument.spreadsheetml.';
const MACRO_WORKBOOK = 'application/vnd.ms-excel.sheet.macroEnabled.main+xml';
const MACRO_TEMPLATE = 'application/vnd.ms-excel.template.macroEnabled.main+xml';
const ADDIN = 'application/vnd.ms-excel.addin.macroEnabled.main+xml';
const VBA_REL = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject';
const MACRO_SHEET_REL = 'http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet';
/** Bytes that stand for a macro part. Nothing reads or runs them. */
const MACRO_BYTES = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 1, 2, 3]);

/** A workbook whose main part has `mainType`, with a macro part when `macros` is set. */
function packageOf(mainType: string, macros = false, rel = VBA_REL): Uint8Array {
  const sheets = [
    { name: 'S', rows: row(1, [c('A1', 1)]) },
    { name: 'T', rows: row(1, [c('A1', 2)]) },
  ];
  const files = unzipSync(buildXlsx({ sheets }));
  let types = strFromU8(files['[Content_Types].xml']!).replace(`${SML}sheet.main+xml`, mainType);
  if (macros) {
    types = types.replace(
      '</Types>',
      '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>'
    );
    const rels = strFromU8(files['xl/_rels/workbook.xml.rels']!).replace(
      '</Relationships>',
      `<Relationship Id="rIdMacros" Type="${rel}" Target="vbaProject.bin"/></Relationships>`
    );
    files['xl/_rels/workbook.xml.rels'] = strToU8(rels);
    files['xl/vbaProject.bin'] = MACRO_BYTES;
  }
  files['[Content_Types].xml'] = strToU8(types);
  return zipSync(files);
}

/** The English catalog, as a translate function. */
function english(key: string, params?: Readonly<Record<string, string | number>>): string {
  let node: unknown = en;
  for (const part of key.split('.')) node = (node as Record<string, unknown>)[part];
  return String(node).replace(/\{(\w+)\}/g, (all, k: string) => String(params?.[k] ?? all));
}

const viewers: XlsxEditorInstance[] = [];
afterEach(() => {
  for (const v of viewers.splice(0)) v.destroy();
});

async function mount(bytes: Uint8Array) {
  const doc = new Window().document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.appendChild(host);
  const viewer = createXlsxEditor({ fonts: FONTS });
  viewers.push(viewer);
  viewer.attach(host);
  await viewer.load(bytes);
  flushEditor(viewer);
  return viewer;
}

describe('macro-enabled packages', () => {
  test('a macro-enabled workbook with no macro part opens', () => {
    const read = readXlsx(packageOf(MACRO_WORKBOOK));
    expect(read.ok ? 'ok' : read.error).toBe('ok');
  });

  test('the main-part content type matches without ASCII case and outer spaces', () => {
    const read = readXlsx(packageOf(`  ${MACRO_WORKBOOK.toUpperCase()} `));
    expect(read.ok ? 'ok' : read.error).toBe('ok');
  });

  test('an add-in package stays refused', () => {
    const read = readXlsx(packageOf(ADDIN));
    expect(read.ok ? null : read.error).toEqual({ reason: 'not-a-workbook', detail: ADDIN });
  });

  test('the snapshot names the kind and the macro-enabled form', async () => {
    const plain = await mount(packageOf(`${SML}sheet.main+xml`));
    expect(plain.snapshot().workbookKind).toBe('workbook');
    expect(plain.snapshot().macroEnabled).toBe(false);
    const book = await mount(packageOf(MACRO_WORKBOOK, true));
    expect(book.snapshot().status).toBe('ready');
    expect(book.snapshot().workbookKind).toBe('workbook');
    expect(book.snapshot().macroEnabled).toBe(true);
    const template = await mount(packageOf(MACRO_TEMPLATE));
    expect(template.snapshot().workbookKind).toBe('template');
    expect(template.snapshot().macroEnabled).toBe(true);
    // With no workbook open, both are cleared.
    template.close();
    expect(template.snapshot().workbookKind).toBeNull();
    expect(template.snapshot().macroEnabled).toBe(false);
  });

  test('an edit and a save keep the content type, the relationship, and the macro bytes', async () => {
    const viewer = await mount(packageOf(MACRO_WORKBOOK, true));
    expect(viewer.enterValue('5').ok).toBe(true);
    const saved = viewer.save()!;
    const files = unzipSync(saved);
    expect(files['xl/vbaProject.bin']).toEqual(MACRO_BYTES);
    const types = strFromU8(files['[Content_Types].xml']!);
    expect(types).toContain(`ContentType="${MACRO_WORKBOOK}"`);
    expect(types).not.toContain(`${SML}sheet.main+xml`);
    expect(strFromU8(files['xl/_rels/workbook.xml.rels']!)).toContain(VBA_REL);
    const reopened = await mount(saved);
    expect(reopened.snapshot().macroEnabled).toBe(true);
    expect(strFromU8(files['xl/worksheets/sheet1.xml']!)).toContain('<v>5</v>');
  });

  test('a macro-enabled template with a macro part keeps it through an edit and a save', async () => {
    const viewer = await mount(packageOf(MACRO_TEMPLATE, true));
    expect(viewer.snapshot().sheetEditRefusal?.about?.kind).toBe('macros');
    expect(viewer.enterValue('5').ok).toBe(true);
    const files = unzipSync(viewer.save()!);
    expect(files['xl/vbaProject.bin']).toEqual(MACRO_BYTES);
    expect(strFromU8(files['[Content_Types].xml']!)).toContain(`ContentType="${MACRO_TEMPLATE}"`);
    const reopened = await mount(zipSync(files));
    expect(reopened.snapshot().workbookKind).toBe('template');
    expect(reopened.snapshot().macroEnabled).toBe(true);
  });

  test('a rewritten content types part keeps the macro-enabled main type', async () => {
    // With no macro part, a sheet add applies: it adds a part and its content type, so the
    // content types part is rewritten. A template keeps its own type too.
    for (const type of [MACRO_WORKBOOK, MACRO_TEMPLATE]) {
      const viewer = await mount(packageOf(type));
      expect(viewer.addSheet().ok).toBe(true);
      const files = unzipSync(viewer.save()!);
      const types = strFromU8(files['[Content_Types].xml']!);
      expect(types).toContain(`ContentType="${type}"`);
      expect(types).not.toContain(`${SML}sheet.main+xml`);
      expect(types).not.toContain(`${SML}template.main+xml`);
      const reopened = await mount(viewer.save()!);
      expect(reopened.snapshot().sheets.length).toBe(3);
      expect(reopened.snapshot().macroEnabled).toBe(true);
    }
  });

  test('macros refuse an edit that moves or renames a reference, and say why', async () => {
    const viewer = await mount(packageOf(MACRO_WORKBOOK, true));
    // Macro code can name any sheet or cell, and the editor never rewrites it.
    expect(viewer.snapshot().sheetEditRefusal?.about?.kind).toBe('macros');
    const refused = [
      viewer.renameSheet(0, 'U'),
      viewer.deleteSheet(1),
      viewer.insertRows(),
      viewer.insertColumns(),
      viewer.deleteRows(),
      viewer.deleteColumns(),
      // Macro code can read a sheet by its position, so the sheet order stays too.
      viewer.addSheet(),
      viewer.moveSheet(0, 1),
    ];
    for (const r of refused) {
      expect(r.ok).toBe(false);
      if (r.ok) continue;
      expect(r.refusal.code).toBe('structure-blocked');
      expect(refusalText(r.refusal, english)).toContain('the macros of the workbook');
    }
    // The menu bar's Insert > Sheet has the same state and refusal.
    expect(toolbarCommandState(viewer.snapshot(), 'sheet.add').disabledReason?.key).toBe(
      'disabledReason.sheetOrderBlocked'
    );
    expect(runToolbarCommand(viewer, 'sheet.add').ok).toBe(false);
    expect(viewer.snapshot().sheets.length).toBe(2);
    // The row and column commands of the menu bar are disabled with the reason of the macros.
    for (const id of ['cells.insertRows', 'cells.insertColumns'] as const) {
      const reason = toolbarCommandState(viewer.snapshot(), id).disabledReason;
      expect(reason?.about?.kind).toBe('macros');
    }
    // A value edit moves no reference.
    expect(viewer.enterValue('7').ok).toBe(true);
    // A macro relationship with other casing names the macros too.
    const cased = await mount(packageOf(MACRO_WORKBOOK, true, VBA_REL.toUpperCase()));
    expect(cased.snapshot().sheetEditRefusal?.about?.kind).toBe('macros');
    // A plain workbook content type with a macro project still names the macros.
    const plainType = await mount(packageOf(`${SML}sheet.main+xml`, true));
    expect(plainType.snapshot().macroEnabled).toBe(false);
    expect(plainType.snapshot().sheetEditRefusal?.about?.kind).toBe('macros');
    expect(plainType.addSheet().ok).toBe(false);
    // Without a macro part, the same edits apply.
    const plain = await mount(packageOf(MACRO_WORKBOOK));
    expect(plain.snapshot().sheetEditRefusal).toBeNull();
    expect(plain.renameSheet(0, 'U').ok).toBe(true);
    expect(plain.insertRows().ok).toBe(true);
  });

  test('a macro sheet with no macro project refuses the same edits, and names the macros', async () => {
    const files = unzipSync(packageOf(MACRO_WORKBOOK));
    const add = (name: string, from: string, to: string) =>
      (files[name] = strToU8(strFromU8(files[name]!).replace(from, to)));
    add(
      'xl/workbook.xml',
      '</sheets>',
      '<sheet name="M" sheetId="9" r:id="rIdMacroSheet"/></sheets>'
    );
    add(
      'xl/_rels/workbook.xml.rels',
      '</Relationships>',
      `<Relationship Id="rIdMacroSheet" Type="${MACRO_SHEET_REL}" Target="macrosheets/sheet1.xml"/></Relationships>`
    );
    add(
      '[Content_Types].xml',
      '</Types>',
      '<Override PartName="/xl/macrosheets/sheet1.xml" ContentType="application/vnd.ms-excel.macrosheet+xml"/></Types>'
    );
    files['xl/macrosheets/sheet1.xml'] = strToU8(
      '<xm:macrosheet xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"/>'
    );
    const viewer = await mount(zipSync(files));
    expect(viewer.snapshot().sheetEditRefusal?.about?.kind).toBe('macros');
    const refused = [
      viewer.addSheet(),
      viewer.moveSheet(0, 1),
      viewer.renameSheet(0, 'U'),
      viewer.deleteSheet(1),
      viewer.insertRows(),
      viewer.deleteColumns(),
    ];
    for (const r of refused) {
      expect(r.ok).toBe(false);
      if (!r.ok) expect(r.refusal.about?.kind).toBe('macros');
    }
    // A value edit applies, and the save keeps the macro sheet part.
    expect(viewer.enterValue('7').ok).toBe(true);
    expect(unzipSync(viewer.save()!)['xl/macrosheets/sheet1.xml']).toBeDefined();
  });

  test('a macro sheet found by its relationship type alone refuses an add and a move', async () => {
    for (const rel of [
      MACRO_SHEET_REL,
      MACRO_SHEET_REL.replace('xlMacrosheet', 'xlIntlMacrosheet'),
    ])
      await expectMacroSheetBlocks(rel);
  });

  test('a sheet part with a macro sheet content type and a worksheet relationship names the macros', async () => {
    const files = unzipSync(packageOf(MACRO_WORKBOOK));
    const add = (name: string, from: string, to: string) =>
      (files[name] = strToU8(strFromU8(files[name]!).replace(from, to)));
    add('xl/workbook.xml', '</sheets>', '<sheet name="M" sheetId="9" r:id="rIdM"/></sheets>');
    add(
      'xl/_rels/workbook.xml.rels',
      '</Relationships>',
      '<Relationship Id="rIdM" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="macrosheets/sheet1.xml"/></Relationships>'
    );
    add(
      '[Content_Types].xml',
      '</Types>',
      '<Override PartName="/xl/macrosheets/sheet1.xml" ContentType="application/vnd.ms-excel.macrosheet+xml"/></Types>'
    );
    files['xl/macrosheets/sheet1.xml'] = strToU8(
      '<xm:macrosheet xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"/>'
    );
    const viewer = await mount(zipSync(files));
    expect(viewer.snapshot().sheetEditRefusal?.about?.kind).toBe('macros');
    const r = viewer.addSheet();
    expect(r.ok ? null : r.refusal.about?.kind).toBe('macros');
  });

  test('a dialog sheet with no macros refuses no add', async () => {
    const files = unzipSync(packageOf(MACRO_WORKBOOK));
    const add = (name: string, from: string, to: string) =>
      (files[name] = strToU8(strFromU8(files[name]!).replace(from, to)));
    add('xl/workbook.xml', '</sheets>', '<sheet name="D" sheetId="9" r:id="rIdDialog"/></sheets>');
    add(
      'xl/_rels/workbook.xml.rels',
      '</Relationships>',
      `<Relationship Id="rIdDialog" Type="http://schemas.microsoft.com/office/2006/relationships/xlDialogsheet" Target="dialogsheets/sheet1.xml"/></Relationships>`
    );
    files['xl/dialogsheets/sheet1.xml'] = strToU8(
      '<dialogsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>'
    );
    const viewer = await mount(zipSync(files));
    expect(viewer.snapshot().sheetEditRefusal?.about?.kind).not.toBe('macros');
    expect(viewer.addSheet().ok).toBe(true);
  });

  async function expectMacroSheetBlocks(rel: string) {
    const files = unzipSync(packageOf(MACRO_WORKBOOK));
    const add = (name: string, from: string, to: string) =>
      (files[name] = strToU8(strFromU8(files[name]!).replace(from, to)));
    add(
      'xl/workbook.xml',
      '</sheets>',
      '<sheet name="M" sheetId="9" r:id="rIdMacroSheet"/></sheets>'
    );
    add(
      'xl/_rels/workbook.xml.rels',
      '</Relationships>',
      `<Relationship Id="rIdMacroSheet" Type="${rel}" Target="macrosheets/sheet1.xml"/></Relationships>`
    );
    // The part has a generic content type, from the `xml` default.
    files['xl/macrosheets/sheet1.xml'] = strToU8(
      '<xm:macrosheet xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"/>'
    );
    const viewer = await mount(zipSync(files));
    expect(viewer.snapshot().sheetEditRefusal?.about?.kind).toBe('macros');
    for (const r of [viewer.addSheet(), viewer.moveSheet(0, 1)]) {
      expect(r.ok).toBe(false);
      if (!r.ok) expect(r.refusal.about?.kind).toBe('macros');
    }
  }

  test('the refusal names the macros when the workbook also holds a chart', async () => {
    const files = unzipSync(packageOf(MACRO_WORKBOOK, true));
    files['xl/charts/chart1.xml'] = strToU8(
      '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>'
    );
    files['[Content_Types].xml'] = strToU8(
      strFromU8(files['[Content_Types].xml']!).replace(
        '</Types>',
        '<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>'
      )
    );
    const viewer = await mount(zipSync(files));
    expect(viewer.snapshot().sheetEditRefusal?.about?.kind).toBe('macros');
  });
});
