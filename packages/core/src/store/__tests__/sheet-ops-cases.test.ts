// More cases of the sheet ops of design D1 of the change `workbook-structure-editing`: the
// whole-part blocker of a rename and a delete, `x:workbookView/@firstSheet`, sheet names with
// control characters, formula text that the tokenizer cannot read, and legacy drawing bytes in
// another encoding.

import { describe, expect, test } from 'bun:test';
import { c, row, type FixtureWorkbook } from '../../../../../scripts/fixtures/builder.ts';
import { attr, child } from '../package/query.ts';
import { parseXmlPart } from '../package/tree.ts';
import { xmlPartOf, type XlsxPackage } from '../package/package.ts';
import { serializeXmlPart } from '../package-write/serialize.ts';
import { workbookIndex } from '../workbook/workbook-index.ts';
import { rewriteSheetPrefixes } from '../structure/sheet-refs.ts';
import { WorkbookDocumentStore } from '../store/document-store.ts';
import type { WorkbookOp } from '../store/ops.ts';
import { open, storeOf } from './structure-helpers.ts';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const rejection = (s: WorkbookDocumentStore, ops: readonly WorkbookOp[]) => {
  const before = s.package;
  const r = s.transact(ops);
  expect(s.package).toBe(before);
  return r.ok ? null : { code: r.rejection.code, subject: r.rejection.subject };
};

/** The refusals of a rename and a delete of `Sheet1`. */
const both = (s: WorkbookDocumentStore) => [
  rejection(s, [{ op: 'renameSheet', sheet: 'Sheet1', name: 'Renamed' }]),
  rejection(s, [{ op: 'deleteSheet', sheet: 'Sheet1' }]),
];

const sheets = (): FixtureWorkbook['sheets'] => [
  { name: 'Sheet1', rows: row(1, [c('A1', 1)]) },
  { name: 'Sheet2', rows: row(1, [c('A1', 2)]) },
  { name: 'Sheet3', rows: row(1, [c('A1', 3)]) },
];

const enc = (text: string) => new TextEncoder().encode(text);

/** The package with its workbook part text changed by `edit`. */
function patchWorkbook(pkg: XlsxPackage, edit: (xml: string) => string): XlsxPackage {
  const main = workbookIndex(pkg)!.mainPart;
  const text = edit(serializeXmlPart(xmlPartOf(pkg, main)!));
  const parsed = parseXmlPart(main, text);
  if (!parsed.ok) throw new Error(parsed.reason);
  const parts = new Map(pkg.parts);
  parts.set(main, { kind: 'xml', partName: main, xml: parsed.part });
  return { parts };
}

describe('the whole-part blocker of a rename and a delete', () => {
  test('a chart that does not name the sheet refuses both, with the chart part as subject', () => {
    const s = storeOf({
      sheets: sheets(),
      extraParts: [
        [
          'xl/charts/chart1.xml',
          'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
          enc('<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>'),
        ],
      ],
    });
    const subject = '/xl/charts/chart1.xml';
    expect(both(s)).toEqual([
      { code: 'structure-blocked', subject },
      { code: 'structure-blocked', subject },
    ]);
  });

  test('a chartsheet refuses both', () => {
    const base = open({
      sheets: sheets(),
      extraParts: [
        [
          'xl/chartsheets/sheet9.xml',
          'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml',
          enc(
            '<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"/></sheetViews></chartsheet>'
          ),
          `${REL}/chartsheet`,
        ],
      ],
    });
    const pkg = patchWorkbook(base, (xml) =>
      xml.replace('</sheets>', '<sheet name="Chart1" sheetId="9" r:id="rIdX1"/></sheets>')
    );
    expect(workbookIndex(pkg)!.sheets.map((x) => x.name)).toContain('Chart1');
    const s = new WorkbookDocumentStore(pkg);
    // The subject is the chartsheet's name, as its tab shows it, not its part name.
    const subject = 'Chart1';
    expect(both(s)).toEqual([
      { code: 'structure-blocked', subject },
      { code: 'structure-blocked', subject },
    ]);
    // A delete of the chartsheet itself names it by its sheet name too.
    expect(rejection(s, [{ op: 'deleteSheet', sheet: 'Chart1' }])).toEqual({
      code: 'structure-blocked',
      subject,
    });
  });

  test('a workbook relationship outside the allowlist refuses both', () => {
    const s = storeOf({
      sheets: sheets(),
      extraParts: [
        [
          'xl/connections.xml',
          'application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml',
          enc('<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>'),
          `${REL}/connections`,
        ],
      ],
    });
    const subject = '/xl/connections.xml';
    expect(both(s)).toEqual([
      { code: 'structure-blocked', subject },
      { code: 'structure-blocked', subject },
    ]);
    // A move does not rewrite references, so the blocker does not refuse it.
    expect(s.transact([{ op: 'moveSheet', sheet: 'Sheet1', index: 2 }]).ok).toBe(true);
  });
});

describe('x:workbookView/@firstSheet', () => {
  const store = (active: number, first: number) =>
    new WorkbookDocumentStore(
      patchWorkbook(open({ sheets: sheets() }), (xml) =>
        xml.replace('activeTab="0"', `activeTab="${active}" firstSheet="${first}"`)
      )
    );
  const firstSheet = (s: WorkbookDocumentStore) =>
    attr(
      child(child(xmlPartOf(s.package, '/xl/workbook.xml')!.root, 'x:bookViews'), 'x:workbookView'),
      'firstSheet'
    );

  test('keeps its index on a move, at most the active tab', () => {
    const s = store(2, 1);
    expect(firstSheet(s)).toBe('1');
    // The first tab is a scroll position: it does not follow the sheet at its index.
    expect(s.transact([{ op: 'moveSheet', sheet: 'Sheet2', index: 2 }]).ok).toBe(true);
    expect(firstSheet(s)).toBe('1');
    // The active sheet moves to the front, so the first tab goes to it.
    expect(s.transact([{ op: 'moveSheet', sheet: 'Sheet3', index: 0 }]).ok).toBe(true);
    expect(firstSheet(s)).toBe('0');
  });

  test('keeps its index on a delete, at most the last position and the active tab', () => {
    const s = store(2, 2);
    expect(s.transact([{ op: 'deleteSheet', sheet: 'Sheet1' }]).ok).toBe(true);
    expect(firstSheet(s)).toBe('1');
    const t = store(2, 1);
    expect(t.transact([{ op: 'deleteSheet', sheet: 'Sheet2' }]).ok).toBe(true);
    expect(firstSheet(t)).toBe('1');
    const u = store(2, 2);
    expect(u.transact([{ op: 'deleteSheet', sheet: 'Sheet3' }]).ok).toBe(true);
    expect(firstSheet(u)).toBe('1');
  });

  test('keeps its index on an add', () => {
    const s = store(2, 1);
    expect(s.transact([{ op: 'addSheet', name: 'New', index: 0 }]).ok).toBe(true);
    expect(firstSheet(s)).toBe('1');
  });
});

describe('sheet names with control characters', () => {
  test('a tab, a line feed, a carriage return, and U+0001 are refused', () => {
    const s = storeOf({ sheets: sheets() });
    for (const name of ['A\tB', 'A\nB', 'A\rB', 'A\u0001B', '\u001f'])
      expect(rejection(s, [{ op: 'renameSheet', sheet: 'Sheet1', name }])?.code).toBe(
        'sheet-name-invalid'
      );
    expect(rejection(s, [{ op: 'addSheet', name: 'X\tY', index: 0 }])?.code).toBe(
      'sheet-name-invalid'
    );
  });
});

describe('formula text that the tokenizer cannot read', () => {
  test('a quoted name with a doubled apostrophe counts as a mention', () => {
    // `"` opens a string that never closes, so the text cannot be read.
    const text = `'It''s'!A1+"`;
    const r = rewriteSheetPrefixes(text, { kind: 'delete', sheet: "IT'S" });
    expect(r.ok).toBe(false);
    // Text that cannot name the sheet keeps its text.
    expect(rewriteSheetPrefixes('Other!A1+"', { kind: 'delete', sheet: "IT'S" })).toEqual({
      ok: true,
      text: 'Other!A1+"',
      changed: false,
    });
  });

  test('a move refuses text with a colon and a bang with sheet-range-blocked', () => {
    const s = storeOf({
      sheets: [
        { name: 'A', rows: row(1, [c('A1', 1)]) },
        { name: 'B', rows: row(1, ['<c r="A1"><f>SUM(A:B!A1)+"</f><v>0</v></c>']) },
      ],
    });
    expect(rejection(s, [{ op: 'moveSheet', sheet: 'A', index: 1 }])).toEqual({
      code: 'sheet-range-blocked',
      subject: 'B!A1',
    });
  });
});

describe('legacy drawing bytes', () => {
  const vml = (body: string, declaration = '') =>
    `${declaration}<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shape><x:ClientData ObjectType="Checkbox">${body}</x:ClientData></v:shape></xml>`;
  const storeWith = (name: string, bytes: Uint8Array) =>
    storeOf({
      sheets: [
        { name, rows: row(1, [c('A1', 1)]) },
        {
          name: 'Sheet2',
          rows: '',
          rels: [['rIdX', `${REL}/vmlDrawing`, '../drawings/vmlDrawing1.vml']],
        },
      ],
      extraParts: [
        [
          '/xl/drawings/vmlDrawing1.vml',
          'application/vnd.openxmlformats-officedocument.vmlDrawing',
          bytes,
        ],
      ],
    });
  const subject = '/xl/drawings/vmlDrawing1.vml';

  test('another declared encoding refuses an edit of a sheet with a non-ASCII name', () => {
    const declared = '<?xml version="1.0" encoding="windows-1252"?>';
    const s = storeWith('Données', enc(vml('<x:Row>0</x:Row>', declared)));
    expect(rejection(s, [{ op: 'renameSheet', sheet: 'Données', name: 'Data' }])).toEqual({
      code: 'structure-blocked',
      subject,
    });
    // An ASCII name is found in any ASCII-compatible encoding, so nothing refuses.
    const ascii = storeWith('Data', enc(vml('<x:Row>0</x:Row>', declared)));
    expect(ascii.transact([{ op: 'renameSheet', sheet: 'Data', name: 'Other' }]).ok).toBe(true);
  });

  test('bytes that are not valid UTF-8 refuse an edit of a sheet with a non-ASCII name', () => {
    const text = enc(vml('<x:Row>0</x:Row>'));
    const bytes = new Uint8Array(text.length + 1);
    bytes.set(text);
    bytes[text.length] = 0xff;
    const s = storeWith('Données', bytes);
    expect(rejection(s, [{ op: 'deleteSheet', sheet: 'Données' }])).toEqual({
      code: 'structure-blocked',
      subject,
    });
    // Valid UTF-8 that does not name the sheet refuses nothing.
    const clean = storeWith('Données', enc(vml('<x:Row>0</x:Row>')));
    expect(clean.transact([{ op: 'renameSheet', sheet: 'Données', name: 'Data' }]).ok).toBe(true);
  });
});

describe('a custom workbook view', () => {
  test('refuses only a delete of the sheet that it shows, by @activeSheetId', () => {
    const view =
      '<customWorkbookViews><customWorkbookView name="Mine" guid="{00000000-0000-0000-0000-000000000001}" windowWidth="800" windowHeight="600" activeSheetId="1"/></customWorkbookViews>';
    const s = storeOf({ sheets: sheets(), afterCalcPr: view });
    expect(rejection(s, [{ op: 'deleteSheet', sheet: 'Sheet1' }])).toEqual({
      code: 'structure-blocked',
      subject: 'x:customWorkbookView Mine',
    });
    // The view names the sheet by its sheetId, which a rename and a move keep.
    expect(s.transact([{ op: 'renameSheet', sheet: 'Sheet1', name: 'Renamed' }]).ok).toBe(true);
    expect(s.transact([{ op: 'moveSheet', sheet: 'Renamed', index: 2 }]).ok).toBe(true);
  });
});
