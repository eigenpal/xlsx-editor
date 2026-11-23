// The sheet ops of design D1 of the change `workbook-structure-editing`: add, rename, delete,
// and move a sheet through the store. Formulas, defined names, and hyperlink locations follow
// the sheet; `@localSheetId` and the active tab follow their sheets; one undo step puts back
// every part as it was; each refusal names its code; and a saved workbook reopens.

import { describe, expect, test } from 'bun:test';
import { c, row, type FixtureWorkbook } from '../../../../../scripts/fixtures/builder.ts';
import { attr, child, children, textOf } from '../package/query.ts';
import { tokenizeFormula } from '../formula/tokenizer.ts';
import { contentTypeOf, readXlsx, relationshipsOf, xmlPartOf } from '../package/package.ts';
import type { XlsxPackage } from '../package/package.ts';
import { serializeXmlPart } from '../package-write/serialize.ts';
import { saveXlsx } from '../package-write/save.ts';
import { workbookIndex } from '../workbook/workbook-index.ts';
import type { WorkbookDocumentStore } from '../store/document-store.ts';
import type { WorkbookOp } from '../store/ops.ts';
import { definedNames, formulaOf, sheetRoot, storeOf } from './structure-helpers.ts';

const names = (pkg: XlsxPackage) => workbookIndex(pkg)!.sheets.map((s) => s.name);

const workbookRoot = (pkg: XlsxPackage) => xmlPartOf(pkg, workbookIndex(pkg)!.mainPart)!.root;

/** Each defined name as `name@localSheetId=text`. */
const scopedNames = (pkg: XlsxPackage) =>
  children(child(workbookRoot(pkg), 'x:definedNames'), 'x:definedName').map(
    (el) => `${attr(el, 'name')}@${attr(el, 'localSheetId') ?? '-'}=${textOf(el)}`
  );

const view = (pkg: XlsxPackage) => child(child(workbookRoot(pkg), 'x:bookViews'), 'x:workbookView');

/** Every part as its serialized text, or its bytes, by part name. */
function snapshot(pkg: XlsxPackage): Map<string, string> {
  const out = new Map<string, string>();
  for (const [name, part] of pkg.parts)
    out.set(name, part.kind === 'xml' ? serializeXmlPart(part.xml) : [...part.bytes].join(','));
  return out;
}

/** Commit the ops, undo, and redo: undo gives back every part as it was. */
function roundTrip(s: WorkbookDocumentStore, ops: readonly WorkbookOp[]): XlsxPackage {
  const before = snapshot(s.package);
  const r = s.transact(ops);
  if (!r.ok) throw new Error(JSON.stringify(r.rejection));
  const after = snapshot(s.package);
  expect(s.undo().ok).toBe(true);
  expect(snapshot(s.package)).toEqual(before);
  expect(s.redo().ok).toBe(true);
  expect(snapshot(s.package)).toEqual(after);
  return s.package;
}

/** Save and reopen. */
function reopen(pkg: XlsxPackage): XlsxPackage {
  const r = readXlsx(saveXlsx(pkg));
  if (!r.ok) throw new Error(JSON.stringify(r.error));
  return r.pkg;
}

const rejection = (s: WorkbookDocumentStore, ops: readonly WorkbookOp[]) => {
  const before = s.package;
  const r = s.transact(ops);
  expect(s.package).toBe(before);
  return r.ok ? null : { code: r.rejection.code, subject: r.rejection.subject };
};

const twoSheets = (extra: Partial<FixtureWorkbook> = {}, second = ''): FixtureWorkbook => ({
  sheets: [
    { name: 'Sheet1', rows: row(2, [c('B2', 5)]) },
    { name: 'Sheet2', rows: row(1, [c('A1', 5, { f: 'Sheet1!B2' })]), afterMerges: second },
  ],
  ...extra,
});

describe('addSheet', () => {
  test('adds a part, a relationship, an override, and a sheetId at the position', () => {
    const s = storeOf({
      sheets: [
        { name: 'A', rows: '' },
        { name: 'B', rows: '' },
      ],
    });
    const pkg = roundTrip(s, [{ op: 'addSheet', name: 'New', index: 1 }]);
    expect(names(pkg)).toEqual(['A', 'New', 'B']);
    const entry = workbookIndex(pkg)!.sheets[1]!;
    expect(entry.partName).toBe('/xl/worksheets/sheet3.xml');
    expect(entry.kind).toBe('worksheet');
    expect(entry.sheetId).toBe(3);
    expect(contentTypeOf(pkg, '/xl/worksheets/sheet3.xml')).toContain('worksheet+xml');
    expect(serializeXmlPart(xmlPartOf(pkg, entry.partName)!)).toContain('<sheetData/>');
    const ids = relationshipsOf(pkg, '/xl/workbook.xml').map((r) => r.id);
    expect(new Set(ids).size).toBe(ids.length);
    const saved = reopen(pkg);
    expect(names(saved)).toEqual(['A', 'New', 'B']);
    expect(workbookIndex(saved)!.sheets[1]!.kind).toBe('worksheet');
  });

  test('the active tab and sheet-scoped names follow their sheets', () => {
    const s = storeOf({
      sheets: [
        { name: 'A', rows: '' },
        { name: 'B', rows: '' },
      ],
      activeTab: 1,
      definedNames: '<definedName name="x" localSheetId="1">B!$A$1</definedName>',
    });
    expect(s.transact([{ op: 'addSheet', name: 'New', index: 0 }]).ok).toBe(true);
    expect(attr(view(s.package), 'activeTab')).toBe('2');
    expect(scopedNames(s.package)).toEqual(['x@2=B!$A$1']);
  });

  test('takes the smallest free part name', () => {
    const s = storeOf(twoSheets());
    expect(s.transact([{ op: 'deleteSheet', sheet: 'Sheet1' }]).ok).toBe(true);
    expect(s.transact([{ op: 'addSheet', name: 'Sheet3', index: 1 }]).ok).toBe(true);
    expect(workbookIndex(s.package)!.sheets[1]!.partName).toBe('/xl/worksheets/sheet1.xml');
    expect(workbookIndex(s.package)!.sheets[1]!.sheetId).toBe(3);
  });

  test('refuses a bad position', () => {
    const s = storeOf(twoSheets());
    expect(rejection(s, [{ op: 'addSheet', name: 'X', index: 3 }])?.code).toBe('bad-position');
  });
});

describe('sheet names', () => {
  const s = storeOf(twoSheets());
  test.each([
    ['', 'sheet-name-invalid'],
    ['A:B', 'sheet-name-invalid'],
    ['a[1]', 'sheet-name-invalid'],
    ['x'.repeat(32), 'sheet-name-invalid'],
    ["'lead", 'sheet-name-invalid'],
    ["trail'", 'sheet-name-invalid'],
    ['bad\u0001', 'sheet-name-invalid'],
    ['_x0041_', 'sheet-name-invalid'],
    ['a_x005F_b', 'sheet-name-invalid'],
    ['\u{1F600}'.repeat(16), 'sheet-name-invalid'],
    ['sheet2', 'sheet-name-taken'],
  ] as const)('%j is refused with %s', (name, code) => {
    expect(rejection(s, [{ op: 'renameSheet', sheet: 'Sheet1', name }])?.code).toBe(code);
    expect(rejection(s, [{ op: 'addSheet', name, index: 0 }])?.code).toBe(code);
  });

  test('a name of 31 UTF-16 code units, and a new case of the same name, are accepted', () => {
    const t = storeOf(twoSheets());
    expect(t.transact([{ op: 'renameSheet', sheet: 'Sheet1', name: 'x'.repeat(31) }]).ok).toBe(
      true
    );
    expect(t.transact([{ op: 'renameSheet', sheet: 'Sheet2', name: 'SHEET2' }]).ok).toBe(true);
    expect(names(t.package)).toEqual(['x'.repeat(31), 'SHEET2']);
    expect(t.transact([{ op: 'renameSheet', sheet: 'SHEET2', name: 'a_x41_b_xZZZZ_' }]).ok).toBe(
      true
    );
  });

  test('the same name changes nothing', () => {
    const t = storeOf(twoSheets());
    expect(t.transact([{ op: 'renameSheet', sheet: 'Sheet1', name: 'Sheet1' }]).ok).toBe(true);
    expect(t.revision).toBe(0);
  });
});

describe('renameSheet', () => {
  const link = '<hyperlinks><hyperlink ref="C1" location="Sheet1!A1"/></hyperlinks>';
  const wb = () =>
    twoSheets(
      {
        definedNames:
          '<definedName name="g">Sheet1!$A$1</definedName>' +
          '<definedName name="l" localSheetId="0">Sheet1!$B$2</definedName>',
      },
      link
    );

  test('rewrites formulas, defined names, and hyperlink locations', () => {
    const s = storeOf(wb());
    const pkg = roundTrip(s, [{ op: 'renameSheet', sheet: 'Sheet1', name: 'Data 1' }]);
    expect(names(pkg)).toEqual(['Data 1', 'Sheet2']);
    expect(formulaOf(pkg, 'Sheet2', 'A1')).toBe("'Data 1'!B2");
    expect(definedNames(pkg)).toEqual({ g: "'Data 1'!$A$1", l: "'Data 1'!$B$2" });
    const h = child(child(sheetRoot(pkg, 'Sheet2'), 'x:hyperlinks'), 'x:hyperlink');
    expect(attr(h, 'location')).toBe("'Data 1'!A1");
    const saved = reopen(pkg);
    expect(formulaOf(saved, 'Sheet2', 'A1')).toBe("'Data 1'!B2");
  });

  test('doubles an apostrophe, and writes a plain name without quotes', () => {
    const s = storeOf(wb());
    expect(s.transact([{ op: 'renameSheet', sheet: 'Sheet1', name: "It's" }]).ok).toBe(true);
    expect(formulaOf(s.package, 'Sheet2', 'A1')).toBe("'It''s'!B2");
    expect(s.transact([{ op: 'renameSheet', sheet: "it's", name: 'Data' }]).ok).toBe(true);
    expect(formulaOf(s.package, 'Sheet2', 'A1')).toBe('Data!B2');
    expect(s.transact([{ op: 'renameSheet', sheet: 'Data', name: 'A1' }]).ok).toBe(true);
    expect(formulaOf(s.package, 'Sheet2', 'A1')).toBe("'A1'!B2");
  });

  test('rewrites a shared formula master, a sheet range, and keeps strings', () => {
    const s = storeOf({
      sheets: [
        { name: 'Sheet1', rows: row(1, [c('A1', 1)]) },
        {
          name: 'Sheet2',
          rows: row(1, [
            '<c r="A1"><f t="shared" ref="A1:A2" si="0">Sheet1!A1</f><v>1</v></c>',
            c('B1', 2, { f: 'SUM(Sheet1:Sheet2!A1)&"Sheet1!A1"' }),
          ]),
        },
      ],
    });
    expect(s.transact([{ op: 'renameSheet', sheet: 'Sheet1', name: 'First' }]).ok).toBe(true);
    expect(formulaOf(s.package, 'Sheet2', 'A1')).toBe('First!A1');
    expect(formulaOf(s.package, 'Sheet2', 'B1')).toBe(`SUM('First:Sheet2'!A1)&"Sheet1!A1"`);
  });

  test('refuses when content that the store does not rewrite names the sheet', () => {
    const cf =
      '<conditionalFormatting sqref="A1"><cfRule type="expression" dxfId="0" priority="1">' +
      '<formula>Sheet1!A1&gt;1</formula></cfRule></conditionalFormatting>';
    const s = storeOf({ ...twoSheets({}, cf), styles: { dxfs: ['<dxf/>'] } });
    expect(rejection(s, [{ op: 'renameSheet', sheet: 'Sheet1', name: 'X' }])).toEqual({
      code: 'structure-blocked',
      subject: 'x:conditionalFormatting (/xl/worksheets/sheet2.xml)',
    });
  });

  test('refuses an unknown sheet', () => {
    const s = storeOf(twoSheets());
    expect(rejection(s, [{ op: 'renameSheet', sheet: 'Nope', name: 'X' }])?.code).toBe(
      'unknown-sheet'
    );
  });
});

describe('deleteSheet', () => {
  test('a reference to the deleted sheet becomes #REF! with its address and its range', () => {
    const s = storeOf({
      sheets: [
        { name: 'Gone', rows: row(1, [c('A1', 1)]) },
        {
          name: 'Kept',
          rows: row(1, [
            c('A1', 1, { f: 'SUM(Gone!A1:B2)+Gone!$A$1+Kept!A2' }),
            c('B1', 1, { f: "'Gone'!A1:A2" }),
          ]),
        },
      ],
    });
    expect(s.transact([{ op: 'deleteSheet', sheet: 'Gone' }]).ok).toBe(true);
    expect(formulaOf(s.package, 'Kept', 'A1')).toBe('SUM(#REF!)+#REF!+Kept!A2');
    expect(formulaOf(s.package, 'Kept', 'B1')).toBe('#REF!');
  });

  test('removes the parts, turns references into #REF!, and moves scoped names', () => {
    const s = storeOf({
      sheets: [
        { name: 'Sheet1', rows: row(2, [c('B2', 5)]) },
        {
          name: 'Sheet2',
          rows: row(1, [c('A1', 5, { f: 'Sheet1!B2+1' })]),
          afterMerges: '<hyperlinks><hyperlink ref="C1" location="Sheet1!A1"/></hyperlinks>',
        },
        { name: 'Sheet3', rows: '' },
      ],
      activeTab: 2,
      definedNames:
        '<definedName name="g">Sheet1!$A$1</definedName>' +
        '<definedName name="l0" localSheetId="0">Sheet2!$A$1</definedName>' +
        '<definedName name="l2" localSheetId="2">Sheet3!$A$1</definedName>',
    });
    const pkg = s.package;
    const after = (() => {
      expect(s.transact([{ op: 'deleteSheet', sheet: 'Sheet1' }]).ok).toBe(true);
      return s.package;
    })();
    expect(names(after)).toEqual(['Sheet2', 'Sheet3']);
    expect(after.parts.has('/xl/worksheets/sheet1.xml')).toBe(false);
    expect(contentTypeOf(after, '/xl/worksheets/sheet1.xml')).toBe('application/xml');
    expect(relationshipsOf(after, '/xl/workbook.xml').some((r) => r.id === 'rId1')).toBe(false);
    expect(formulaOf(after, 'Sheet2', 'A1')).toBe('#REF!+1');
    expect(scopedNames(after)).toEqual(['g@-=#REF!', 'l2@1=Sheet3!$A$1']);
    expect(attr(view(after), 'activeTab')).toBe('1');
    const h = child(child(sheetRoot(after, 'Sheet2'), 'x:hyperlinks'), 'x:hyperlink');
    expect(attr(h, 'location')).toBe('#REF!');
    expect(workbookIndex(after)!.fullCalcOnLoad).toBe(true);
    // The reference goes with its address, so a later row edit has nothing to move.
    expect(s.transact([{ op: 'insertRows', sheet: 'Sheet2', start: 1, count: 2 }]).ok).toBe(true);
    expect(formulaOf(s.package, 'Sheet2', 'A3')).toBe('#REF!+1');
    expect(s.undo().ok).toBe(true);
    // Undo puts back every part, with the stale flag that undo never clears.
    expect(s.undo().ok).toBe(true);
    expect(names(s.package)).toEqual(['Sheet1', 'Sheet2', 'Sheet3']);
    for (const [name, part] of pkg.parts)
      if (name !== '/xl/workbook.xml') expect(s.package.parts.get(name)).toBe(part);
    expect(formulaOf(s.package, 'Sheet2', 'A1')).toBe('Sheet1!B2+1');
    expect(s.redo().ok).toBe(true);
    expect(names(reopen(s.package))).toEqual(['Sheet2', 'Sheet3']);
  });

  test('undo of a delete with no formula gives back every part', () => {
    const s = storeOf({
      sheets: [
        {
          name: 'A',
          rows: row(1, [c('A1', 1)]),
          afterMerges: '<hyperlinks><hyperlink ref="A1" r:id="rIdH"/></hyperlinks>',
          rels: [['rIdH', `${REL}/hyperlink`, 'https://example.com/', true]],
        },
        { name: 'B', rows: '' },
      ],
    });
    const pkg = roundTrip(s, [{ op: 'deleteSheet', sheet: 'A' }]);
    expect(pkg.parts.has('/xl/worksheets/_rels/sheet1.xml.rels')).toBe(false);
  });

  test('refuses the last visible sheet', () => {
    const one = storeOf({ sheets: [{ name: 'A', rows: '' }] });
    expect(rejection(one, [{ op: 'deleteSheet', sheet: 'A' }])?.code).toBe('last-sheet');
    const hidden = storeOf({
      sheets: [
        { name: 'A', rows: '' },
        { name: 'B', rows: '', state: 'hidden' },
      ],
    });
    expect(rejection(hidden, [{ op: 'deleteSheet', sheet: 'A' }])?.code).toBe('last-sheet');
    expect(hidden.transact([{ op: 'deleteSheet', sheet: 'B' }]).ok).toBe(true);
  });

  test('an absent active tab on a deleted first sheet goes to the next visible sheet', () => {
    for (const bookViews of ['<workbookView/>', '']) {
      const s = storeOf({
        sheets: [
          { name: 'A', rows: '' },
          { name: 'B', rows: '', state: 'hidden' },
          { name: 'C', rows: '' },
        ],
        bookViews,
      });
      expect(s.transact([{ op: 'deleteSheet', sheet: 'A' }]).ok).toBe(true);
      expect(attr(view(s.package), 'activeTab')).toBe('1');
    }
  });

  test('refuses a sheet with a related part other than a hyperlink or printer settings', () => {
    const s = storeOf({
      sheets: [
        {
          name: 'A',
          rows: '',
          rels: [['rIdC', `${REL}/comments`, '../comments1.xml']],
        },
        { name: 'B', rows: '' },
      ],
    });
    expect(rejection(s, [{ op: 'deleteSheet', sheet: 'A' }])).toEqual({
      code: 'structure-blocked',
      subject: '/xl/comments1.xml',
    });
  });

  test('removes printer settings that only the deleted sheet uses, and keeps shared ones', () => {
    const bin = new Uint8Array([1, 2, 3]);
    const type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings';
    const printer = (target: string) =>
      [['rIdP', `${REL}/printerSettings`, `../printerSettings/${target}`]] as const;
    const s = storeOf({
      sheets: [
        { name: 'A', rows: '', rels: printer('p1.bin') },
        { name: 'B', rows: '', rels: printer('p2.bin') },
        { name: 'C', rows: '', rels: printer('p2.bin') },
      ],
      extraParts: [
        ['/xl/printerSettings/p1.bin', type, bin],
        ['/xl/printerSettings/p2.bin', type, bin],
      ],
    });
    const pkg = roundTrip(s, [{ op: 'deleteSheet', sheet: 'A' }]);
    expect(pkg.parts.has('/xl/printerSettings/p1.bin')).toBe(false);
    expect(contentTypeOf(pkg, '/xl/printerSettings/p1.bin')).toBeUndefined();
    expect(s.transact([{ op: 'deleteSheet', sheet: 'B' }]).ok).toBe(true);
    expect(s.package.parts.has('/xl/printerSettings/p2.bin')).toBe(true);
    expect(contentTypeOf(s.package, '/xl/printerSettings/p2.bin')).toBe(type);
    expect(names(reopen(s.package))).toEqual(['C']);
  });

  test('refuses a sheet that a custom workbook view shows', () => {
    const s = storeOf({
      sheets: [
        { name: 'A', rows: '' },
        { name: 'B', rows: '' },
      ],
      afterCalcPr:
        '<customWorkbookViews><customWorkbookView name="Mine" guid="{00000000-0000-0000-0000-000000000001}" windowWidth="800" windowHeight="600" activeSheetId="1"/></customWorkbookViews>',
    });
    expect(rejection(s, [{ op: 'deleteSheet', sheet: 'A' }])).toEqual({
      code: 'structure-blocked',
      subject: 'x:customWorkbookView Mine',
    });
    expect(s.transact([{ op: 'deleteSheet', sheet: 'B' }]).ok).toBe(true);
  });

  test('a range operand joined by a colon becomes one #REF!', () => {
    const s = storeOf({
      sheets: [
        { name: 'Gone', rows: '' },
        {
          name: 'Kept',
          rows: row(1, [
            c('A1', 0, { f: 'Gone!A1:Gone!B2' }),
            c('B1', 0, { f: 'SUM(Gone!A1:Gone!B2,Kept!A2)' }),
            c('C1', 0, { f: 'Gone!A1 Gone!B1' }),
          ]),
        },
      ],
    });
    expect(s.transact([{ op: 'deleteSheet', sheet: 'Gone' }]).ok).toBe(true);
    expect(formulaOf(s.package, 'Kept', 'A1')).toBe('#REF!');
    expect(formulaOf(s.package, 'Kept', 'B1')).toBe('SUM(#REF!,Kept!A2)');
    // An intersection of two references keeps both operands.
    expect(formulaOf(s.package, 'Kept', 'C1')).toBe('#REF! #REF!');
  });

  test('scenario "A range operand": Sheet2!A1 with =Sheet1!A1:Sheet1!B2 holds =#REF!', () => {
    const s = storeOf({
      sheets: [
        { name: 'Sheet1', rows: '' },
        { name: 'Sheet2', rows: row(1, [c('A1', 0, { f: 'Sheet1!A1:Sheet1!B2' })]) },
      ],
    });
    expect(s.transact([{ op: 'deleteSheet', sheet: 'Sheet1' }]).ok).toBe(true);
    expect(formulaOf(s.package, 'Sheet2', 'A1')).toBe('#REF!');
  });

  test('refuses a sheet range in a formula or a defined name', () => {
    const f = storeOf({
      sheets: [
        { name: 'A', rows: '' },
        { name: 'B', rows: row(1, [c('A1', 0, { f: 'SUM(A:C!A1)' })]) },
        { name: 'C', rows: '' },
      ],
    });
    expect(rejection(f, [{ op: 'deleteSheet', sheet: 'C' }])).toEqual({
      code: 'sheet-range-blocked',
      subject: 'B!A1',
    });
    const n = storeOf({
      sheets: [
        { name: 'A', rows: '' },
        { name: 'B', rows: '' },
      ],
      definedNames: '<definedName name="r">A:B!$A$1</definedName>',
    });
    expect(rejection(n, [{ op: 'deleteSheet', sheet: 'A' }])?.subject).toBe('x:definedName r');
  });
});

describe('sheet names in formula text', () => {
  test('a name with a digit, a dot, or a reference-like start is quoted and reads back', () => {
    const cases: [string, string][] = [
      ['A1.B', "'A1.B'!B2"],
      ['R1C1x', "'R1C1x'!B2"],
      ['TRUEX', "'TRUEX'!B2"],
      ['name1', "'name1'!B2"],
      ['a.b', "'a.b'!B2"],
      ['RC', "'RC'!B2"],
      ['_x', '_x!B2'],
      ['Data', 'Data!B2'],
    ];
    for (const [name, text] of cases) {
      const s = storeOf(twoSheets());
      expect(s.transact([{ op: 'renameSheet', sheet: 'Sheet1', name }]).ok).toBe(true);
      const f = formulaOf(s.package, 'Sheet2', 'A1')!;
      expect(f).toBe(text);
      const read = tokenizeFormula(f);
      if (!read.ok) throw new Error(read.problem);
      const ref = read.tokens[0]!;
      expect(ref.kind === 'reference' ? ref.sheet?.sheets : undefined).toEqual([name]);
    }
  });
});

describe('content that names the sheet', () => {
  const shape = (link: string) =>
    '<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
    '<xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
    `<xdr:sp macro="" textlink="${link}"><xdr:nvSpPr><xdr:cNvPr id="2" name="T"/><xdr:cNvSpPr/></xdr:nvSpPr>` +
    '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>';
  const both = (s: WorkbookDocumentStore) =>
    (['renameSheet', 'deleteSheet'] as const).map((op) =>
      rejection(
        s,
        op === 'renameSheet'
          ? [{ op, sheet: 'Sheet1', name: 'Renamed' }]
          : [{ op, sheet: 'Sheet1' }]
      )
    );
  const sheet1 = { name: 'Sheet1', rows: row(1, [c('A1', 1)]) };

  test('a shape text link in a drawing of another sheet refuses a rename and a delete', () => {
    const s = storeOf({
      sheets: [sheet1, { name: 'Sheet2', rows: '', drawing: { anchors: shape('Sheet1!$A$1') } }],
    });
    const subject = '/xl/drawings/drawing2.xml';
    expect(both(s)).toEqual([
      { code: 'structure-blocked', subject },
      { code: 'structure-blocked', subject },
    ]);
    // A text link to the sheet's own cells refuses nothing.
    const own = storeOf({
      sheets: [sheet1, { name: 'Sheet2', rows: '', drawing: { anchors: shape('$A$1') } }],
    });
    expect(own.transact([{ op: 'renameSheet', sheet: 'Sheet1', name: 'Renamed' }]).ok).toBe(true);
  });

  test('an internal hyperlink target of a drawing refuses a rename and a delete', () => {
    const rels =
      '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rIdL" Type="${REL}/hyperlink" Target="#Sheet1!A1"/></Relationships>`;
    const s = storeOf({
      sheets: [sheet1, { name: 'Sheet2', rows: '', drawing: { anchors: shape('') } }],
      extraParts: [
        [
          '/xl/drawings/_rels/drawing2.xml.rels',
          'application/vnd.openxmlformats-package.relationships+xml',
          new TextEncoder().encode(rels),
        ],
      ],
    });
    const subject = '/xl/drawings/_rels/drawing2.xml.rels';
    expect(both(s)).toEqual([
      { code: 'structure-blocked', subject },
      { code: 'structure-blocked', subject },
    ]);
  });

  test('a control property link and a legacy drawing link refuse a rename and a delete', () => {
    const ctrl = `<formControlPr xmlns="${NS_X14}" objectType="CheckBox" fmlaLink="Sheet1!$A$1"/>`;
    const vml =
      '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
      '<v:shape><x:ClientData ObjectType="Checkbox"><x:FmlaLink>Sheet1!$A$1</x:FmlaLink></x:ClientData></v:shape></xml>';
    const enc = (text: string) => new TextEncoder().encode(text);
    const withPart = (relType: string, part: string, type: string, text: string) =>
      storeOf({
        sheets: [
          sheet1,
          { name: 'Sheet2', rows: '', rels: [['rIdX', `${REL}/${relType}`, `..${part.slice(3)}`]] },
        ],
        extraParts: [[part, type, enc(text)]],
      });
    const control = withPart('ctrlProp', '/xl/ctrlProps/ctrlProp1.xml', 'application/xml', ctrl);
    const legacy = withPart(
      'vmlDrawing',
      '/xl/drawings/vmlDrawing1.vml',
      'application/vnd.openxmlformats-officedocument.vmlDrawing',
      vml
    );
    for (const [s, subject] of [
      [control, '/xl/ctrlProps/ctrlProp1.xml'],
      [legacy, '/xl/drawings/vmlDrawing1.vml'],
    ] as const)
      expect(both(s)).toEqual([
        { code: 'structure-blocked', subject },
        { code: 'structure-blocked', subject },
      ]);
  });

  test('a whole column of a one-letter sheet name is not a mention of that sheet', () => {
    const cf =
      '<conditionalFormatting sqref="A:A"><cfRule type="expression" dxfId="0" priority="1">' +
      '<formula>A1&gt;1</formula></cfRule></conditionalFormatting>';
    const s = storeOf({
      sheets: [
        { name: 'A', rows: '' },
        { name: 'B', rows: '', afterMerges: cf },
      ],
      styles: { dxfs: ['<dxf/>'] },
    });
    expect(s.transact([{ op: 'renameSheet', sheet: 'A', name: 'Renamed' }]).ok).toBe(true);
    // A sheet range that starts at the sheet still refuses.
    const range = storeOf({
      sheets: [
        { name: 'A', rows: '' },
        { name: 'B', rows: '', afterMerges: cf.replace('A1&gt;1', 'SUM(A:Z!A1)&gt;1') },
        { name: 'Z', rows: '' },
      ],
      styles: { dxfs: ['<dxf/>'] },
    });
    expect(rejection(range, [{ op: 'renameSheet', sheet: 'A', name: 'Renamed' }])?.code).toBe(
      'structure-blocked'
    );
  });
});

describe('moveSheet', () => {
  const wb = (): FixtureWorkbook => ({
    sheets: [
      { name: 'A', rows: '' },
      { name: 'B', rows: row(1, [c('A1', 0, { f: 'A!A1' })]) },
      { name: 'C', rows: '' },
    ],
    activeTab: 1,
    definedNames:
      '<definedName name="a" localSheetId="0">A!$A$1</definedName>' +
      '<definedName name="b" localSheetId="1">B!$A$1</definedName>',
  });

  test('moves the sheet, and scoped names and the active tab follow', () => {
    const s = storeOf(wb());
    const pkg = roundTrip(s, [{ op: 'moveSheet', sheet: 'B', index: 0 }]);
    expect(names(pkg)).toEqual(['B', 'A', 'C']);
    expect(scopedNames(pkg)).toEqual(['a@1=A!$A$1', 'b@0=B!$A$1']);
    expect(attr(view(pkg), 'activeTab')).toBe('0');
    expect(formulaOf(pkg, 'B', 'A1')).toBe('A!A1');
    expect(names(reopen(pkg))).toEqual(['B', 'A', 'C']);
    expect(s.transact([{ op: 'moveSheet', sheet: 'A', index: 2 }]).ok).toBe(true);
    expect(names(s.package)).toEqual(['B', 'C', 'A']);
    expect(scopedNames(s.package)).toEqual(['a@2=A!$A$1', 'b@0=B!$A$1']);
  });

  test('an absent active tab or view reads as the first sheet and moves with it', () => {
    const abc = (bookViews: string): FixtureWorkbook => ({
      sheets: [
        { name: 'A', rows: '' },
        { name: 'B', rows: '' },
        { name: 'C', rows: '' },
      ],
      bookViews,
    });
    for (const bookViews of ['<workbookView/>', '']) {
      const s = storeOf(abc(bookViews));
      const pkg = roundTrip(s, [{ op: 'moveSheet', sheet: 'A', index: 2 }]);
      expect(attr(view(pkg), 'activeTab')).toBe('2');
      // The first tab of the tab strip is a scroll position, not a sheet: none is written.
      expect(attr(view(pkg), 'firstSheet')).toBeUndefined();
    }
    // A move that leaves the first sheet in place writes no view.
    const s = storeOf(abc(''));
    expect(s.transact([{ op: 'moveSheet', sheet: 'C', index: 1 }]).ok).toBe(true);
    expect(child(workbookRoot(s.package), 'x:bookViews')).toBeUndefined();
    // An add at the front moves the active sheet, and writes no first tab.
    const t = storeOf(abc(''));
    expect(t.transact([{ op: 'addSheet', name: 'New', index: 0 }]).ok).toBe(true);
    expect(attr(view(t.package), 'activeTab')).toBe('1');
    expect(attr(view(t.package), 'firstSheet')).toBeUndefined();
  });

  test('a move to the front of the active sheet writes no first tab', () => {
    const s = storeOf({ ...wb(), activeTab: 2 });
    expect(s.transact([{ op: 'moveSheet', sheet: 'C', index: 0 }]).ok).toBe(true);
    expect(attr(view(s.package), 'activeTab')).toBe('0');
    expect(attr(view(s.package), 'firstSheet')).toBeUndefined();
    expect(s.transact([{ op: 'addSheet', name: 'New', index: 0 }]).ok).toBe(true);
    expect(attr(view(s.package), 'activeTab')).toBe('1');
    expect(attr(view(s.package), 'firstSheet')).toBeUndefined();
  });

  test('a view after the sheets moves, and no second view goes in', () => {
    const s = storeOf({
      ...wb(),
      bookViews: '',
      afterCalcPr: '<bookViews><workbookView activeTab="1"/></bookViews>',
    });
    expect(s.transact([{ op: 'moveSheet', sheet: 'B', index: 0 }]).ok).toBe(true);
    expect(children(workbookRoot(s.package), 'x:bookViews')).toHaveLength(1);
    expect(attr(view(s.package), 'activeTab')).toBe('0');
  });

  test('a move to the same place commits nothing', () => {
    const s = storeOf(wb());
    expect(s.transact([{ op: 'moveSheet', sheet: 'B', index: 1 }]).ok).toBe(true);
    expect(s.revision).toBe(0);
    expect(s.canUndo).toBe(false);
  });

  test('refuses a sheet range, and a bad position', () => {
    const s = storeOf({
      sheets: [
        { name: 'A', rows: '' },
        { name: 'B', rows: row(1, [c('A1', 0, { f: "SUM('A:C'!A1)" })]) },
        { name: 'C', rows: '' },
      ],
    });
    expect(rejection(s, [{ op: 'moveSheet', sheet: 'A', index: 2 }])?.code).toBe(
      'sheet-range-blocked'
    );
    expect(rejection(s, [{ op: 'moveSheet', sheet: 'A', index: 3 }])?.code).toBe('bad-position');
  });
});

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
