// The formula scan of the stale flag (design D3): an edit marks the workbook stale when any
// sheet part holds a formula cell, in every form that the file can store it. The scan reads
// every sheet part that `x:sheets` names, whatever its root element (a macro sheet too), and
// the cells that the sheet index skips: rows and cells inside `mc:AlternateContent`, and a
// `x:c` that the typed tree keeps generic.

import { describe, expect, test } from 'bun:test';
import { c, NS_X, row, type FixtureWorkbook } from '../../../../../scripts/fixtures/builder.ts';
import { contentTypeOf, readXlsx, xmlPartOf, type XlsxPackage } from '../package/package.ts';
import { attr, child } from '../package/query.ts';
import { saveXlsx } from '../package-write/save.ts';
import { sheetByName, workbookIndex } from '../workbook/workbook-index.ts';
import { sheetIndex } from '../workbook/sheet-index.ts';
import {
  cellOrNestedHasFormula,
  partHasFormula,
  rowHasFormula,
  sheetMayHoldFormula,
  workbookHasFormula,
} from '../store/calc.ts';
import { cellHasFormula, wrappedFormula } from '../workbook/cell-formula.ts';
import { parseXmlPart } from '../package/tree.ts';
import type { CellInput, WorkbookOp } from '../store/ops.ts';
import { op, storeOf } from './structure-helpers.ts';

const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
// The prefix that `@Requires` names is declared, as the reader requires.
const X14 = 'urn:example:x14';
const enc = new TextEncoder();

const fullCalc = (pkg: XlsxPackage) =>
  attr(child(xmlPartOf(pkg, workbookIndex(pkg)!.mainPart)!.root, 'x:calcPr'), 'fullCalcOnLoad');

/** Row 2 inside `mc:AlternateContent`, with the cell content `cell` in both branches. */
const wrappedRows = (cell: string) =>
  `<mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}">` +
  `<mc:Choice Requires="x14"><row r="2">${cell}</row></mc:Choice>` +
  `<mc:Fallback><row r="2">${cell}</row></mc:Fallback>` +
  '</mc:AlternateContent>';

/** A typed row 2 that holds one cell inside `mc:AlternateContent`. */
const wrappedCell = (cell: string) =>
  `<row r="2"><mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}">` +
  `<mc:Choice Requires="x14">${cell}</mc:Choice><mc:Fallback>${cell}</mc:Fallback>` +
  '</mc:AlternateContent></row>';

const FORMULA = '<c r="A2"><f>1+1</f><v>2</v></c>';
const VALUE = '<c r="A2"><v>2</v></c>';

/** A workbook with a plain sheet `S` and a sheet `T` whose rows are `rows`. */
const book = (rows: string): FixtureWorkbook => ({
  sheets: [
    { name: 'S', rows: row(1, [c('A1', 1)]) },
    { name: 'T', rows: row(1, [c('A1', 1)]) + rows },
  ],
});

/** An `x:c` directly in `x:sheetData`: not in a row, so the tree keeps it generic. */
const looseCell = (cell: string) => cell.replace('A2', 'A3');

/** A typed cell whose only `x:f` is inside `mc:AlternateContent`. */
const wrappedFormulaRow = (f: string) =>
  `<row r="2"><c r="A2"><mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}">` +
  `<mc:Choice Requires="x14">${f}</mc:Choice></mc:AlternateContent><v>2</v></c></row>`;

/** Each form of a formula that the sheet index does not count, with the same form without one. */
const FORMS: readonly (readonly [string, FixtureWorkbook, FixtureWorkbook])[] = [
  ['rows inside mc:AlternateContent', book(wrappedRows(FORMULA)), book(wrappedRows(VALUE))],
  ['a cell inside mc:AlternateContent', book(wrappedCell(FORMULA)), book(wrappedCell(VALUE))],
  ['a cell that the typed tree keeps generic', book(looseCell(FORMULA)), book(looseCell(VALUE))],
  [
    'a formula inside mc:AlternateContent in a cell',
    book(wrappedFormulaRow('<f>1+1</f>')),
    book(wrappedFormulaRow('<is><t>x</t></is>')),
  ],
];

// A macro sheet: an extension-namespace root that holds `CT_Macrosheet` content. Its part name,
// relationship type, and content type are the ones that workbooks store for it. The reader
// keeps it as unread bytes.
const MACRO_PART = '/xl/macrosheets/sheet1.xml';
const MACRO_TYPE = 'application/vnd.ms-excel.macrosheet+xml';
const MACRO_REL = 'http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet';
const MACRO_NS = 'http://schemas.microsoft.com/office/excel/2006/main';
const MACRO_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  `<xm:macrosheet xmlns="${NS_X}" xmlns:xm="${MACRO_NS}">` +
  '<sheetData><row r="1"><c r="A1"><f>1+1</f><v>2</v></c></row></sheetData></xm:macrosheet>';
const CHART_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml';
const CHART_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet';
const CHART_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  `<chartsheet xmlns="${NS_X}"><sheetViews><sheetView workbookViewId="0"/></sheetViews></chartsheet>`;

/** A workbook with a plain sheet `S` and one extra sheet part `M`. */
const extraSheetBook = (part: string, type: string, rel: string, xml: string): FixtureWorkbook => ({
  sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }],
  extraParts: [[part, type, enc.encode(xml), rel]],
  extraSheets: '<sheet name="M" sheetId="9" r:id="rIdX1"/>',
});

const setA1: readonly WorkbookOp[] = [
  { op: 'setCellValue', sheet: 'S', ref: 'A1', value: { type: 'number', value: 5 } },
];

describe('the formula scan of the stale flag', () => {
  for (const [name, withFormula, without] of FORMS) {
    test(`${name}: the sheet index counts no formula, the scan finds one`, () => {
      const store = storeOf(withFormula);
      const pkg = store.package;
      for (const sheet of workbookIndex(pkg)!.sheets) {
        const part = xmlPartOf(pkg, sheet.partName)!;
        if (sheet.kind === 'worksheet') expect(sheetIndex(part).formulas).toBe(0);
      }
      expect(workbookHasFormula(pkg)).toBe(true);
      expect(workbookHasFormula(storeOf(without).package)).toBe(false);
    });

    test(`${name}: a value edit on another sheet marks the workbook stale`, () => {
      const store = storeOf(withFormula);
      expect(fullCalc(store.package)).toBeUndefined();
      expect(store.transact(setA1).ok).toBe(true);
      expect(fullCalc(store.package)).toBe('1');
    });

    test(`${name}: the same edit with no formula leaves the flag unset`, () => {
      const store = storeOf(without);
      expect(store.transact(setA1).ok).toBe(true);
      expect(fullCalc(store.package)).toBeUndefined();
    });

    test(`${name}: the saved workbook keeps the flag and the formula after reopen`, () => {
      const store = storeOf(withFormula);
      expect(store.transact(setA1).ok).toBe(true);
      const reopened = readXlsx(saveXlsx(store.package));
      if (!reopened.ok) throw new Error(JSON.stringify(reopened.error));
      expect(workbookIndex(reopened.pkg)!.fullCalcOnLoad).toBe(true);
      expect(workbookHasFormula(reopened.pkg)).toBe(true);
    });
  }

  test('a structural edit on another sheet marks the workbook stale', () => {
    // A structural edit refuses wrapped content and a macro sheet, so the generic cell is the
    // form that reaches it.
    const store = storeOf(book(looseCell(FORMULA)));
    expect(store.transact([op('insertRows', 'S', 1)]).ok).toBe(true);
    expect(fullCalc(store.package)).toBe('1');
  });

  test('a cell nested in another cell counts, in a generic and in a typed cell', () => {
    for (const rows of [
      '<c r="A3"><c r="B9"><f>1+1</f><v>2</v></c></c>',
      '<row r="2"><c r="A2"><v>1</v><c r="B2"><f>1+1</f><v>2</v></c></c></row>',
    ]) {
      const store = storeOf(book(rows));
      const part = xmlPartOf(
        store.package,
        sheetByName(workbookIndex(store.package)!, 'T')!.partName
      )!;
      expect(partHasFormula(part)).toBe(true);
    }
  });

  test('a formula at the end of a deep chain of nested cells is found', () => {
    let rows = '<f>1</f>';
    for (let i = 0; i < 200; i += 1) rows = `<c r="A${i + 3}">${rows}</c>`;
    const store = storeOf(book(rows));
    const part = xmlPartOf(
      store.package,
      sheetByName(workbookIndex(store.package)!, 'T')!.partName
    )!;
    expect(partHasFormula(part)).toBe(true);
  });

  test('a macro sheet that the reader keeps as bytes counts as one with formulas', () => {
    const store = storeOf(extraSheetBook(MACRO_PART, MACRO_TYPE, MACRO_REL, MACRO_XML));
    expect(store.package.parts.get(MACRO_PART)?.kind).toBe('binary');
    expect(workbookHasFormula(store.package)).toBe(true);
    expect(store.transact(setA1).ok).toBe(true);
    expect(fullCalc(store.package)).toBe('1');
    // The save keeps the macro sheet byte for byte, and the reopened workbook is stale.
    const reopened = readXlsx(saveXlsx(store.package));
    if (!reopened.ok) throw new Error(JSON.stringify(reopened.error));
    expect(workbookIndex(reopened.pkg)!.fullCalcOnLoad).toBe(true);
    const bytes = reopened.pkg.parts.get(MACRO_PART);
    expect(bytes?.kind === 'binary' && new TextDecoder().decode(bytes.bytes)).toBe(MACRO_XML);
  });

  test('a content type is read without its parameters and without case', () => {
    const CASES: readonly (readonly [string, string, string, string, boolean])[] = [
      [MACRO_PART, `${MACRO_TYPE}; charset=UTF-8`, MACRO_REL, MACRO_XML, true],
      [MACRO_PART, MACRO_TYPE.toUpperCase(), MACRO_REL, MACRO_XML, true],
      ['/xl/chartsheets/sheet1.xml', `${CHART_TYPE}; v=1`, CHART_REL, CHART_XML, false],
      ['/xl/chartsheets/sheet1.xml', CHART_TYPE.toUpperCase(), CHART_REL, CHART_XML, false],
    ];
    for (const [part, type, rel, xml, counts] of CASES) {
      const store = storeOf(extraSheetBook(part, type, rel, xml));
      expect(contentTypeOf(store.package, part)).toBe(type);
      expect(workbookHasFormula(store.package)).toBe(counts);
    }
  });

  test('a sheet that names a part that is not XML does not count', () => {
    const store = storeOf(
      extraSheetBook(
        '/xl/media/image1.png',
        'image/png',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
        'not a sheet'
      )
    );
    expect(workbookHasFormula(store.package)).toBe(false);
  });

  test('a chartsheet with a root in another namespace does not count either', () => {
    const strict = CHART_XML.replace(NS_X, 'http://purl.oclc.org/ooxml/spreadsheetml/main');
    const store = storeOf(
      extraSheetBook('/xl/chartsheets/sheet1.xml', CHART_TYPE, CHART_REL, strict)
    );
    expect(workbookHasFormula(store.package)).toBe(false);
  });

  test('a chartsheet holds no cells, so it does not count', () => {
    const store = storeOf(
      extraSheetBook('/xl/chartsheets/sheet1.xml', CHART_TYPE, CHART_REL, CHART_XML)
    );
    expect(workbookHasFormula(store.package)).toBe(false);
    expect(store.transact(setA1).ok).toBe(true);
    expect(fullCalc(store.package)).toBeUndefined();
  });

  test('a parsed sheet part is scanned whatever its root element', () => {
    // A worksheet content type with a macro sheet root: the tree keeps the root generic, and
    // the workbook index lists the sheet as `other`.
    const store = storeOf(
      extraSheetBook(
        '/xl/worksheets/sheet9.xml',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
        MACRO_XML
      )
    );
    const entry = sheetByName(workbookIndex(store.package)!, 'M')!;
    expect(entry.kind).toBe('other');
    expect(partHasFormula(xmlPartOf(store.package, entry.partName)!)).toBe(true);
    expect(store.transact(setA1).ok).toBe(true);
    expect(fullCalc(store.package)).toBe('1');
  });

  test('a parsed sheet part in another namespace counts as one with formulas', () => {
    // A root in the Strict namespace: the scan cannot read its cells, as for unread bytes.
    const strict =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><sheetData>' +
      '<row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>';
    const store = storeOf(
      extraSheetBook(
        '/xl/worksheets/sheet9.xml',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
        strict
      )
    );
    expect(workbookHasFormula(store.package)).toBe(true);
    // The stale notice reads the same rule. The editor does not show such a sheet.
    expect(sheetMayHoldFormula(xmlPartOf(store.package, '/xl/worksheets/sheet9.xml')!)).toBe(true);
  });

  test('a dialog sheet holds no cells, so it does not count', () => {
    const store = storeOf(
      extraSheetBook(
        '/xl/dialogsheets/sheet1.xml',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet',
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<dialogsheet xmlns="${NS_X}"/>`
      )
    );
    expect(store.package.parts.get('/xl/dialogsheets/sheet1.xml')?.kind).toBe('binary');
    expect(workbookHasFormula(store.package)).toBe(false);
  });

  test('an unread sheet part with no content type counts as one with formulas', () => {
    // No default and no override names the `.dat` extension, so the part has no content type.
    const store = storeOf(extraSheetBook('/xl/macrosheets/sheet1.dat', '', MACRO_REL, CHART_XML));
    expect(contentTypeOf(store.package, '/xl/macrosheets/sheet1.dat')).toBeUndefined();
    expect(workbookHasFormula(store.package)).toBe(true);
  });

  test('a rule formula (x:formula) is not a cell formula', () => {
    // `x:formula` in a rule and `xm:f` in an extension have no cached value.
    const store = storeOf({
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 1)]),
          afterMerges:
            '<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1">' +
            '<formula>A1&gt;0</formula></cfRule></conditionalFormatting>',
          ext:
            '<ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
            '<x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">' +
            '<x14:cfRule type="expression" priority="2" id="{00000000-0000-0000-0000-000000000001}"><xm:f>A1&gt;1</xm:f></x14:cfRule>' +
            '<xm:sqref>A1</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext>',
        },
      ],
    });
    expect(workbookHasFormula(store.package)).toBe(false);
  });
});

// Design D5: an edit that can touch a formula drops the calculation chain. The same formula forms
// count there.
describe('the formula scan of the calculation chain', () => {
  const CHAIN = '/xl/calcChain.xml';
  const withChain = (wb: FixtureWorkbook, sheetId = 2): FixtureWorkbook => ({
    ...wb,
    extraParts: [
      [
        CHAIN,
        'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml',
        enc.encode(
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
            `<calcChain xmlns="${NS_X}"><c r="A2" i="${sheetId}"/></calcChain>`
        ),
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain',
      ],
    ],
  });

  describe('a shared master with a child inside a wrapper', () => {
    const rows =
      '<row r="2"><c r="A2"><f t="shared" ref="A2:A3" si="0">B2+1</f><v>2</v></c></row>' +
      `<row r="3"><c r="A3"><mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}"><mc:Choice Requires="x14">` +
      '<f t="shared" si="0"/><v>3</v></mc:Choice></mc:AlternateContent></c></row>';
    const set = (ref: string, value: CellInput): WorkbookOp => ({
      op: 'setCellValue',
      sheet: 'T',
      ref,
      value,
    });
    const EDITS: readonly (readonly [string, WorkbookOp[]])[] = [
      ['a value edit', [set('A2', { type: 'number', value: 9 })]],
      ['a formula edit', [set('A2', { type: 'formula', text: '1+1' })]],
      [
        'a batch, as a paste writes',
        [set('A2', { type: 'number', value: 9 }), set('B2', { type: 'number', value: 1 })],
      ],
      ['a clear of the master', [{ op: 'clearValues', sheet: 'T', ref: 'A2' }]],
      [
        'a paste over both cells, master first',
        [set('A2', { type: 'number', value: 9 }), set('A3', { type: 'number', value: 1 })],
      ],
    ];
    for (const [edit, ops] of EDITS)
      test(`${edit} is refused, so the child keeps its master`, () => {
        const store = storeOf(book(rows));
        const before = store.package;
        const result = store.transact(ops);
        expect(result.ok).toBe(false);
        if (!result.ok) expect(result.rejection.code).toBe('shared-formula');
        expect(store.package).toBe(before);
      });

    test('the refusal names the shared range, which holds the wrapped child', () => {
      const result = storeOf(book(rows)).transact([set('A2', { type: 'number', value: 9 })]);
      expect(result.ok ? undefined : result.rejection.subject).toBe('T!A2:A3');
      // The cause names the wrapped child, and the master is the cell to leave out.
      expect(result.ok ? undefined : result.rejection.causeCell).toBe('T!A3');
      expect(result.ok ? undefined : result.rejection.masterCell).toBe('T!A2');
    });

    test('a clear of the master and the child together is refused for the wrapped child', () => {
      const store = storeOf(book(rows));
      const before = store.package;
      const result = store.transact([{ op: 'clearValues', sheet: 'T', ref: 'A2:A3' }]);
      expect(result.ok ? undefined : result.rejection.code).toBe('cell-alternate-content');
      expect(result.ok ? undefined : result.rejection.subject).toBe('T!A3');
      expect(store.package).toBe(before);
    });

    test("a wrapped child in the master's own cell refuses the edit that replaces it", () => {
      const own =
        '<row r="2"><c r="A2"><f t="shared" ref="A2:A2" si="0">B2+1</f><v>2</v>' +
        `<mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}"><mc:Choice Requires="x14">` +
        '<f t="shared" si="0"/><v>2</v></mc:Choice></mc:AlternateContent></c></row>';
      const result = storeOf(book(own)).transact([set('A2', { type: 'number', value: 9 })]);
      expect(result.ok ? undefined : result.rejection.code).toBe('cell-alternate-content');
    });
  });

  test('a value edit of a wrapped formula cell is refused, and the cell and the chain stay', () => {
    const rows =
      `<row r="2"><c r="A2"><mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}"><mc:Choice Requires="x14">` +
      '<f>1+1</f><v>2</v></mc:Choice></mc:AlternateContent></c><c r="B2"><v>1</v></c></row>';
    const store = storeOf(withChain(book(rows)));
    const before = store.package;
    const result = store.transact([
      { op: 'setCellValue', sheet: 'T', ref: 'A2', value: { type: 'number', value: 5 } },
    ]);
    expect(result.ok ? undefined : result.rejection.code).toBe('cell-alternate-content');
    expect(store.package).toBe(before);
    expect(store.package.parts.has(CHAIN)).toBe(true);
  });

  test('an edit of a cell with a formula cell nested in it drops the chain', () => {
    const nested = '<row r="2"><c r="A2"><v>1</v><c r="B2"><f>1+1</f><v>2</v></c></c></row>';
    for (const op of [
      { op: 'setCellValue', sheet: 'T', ref: 'A2', value: { type: 'number', value: 7 } },
      { op: 'clearValues', sheet: 'T', ref: 'A2' },
    ] as const satisfies readonly WorkbookOp[]) {
      const store = storeOf(withChain(book(nested)));
      expect(store.transact([op]).ok).toBe(true);
      expect(store.package.parts.has(CHAIN)).toBe(false);
    }
  });

  test('a clear of a cell whose only content is a wrapped formula is refused', () => {
    const only =
      `<row r="2"><c r="A2"><mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}"><mc:Choice Requires="x14">` +
      '<f>1+1</f></mc:Choice></mc:AlternateContent></c></row>';
    const store = storeOf(withChain(book(only)));
    const before = store.package;
    const result = store.transact([{ op: 'clearValues', sheet: 'T', ref: 'A2' }]);
    expect(result.ok ? undefined : result.rejection.code).toBe('cell-alternate-content');
    expect(store.package).toBe(before);
  });

  // The edit cannot keep or replace the wrapper, so it names the cell (`cellChildBlocker`).
  describe('a formula range master inside a wrapper refuses each edit that replaces its cell', () => {
    const W = (inner: string) =>
      `<mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}">${inner}</mc:AlternateContent>`;
    const M = '<f t="shared" ref="A2:A3" si="0">B2+1</f><v>2</v>';
    const LAYOUTS: readonly (readonly [string, string])[] = [
      ['the only wrapped formula', W(`<mc:Choice Requires="x14">${M}</mc:Choice>`)],
      [
        'a Choice after a Fallback formula',
        W(`<mc:Choice Requires="x14">${M}</mc:Choice><mc:Fallback><f>A1+1</f></mc:Fallback>`),
      ],
      ['next to a direct formula', `<f>A1+1</f>${W(`<mc:Choice Requires="x14">${M}</mc:Choice>`)}`],
      [
        'a second sibling wrapper',
        W('<mc:Choice Requires="x14"><f>A1+1</f></mc:Choice>') +
          W(`<mc:Choice Requires="x14">${M}</mc:Choice>`),
      ],
      [
        'an array master',
        W('<mc:Choice Requires="x14"><f t="array" ref="A2:A3">{1;2}</f><v>1</v></mc:Choice>'),
      ],
      [
        'a data table master',
        W('<mc:Choice Requires="x14"><f t="dataTable" ref="A2:A3" r1="B1"/><v>1</v></mc:Choice>'),
      ],
      [
        'a wrapper nested in a branch',
        W(
          `<mc:Choice Requires="x14"><f>A1+1</f>${W(`<mc:Fallback>${M}</mc:Fallback>`)}</mc:Choice>`
        ),
      ],
    ];
    const EDITS: readonly (readonly [string, WorkbookOp[]])[] = [
      ['a clear', [{ op: 'clearValues', sheet: 'T', ref: 'A2:A3' }]],
      [
        'a value edit',
        [{ op: 'setCellValue', sheet: 'T', ref: 'A2', value: { type: 'number', value: 1 } }],
      ],
      [
        'a batch of value edits',
        [
          { op: 'setCellValue', sheet: 'T', ref: 'A2', value: { type: 'number', value: 1 } },
          { op: 'setCellValue', sheet: 'T', ref: 'B2', value: { type: 'number', value: 1 } },
        ],
      ],
    ];
    for (const [layout, inner] of LAYOUTS)
      for (const [edit, ops] of EDITS)
        test(`${layout}: ${edit} is refused`, () => {
          const rows =
            `<row r="2"><c r="A2">${inner}</c></row>` +
            '<row r="3"><c r="A3"><f t="shared" si="0"/><v>3</v></c></row>';
          const store = storeOf(book(rows));
          const before = store.package;
          const result = store.transact(ops);
          expect(result.ok).toBe(false);
          if (!result.ok) {
            expect(result.rejection.code).toBe('cell-alternate-content');
            expect(result.rejection.subject).toBe('T!A2');
          }
          expect(store.package).toBe(before);
        });
  });

  describe('a member of an array range whose master is inside a wrapper refuses each edit', () => {
    const rows =
      `<row r="2"><c r="A2"><mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}"><mc:Choice Requires="x14">` +
      '<f t="array" ref="A2:A3">{1;2}</f><v>1</v></mc:Choice></mc:AlternateContent></c></row>' +
      '<row r="3"><c r="A3"><v>2</v></c></row>';
    const set = (value: CellInput, ref = 'A3'): WorkbookOp => ({
      op: 'setCellValue',
      sheet: 'T',
      ref,
      value,
    });
    const EDITS: readonly (readonly [string, WorkbookOp[]])[] = [
      ['a value edit', [set({ type: 'number', value: 9 })]],
      ['a formula edit', [set({ type: 'formula', text: '5+5' })]],
      [
        'a batch, as a paste writes',
        [set({ type: 'number', value: 9 }), set({ type: 'number', value: 1 }, 'B3')],
      ],
      ['a clear', [{ op: 'clearValues', sheet: 'T', ref: 'A3' }]],
    ];
    for (const [edit, ops] of EDITS)
      test(`${edit} is refused`, () => {
        const store = storeOf(book(rows));
        const before = store.package;
        const result = store.transact(ops);
        expect(result.ok).toBe(false);
        if (!result.ok) expect(result.rejection.code).toBe('shared-formula');
        expect(store.package).toBe(before);
      });

    test('the refusal names the range as the reader reads it, not the raw @ref', () => {
      const store = storeOf(book(rows.replace('ref="A2:A3"', 'ref="A2:A3:XFD9"')));
      const result = store.transact([set({ type: 'number', value: 9 })]);
      expect(result.ok ? undefined : result.rejection.subject).toBe('T!A2:A3');
    });
  });

  test('a structural edit in a workbook whose only formula is a generic cell drops the chain', () => {
    const store = storeOf(withChain(book(looseCell(FORMULA))));
    expect(store.transact([op('insertRows', 'S', 1)]).ok).toBe(true);
    expect(store.package.parts.has(CHAIN)).toBe(false);
  });

  test('a sheet delete of the sheet whose only formula is a generic cell drops the chain', () => {
    const store = storeOf(withChain(book(looseCell(FORMULA))));
    expect(store.transact([{ op: 'deleteSheet', sheet: 'T' }]).ok).toBe(true);
    expect(store.package.parts.has(CHAIN)).toBe(false);
  });
});

// One rule: the cell rule, the scan of a part, the row check of undo, and the formula bar agree
// for every form of a formula cell.
describe('the formula cell rule', () => {
  const W = (inner: string) =>
    `<mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}">${inner}</mc:AlternateContent>`;
  const FORMS: readonly (readonly [string, string, boolean, string | undefined])[] = [
    ['an x:f child', '<f>1+1</f><v>2</v>', true, undefined],
    [
      'an x:f in a Choice',
      W('<mc:Choice Requires="x14"><f>1+2</f></mc:Choice>') + '<v>3</v>',
      true,
      '1+2',
    ],
    [
      'an x:f in a Fallback wins over a Choice',
      W('<mc:Choice Requires="x14"><f>NEW()</f></mc:Choice><mc:Fallback><f>1+3</f></mc:Fallback>'),
      true,
      '1+3',
    ],
    [
      'an x:f in a nested wrapper',
      W(`<mc:Choice Requires="x14">${W('<mc:Fallback><f>1+4</f></mc:Fallback>')}</mc:Choice>`),
      true,
      '1+4',
    ],
    [
      'an x:f under another element of a branch',
      W('<mc:Choice Requires="x14"><extLst><f>1</f></extLst></mc:Choice>'),
      false,
      undefined,
    ],
    [
      'the first of two sibling wrappers wins',
      W('<mc:Choice Requires="x14"><f>FIRST()</f></mc:Choice>') +
        W('<mc:Choice Requires="x14"><f>SECOND()</f></mc:Choice>'),
      true,
      'FIRST()',
    ],
    [
      'a branch wins over the wrappers nested in it',
      W(
        `<mc:Choice Requires="x14"><f>OUTER()</f>${W('<mc:Fallback><f>INNER()</f></mc:Fallback>')}</mc:Choice>`
      ),
      true,
      'OUTER()',
    ],
    ['a value only', '<v>2</v>', false, undefined],
    [
      'a value in a wrapper with no formula',
      W('<mc:Choice Requires="x14"><v>2</v><is><t>x</t></is></mc:Choice>'),
      false,
      undefined,
    ],
  ];
  for (const [name, inner, formula, text] of FORMS)
    test(`${name}: ${formula ? 'a formula cell' : 'no formula cell'}`, () => {
      const xml = `<worksheet xmlns="${NS_X}"><sheetData><row r="1"><c r="A1">${inner}</c></row></sheetData></worksheet>`;
      const read = parseXmlPart('/xl/worksheets/sheet1.xml', xml);
      if (!read.ok) throw new Error(read.reason);
      const rowEl = child(child(read.part.root, 'x:sheetData')!, 'x:row')!;
      const cell = child(rowEl, 'x:c')!;
      expect(cellHasFormula(cell)).toBe(formula);
      expect(cellOrNestedHasFormula(cell)).toBe(formula);
      expect(rowHasFormula(rowEl)).toBe(formula);
      expect(partHasFormula(read.part)).toBe(formula);
      const f = wrappedFormula(cell);
      expect(
        f ? f.children.map((n) => (n.type === 'text' ? n.value : '')).join('') : undefined
      ).toBe(text);
    });
});
