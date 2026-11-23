// The fixtures of the change `workbook-structure-editing`: each rule of designs D1, D2, and D3
// that a fixture covers holds in the edited bytes that the save check opens.

import { describe, expect, test } from 'bun:test';
import {
  attr,
  cellKey,
  child,
  children,
  parseCellRef,
  readXlsx,
  sheetByName,
  sheetIndex,
  textOf,
  workbookIndex,
  xmlPartOf,
  WorkbookDocumentStore,
  type Element,
  type XlsxPackage,
} from '../../packages/core/src/store/index.ts';
import { fixtureCopies } from '../save-check/records.ts';
import { FIXTURES, fixtureEditOps, type Fixture } from './catalog.ts';
import { buildXlsx, c, row } from './builder.ts';

function edited(name: string): XlsxPackage {
  const read = readXlsx(fixtureCopies(FIXTURES.find((f) => f.name === name)!).edited);
  if (!read.ok) throw new Error(JSON.stringify(read.error));
  return read.pkg;
}

function partOf(pkg: XlsxPackage, sheet: string) {
  return xmlPartOf(pkg, sheetByName(workbookIndex(pkg)!, sheet)!.partName!)!;
}

function cell(pkg: XlsxPackage, sheet: string, ref: string): Element | undefined {
  const a = parseCellRef(ref)!;
  return sheetIndex(partOf(pkg, sheet)).cells.get(cellKey(a.row, a.col))?.el;
}

const formula = (pkg: XlsxPackage, sheet: string, ref: string) =>
  textOf(child(cell(pkg, sheet, ref), 'x:f'));

describe('the sheet-add fixture', () => {
  test('the edit adds a third sheet, and the extended properties keep two titles', () => {
    const pkg = edited('sheet-add');
    expect(workbookIndex(pkg)!.sheets.map((s) => s.name)).toEqual(['Report', 'Data', 'Added']);
    const part = pkg.parts.get('/docProps/app.xml')!;
    const app = part.kind === 'binary' ? new TextDecoder().decode(part.bytes) : '';
    expect(app).toContain('<vt:vector size="2" baseType="lpstr">');
    expect(app).not.toContain('Added');
  });
});

describe('the sheet-ops fixture', () => {
  const pkg = edited('sheet-ops');
  const index = workbookIndex(pkg)!;
  const book = xmlPartOf(pkg, index.mainPart)!.root;

  test('the sheets are renamed, added, moved, and deleted', () => {
    expect(index.sheets.map((s) => s.name)).toEqual(['Notes', 'Sales Data', 'Summary']);
    expect(attr(child(child(book, 'x:bookViews'), 'x:workbookView'), 'activeTab')).toBe('2');
    expect([...pkg.parts.keys()].some((n) => n.endsWith('/sheet3.xml'))).toBe(false);
  });

  test('formulas, names, and hyperlinks follow the rename and the delete', () => {
    expect(formula(pkg, 'Summary', 'B2')).toBe("'Sales Data'!B2");
    expect(formula(pkg, 'Summary', 'B5')).toBe('#REF!*2');
    expect(formula(pkg, 'Summary', 'B6')).toBe("SUM('Sales Data'!B2:B3)/2");
    const names = children(child(book, 'x:definedNames'), 'x:definedName').map(
      (n) => `${attr(n, 'name')}@${attr(n, 'localSheetId') ?? '-'}=${textOf(n)}`
    );
    expect(names).toEqual(["Rate@1='Sales Data'!$B$2", "Total@-=SUM('Sales Data'!$B$2:$B$3)"]);
    const links = children(child(partOf(pkg, 'Sales Data').root, 'x:hyperlinks'), 'x:hyperlink');
    expect(links.map((h) => attr(h, 'location'))).toEqual(["'Sales Data'!B2", '#REF!']);
    expect(index.fullCalcOnLoad).toBe(true);
  });

  test('the extended properties keep their sheet titles', () => {
    // The store does not model the extended properties: the part stays as its bytes.
    const part = pkg.parts.get('/docProps/app.xml')!;
    const app = part.kind === 'binary' ? new TextDecoder().decode(part.bytes) : '';
    expect(app).toContain('<vt:lpstr>Scratch</vt:lpstr>');
  });
});

describe('the clipboard-edit fixture', () => {
  const pkg = edited('clipboard-edit');

  test('a copy pastes values, moved formulas, and cell formats', () => {
    expect(formula(pkg, 'Clip', 'D8')).toBe('B8*C8');
    expect(attr(cell(pkg, 'Clip', 'A7'), 's')).toBe('1');
    expect(attr(cell(pkg, 'Clip', 'C9'), 's')).toBe('2');
  });

  test('a cut clears its source and keeps its formula text', () => {
    expect(cell(pkg, 'Clip', 'A4')).toBeUndefined();
    expect(formula(pkg, 'Clip', 'I2')).toBe('B4*C4');
  });

  test('plain text pastes as typed input, and a typed form sets its number format', () => {
    expect(formula(pkg, 'Clip', 'C11')).toBe('B11*2');
    expect(textOf(child(cell(pkg, 'Clip', 'A12'), 'x:v'))).toBe('-5');
    expect(attr(cell(pkg, 'Clip', 'B11'), 's')).toBeUndefined();
    // A date, a percent, and a currency amount store numbers with the format of their form.
    const index = workbookIndex(pkg)!;
    const numFmt = (ref: string) => {
      const xf = children(child(xmlPartOf(pkg, index.stylesPart)!.root, 'x:cellXfs'), 'x:xf');
      return attr(xf[Number(attr(cell(pkg, 'Clip', ref), 's'))], 'numFmtId');
    };
    for (const [ref, v, id] of [
      ['C12', '45352', '14'],
      ['A13', '0.12', '9'],
      ['B13', '5.5', undefined],
    ] as const) {
      expect(attr(cell(pkg, 'Clip', ref), 't')).toBeUndefined();
      expect(textOf(child(cell(pkg, 'Clip', ref), 'x:v'))).toBe(v);
      if (id) expect(numFmt(ref)).toBe(id);
      else expect(Number(numFmt(ref))).toBeGreaterThanOrEqual(164);
    }
  });
});

describe('the resize-edit fixture', () => {
  const pkg = edited('resize-edit');
  const root = partOf(pkg, 'Sizes').root;

  test('column widths split the column entry and add one', () => {
    const cols = children(child(root, 'x:cols'), 'x:col').map((c) =>
      ['min', 'max', 'width', 'customWidth'].map((a) => attr(c, a)).join(' ')
    );
    expect(cols).toEqual(['1 1 12 1', '2 3 20 1', '4 4 12 1', '6 6 4 1']);
  });

  test('row heights set existing rows and add an empty row', () => {
    const rows = children(child(root, 'x:sheetData'), 'x:row').map((r) =>
      [attr(r, 'r'), attr(r, 'ht') ?? '-', attr(r, 'customHeight') ?? '-'].join(' ')
    );
    expect(rows).toEqual(['1 - -', '2 30 1', '3 30 1', '4 - -', '8 40 1']);
  });
});

describe('a clipboard edit after other edits', () => {
  test('reads its source after the earlier edits', () => {
    const read = readXlsx(buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] }));
    if (!read.ok) throw new Error('read');
    const fixture: Fixture = {
      name: 'probe',
      purpose: 'A value edit, then a copy of the edited cell.',
      workbook: () => ({ sheets: [] }),
      edits: [
        { sheet: 'S', ref: 'A1', value: 7 },
        { sheet: 'S', copy: 'A1', paste: 'C3' },
      ],
    };
    const store = new WorkbookDocumentStore(read.pkg);
    expect(store.transact(fixtureEditOps(fixture, read.pkg)).ok).toBe(true);
    expect(textOf(child(cell(store.package, 'S', 'C3'), 'x:v'))).toBe('7');
  });
});
