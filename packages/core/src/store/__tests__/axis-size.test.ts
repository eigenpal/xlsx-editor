// The resize ops of design D3 of the change `workbook-structure-editing`: `setColumnWidth`
// splits the `x:col` elements that it crosses and keeps their other attributes, `setRowHeight`
// creates the rows that lack an `x:row`, each refusal names its code, one undo gives back the
// worksheet part byte for byte. The editor tests check that layout geometry reads the result.

import { describe, expect, test } from 'bun:test';
import { c, row } from '../../../../../scripts/fixtures/builder.ts';
import { attr, child, children } from '../package/query.ts';
import { readXlsx, xmlPartOf, type XlsxPackage } from '../package/package.ts';
import { serializeXmlPart } from '../package-write/serialize.ts';
import { saveXlsx } from '../package-write/save.ts';
import { sheetByName, workbookIndex } from '../workbook/workbook-index.ts';
import type { WorkbookDocumentStore } from '../store/document-store.ts';
import type { WorkbookOp } from '../store/ops.ts';
import { sheetRoot, storeOf } from './structure-helpers.ts';

const partText = (pkg: XlsxPackage) => {
  const entry = sheetByName(workbookIndex(pkg)!, 'S')!;
  return serializeXmlPart(xmlPartOf(pkg, entry.partName)!);
};

/** Each `x:col` as `min-max` with its other attributes. */
const cols = (pkg: XlsxPackage) =>
  children(child(sheetRoot(pkg, 'S'), 'x:cols'), 'x:col').map((el) =>
    el.attrs.map((a) => `${a.local}=${a.value}`).join(' ')
  );

/** Each `x:row` as `r ht customHeight`. */
const rows = (pkg: XlsxPackage) =>
  children(child(sheetRoot(pkg, 'S'), 'x:sheetData'), 'x:row').map(
    (el) => `${attr(el, 'r')}:${attr(el, 'ht') ?? '-'}:${attr(el, 'customHeight') ?? '-'}`
  );

/** Commit, then undo gives back the part byte for byte, and redo the edit. */
function commit(s: WorkbookDocumentStore, op: WorkbookOp): XlsxPackage {
  const before = partText(s.package);
  const r = s.transact([op]);
  if (!r.ok) throw new Error(JSON.stringify(r.rejection));
  const after = partText(s.package);
  expect(s.undo().ok).toBe(true);
  expect(partText(s.package)).toBe(before);
  expect(s.redo().ok).toBe(true);
  expect(partText(s.package)).toBe(after);
  return s.package;
}

const refusal = (s: WorkbookDocumentStore, op: WorkbookOp) => {
  const before = s.package;
  const r = s.transact([op]);
  expect(s.package).toBe(before);
  return r.ok ? null : r.rejection.code;
};

const width = (sheet: string, first: number, last: number, w: number): WorkbookOp => ({
  op: 'setColumnWidth',
  sheet,
  first,
  last,
  width: w,
});

const height = (sheet: string, first: number, last: number, h: number): WorkbookOp => ({
  op: 'setRowHeight',
  sheet,
  first,
  last,
  height: h,
});

describe('setColumnWidth', () => {
  test('splits a crossed x:col and keeps its other attributes', () => {
    const s = storeOf({
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 1)]),
          cols: '<cols><col min="2" max="6" width="12" style="0" hidden="1" customWidth="1"/></cols>',
        },
      ],
    });
    const pkg = commit(s, width('S', 4, 4, 20));
    expect(cols(pkg)).toEqual([
      'min=2 max=3 width=12 style=0 hidden=1 customWidth=1',
      'min=4 max=4 width=20 style=0 hidden=1 customWidth=1',
      'min=5 max=6 width=12 style=0 hidden=1 customWidth=1',
    ]);
  });

  test('the resized piece drops bestFit, and the other pieces keep it', () => {
    const s = storeOf({
      sheets: [
        {
          name: 'S',
          rows: '',
          cols: '<cols><col min="1" max="3" width="9" bestFit="1" customWidth="1"/></cols>',
        },
      ],
    });
    const pkg = commit(s, width('S', 2, 2, 20));
    expect(cols(pkg)).toEqual([
      'min=1 max=1 width=9 bestFit=1 customWidth=1',
      'min=2 max=2 width=20 customWidth=1',
      'min=3 max=3 width=9 bestFit=1 customWidth=1',
    ]);
  });

  test('fills the gaps of the range with new x:col elements, in order', () => {
    const s = storeOf({
      sheets: [
        {
          name: 'S',
          rows: '',
          cols: '<cols><col min="3" max="3" width="5" style="0" customWidth="1"/></cols>',
        },
      ],
    });
    const pkg = commit(s, width('S', 2, 5, 9.5));
    expect(cols(pkg)).toEqual([
      'min=2 max=2 width=9.5 customWidth=1',
      'min=3 max=3 width=9.5 style=0 customWidth=1',
      'min=4 max=5 width=9.5 customWidth=1',
    ]);
  });

  test('creates x:cols before x:sheetData and keeps x:sheetFormatPr', () => {
    const s = storeOf({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] });
    const formatPr = partText(s.package).match(/<sheetFormatPr[^>]*>/)?.[0];
    const pkg = commit(s, width('S', 2, 2, 30));
    const names = sheetRoot(pkg, 'S').children.flatMap((n) =>
      n.type === 'element' ? [n.local] : []
    );
    expect(names.indexOf('cols')).toBe(names.indexOf('sheetData') - 1);
    expect(cols(pkg)).toEqual(['min=2 max=2 width=30 customWidth=1']);
    expect(partText(pkg).match(/<sheetFormatPr[^>]*>/)?.[0]).toBe(formatPr);
  });

  test('a width that is already set changes nothing', () => {
    const s = storeOf({
      sheets: [
        {
          name: 'S',
          rows: '',
          cols: '<cols><col min="1" max="1" width="7" customWidth="1"/></cols>',
        },
      ],
    });
    const r = s.transact([width('S', 1, 1, 7)]);
    expect(r.ok).toBe(true);
    expect(s.canUndo).toBe(false);
  });

  test('refuses widths and ranges outside their limits', () => {
    const s = storeOf({ sheets: [{ name: 'S', rows: '' }] });
    expect(refusal(s, width('S', 1, 1, 256))).toBe('bad-size');
    expect(refusal(s, width('S', 1, 1, -1))).toBe('bad-size');
    expect(refusal(s, width('S', 1, 1, Number.NaN))).toBe('bad-size');
    expect(refusal(s, width('S', 0, 1, 5))).toBe('bad-reference');
    expect(refusal(s, width('S', 3, 2, 5))).toBe('bad-reference');
    expect(refusal(s, width('S', 1, 16_385, 5))).toBe('bad-reference');
    expect(refusal(s, width('Nope', 1, 1, 5))).toBe('unknown-sheet');
  });

  test('a width is written with no exponent and no trailing zeros', () => {
    const s = storeOf({ sheets: [{ name: 'S', rows: '' }] });
    expect(s.transact([width('S', 1, 1, 1e-7)]).ok).toBe(true);
    expect(s.transact([width('S', 2, 2, 12.5)]).ok).toBe(true);
    expect(s.transact([width('S', 3, 3, 1.000000004)]).ok).toBe(true);
    expect(cols(s.package)).toEqual([
      'min=1 max=1 width=0.0000001 customWidth=1',
      'min=2 max=2 width=12.5 customWidth=1',
      'min=3 max=3 width=1 customWidth=1',
    ]);
  });

  test('a saved workbook reopens with the width', () => {
    const s = storeOf({ sheets: [{ name: 'S', rows: '' }] });
    s.transact([width('S', 1, 1, 0)]);
    const read = readXlsx(saveXlsx(s.package));
    if (!read.ok) throw new Error('reopen');
    expect(cols(read.pkg)).toEqual(['min=1 max=1 width=0 customWidth=1']);
  });
});

describe('drawing anchors', () => {
  test('a resize leaves the drawing part as it was (design D3)', () => {
    const anchor =
      '<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
      '<xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
      '<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="2" name="T"/><xdr:cNvSpPr/></xdr:nvSpPr>' +
      '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>';
    const s = storeOf({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]), drawing: { anchors: anchor } }],
    });
    const drawing = s.package.parts.get('/xl/drawings/drawing1.xml');
    expect(drawing).toBeDefined();
    expect(s.transact([width('S', 2, 3, 30)]).ok).toBe(true);
    expect(s.transact([height('S', 2, 3, 40)]).ok).toBe(true);
    expect(s.package.parts.get('/xl/drawings/drawing1.xml')).toBe(drawing);
  });

  // Columns B to D and rows 2 to 4 (0-based 1 to 3).
  const kept = (editAs: string, rows = row(1, [c('A1', 1)]), cols?: string) =>
    storeOf({
      sheets: [
        {
          name: 'S',
          rows,
          ...(cols ? { cols } : {}),
          drawing: {
            anchors:
              `<xdr:twoCellAnchor editAs="${editAs}"><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
              '<xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
              '<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="2" name="T"/><xdr:cNvSpPr/></xdr:nvSpPr>' +
              '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>',
          },
        },
      ],
    });
  const blocked = { code: 'resize-anchor', subject: 'T' } as const;
  const rejection = (s: WorkbookDocumentStore, op: WorkbookOp) => {
    const before = s.package;
    const r = s.transact([op]);
    expect(s.package).toBe(before);
    return r.ok ? null : { code: r.rejection.code, subject: r.rejection.subject };
  };

  test('an anchor that keeps its size refuses a resize that meets it, not one before it', () => {
    const s = kept('oneCell');
    expect(rejection(s, width('S', 3, 3, 30))).toEqual(blocked);
    expect(rejection(s, height('S', 3, 9, 40))).toEqual(blocked);
    expect(s.transact([width('S', 1, 1, 30)]).ok).toBe(true);
    expect(s.transact([height('S', 1, 1, 40)]).ok).toBe(true);
    expect(s.transact([width('S', 5, 6, 30)]).ok).toBe(true);
  });

  test('an anchor that keeps its position also refuses a resize before it', () => {
    const s = kept('absolute');
    expect(rejection(s, width('S', 1, 1, 30))).toEqual(blocked);
    expect(rejection(s, height('S', 1, 1, 40))).toEqual(blocked);
    expect(rejection(s, width('S', 2, 2, 30))).toEqual(blocked);
    expect(s.transact([width('S', 5, 6, 30)]).ok).toBe(true);
    expect(s.transact([height('S', 5, 5, 40)]).ok).toBe(true);
  });

  test('an anchor refuses a fitted row height that meets it, not one after it', () => {
    const s = kept('oneCell', row(3, [c('A3', 1)]) + row(6, [c('A6', 2)]));
    const fitted = (row: number): WorkbookOp => ({
      op: 'setFittedRowHeights',
      sheet: 'S',
      rows: [{ row, height: 15 }],
    });
    expect(rejection(s, fitted(3))).toEqual(blocked);
    expect(s.transact([fitted(6)]).ok).toBe(true);
  });

  test('an anchor refuses a column band that meets it, not one away from it', () => {
    const s = kept('oneCell');
    const band = (width: number): WorkbookOp => ({
      op: 'setColumnWidths',
      sheet: 'S',
      columns: [
        { column: 3, width },
        { column: 6, width: 30 },
      ],
    });
    expect(rejection(s, band(20))).toEqual(blocked);
    // A band that does not meet the anchor applies.
    const after: WorkbookOp = {
      op: 'setColumnWidths',
      sheet: 'S',
      columns: [{ column: 6, width: 30 }],
    };
    expect(s.transact([after]).ok).toBe(true);
  });

  test('an anchor lets a column fit add only its custom flag, and refuses a row that drops it', () => {
    const s = kept(
      'absolute',
      row(3, [c('A3', 1)], 'ht="15" customHeight="1"'),
      '<cols><col min="2" max="2" width="12"/></cols>'
    );
    const fitted: WorkbookOp = {
      op: 'setFittedRowHeights',
      sheet: 'S',
      rows: [{ row: 3, height: 15 }],
    };
    // A row without the flag would take its automatic height after a value edit.
    expect(rejection(s, fitted)).toEqual(blocked);
    expect(rows(s.package)).toEqual(['3:15:1']);
    // A false `@customHeight` is no custom height, so its loss changes no size.
    const off = kept('absolute', row(3, [c('A3', 1)], 'ht="15" customHeight="0"'));
    expect(off.transact([fitted]).ok).toBe(true);
    expect(rows(off.package)).toEqual(['3:15:-']);
    const band: WorkbookOp = {
      op: 'setColumnWidths',
      sheet: 'S',
      columns: [{ column: 2, width: 12 }],
    };
    expect(s.transact([band]).ok).toBe(true);
    expect(cols(s.package)).toEqual(['min=2 max=2 width=12 customWidth=1']);
  });
});

describe('the anchors that keep their size or position', () => {
  const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
  const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
  /** A marker: 0-based column and row, with offsets. */
  const marker = (tag: string, col: number, colOff: number, row: number, rowOff: number) =>
    `<xdr:${tag}><xdr:col>${col}</xdr:col><xdr:colOff>${colOff}</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>${rowOff}</xdr:rowOff></xdr:${tag}>`;
  const drawn = (editAs: string, toOff: number, name = 'Pic') =>
    storeOf({
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 1)]),
          drawing: {
            anchors:
              `<xdr:twoCellAnchor editAs="${editAs}">${marker('from', 1, 0, 1, 0)}${marker('to', 3, toOff, 3, toOff)}` +
              `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="2" name="${name}"/><xdr:cNvSpPr/></xdr:nvSpPr>` +
              '<xdr:spPr/></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>',
          },
        },
      ],
    });
  const control = (anchorAttrs: string, wrap = false, name = ' name="Button 1"') => {
    const list =
      `<controls><control shapeId="1025" r:id="rIdC"${name}><controlPr><anchor${anchorAttrs} xmlns:xdr="${XDR}">` +
      `${marker('from', 1, 0, 1, 0)}${marker('to', 3, 5, 3, 5)}</anchor></controlPr></control></controls>`;
    return storeOf({
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 1)]),
          afterPageSetup: wrap
            ? `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:x14="urn:example:x14" Requires="x14">${list}</mc:Choice></mc:AlternateContent>`
            : list,
        },
      ],
    });
  };
  const code = (s: WorkbookDocumentStore, op: WorkbookOp) => {
    const r = s.transact([op]);
    return r.ok ? null : [r.rejection.code, r.rejection.subject];
  };

  test('a to marker with offset 0 ends the span at the cell before it', () => {
    expect(code(drawn('oneCell', 0), width('S', 4, 4, 30))).toBeNull();
    expect(code(drawn('oneCell', 0), height('S', 4, 4, 30))).toBeNull();
    expect(code(drawn('oneCell', 7), width('S', 4, 4, 30))).toEqual(['resize-anchor', 'Pic']);
    expect(code(drawn('absolute', 0), width('S', 4, 4, 30))).toBeNull();
  });

  test('@editAs is read as a token, without white space around it', () => {
    expect(code(drawn(' oneCell ', 7), width('S', 3, 3, 30))).toEqual(['resize-anchor', 'Pic']);
    expect(code(drawn(' absolute', 7), width('S', 1, 1, 30))).toEqual(['resize-anchor', 'Pic']);
  });

  test('a shape with no name gives no subject', () => {
    expect(code(drawn('oneCell', 7, ''), width('S', 3, 3, 30))).toEqual([
      'resize-anchor',
      undefined,
    ]);
  });

  test('a control anchor keeps its size and position unless the file says it moves', () => {
    const named = ['resize-anchor', 'Button 1'];
    expect(code(control(''), width('S', 3, 3, 30))).toEqual(named);
    expect(code(control(''), width('S', 1, 1, 30))).toEqual(named);
    expect(code(control('', true), height('S', 2, 2, 30))).toEqual(named);
    expect(code(control(' moveWithCells="1"'), width('S', 1, 1, 30))).toBeNull();
    expect(code(control(' moveWithCells="1"'), width('S', 3, 3, 30))).toEqual(named);
    expect(code(control(' moveWithCells="1" sizeWithCells="1"'), width('S', 3, 3, 30))).toBeNull();
    expect(code(control(' sizeWithCells="true"'), width('S', 3, 3, 30))).toBeNull();
    expect(code(control(''), width('S', 6, 6, 30))).toBeNull();
    expect(code(control('', false, ''), width('S', 3, 3, 30))).toEqual([
      'resize-anchor',
      undefined,
    ]);
  });

  test('an embedded object anchor refuses a resize that would change it', () => {
    const s = storeOf({
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 1)]),
          afterPageSetup:
            `<oleObjects><oleObject progId="Package" shapeId="1026" r:id="rIdO"><objectPr><anchor moveWithCells="1" xmlns:xdr="${XDR}">` +
            `${marker('from', 1, 0, 1, 0)}${marker('to', 3, 5, 3, 5)}</anchor></objectPr></oleObject></oleObjects>`,
        },
      ],
    });
    expect(code(s, width('S', 2, 2, 30))).toEqual(['resize-anchor', undefined]);
    expect(code(s, width('S', 1, 1, 30))).toBeNull();
  });

  test('a legacy drawing anchor does not refuse a resize (legacy-drawing-anchors)', () => {
    const s = storeOf({
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 1)]),
          afterPageSetup: '<legacyDrawing r:id="rIdV"/>',
          rels: [
            [
              'rIdV',
              'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
              '../drawings/vmlDrawing1.vml',
            ],
          ],
        },
      ],
    });
    expect(code(s, width('S', 1, 9, 30))).toBeNull();
  });
});

describe('setRowHeight', () => {
  test('sets existing rows and creates missing ones in row order', () => {
    const s = storeOf({
      sheets: [{ name: 'S', rows: row(2, [c('A2', 1)]) + row(5, [c('A5', 2)]) }],
    });
    const pkg = commit(s, height('S', 1, 3, 24.75));
    expect(rows(pkg)).toEqual(['1:24.75:1', '2:24.75:1', '3:24.75:1', '5:-:-']);
    const a2 = children(child(sheetRoot(pkg, 'S'), 'x:sheetData'), 'x:row')[1]!;
    expect(children(a2, 'x:c')).toHaveLength(1);
  });

  test('appends rows past the last row', () => {
    const s = storeOf({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] });
    const pkg = commit(s, height('S', 3, 4, 0));
    expect(rows(pkg)).toEqual(['1:-:-', '3:0:1', '4:0:1']);
  });

  test('refuses rows out of order, so no row gets a second x:row', () => {
    const s = storeOf({
      sheets: [{ name: 'S', rows: [row(5, [c('A5', 1)]), row(2, [c('A2', 2)])].join('') }],
    });
    expect(refusal(s, height('S', 2, 4, 30))).toBe('invalid-sheet');
  });

  test('refuses heights and ranges outside their limits', () => {
    const s = storeOf({ sheets: [{ name: 'S', rows: '' }] });
    expect(refusal(s, height('S', 1, 1, 409.5))).toBe('bad-size');
    expect(refusal(s, height('S', 1, 1, -0.5))).toBe('bad-size');
    expect(refusal(s, height('S', 0, 1, 5))).toBe('bad-reference');
    expect(refusal(s, height('S', 1, 1_048_577, 5))).toBe('bad-reference');
    expect(refusal(s, height('S', 1, 20_000, 5))).toBe('too-many-rows');
  });
});

describe('setColumnWidths', () => {
  const widths = (columns: { column: number; width: number }[]): WorkbookOp => ({
    op: 'setColumnWidths',
    sheet: 'S',
    columns,
  });

  test('sets each column in one pass: splits, fills gaps, and joins equal neighbors', () => {
    const s = storeOf({
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 1)]),
          cols: '<cols><col min="2" max="6" width="9" style="1" bestFit="1"/></cols>',
        },
      ],
    });
    const pkg = commit(
      s,
      widths([
        { column: 1, width: 12 },
        { column: 3, width: 20 },
        { column: 4, width: 20 },
        { column: 8, width: 5.5 },
      ])
    );
    expect(cols(pkg)).toEqual([
      'min=1 max=1 width=12 customWidth=1',
      'min=2 max=2 width=9 style=1 bestFit=1',
      'min=3 max=4 width=20 style=1 customWidth=1',
      'min=5 max=6 width=9 style=1 bestFit=1',
      'min=8 max=8 width=5.5 customWidth=1',
    ]);
    // The same widths again change nothing.
    const before = s.package;
    expect(
      s.transact([
        widths([
          { column: 3, width: 20 },
          { column: 4, width: 20 },
        ]),
      ]).ok
    ).toBe(true);
    expect(s.package).toBe(before);
    // One column inside a wider `x:col` at the same width: the element does not split.
    expect(s.transact([widths([{ column: 3, width: 20 }])]).ok).toBe(true);
    expect(s.package).toBe(before);
  });

  test('refuses columns out of order or range, an empty or malformed list, and widths outside the limits', () => {
    const s = storeOf({ sheets: [{ name: 'S', rows: '' }] });
    expect(
      refusal(
        s,
        widths([
          { column: 2, width: 9 },
          { column: 2, width: 9 },
        ])
      )
    ).toBe('bad-reference');
    expect(refusal(s, widths([{ column: 16_385, width: 9 }]))).toBe('bad-reference');
    expect(refusal(s, widths([]))).toBe('bad-reference');
    expect(refusal(s, widths([{ column: 1, width: 256 }]))).toBe('bad-size');
    const malformed = {
      op: 'setColumnWidths',
      sheet: 'S',
      columns: [null],
    } as unknown as WorkbookOp;
    expect(refusal(s, malformed)).toBe('bad-reference');
  });

  test('sets every column of the grid in one pass', () => {
    const many = Array.from(
      { length: 4_000 },
      (_, i) => `<col min="${i * 4 + 1}" max="${i * 4 + 2}" width="7"/>`
    ).join('');
    const s = storeOf({ sheets: [{ name: 'S', rows: '', cols: `<cols>${many}</cols>` }] });
    const started = performance.now();
    const r = s.transact([
      widths(Array.from({ length: 16_384 }, (_, i) => ({ column: i + 1, width: 8 + (i % 5) }))),
    ]);
    expect(r.ok).toBe(true);
    // One pass: one op for each column took minutes on such a band.
    expect(performance.now() - started).toBeLessThan(20_000);
    expect(cols(s.package).slice(0, 2)).toEqual([
      'min=1 max=1 width=8 customWidth=1',
      'min=2 max=2 width=9 customWidth=1',
    ]);
    expect(cols(s.package)).toHaveLength(16_384);
  });
});

describe('setFittedRowHeights', () => {
  const fitted = (rows: { row: number; height: number }[]): WorkbookOp => ({
    op: 'setFittedRowHeights',
    sheet: 'S',
    rows,
  });

  test('writes @ht without @customHeight, and drops an existing one', () => {
    const s = storeOf({
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 1)], 'ht="40" customHeight="1"') + row(3, [c('A3', 2)]),
        },
      ],
    });
    const pkg = commit(
      s,
      fitted([
        { row: 1, height: 15 },
        { row: 3, height: 18.75 },
      ])
    );
    expect(rows(pkg)).toEqual(['1:15:-', '3:18.75:-']);
    // A height set by hand after the fit is custom again.
    expect(rows(commit(s, height('S', 1, 1, 20)))).toEqual(['1:20:1', '3:18.75:-']);
  });

  test('refuses a row with no x:row, rows out of order in the op or the file, an empty or malformed list, and heights outside the limits', () => {
    const s = storeOf({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) + row(2, []) }] });
    expect(refusal(s, fitted([{ row: 3, height: 15 }]))).toBe('bad-reference');
    expect(
      refusal(
        s,
        fitted([
          { row: 2, height: 15 },
          { row: 1, height: 15 },
        ])
      )
    ).toBe('bad-reference');
    expect(refusal(s, fitted([]))).toBe('bad-reference');
    // Rows out of order in the file, so one row number could have two x:row.
    const unordered = storeOf({
      sheets: [{ name: 'S', rows: [row(5, [c('A5', 1)]), row(2, [c('A2', 2)])].join('') }],
    });
    expect(refusal(unordered, fitted([{ row: 2, height: 15 }]))).toBe('invalid-sheet');
    expect(refusal(s, fitted([{ row: 1, height: 409.5 }]))).toBe('bad-size');
    const malformed = {
      op: 'setFittedRowHeights',
      sheet: 'S',
      rows: [null],
    } as unknown as WorkbookOp;
    expect(refusal(s, malformed)).toBe('bad-reference');
  });

  test('sets many rows in one pass', () => {
    const many = Array.from({ length: 20_000 }, (_, i) => row(i + 1, [c(`A${i + 1}`, i)])).join('');
    const s = storeOf({ sheets: [{ name: 'S', rows: many }] });
    const started = performance.now();
    const r = s.transact([
      fitted(Array.from({ length: 20_000 }, (_, i) => ({ row: i + 1, height: 15 + (i % 3) }))),
    ]);
    expect(r.ok).toBe(true);
    // One pass: well under the seconds that one op for each row took.
    expect(performance.now() - started).toBeLessThan(20_000);
    expect(rows(s.package).slice(0, 3)).toEqual(['1:15:-', '2:16:-', '3:17:-']);
  });
});
