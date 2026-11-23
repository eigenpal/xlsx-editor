// Structural edits that refuse, and drawing anchors (task 7.5, design D4): every element or part
// that can hold a reference which the edit can change refuses it with a reason that names it;
// formula text that the tokenizer cannot read refuses it and names the cell; a refusal never
// locks other editing. Drawing anchors move by their kind, the dimension moves, and an anchor
// that keeps its size or position in points refuses an edit that would move it.

import { describe, expect, test } from 'bun:test';
import { c, row, type FixtureSheet } from '../../../../../scripts/fixtures/builder.ts';
import { attr, child, children, textOf } from '../package/query.ts';
import { xmlPartOf, type XlsxPackage } from '../package/package.ts';
import { NS } from '../package/namespaces.ts';
import type { Element } from '../package/tree.ts';
import { MAX_ROWS } from '../cell-ref.ts';
import { WorkbookDocumentStore } from '../store/document-store.ts';
import { op, open, sheetRoot, storeOf, valueOf } from './structure-helpers.ts';

const base = (extra: Partial<FixtureSheet> = {}): FixtureSheet => ({
  name: 'Data',
  rows: row(1, [c('A1', 1)]) + row(3, [c('A3', 3)]),
  ...extra,
});

function refusal(store: WorkbookDocumentStore, edit = op('insertRows', 'Data', 2)) {
  const before = store.package;
  const r = store.transact([edit]);
  expect(store.package).toBe(before);
  return r.ok ? null : { code: r.rejection.code, subject: r.rejection.subject };
}

describe('content that refuses a structural edit', () => {
  const sheet1 = '/xl/worksheets/sheet1.xml';
  const cases: readonly (readonly [string, Partial<FixtureSheet>, string])[] = [
    [
      'conditional formatting',
      {
        afterMerges:
          '<conditionalFormatting sqref="A1:A3"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>2</formula></cfRule></conditionalFormatting>',
      },
      `x:conditionalFormatting (${sheet1})`,
    ],
    [
      'data validation',
      {
        afterMerges:
          '<dataValidations count="1"><dataValidation type="whole" sqref="A1"><formula1>1</formula1></dataValidation></dataValidations>',
      },
      `x:dataValidations (${sheet1})`,
    ],
    ['an auto filter', { beforeMerges: '<autoFilter ref="A1:A3"/>' }, `x:autoFilter (${sheet1})`],
    [
      'protected ranges',
      { beforeMerges: '<protectedRanges><protectedRange sqref="A1" name="p"/></protectedRanges>' },
      `x:protectedRanges (${sheet1})`,
    ],
    [
      'page breaks',
      { afterPageSetup: '<rowBreaks count="1"><brk id="2" max="16383" man="1"/></rowBreaks>' },
      `x:rowBreaks (${sheet1})`,
    ],
    [
      'extension content',
      {
        ext: '<ext uri="{00000000-0000-0000-0000-000000000001}"><x14:any xmlns:x14="urn:t"/></ext>',
      },
      `x:extLst (${sheet1})`,
    ],
    [
      'a table',
      {
        afterPageSetup: '<tableParts count="1"><tablePart r:id="rIdT"/></tableParts>',
        rels: [
          [
            'rIdT',
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
            '../tables/table1.xml',
          ],
        ],
      },
      `x:tableParts (${sheet1})`,
    ],
    [
      'a comment part',
      {
        rels: [
          [
            'rIdC',
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
            '../comments1.xml',
          ],
        ],
      },
      '/xl/comments1.xml',
    ],
  ];
  for (const [what, extra, subject] of cases)
    test(`${what} on the edited sheet refuses and is named`, () => {
      const s = storeOf({ sheets: [base(extra)], styles: { dxfs: ['<dxf/>'] } });
      expect(refusal(s)).toEqual({ code: 'structure-blocked', subject });
    });

  for (const local of [
    'sortState',
    'ignoredErrors',
    'cellWatches',
    'scenarios',
    'customSheetViews',
    'legacyDrawing',
    'oleObjects',
    'controls',
    'colBreaks',
  ])
    test(`x:${local} on the edited sheet refuses: no allowlist names it`, () => {
      const s = storeOf({ sheets: [base({ afterMerges: `<${local}/>` })] });
      expect(refusal(s, op('deleteColumns', 'Data', 1))).toEqual({
        code: 'structure-blocked',
        subject: `x:${local} (${sheet1})`,
      });
    });

  test('a pivot part anywhere refuses', () => {
    const s = storeOf({
      sheets: [base()],
      extraParts: [
        [
          'xl/pivotCache/pivotCacheDefinition1.xml',
          'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
          new TextEncoder().encode('<pivotCacheDefinition xmlns="urn:p"/>'),
        ],
      ],
    });
    expect(refusal(s)).toEqual({
      code: 'structure-blocked',
      subject: '/xl/pivotCache/pivotCacheDefinition1.xml',
    });
  });

  test('content of another sheet refuses only when it names the edited sheet', () => {
    const cf = (formula: string) =>
      `<conditionalFormatting sqref="A1"><cfRule type="expression" dxfId="0" priority="1"><formula>${formula}</formula></cfRule></conditionalFormatting>`;
    const named = storeOf({
      sheets: [base(), { name: 'Other', rows: '', afterMerges: cf('Data!A1&gt;1') }],
      styles: { dxfs: ['<dxf/>'] },
    });
    expect(refusal(named)).toEqual({
      code: 'structure-blocked',
      subject: 'x:conditionalFormatting (/xl/worksheets/sheet2.xml)',
    });
    const quoted = storeOf({
      sheets: [
        { ...base(), name: "It's" },
        { name: 'Other', rows: '', afterMerges: cf("'It''s'!A1&gt;1") },
      ],
      styles: { dxfs: ['<dxf/>'] },
    });
    expect(refusal(quoted, op('insertRows', "It's", 2))?.code).toBe('structure-blocked');
    const own = storeOf({
      sheets: [base(), { name: 'Other', rows: '', afterMerges: cf('A1&gt;1') }],
      styles: { dxfs: ['<dxf/>'] },
    });
    expect(own.transact([op('insertRows', 'Data', 2)]).ok).toBe(true);
  });

  test('a chart part anywhere refuses', () => {
    const s = storeOf({
      sheets: [base()],
      extraParts: [
        [
          'xl/charts/chart1.xml',
          'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
          new TextEncoder().encode('<c:chartSpace xmlns:c="urn:c"/>'),
        ],
      ],
    });
    expect(refusal(s)).toEqual({ code: 'structure-blocked', subject: '/xl/charts/chart1.xml' });
  });

  test('formula text from the file that the tokenizer cannot read refuses and names the cell', () => {
    const s = storeOf({
      sheets: [base(), { name: 'Other', rows: row(4, [c('B4', 0, { f: 'SUM(A1' })]) }],
    });
    expect(refusal(s)).toEqual({ code: 'structure-blocked', subject: 'Other!B4' });
  });

  test('a data table formula on the edited sheet refuses', () => {
    const s = storeOf({
      sheets: [
        base({ rows: row(2, ['<c r="B2"><f t="dataTable" ref="B2:B3" r1="A1"/><v>0</v></c>']) }),
      ],
    });
    expect(refusal(s)).toEqual({ code: 'structure-blocked', subject: 'Data!B2' });
  });

  test('a refusal does not lock other editing', () => {
    const s = storeOf({
      sheets: [
        base({
          afterMerges:
            '<conditionalFormatting sqref="A1"><cfRule type="expression" dxfId="0" priority="1"><formula>1</formula></cfRule></conditionalFormatting>',
        }),
      ],
      styles: { dxfs: ['<dxf/>'] },
    });
    expect(refusal(s)?.code).toBe('structure-blocked');
    expect(
      s.transact([
        { op: 'setCellValue', sheet: 'Data', ref: 'A2', value: { type: 'number', value: 5 } },
      ]).ok
    ).toBe(true);
    expect(
      s.transact([{ op: 'applyFormat', sheet: 'Data', ref: 'A2', patch: { bold: true } }]).ok
    ).toBe(true);
    expect(s.transact([{ op: 'clearValues', sheet: 'Data', ref: 'A1:A3' }]).ok).toBe(true);
    expect(valueOf(s.package, 'Data', 'A2')).toBeUndefined();
  });
});

const point = (tag: string, col: number, r: number, off = 0) =>
  `<xdr:${tag}><xdr:col>${col}</xdr:col><xdr:colOff>${off}</xdr:colOff><xdr:row>${r}</xdr:row><xdr:rowOff>${off}</xdr:rowOff></xdr:${tag}>`;
const sp = (id: number, textlink = '') =>
  `<xdr:sp macro="" textlink="${textlink}"><xdr:nvSpPr><xdr:cNvPr id="${id}" name="S${id}"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp>`;

function anchors(pkg: XlsxPackage): Element[] {
  const part = xmlPartOf(pkg, '/xl/drawings/drawing1.xml')!;
  return part.root.children.filter((n): n is Element => n.type === 'element');
}

const corner = (anchor: Element, tag: 'from' | 'to') => {
  const p = child(anchor, { ns: NS.xdr, local: tag })!;
  const get = (k: string) => Number(textOf(child(p, { ns: NS.xdr, local: k })));
  return [get('col'), get('colOff'), get('row'), get('rowOff')];
};

describe('drawing anchors', () => {
  const drawing = (list: string) => storeOf({ sheets: [base({ drawing: { anchors: list } })] });

  test('two-cell anchors move and resize; one-cell anchors move; absolute anchors stay', () => {
    const s = drawing(
      `<xdr:twoCellAnchor>${point('from', 1, 2, 5)}${point('to', 3, 6, 7)}${sp(2)}<xdr:clientData/></xdr:twoCellAnchor>` +
        `<xdr:oneCellAnchor>${point('from', 1, 8, 5)}<xdr:ext cx="100" cy="100"/>${sp(3)}<xdr:clientData/></xdr:oneCellAnchor>` +
        `<xdr:absoluteAnchor><xdr:pos x="10" y="10"/><xdr:ext cx="100" cy="100"/>${sp(4)}<xdr:clientData/></xdr:absoluteAnchor>`
    );
    const before = anchors(s.package)[2];
    expect(s.transact([op('insertRows', 'Data', 5, 2)]).ok).toBe(true);
    const [two, one, abs] = anchors(s.package);
    expect([corner(two!, 'from'), corner(two!, 'to')]).toEqual([
      [1, 5, 2, 5],
      [3, 7, 8, 7],
    ]);
    expect(corner(one!, 'from')).toEqual([1, 5, 10, 5]);
    expect(attr(child(one!, { ns: NS.xdr, local: 'ext' }), 'cx')).toBe('100');
    expect(abs).toBe(before);
    // A corner in a deleted band goes to the start of the first row after it.
    expect(s.transact([op('deleteRows', 'Data', 2, 2)]).ok).toBe(true);
    expect(corner(anchors(s.package)[0]!, 'from')).toEqual([1, 5, 1, 0]);
    expect(s.transact([op('deleteColumns', 'Data', 1)]).ok).toBe(true);
    expect(corner(anchors(s.package)[0]!, 'to')).toEqual([2, 7, 6, 7]);
  });

  test('an anchor that keeps its size in points refuses an edit that moves it', () => {
    const s = drawing(
      `<xdr:twoCellAnchor editAs="oneCell">${point('from', 1, 4)}${point('to', 3, 6)}${sp(2)}<xdr:clientData/></xdr:twoCellAnchor>`
    );
    expect(refusal(s)).toEqual({ code: 'structure-blocked', subject: '/xl/drawings/drawing1.xml' });
    expect(s.transact([op('insertRows', 'Data', 10)]).ok).toBe(true);
  });

  test('a shape linked to a cell refuses; an anchor pushed past the grid refuses', () => {
    const linked = drawing(
      `<xdr:oneCellAnchor>${point('from', 1, 1)}<xdr:ext cx="1" cy="1"/>${sp(2, '$A$1')}<xdr:clientData/></xdr:oneCellAnchor>`
    );
    expect(refusal(linked)?.code).toBe('structure-blocked');
    const far = drawing(
      `<xdr:oneCellAnchor>${point('from', 1, MAX_ROWS - 1)}<xdr:ext cx="1" cy="1"/>${sp(2)}<xdr:clientData/></xdr:oneCellAnchor>`
    );
    expect(refusal(far)?.code).toBe('grid-full');
    // The refusal names the drawing in words for user text, not its part path.
    for (const store of [far, linked]) {
      const r = store.transact([op('insertRows', 'Data', 2)]);
      expect(r.ok ? null : r.rejection.about).toEqual({ kind: 'drawing' });
    }
  });
});

describe('the dimension', () => {
  test('moves with the cells, and is A1 when no cell of it is left', () => {
    const pkg = open({ sheets: [base()] });
    const part = xmlPartOf(pkg, '/xl/worksheets/sheet1.xml')!;
    const dimension: Element = {
      type: 'element',
      id: 'd1',
      kind: 'dimension',
      ns: NS.x,
      local: 'dimension',
      nsDecls: [],
      attrs: [{ ns: '', local: 'ref', value: 'A1:C3' }],
      children: [],
    };
    const root = { ...part.root, children: [dimension, ...part.root.children] };
    const parts = new Map(pkg.parts);
    parts.set(part.partName, { kind: 'xml', partName: part.partName, xml: { ...part, root } });
    const s = new WorkbookDocumentStore({ parts });
    expect(s.transact([op('insertColumns', 'Data', 2)]).ok).toBe(true);
    const ref = () => attr(child(sheetRoot(s.package, 'Data'), 'x:dimension'), 'ref');
    expect(ref()).toBe('A1:D3');
    expect(s.transact([op('deleteRows', 'Data', 1, 3)]).ok).toBe(true);
    expect(ref()).toBe('A1');
    expect(children(child(sheetRoot(s.package, 'Data'), 'x:sheetData'), 'x:row')).toEqual([]);
  });
});
