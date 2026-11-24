// More structural edit cases (task 7.5, design D4): table and query table parts of other
// sheets that name the edited sheet, workbook and package relationships outside an allowlist,
// inverse ops that only undo and redo apply, frozen panes, the edge rule of an insertion, the
// external workbook prefix, sheet ranges that span the edited sheet, and the relationship of a
// deleted hyperlink.

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildXlsx, c, row, type FixtureSheet } from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx, relationshipsOf, type XlsxPackage } from '../package/package.ts';
import { attr, child, children } from '../package/query.ts';
import { REL } from '../package/namespaces.ts';
import { WorkbookDocumentStore } from '../store/document-store.ts';
import type { WorkbookOp } from '../store/ops.ts';
import { moveReferences, type ReferenceContext } from '../structure/references.ts';
import { moveSpan, type AxisEdit } from '../structure/axis.ts';
import { op, sheetRoot, storeOf } from './structure-helpers.ts';

const enc = (s: string) => new TextEncoder().encode(s);
const TABLE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml';
const QUERY = 'application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

const data = (extra: Partial<FixtureSheet> = {}): FixtureSheet => ({
  name: 'Data',
  rows: row(1, [c('A1', 1)]) + row(3, [c('A3', 3)]),
  ...extra,
});

function refusal(store: WorkbookDocumentStore, edit: WorkbookOp = op('insertRows', 'Data', 2)) {
  const before = store.package;
  const r = store.transact([edit]);
  expect(store.package).toBe(before);
  return r.ok ? null : { code: r.rejection.code, subject: r.rejection.subject };
}

const table = (formula: string) =>
  enc(
    '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" displayName="T" ref="A1:B3">' +
      '<tableColumns count="2"><tableColumn id="1" name="A"/>' +
      `<tableColumn id="2" name="B"><calculatedColumnFormula>${formula}</calculatedColumnFormula></tableColumn>` +
      '</tableColumns></table>'
  );

const other = (): FixtureSheet => ({
  name: 'Other',
  rows: row(1, [c('A1', 'A'), c('B1', 'B')]),
  afterPageSetup: '<tableParts count="1"><tablePart r:id="rIdT"/></tableParts>',
  rels: [['rIdT', `${R}table`, '../tables/table1.xml']],
});

describe('parts that can name the edited sheet', () => {
  test('a table on another sheet whose formula names the edited sheet refuses', () => {
    const s = storeOf({
      sheets: [data(), other()],
      extraParts: [['xl/tables/table1.xml', TABLE, table('Data!A1*2')]],
    });
    expect(refusal(s)).toEqual({ code: 'structure-blocked', subject: '/xl/tables/table1.xml' });
  });

  test('a table on another sheet that does not name the edited sheet does not refuse', () => {
    const s = storeOf({
      sheets: [data(), other()],
      extraParts: [['xl/tables/table1.xml', TABLE, table('[@A]*2')]],
    });
    expect(s.transact([op('insertRows', 'Data', 2)]).ok).toBe(true);
  });

  test('a table formula that names the edited sheet in quotes refuses', () => {
    const s = storeOf({
      sheets: [{ ...data(), name: "It's" }, other()],
      extraParts: [['xl/tables/table1.xml', TABLE, table('&apos;It&apos;&apos;s&apos;!A1*2')]],
    });
    expect(refusal(s, op('insertRows', "It's", 2))?.subject).toBe('/xl/tables/table1.xml');
  });

  test('a table formula that a CDATA section or a comment splits still names the edited sheet', () => {
    for (const formula of ['Da<![CDATA[ta]]>!A1*2', 'Dat<!--x-->a!A1*2'] as const) {
      const s = storeOf({
        sheets: [data(), other()],
        extraParts: [['xl/tables/table1.xml', TABLE, table(formula)]],
      });
      expect([formula, refusal(s)?.subject]).toEqual([formula, '/xl/tables/table1.xml']);
    }
  });

  test('a query table part that names the edited sheet refuses', () => {
    const s = storeOf({
      sheets: [data()],
      extraParts: [
        [
          'xl/queryTables/queryTable1.xml',
          QUERY,
          enc('<queryTable xmlns="urn:q" name="Q" connectionId="1" note="Data!A1"/>'),
        ],
      ],
    });
    expect(refusal(s)?.subject).toBe('/xl/queryTables/queryTable1.xml');
  });

  for (const [what, name, rel] of [
    ['connections', 'xl/connections.xml', `${R}connections`],
    ['volatile dependencies', 'xl/volatileDependencies.xml', `${R}volatileDependencies`],
  ] as const)
    test(`a workbook relationship to ${what} refuses and names the part`, () => {
      const s = storeOf({
        sheets: [data()],
        extraParts: [[name, 'application/xml', enc('<x xmlns="urn:x"/>'), rel]],
      });
      expect(refusal(s)).toEqual({ code: 'structure-blocked', subject: `/${name}` });
    });

  test('custom XML data and cell metadata do not refuse', () => {
    const s = storeOf({
      sheets: [data()],
      extraParts: [
        ['customXml/item1.xml', 'application/xml', enc('<x xmlns="urn:x"/>'), `${R}customXml`],
        ['xl/metadata.xml', 'application/xml', enc('<x xmlns="urn:x"/>'), `${R}sheetMetadata`],
      ],
    });
    expect(s.transact([op('insertRows', 'Data', 2)]).ok).toBe(true);
  });

  /** The workbook with a package relationship of `type` to `target`. */
  function withPackageRel(type: string, target: string): XlsxPackage {
    const files = unzipSync(buildXlsx({ sheets: [data()] }));
    files[target] = enc('<x xmlns="urn:x"/>');
    files['_rels/.rels'] = strToU8(
      strFromU8(files['_rels/.rels']!).replace(
        '</Relationships>',
        `<Relationship Id="rIdP" Type="${type}" Target="${target}"/></Relationships>`
      )
    );
    const read = readXlsx(zipSync(files));
    if (!read.ok) throw new Error('read');
    return read.pkg;
  }

  test('package properties do not refuse; another package relationship does', () => {
    const props = new WorkbookDocumentStore(
      withPackageRel(
        'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
        'docProps/core.xml'
      )
    );
    expect(props.transact([op('insertRows', 'Data', 2)]).ok).toBe(true);
    const other = new WorkbookDocumentStore(withPackageRel('urn:example:custom', 'extra.xml'));
    expect(refusal(other)).toEqual({ code: 'structure-blocked', subject: '/extra.xml' });
  });

  const cf = (formula: string) =>
    `<conditionalFormatting sqref="A1"><cfRule type="expression" dxfId="0" priority="1"><formula>${formula}</formula></cfRule></conditionalFormatting>`;
  const ranged = () =>
    storeOf({
      sheets: [
        { name: 'First', rows: '' },
        data(),
        { name: 'Last', rows: '' },
        { name: 'Zed', rows: '' },
        { name: 'Other', rows: '', afterMerges: cf('SUM(First:Last!A1)&gt;1') },
      ],
      styles: { dxfs: ['<dxf/>'] },
    });

  test('a sheet range that spans the edited sheet in its middle refuses', () => {
    expect(refusal(ranged())).toEqual({
      code: 'structure-blocked',
      subject: 'x:conditionalFormatting (/xl/worksheets/sheet5.xml)',
    });
  });

  test('a sheet range that does not span the edited sheet does not refuse', () => {
    const s = ranged();
    expect(s.transact([op('insertRows', 'Zed', 1)]).ok).toBe(true);
  });
});

describe('inverse ops', () => {
  test('transact refuses restoreParts, restoreStyles, and restoreRow', () => {
    const s = storeOf({ sheets: [data()] });
    const before = s.package;
    const evil = { kind: 'binary', partName: '/evil.bin', bytes: enc('x') };
    const inverse = [
      { op: 'restoreParts', parts: [{ partName: '/evil.bin', part: evil }] },
      { op: 'restoreParts', parts: [{ partName: '/xl/worksheets/sheet1.xml', part: null }] },
      { op: 'restoreStyles', part: null },
      { op: 'restoreRow', sheet: 'Data', row: 1, node: null },
    ];
    for (const o of inverse) {
      const r = s.transact([op('insertRows', 'Data', 1), o as never]);
      expect(r.ok ? null : r.rejection.code).toBe('bad-value');
      expect(s.dryRun([o as never]).ok).toBe(false);
    }
    expect(s.package).toBe(before);
    expect(s.revision).toBe(0);
    expect(s.package.parts.has('/evil.bin')).toBe(false);
  });

  test('undo and redo still apply the inverse of a structural edit', () => {
    const s = storeOf({ sheets: [data()] });
    const before = s.package;
    expect(s.transact([op('insertRows', 'Data', 1)]).ok).toBe(true);
    expect(s.undo().ok).toBe(true);
    expect(s.package.parts.get('/xl/worksheets/sheet1.xml')).toBe(
      before.parts.get('/xl/worksheets/sheet1.xml')
    );
    expect(s.redo().ok).toBe(true);
  });
});

describe('frozen panes', () => {
  const frozen = (pane: string, selections: string, viewAttrs = '') =>
    storeOf({
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 'Head')]) + row(2, [c('A2', 1)]) + row(5, [c('A5', 5)]),
          viewAttrs,
          viewChildren: pane + selections,
        },
      ],
    });
  const view = (s: WorkbookDocumentStore) =>
    child(child(sheetRoot(s.package, 'S'), 'x:sheetViews'), 'x:sheetView')!;
  const pane = (s: WorkbookDocumentStore) => {
    const p = child(view(s), 'x:pane');
    return p && ['xSplit', 'ySplit', 'topLeftCell', 'activePane'].map((a) => attr(p, a));
  };
  const panes = (s: WorkbookDocumentStore) =>
    children(view(s), 'x:selection').map((el) => attr(el, 'pane') ?? 'topLeft');

  test('an insertion inside the band grows it, and the view keeps its top-left cell', () => {
    const s = frozen(
      '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>',
      '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>',
      'topLeftCell="A1"'
    );
    expect(s.transact([op('insertRows', 'S', 1)]).ok).toBe(true);
    expect(pane(s)).toEqual([undefined, '2', 'A3', 'bottomLeft']);
    expect(attr(view(s), 'topLeftCell')).toBe('A1');
    // An insertion after the band does not grow it.
    expect(s.transact([op('insertRows', 'S', 4)]).ok).toBe(true);
    expect(pane(s)).toEqual([undefined, '2', 'A3', 'bottomLeft']);
  });

  // The band starts at `x:sheetView/@topLeftCell` (ECMA-376 Part 1, 18.3.1.87): rows 10 and 11.
  const scrolled = () =>
    frozen(
      '<pane ySplit="2" topLeftCell="A12" activePane="bottomLeft" state="frozen"/>',
      '<selection pane="bottomLeft" activeCell="A12" sqref="A12"/>',
      'topLeftCell="A10"'
    );
  const band = (s: WorkbookDocumentStore) => [attr(view(s), 'topLeftCell'), ...pane(s)!];

  test('an edit before a band that starts at row 10 moves the band, and keeps its size', () => {
    const s = scrolled();
    expect(s.transact([op('insertRows', 'S', 3, 2)]).ok).toBe(true);
    expect(band(s)).toEqual(['A12', undefined, '2', 'A14', 'bottomLeft']);
    expect(s.transact([op('deleteRows', 'S', 2, 3)]).ok).toBe(true);
    expect(band(s)).toEqual(['A9', undefined, '2', 'A11', 'bottomLeft']);
  });

  test('an insertion at the first row of a band that starts at row 10 grows it', () => {
    const s = scrolled();
    expect(s.transact([op('insertRows', 'S', 10)]).ok).toBe(true);
    expect(band(s)).toEqual(['A10', undefined, '3', 'A13', 'bottomLeft']);
    // Right after the band: no growth, and the body moves down.
    expect(s.transact([op('insertRows', 'S', 13)]).ok).toBe(true);
    expect(band(s)).toEqual(['A10', undefined, '3', 'A14', 'bottomLeft']);
  });

  test('a deletion over the first row of a band that starts at row 10 shrinks it', () => {
    const s = scrolled();
    // Rows 9 and 10 go: row 11 is left in the band, and it is row 9 now.
    expect(s.transact([op('deleteRows', 'S', 9, 2)]).ok).toBe(true);
    expect(band(s)).toEqual(['A9', undefined, '1', 'A10', 'bottomLeft']);
  });

  test('a column band that starts at column C moves, grows, and shrinks on the columns axis', () => {
    const s = frozen(
      '<pane xSplit="2" topLeftCell="E1" activePane="topRight" state="frozen"/>',
      '<selection pane="topRight" activeCell="E1" sqref="E1"/>',
      'topLeftCell="C1"'
    );
    // Columns C and D are frozen. An insertion before them moves them.
    expect(s.transact([op('insertColumns', 'S', 1)]).ok).toBe(true);
    expect(band(s)).toEqual(['D1', '2', undefined, 'F1', 'topRight']);
    // An insertion inside the band grows it.
    expect(s.transact([op('insertColumns', 'S', 5)]).ok).toBe(true);
    expect(band(s)).toEqual(['D1', '3', undefined, 'G1', 'topRight']);
    // A deletion over the first column of the band shrinks it and moves it back.
    expect(s.transact([op('deleteColumns', 'S', 3, 2)]).ok).toBe(true);
    expect(band(s)).toEqual(['C1', '2', undefined, 'E1', 'topRight']);
  });

  test('a split past the end of the grid counts as layout cuts it', () => {
    // Rows 1048570 to 1048575 are the band, as layout cuts it, and row 1048576 is the body.
    const s = frozen(
      '<pane ySplit="4294967295" topLeftCell="A1048576" activePane="bottomLeft" state="frozen"/>',
      '<selection pane="bottomLeft" activeCell="A1048576" sqref="A1048576"/>',
      'topLeftCell="A1048570"'
    );
    // A deletion inside the band shrinks the count that layout reads, 6, to 5, not the split.
    expect(s.transact([op('deleteRows', 'S', 1048571)]).ok).toBe(true);
    expect(band(s)).toEqual(['A1048570', undefined, '5', 'A1048575', 'bottomLeft']);
  });

  test('an edit before a band with a split past the grid writes the split of the band', () => {
    const s = frozen(
      '<pane ySplit="4294967295" topLeftCell="A12" activePane="bottomLeft" state="frozen"/>',
      '<selection pane="bottomLeft" activeCell="A12" sqref="A12"/>',
      'topLeftCell="A10"'
    );
    // Layout reads rows 10 to 1048575 as the band. The deletion moves it up two rows, and the
    // split keeps its count, so the band does not grow into the last rows of the grid.
    expect(s.transact([op('deleteRows', 'S', 1, 2)]).ok).toBe(true);
    expect(band(s)).toEqual(['A8', undefined, String(1_048_576 - 1 - 9), 'A1048574', 'bottomLeft']);
  });

  test('a frozen view whose top-left cell is not one cell keeps it on an edit', () => {
    const s = frozen(
      '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>',
      '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>',
      'topLeftCell="A10:A11"'
    );
    // Layout reads the band from row 1, so the insertion at row 1 grows it.
    expect(s.transact([op('insertRows', 'S', 1)]).ok).toBe(true);
    expect(band(s)).toEqual(['A10:A11', undefined, '2', 'A3', 'bottomLeft']);
    // Right after the band: the body moves down.
    expect(s.transact([op('insertRows', 'S', 3)]).ok).toBe(true);
    expect(band(s)).toEqual(['A10:A11', undefined, '2', 'A4', 'bottomLeft']);
  });

  test('a deletion inside the band shrinks it; the scrolling pane starts after the band', () => {
    const s = frozen(
      '<pane ySplit="3" topLeftCell="A1" activePane="bottomLeft" state="frozen"/>',
      '<selection pane="bottomLeft" activeCell="A5" sqref="A5"/>'
    );
    expect(s.transact([op('deleteRows', 'S', 2)]).ok).toBe(true);
    expect(pane(s)).toEqual([undefined, '2', 'A3', 'bottomLeft']);
  });

  test('a band that reaches 0 goes, with its pane and the panes of its selections', () => {
    const s = frozen(
      '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>',
      '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>'
    );
    expect(s.transact([op('deleteRows', 'S', 1)]).ok).toBe(true);
    expect(pane(s)).toBeUndefined();
    expect(panes(s)).toEqual(['topLeft']);
    expect(s.undo().ok).toBe(true);
    expect(pane(s)).toEqual([undefined, '1', 'A2', 'bottomLeft']);
  });

  test('one axis that goes keeps the other, and one selection for each remaining pane', () => {
    const s = frozen(
      '<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>',
      '<selection pane="topRight" activeCell="B1" sqref="B1"/>' +
        '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>' +
        '<selection pane="bottomRight" activeCell="B2" sqref="B2"/>'
    );
    expect(s.transact([op('deleteColumns', 'S', 1)]).ok).toBe(true);
    expect(pane(s)).toEqual([undefined, '1', 'A2', 'bottomLeft']);
    expect(panes(s)).toEqual(['topLeft', 'bottomLeft']);
    const active = children(view(s), 'x:selection')[1]!;
    expect(attr(active, 'activeCell')).toBe('A2');
  });
});

describe('the edge rule of an insertion (design D4)', () => {
  const ins = (at: number): AxisEdit => ({ axis: 'rows', kind: 'insert', at, count: 1 });
  const ctx = (at: number): ReferenceContext => ({
    edit: ins(at),
    sheet: 'DATA',
    host: 'DATA',
    order: new Map([['DATA', 0]]),
  });
  const moved = (text: string, at: number) => {
    const m = moveReferences(text, ctx(at));
    return m.ok ? m.text : m.reason;
  };

  test('at the first row of a range it moves it', () => {
    expect(moveSpan(ins(1), 1, 3)).toEqual([2, 4]);
    expect(moved('SUM(A2:A4)', 1)).toBe('SUM(A3:A5)');
  });

  test('right after the last row of a range it does not grow it', () => {
    expect(moveSpan(ins(4), 1, 3)).toEqual([1, 3]);
    expect(moved('SUM(A2:A4)', 4)).toBe('SUM(A2:A4)');
  });

  test('strictly inside a range it grows it', () => {
    expect(moveSpan(ins(2), 1, 3)).toEqual([1, 4]);
    expect(moved('SUM(A2:A4)', 3)).toBe('SUM(A2:A5)');
  });

  test('an external workbook prefix, by index or by name, is not moved', () => {
    expect(moved('[Book.xlsx]Data!A5+[1]Data!A5+Data!A5', 0)).toBe(
      '[Book.xlsx]Data!A5+[1]Data!A5+Data!A6'
    );
  });

  test('a reference pushed off the grid becomes #REF!', () => {
    expect(moved('Data!A1048576', 0)).toBe('Data!#REF!');
  });
});

describe('hyperlink relationships', () => {
  test('a deleted hyperlink takes its relationship; undo puts it back', () => {
    const s = storeOf({
      sheets: [
        data({
          afterMerges:
            '<hyperlinks><hyperlink ref="A3" r:id="rIdH"/><hyperlink ref="A1" r:id="rIdK"/></hyperlinks>',
          rels: [
            ['rIdH', REL.hyperlink, 'https://example.com/a', true],
            ['rIdK', REL.hyperlink, 'https://example.com/b', true],
          ],
        }),
      ],
    });
    const ids = () =>
      relationshipsOf(s.package, '/xl/worksheets/sheet1.xml')
        .map((r) => r.id)
        .sort();
    expect(s.transact([op('deleteRows', 'Data', 3)]).ok).toBe(true);
    expect(ids()).toEqual(['rIdK']);
    expect(s.undo().ok).toBe(true);
    expect(ids()).toEqual(['rIdH', 'rIdK']);
  });
});
