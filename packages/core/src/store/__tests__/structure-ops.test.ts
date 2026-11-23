// Structural edits through the store (task 7.5, design D4): rows and columns move, every
// formula reference moves on every sheet, a reference to a deleted cell becomes `#REF!`, ranges
// shrink, defined names, shared and array formulas, merged ranges, column entries, hyperlinks,
// views, one undo step that restores the parts exactly, and the stale flag of design D3.

import { describe, expect, test } from 'bun:test';
import { c, row } from '../../../../../scripts/fixtures/builder.ts';
import { attr, child, children } from '../package/query.ts';
import { serializeXmlPart } from '../package-write/serialize.ts';
import { xmlPartOf } from '../package/package.ts';
import { workbookIndex } from '../workbook/workbook-index.ts';
import { MAX_COLS, MAX_ROWS } from '../cell-ref.ts';
import {
  cell,
  definedNames,
  formulaOf,
  op,
  refsOf,
  sheetRoot,
  storeOf,
  valueOf,
} from './structure-helpers.ts';

const formulas = (cells: readonly (readonly [string, string])[]) =>
  cells.map(([ref, f]) => c(ref, 0, { f }));

describe('the scenarios of the cell editing spec', () => {
  test('a formula follows an inserted row, and its result is stale', () => {
    const s = storeOf({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('C1', 1, { f: 'SUM(A1:A10)' })]) }],
    });
    expect(s.transact([op('insertRows', 'S', 5)]).ok).toBe(true);
    expect(formulaOf(s.package, 'S', 'C1')).toBe('SUM(A1:A11)');
    expect(workbookIndex(s.package)!.fullCalcOnLoad).toBe(true);
  });

  test('a deleted cell under a reference becomes #REF!, and a range shrinks', () => {
    const s = storeOf({
      sheets: [{ name: 'S', rows: row(1, formulas([['C1', 'A5+SUM(A4:A6)']])) }],
    });
    expect(s.transact([op('deleteRows', 'S', 5)]).ok).toBe(true);
    expect(formulaOf(s.package, 'S', 'C1')).toBe('#REF!+SUM(A4:A5)');
  });

  test('an insertion past the last row is refused and changes nothing', () => {
    const s = storeOf({ sheets: [{ name: 'S', rows: row(MAX_ROWS, [c(`A${MAX_ROWS}`, 1)]) }] });
    const before = s.package;
    const r = s.transact([op('insertRows', 'S', 2)]);
    expect(r.ok ? null : r.rejection.code).toBe('grid-full');
    expect(r.ok ? null : r.rejection.subject).toBe(`S!${MAX_ROWS}:${MAX_ROWS}`);
    expect(r.ok ? null : r.rejection.about).toEqual({
      kind: 'row',
      name: `S!${MAX_ROWS}:${MAX_ROWS}`,
    });
    expect(s.package).toBe(before);
    expect(s.revision).toBe(0);
  });

  test('a reference to an empty row moves, and nothing is sized by the grid', () => {
    const s = storeOf({ sheets: [{ name: 'S', rows: row(1, formulas([['C1', 'A900000']])) }] });
    expect(s.transact([op('insertRows', 'S', 2)]).ok).toBe(true);
    expect(formulaOf(s.package, 'S', 'C1')).toBe('A900001');
    expect(children(child(sheetRoot(s.package, 'S'), 'x:sheetData'), 'x:row').length).toBe(1);
  });
});

describe('cells and references', () => {
  const wb = () => ({
    sheets: [
      {
        name: 'Data',
        rows:
          row(1, [c('A1', 1), c('B1', 2), c('C1', 3), c('D1', 4)]) +
          row(3, [c('A3', 30)], 'ht="20" customHeight="1"') +
          row(5, [c('A5', 50), ...formulas([['B5', '$A$5+A3+$A3+A$3']])]) +
          row(
            6,
            formulas([
              ['A6', 'SUM(A1:A5)&"A5"'],
              ['B6', 'SUM(B:B)+SUM(2:4)+SUM(B1:D1)'],
            ])
          ),
      },
      {
        name: 'Other',
        rows: row(1, formulas([['A1', "Data!A5+'Data'!A3+A5+SUM(Data!A1:A5)"]])),
      },
    ],
  });

  test('rows move with their properties; cells of a deleted row go', () => {
    const s = storeOf(wb());
    expect(s.transact([op('deleteRows', 'Data', 2, 2)]).ok).toBe(true);
    const pkg = s.package;
    expect(valueOf(pkg, 'Data', 'A3')).toBe('50');
    expect(valueOf(pkg, 'Data', 'A1')).toBe('1');
    const rows = children(child(sheetRoot(pkg, 'Data'), 'x:sheetData'), 'x:row');
    expect(rows.map((r) => attr(r, 'r'))).toEqual(['1', '3', '4']);
    expect(rows.some((r) => attr(r, 'ht') === '20')).toBe(false);
  });

  test('absolute and relative references move, strings and other sheets stay', () => {
    const s = storeOf(wb());
    expect(s.transact([op('insertRows', 'Data', 2, 2)]).ok).toBe(true);
    const pkg = s.package;
    expect(attr(child(sheetRoot(pkg, 'Data'), 'x:sheetData')!.children[1] as never, 'ht')).toBe(
      '20'
    );
    expect(formulaOf(pkg, 'Data', 'B7')).toBe('$A$7+A5+$A5+A$5');
    expect(formulaOf(pkg, 'Data', 'A8')).toBe('SUM(A1:A7)&"A5"');
    expect(formulaOf(pkg, 'Data', 'B8')).toBe('SUM(B:B)+SUM(4:6)+SUM(B1:D1)');
    expect(formulaOf(pkg, 'Other', 'A1')).toBe("Data!A7+'Data'!A5+A5+SUM(Data!A1:A7)");
  });

  test('a column deletion moves cells and whole-column references', () => {
    const s = storeOf(wb());
    expect(s.transact([op('deleteColumns', 'Data', 2)]).ok).toBe(true);
    const pkg = s.package;
    expect(valueOf(pkg, 'Data', 'B1')).toBe('3');
    expect(valueOf(pkg, 'Data', 'C1')).toBe('4');
    expect(cell(pkg, 'Data', 'D1')).toBeUndefined();
    expect(formulaOf(pkg, 'Data', 'A6')).toBe('SUM(A1:A5)&"A5"');
    expect(formulaOf(pkg, 'Data', 'B6')).toBeUndefined();
    expect(formulaOf(pkg, 'Data', 'B5')).toBeUndefined();
    s.undo();
    expect(s.transact([op('deleteColumns', 'Data', 3)]).ok).toBe(true);
    expect(formulaOf(s.package, 'Data', 'B6')).toBe('SUM(B:B)+SUM(2:4)+SUM(B1:C1)');
  });

  test('a deleted whole column becomes #REF!, with its sheet prefix', () => {
    const s = storeOf({
      sheets: [{ name: 'Data', rows: row(1, formulas([['A1', 'SUM(B:B)+Data!B2+Data!A1:C1']])) }],
    });
    expect(s.transact([op('deleteColumns', 'Data', 2)]).ok).toBe(true);
    expect(formulaOf(s.package, 'Data', 'A1')).toBe('SUM(#REF!)+Data!#REF!+Data!A1:B1');
  });

  test('a range end pushed past the grid is cut; a cell reference there becomes #REF!', () => {
    const s = storeOf({
      sheets: [
        { name: 'S', rows: row(1, formulas([['A1', `SUM(B1:B${MAX_ROWS})+B${MAX_ROWS}`]])) },
      ],
    });
    expect(s.transact([op('insertRows', 'S', 3)]).ok).toBe(true);
    expect(formulaOf(s.package, 'S', 'A1')).toBe(`SUM(B1:B${MAX_ROWS})+#REF!`);
  });

  test('a reference that spans the edited sheet in a sheet range refuses', () => {
    const s = storeOf({
      sheets: [
        { name: 'A', rows: row(1, [c('A1', 1)]) },
        { name: 'B', rows: row(1, [c('A1', 1)]) },
        { name: 'C', rows: row(1, [c('A1', 1)]) },
        { name: 'D', rows: row(1, formulas([['A1', 'SUM(A:B!B2)']])) },
      ],
    });
    const r = s.transact([op('insertRows', 'B', 1)]);
    expect(r.ok ? null : [r.rejection.code, r.rejection.subject]).toEqual([
      'structure-blocked',
      'D!A1',
    ]);
    expect(s.transact([op('insertRows', 'C', 1)]).ok).toBe(true);
    expect(formulaOf(s.package, 'D', 'A1')).toBe('SUM(A:B!B2)');
  });

  test('cell @r is written on moved rows; spans go on column edits', () => {
    const s = storeOf({
      sheets: [{ name: 'S', rows: '<row r="2" spans="1:2"><c><v>1</v></c><c><v>2</v></c></row>' }],
    });
    expect(s.transact([op('insertColumns', 'S', 1)]).ok).toBe(true);
    const r = children(child(sheetRoot(s.package, 'S'), 'x:sheetData'), 'x:row')[0]!;
    expect(attr(r, 'spans')).toBeUndefined();
    expect(children(r, 'x:c').map((x) => attr(x, 'r'))).toEqual(['B2', 'C2']);
  });
});

describe('defined names', () => {
  const wb = (names: string) => ({
    definedNames: names,
    sheets: [
      { name: 'Data', rows: row(1, [c('A1', 1)]) },
      { name: 'Two', rows: '' },
    ],
  });

  test('move, shrink, and become #REF! like formula references', () => {
    const s = storeOf(
      wb(
        '<definedName name="Block">Data!$A$1:$A$5</definedName>' +
          '<definedName name="One">Data!$A$3</definedName>' +
          '<definedName name="Away">Two!$A$3</definedName>' +
          '<definedName name="_xlnm.Print_Titles" localSheetId="0">Data!$1:$3</definedName>'
      )
    );
    expect(s.transact([op('deleteRows', 'Data', 3)]).ok).toBe(true);
    expect(definedNames(s.package)).toEqual({
      Block: 'Data!$A$1:$A$4',
      One: 'Data!#REF!',
      Away: 'Two!$A$3',
      '_xlnm.Print_Titles': 'Data!$1:$2',
    });
  });

  test('a defined name with a reference and no sheet refuses', () => {
    const s = storeOf(wb('<definedName name="Loose">$A$1</definedName>'));
    const r = s.transact([op('insertRows', 'Data', 1)]);
    expect(r.ok ? null : [r.rejection.code, r.rejection.subject]).toEqual([
      'structure-blocked',
      'x:definedName Loose',
    ]);
  });
});

describe('shared and array formulas', () => {
  const shared = (master: string, children: readonly string[], text: string, ref: string) =>
    `<c r="${master}"><f t="shared" ref="${ref}" si="0">${text}</f><v>0</v></c>` +
    children.map((r) => `<c r="${r}"><f t="shared" si="0"/><v>0</v></c>`).join('');

  test('a shared formula whose children move alike stays shared, and its range moves', () => {
    const s = storeOf({
      sheets: [
        {
          name: 'S',
          rows:
            row(1, [shared('B1', [], 'A1*2+$A$1', 'B1:B4')]) +
            ['B2', 'B3', 'B4']
              .map((r, i) => row(i + 2, [`<c r="${r}"><f t="shared" si="0"/><v>0</v></c>`]))
              .join(''),
        },
      ],
    });
    expect(s.transact([op('insertRows', 'S', 3)]).ok).toBe(true);
    const pkg = s.package;
    const master = child(cell(pkg, 'S', 'B1')!, 'x:f')!;
    expect([attr(master, 'ref'), attr(master, 't')]).toEqual(['B1:B5', 'shared']);
    for (const r of ['B2', 'B4', 'B5'])
      expect(attr(child(cell(pkg, 'S', r)!, 'x:f'), 'si')).toBe('0');
  });

  test('children that move differently from the master are unshared', () => {
    const s = storeOf({
      sheets: [
        { name: 'Data', rows: row(1, [c('A1', 1)]) },
        {
          name: 'Refs',
          rows:
            row(1, [shared('B1', [], 'Data!A1', 'B1:B3')]) +
            row(2, ['<c r="B2"><f t="shared" si="0"/><v>0</v></c>']) +
            row(3, ['<c r="B3"><f t="shared" si="0"/><v>0</v></c>']),
        },
      ],
    });
    expect(s.transact([op('deleteRows', 'Data', 3)]).ok).toBe(true);
    const pkg = s.package;
    expect(attr(child(cell(pkg, 'Refs', 'B2')!, 'x:f'), 'si')).toBe('0');
    const b3 = child(cell(pkg, 'Refs', 'B3')!, 'x:f')!;
    expect([formulaOf(pkg, 'Refs', 'B3'), attr(b3, 't'), attr(b3, 'si')]).toEqual([
      'Data!#REF!',
      undefined,
      undefined,
    ]);
    s.undo();
    expect(s.transact([op('deleteRows', 'Data', 1)]).ok).toBe(true);
    const after = s.package;
    // Every child moved differently: each is its own formula, and the master is a normal one.
    expect(formulaOf(after, 'Refs', 'B1')).toBe('Data!#REF!');
    expect(attr(child(cell(after, 'Refs', 'B1')!, 'x:f'), 't')).toBeUndefined();
    expect(formulaOf(after, 'Refs', 'B2')).toBe('Data!A1');
    expect(formulaOf(after, 'Refs', 'B3')).toBe('Data!A2');
  });

  test('a deleted master unshares its children first', () => {
    const s = storeOf({
      sheets: [
        {
          name: 'S',
          rows:
            row(1, [shared('B1', [], 'A1*2', 'B1:B3')]) +
            row(2, ['<c r="B2"><f t="shared" si="0"/><v>0</v></c>']) +
            row(3, ['<c r="B3"><f t="shared" si="0"/><v>0</v></c>']),
        },
      ],
    });
    expect(s.transact([op('deleteRows', 'S', 1)]).ok).toBe(true);
    expect([formulaOf(s.package, 'S', 'B1'), formulaOf(s.package, 'S', 'B2')]).toEqual([
      'A1*2',
      'A2*2',
    ]);
    expect(attr(child(cell(s.package, 'S', 'B1')!, 'x:f'), 't')).toBeUndefined();
  });

  test('an array formula range moves as a whole, and an edit that splits it refuses', () => {
    const wb = {
      sheets: [
        {
          name: 'S',
          rows:
            row(2, ['<c r="C2"><f t="array" ref="C2:C3">A2:A3*2</f><v>0</v></c>']) +
            row(3, ['<c r="C3"><v>0</v></c>']),
        },
      ],
    };
    const s = storeOf(wb);
    for (const edit of [
      op('insertRows', 'S', 3),
      op('deleteRows', 'S', 3),
      op('deleteRows', 'S', 1, 2),
    ]) {
      const r = s.transact([edit]);
      expect(r.ok ? null : [r.rejection.code, r.rejection.subject]).toEqual([
        'structure-blocked',
        'S!C2:C3',
      ]);
    }
    expect(s.transact([op('insertRows', 'S', 1)]).ok).toBe(true);
    const f = child(cell(s.package, 'S', 'C3')!, 'x:f')!;
    expect([attr(f, 'ref'), formulaOf(s.package, 'S', 'C3')]).toEqual(['C3:C4', 'A3:A4*2']);
    expect(s.transact([op('deleteRows', 'S', 3, 2)]).ok).toBe(true);
    expect(cell(s.package, 'S', 'C3')).toBeUndefined();
  });
});

describe('merged ranges, column entries, hyperlinks, and views', () => {
  const wb = () => ({
    sheets: [
      {
        name: 'S',
        cols:
          '<cols><col min="2" max="4" width="20" customWidth="1"/>' +
          `<col min="6" max="${MAX_COLS}" width="5" customWidth="1"/></cols>`,
        rows: row(1, [c('A1', 1)]),
        merges: ['B2:C4', 'A6:A7', 'E2:F2'],
        viewAttrs: 'topLeftCell="A5"',
        viewChildren:
          '<pane ySplit="1" topLeftCell="A6" activePane="bottomLeft" state="frozen"/>' +
          '<selection pane="bottomLeft" activeCell="A7" sqref="A7 B9:C10"/>',
        afterMerges:
          '<hyperlinks><hyperlink ref="A7" location="S!A7" display="x"/>' +
          '<hyperlink ref="B9:C10" location="S!B10"/></hyperlinks>',
      },
    ],
  });

  test('merged ranges shrink, grow, and go when one cell is left', () => {
    const s = storeOf(wb());
    expect(s.transact([op('deleteRows', 'S', 3)]).ok).toBe(true);
    expect(refsOf(s.package, 'S', 'x:mergeCells', 'x:mergeCell')).toEqual([
      'B2:C3',
      'A5:A6',
      'E2:F2',
    ]);
    expect(s.transact([op('deleteRows', 'S', 6)]).ok).toBe(true);
    expect(refsOf(s.package, 'S', 'x:mergeCells', 'x:mergeCell')).toEqual(['B2:C3', 'E2:F2']);
    expect(attr(child(sheetRoot(s.package, 'S'), 'x:mergeCells'), 'count')).toBe('2');
    expect(s.transact([op('insertRows', 'S', 3, 2)]).ok).toBe(true);
    expect(refsOf(s.package, 'S', 'x:mergeCells', 'x:mergeCell')).toEqual(['B2:C5', 'E2:F2']);
    expect(s.transact([op('deleteRows', 'S', 2, 4)]).ok).toBe(true);
    expect(child(sheetRoot(s.package, 'S'), 'x:mergeCells')).toBeUndefined();
  });

  test('a selection keeps its active cell inside a kept range, and an unreadable one resets', () => {
    const selection = (store: ReturnType<typeof storeOf>) =>
      child(
        child(child(sheetRoot(store.package, 'S'), 'x:sheetViews'), 'x:sheetView'),
        'x:selection'
      );
    const s = storeOf(wb());
    // Row 7 holds A7, the active cell and the first range: B9:C10 is the range that is left.
    expect(s.transact([op('deleteRows', 'S', 7)]).ok).toBe(true);
    expect(attr(selection(s), 'sqref')).toBe('B8:C9');
    expect(attr(selection(s), 'activeCell')).toBe('B8');
    const broken = storeOf({
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 1)]),
          viewChildren: '<selection activeCell="B5" activeCellId="1" sqref="B5 nonsense"/>',
        },
      ],
    });
    expect(broken.transact([op('insertRows', 'S', 2)]).ok).toBe(true);
    expect(attr(selection(broken), 'sqref')).toBe('B6');
    expect(attr(selection(broken), 'activeCell')).toBe('B6');
    expect(attr(selection(broken), 'activeCellId')).toBeUndefined();
  });

  test('column entries split on an insertion, shift, and are cut at the last column', () => {
    const s = storeOf(wb());
    expect(s.transact([op('insertColumns', 'S', 3, 2)]).ok).toBe(true);
    const cols = () =>
      children(child(sheetRoot(s.package, 'S'), 'x:cols'), 'x:col').map((el) => [
        attr(el, 'min'),
        attr(el, 'max'),
      ]);
    expect(cols()).toEqual([
      ['2', '2'],
      ['5', '6'],
      ['8', String(MAX_COLS)],
    ]);
    expect(s.transact([op('deleteColumns', 'S', 2, 5)]).ok).toBe(true);
    expect(cols()).toEqual([['3', String(MAX_COLS - 5)]]);
  });

  test('hyperlinks move with their cells and go with them; views move', () => {
    const s = storeOf(wb());
    expect(s.transact([op('deleteRows', 'S', 7)]).ok).toBe(true);
    const root = sheetRoot(s.package, 'S');
    const links = children(child(root, 'x:hyperlinks'), 'x:hyperlink');
    expect(links.map((l) => [attr(l, 'ref'), attr(l, 'location')])).toEqual([['B8:C9', 'S!B9']]);
    const view = child(child(root, 'x:sheetViews'), 'x:sheetView')!;
    expect(attr(view, 'topLeftCell')).toBe('A5');
    expect(attr(child(view, 'x:pane'), 'topLeftCell')).toBe('A6');
    const sel = child(view, 'x:selection')!;
    // The active cell was deleted with its range: it moves into the range that is left.
    expect([attr(sel, 'activeCell'), attr(sel, 'sqref')]).toEqual(['B8', 'B8:C9']);
    // Row 1 is the frozen band: the insertion grows it, and the view keeps its top-left cell.
    expect(s.transact([op('insertRows', 'S', 1)]).ok).toBe(true);
    const view2 = child(child(sheetRoot(s.package, 'S'), 'x:sheetViews'), 'x:sheetView')!;
    expect(attr(view2, 'topLeftCell')).toBe('A5');
    expect(attr(child(view2, 'x:pane'), 'ySplit')).toBe('2');
    expect(attr(child(view2, 'x:pane'), 'topLeftCell')).toBe('A7');
    expect(attr(child(sheetRoot(s.package, 'S'), 'x:dimension'), 'ref')).toBeUndefined();
  });

  test('a hyperlink pushed past the grid refuses', () => {
    const s = storeOf({
      sheets: [
        {
          name: 'S',
          rows: '',
          afterMerges: `<hyperlinks><hyperlink ref="A${MAX_ROWS}" location="S!A1"/></hyperlinks>`,
        },
      ],
    });
    const r = s.transact([op('insertRows', 'S', 1)]);
    expect(r.ok ? null : r.rejection.code).toBe('grid-full');
    expect(r.ok ? null : r.rejection.about).toEqual({ kind: 'hyperlink', name: 'S' });
  });
});

describe('undo, redo, and the arguments', () => {
  const wb = () => ({
    definedNames: '<definedName name="N">S!$A$2</definedName>',
    sheets: [
      {
        name: 'S',
        rows: row(1, formulas([['B1', 'A2']])) + row(2, [c('A2', 2)]),
        merges: ['C2:D3'],
      },
    ],
  });

  test('one undo step restores every part exactly, and redo applies it again', () => {
    const s = storeOf(wb());
    const before = s.package;
    const sheet = serializeXmlPart(xmlPartOf(before, '/xl/worksheets/sheet1.xml')!);
    expect(s.transact([op('deleteRows', 'S', 2), op('insertColumns', 'S', 1)]).ok).toBe(true);
    const edited = serializeXmlPart(xmlPartOf(s.package, '/xl/worksheets/sheet1.xml')!);
    expect(s.undo().ok).toBe(true);
    expect(s.package.parts.get('/xl/worksheets/sheet1.xml')).toBe(
      before.parts.get('/xl/worksheets/sheet1.xml')
    );
    expect(definedNames(s.package)).toEqual({ N: 'S!$A$2' });
    // Undo never clears the stale flag (design D3).
    expect(workbookIndex(s.package)!.fullCalcOnLoad).toBe(true);
    expect(s.canUndo).toBe(false);
    expect(s.redo().ok).toBe(true);
    expect(serializeXmlPart(xmlPartOf(s.package, '/xl/worksheets/sheet1.xml')!)).toBe(edited);
    expect(edited).not.toBe(sheet);
  });

  test('undo puts back rows and cells that have no reference', () => {
    const s = storeOf({
      sheets: [
        { name: 'S', rows: '<row><c><v>1</v></c><c><v>2</v></c></row><row><c><v>3</v></c></row>' },
      ],
    });
    const before = s.package.parts.get('/xl/worksheets/sheet1.xml');
    expect(s.transact([op('insertRows', 'S', 1)]).ok).toBe(true);
    expect(valueOf(s.package, 'S', 'B2')).toBe('2');
    expect(s.undo().ok).toBe(true);
    expect(s.package.parts.get('/xl/worksheets/sheet1.xml')).toBe(before);
    expect(s.canRedo).toBe(true);
  });

  test('a shared master with no child keeps its form; an unrelated sheet keeps its part', () => {
    const s = storeOf({
      sheets: [
        { name: 'S', rows: row(1, [c('A1', 1)]) },
        { name: 'T', rows: row(1, ['<c r="B1"><f t="shared" ref="B1" si="0">A1</f><v>0</v></c>']) },
      ],
    });
    const t = s.package.parts.get('/xl/worksheets/sheet2.xml');
    expect(s.transact([op('insertRows', 'S', 1)]).ok).toBe(true);
    expect(s.package.parts.get('/xl/worksheets/sheet2.xml')).toBe(t);
  });

  test('bad positions and counts refuse with bad-reference', () => {
    const s = storeOf(wb());
    for (const edit of [
      op('insertRows', 'S', 0),
      op('insertRows', 'S', 1, 0),
      op('deleteRows', 'S', MAX_ROWS, 2),
      op('insertColumns', 'S', MAX_COLS + 1),
      op('deleteColumns', 'S', 1.5),
    ]) {
      const r = s.transact([edit]);
      expect(r.ok ? null : r.rejection.code).toBe('bad-reference');
    }
    const r = s.transact([op('insertRows', 'Nope', 1)]);
    expect(r.ok ? null : r.rejection.code).toBe('unknown-sheet');
  });

  test('ops read at an older revision refuse with stale-revision', () => {
    const s = storeOf(wb());
    const read = s.revision;
    expect(s.transact([op('insertRows', 'S', 1)], { baseRevision: read }).ok).toBe(true);
    const r = s.transact([op('deleteRows', 'S', 1)], { baseRevision: read });
    expect(r.ok ? null : r.rejection.code).toBe('stale-revision');
  });

  test('a workbook with no formula is not marked stale', () => {
    const s = storeOf({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] });
    expect(s.transact([op('insertRows', 'S', 1)]).ok).toBe(true);
    expect(workbookIndex(s.package)!.fullCalcOnLoad).toBe(false);
  });
});
