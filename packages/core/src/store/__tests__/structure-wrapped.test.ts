// Formulas and defined names inside `mc:AlternateContent` under row and column inserts and
// deletes, and under a sheet rename, delete, and move (`structure/wrapped-formulas.ts`). An edit
// that would change a wrapped copy is refused, and the refusal names the cell or the name. An
// edit that leaves every copy as it is goes on, and a save and reopen keeps each branch.

import { describe, expect, test } from 'bun:test';
import { c, row, type FixtureWorkbook } from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx, xmlPartOf, type XlsxPackage } from '../package/package.ts';
import { saveXlsx } from '../package-write/save.ts';
import { textOf } from '../package/query.ts';
import type { Element } from '../package/tree.ts';
import { wrappedFormulas } from '../workbook/cell-formula.ts';
import { workbookIndex } from '../workbook/workbook-index.ts';
import { firstValueLost } from '../store/merge-cells.ts';
import type { WorkbookDocumentStore } from '../store/document-store.ts';
import type { WorkbookOp } from '../store/ops.ts';
import { cell, formulaOf, op, storeOf } from './structure-helpers.ts';

const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
// `mc:Choice/@Requires` names a prefix, which must be declared (ECMA-376 Part 3, 7.6).
const X14 = 'urn:example:x14';

/** An `mc:AlternateContent` with a Choice and a Fallback branch. */
const wrapper = (choice: string, fallback: string) =>
  `<mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}">` +
  `<mc:Choice Requires="x14">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback>` +
  '</mc:AlternateContent>';

/** A cell whose formula is only in the branches of a wrapper, with a cached value in each. */
const wrappedCell = (ref: string, choice: string, fallback = choice, f = '<f>') =>
  `<c r="${ref}">${wrapper(`${f}${choice}</f><v>1</v>`, `${f}${fallback}</f><v>1</v>`)}</c>`;

/** The text of each wrapped `x:f` of a cell, in branch order: Fallback first, then Choice. */
const wrappedTexts = (pkg: XlsxPackage, sheet: string, ref: string) => {
  const el = cell(pkg, sheet, ref);
  return el ? wrappedFormulas(el).map((f) => textOf(f)) : undefined;
};

/** The refusal of the ops, with the package unchanged, or null when they apply. */
function refusal(s: WorkbookDocumentStore, ops: readonly WorkbookOp[]) {
  const before = s.package;
  const r = s.transact(ops);
  if (r.ok) return null;
  expect(s.package).toBe(before);
  const { code, subject, about } = r.rejection;
  return { code, subject, about };
}

const reopen = (s: WorkbookDocumentStore): XlsxPackage => {
  const read = readXlsx(saveXlsx(s.package));
  if (!read.ok) throw new Error('reopen');
  return read.pkg;
};

/** Sheet S with values in A1 and A5, and the extra cells of row 2. */
const book = (row2: string, extra: Partial<FixtureWorkbook> = {}): FixtureWorkbook => ({
  sheets: [
    {
      name: 'S',
      rows: row(1, [c('A1', 1)]) + `<row r="2">${row2}</row>` + row(5, [c('A5', 5)]),
    },
  ],
  ...extra,
});

describe('a wrapped formula under row and column edits', () => {
  test('an edit that moves a reference of the formula is refused and names the cell', () => {
    for (const edit of [
      op('insertRows', 'S', 3),
      op('deleteRows', 'S', 3),
      op('deleteRows', 'S', 1),
      op('insertColumns', 'S', 1),
      op('deleteColumns', 'S', 1),
    ]) {
      const s = storeOf(book(wrappedCell('B2', 'A5')));
      expect(refusal(s, [edit])).toEqual({
        code: 'structure-blocked',
        subject: 'S!B2',
        about: { kind: 'wrapped-formula', name: 'S!B2' },
      });
    }
  });

  test('a change to the copy of one branch only is refused', () => {
    // The Choice names A1, which an insert at row 3 keeps; the Fallback names A5, which moves.
    for (const [choice, fallback] of [
      ['A1', 'A5'],
      ['A5', 'A1'],
    ] as const) {
      const s = storeOf(book(wrappedCell('B2', choice, fallback)));
      expect(refusal(s, [op('insertRows', 'S', 3)])?.subject).toBe('S!B2');
    }
  });

  test('an edit that keeps every copy goes on, and a save and reopen keeps each branch', () => {
    const s = storeOf(book(wrappedCell('B2', 'A1', 'SUM(A1)') + c('C2', 1, { f: 'A5' })));
    expect(refusal(s, [op('insertRows', 'S', 3), op('insertColumns', 'S', 5)])).toBeNull();
    // The direct formula moves; the wrapped copies keep their text.
    expect(formulaOf(s.package, 'S', 'C2')).toBe('A6');
    expect(wrappedTexts(s.package, 'S', 'B2')).toEqual(['SUM(A1)', 'A1']);
    const back = reopen(s);
    expect(formulaOf(back, 'S', 'C2')).toBe('A6');
    expect(wrappedTexts(back, 'S', 'B2')).toEqual(['SUM(A1)', 'A1']);
  });

  test('a cell whose wrapped formula keeps its text moves with its row', () => {
    const s = storeOf(book(wrappedCell('B2', 'A1')));
    expect(refusal(s, [op('insertRows', 'S', 2)])).toBeNull();
    expect(cell(s.package, 'S', 'B2')).toBeUndefined();
    expect(wrappedTexts(s.package, 'S', 'B3')).toEqual(['A1', 'A1']);
    expect(wrappedTexts(reopen(s), 'S', 'B3')).toEqual(['A1', 'A1']);
  });

  test('an edit that keeps a wrapped formula marks the saved copy for a recalculation', () => {
    // ROW() has no reference to move, but its value changes when its row moves.
    const s = storeOf(book(wrappedCell('B2', 'ROW()')));
    expect(refusal(s, [op('insertRows', 'S', 1)])).toBeNull();
    const back = reopen(s);
    expect(wrappedTexts(back, 'S', 'B3')).toEqual(['ROW()', 'ROW()']);
    expect(workbookIndex(back)!.fullCalcOnLoad).toBe(true);
  });

  test('a delete of the row of a wrapped formula goes on, and the cell goes with it', () => {
    // The delete would move A5, the reference of B2, but it deletes B2 too.
    const s = storeOf(book(wrappedCell('B2', 'A5')));
    expect(refusal(s, [op('deleteRows', 'S', 2)])).toBeNull();
    expect(cell(s.package, 'S', 'B2')).toBeUndefined();
    expect(wrappedTexts(s.package, 'S', 'B1')).toBeUndefined();
  });

  test('a wrapped formula that names the edited sheet from another sheet is refused', () => {
    const wb: FixtureWorkbook = {
      sheets: [
        { name: 'S', rows: row(5, [c('A5', 5)]) },
        { name: 'T', rows: `<row r="2">${wrappedCell('B2', 'S!A5')}</row>` },
      ],
    };
    const s = storeOf(wb);
    expect(refusal(s, [op('insertRows', 'S', 3)])).toEqual({
      code: 'structure-blocked',
      subject: 'T!B2',
      about: { kind: 'wrapped-formula', name: 'T!B2' },
    });
    // A reference without a sheet names T, which the edit of S does not move.
    const own = storeOf({
      sheets: [wb.sheets[0]!, { name: 'T', rows: `<row r="2">${wrappedCell('B2', 'A5')}</row>` }],
    });
    expect(refusal(own, [op('insertRows', 'S', 3)])).toBeNull();
    expect(wrappedTexts(own.package, 'T', 'B2')).toEqual(['A5', 'A5']);
  });

  test('a formula text that the move cannot read is refused', () => {
    const s = storeOf(book(wrappedCell('B2', 'SUM(')));
    expect(refusal(s, [op('insertRows', 'S', 3)])?.subject).toBe('S!B2');
  });
});

describe('wrapped formula ranges under row and column edits', () => {
  const array = (ref: string) => `<f t="array" ref="${ref}">`;

  test('a wrapped array master is refused when the edit moves, resizes, or deletes its range', () => {
    for (const edit of [
      op('insertRows', 'S', 1),
      op('insertRows', 'S', 3),
      op('deleteRows', 'S', 3),
      op('deleteRows', 'S', 2),
    ]) {
      const s = storeOf(book(wrappedCell('B2', '1', '1', array('B2:B3'))));
      expect(refusal(s, [edit])).toMatchObject({ code: 'structure-blocked', subject: 'S!B2' });
    }
    // An edit after the range keeps it.
    const s = storeOf(book(wrappedCell('B2', '1', '1', array('B2:B3'))));
    expect(refusal(s, [op('insertRows', 'S', 10)])).toBeNull();
  });

  test('a wrapped data table formula refuses every edit of its sheet', () => {
    const table = '<f t="dataTable" ref="B2:B3" dt2D="0" dtr="0" r1="A1">';
    const s = storeOf(book(wrappedCell('B2', '', '', table)));
    expect(refusal(s, [op('insertRows', 'S', 10)])).toMatchObject({ subject: 'S!B2' });
    expect(refusal(s, [op('insertColumns', 'S', 10)])).toMatchObject({ subject: 'S!B2' });
  });

  test('a shared formula with a wrapped child is refused when the edit moves a cell of it', () => {
    // A1 is a direct shared master; its child A2 keeps `x:f` inside the wrapper.
    const master = '<c r="A1"><f t="shared" ref="A1:A2" si="0">B1</f><v>1</v></c>';
    const child = wrappedCell('A2', '', '', '<f t="shared" si="0">');
    const wb: FixtureWorkbook = {
      sheets: [{ name: 'S', rows: `<row r="1">${master}</row><row r="2">${child}</row>` }],
    };
    for (const edit of [op('deleteRows', 'S', 1), op('insertRows', 'S', 2)]) {
      const s = storeOf(wb);
      expect(refusal(s, [edit])).toMatchObject({ code: 'structure-blocked', subject: 'S!A2' });
    }
    // An insert of column A moves the reference B1 of the master text.
    expect(refusal(storeOf(wb), [op('insertColumns', 'S', 2)])?.subject).toBe('S!A2');
    // An edit that moves no cell of the formula and no reference of its master goes on.
    expect(refusal(storeOf(wb), [op('insertRows', 'S', 10)])).toBeNull();
  });

  test('a shared formula with a wrapped member is refused when the edit changes a member', () => {
    const sharedChild = '<f t="shared" si="0">';
    const directChild = (ref: string) => `<c r="${ref}"><f t="shared" si="0"/><v>1</v></c>`;
    const rows = (...cells: string[]) =>
      cells.map((cellXml, i) => `<row r="${i + 1}">${cellXml}</row>`).join('');
    // A direct master, a wrapped child, and a direct child. An edit that moves only the reference
    // of A3 would make the store write A3 as its own formula and drop the master's sharing.
    const directMaster = rows(
      '<c r="A1"><f t="shared" ref="A1:A3" si="0">B10</f><v>1</v></c>',
      wrappedCell('A2', '', '', sharedChild),
      directChild('A3')
    );
    // A wrapped master with two direct children: the children's references move.
    const wrappedMaster = rows(
      wrappedCell('A1', 'B10', 'B10', '<f t="shared" ref="A1:A3" si="0">'),
      directChild('A2'),
      directChild('A3')
    );
    // An insert at row 11 moves B11, the reference of A2; one at row 12 moves only B12, the
    // reference of A3.
    for (const rows2 of [directMaster, wrappedMaster]) {
      const wb: FixtureWorkbook = { sheets: [{ name: 'S', rows: rows2 }] };
      for (const [at, subject] of [
        [11, 'S!A2'],
        [12, 'S!A3'],
      ] as const)
        expect(refusal(storeOf(wb), [op('insertRows', 'S', at)])).toMatchObject({
          code: 'structure-blocked',
          subject,
          about: { kind: 'wrapped-formula', name: subject },
        });
      expect(refusal(storeOf(wb), [op('insertRows', 'S', 20)])).toBeNull();
    }
    // On another sheet, a wrapped child whose reference to the edited sheet moves.
    const other: FixtureWorkbook = {
      sheets: [
        { name: 'S', rows: row(1, [c('A1', 1)]) },
        {
          name: 'T',
          rows: rows(
            '<c r="A1"><f t="shared" ref="A1:A2" si="0">S!B10</f><v>1</v></c>',
            wrappedCell('A2', '', '', sharedChild)
          ),
        },
      ],
    };
    expect(refusal(storeOf(other), [op('insertRows', 'S', 11)])?.subject).toBe('T!A2');
    expect(refusal(storeOf(other), [op('insertRows', 'S', 20)])).toBeNull();
  });

  test('a shared formula with a wrapped member and a master text that cannot be read is refused', () => {
    const wb: FixtureWorkbook = {
      sheets: [
        {
          name: 'S',
          rows:
            '<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">SUM(</f><v>1</v></c></row>' +
            `<row r="2">${wrappedCell('A2', '', '', '<f t="shared" si="0">')}</row>`,
        },
      ],
    };
    expect(refusal(storeOf(wb), [op('insertRows', 'S', 10)])).toMatchObject({
      code: 'structure-blocked',
      subject: 'S!A1',
    });
  });

  test('wrapped master copies in two cells, or too many master texts, are refused', () => {
    const master = (ref: string, text: string) =>
      wrappedCell(ref, text, text, '<f t="shared" ref="A1:A2" si="0">');
    // Each of many cells claims to be the master: the check stops at the second one.
    const rows = Array.from(
      { length: 2000 },
      (_, i) => `<row r="${i + 1}">${master(`A${i + 1}`, 'B5000')}</row>`
    );
    const many: FixtureWorkbook = {
      sheets: [
        { name: 'S', rows: row(1, [c('A1', 1)]) },
        { name: 'T', rows: rows.join('') },
      ],
    };
    expect(refusal(storeOf(many), [op('insertRows', 'S', 3)])).toMatchObject({
      code: 'structure-blocked',
      subject: 'T!A2',
    });
    // One cell with five distinct master texts in nested wrappers.
    const f = (text: string) => `<f t="shared" ref="A1:A2" si="0">${text}</f>`;
    const nested = wrapper(
      f('B1') + wrapper(f('B2') + wrapper(f('B3'), f('B4')), f('B5')),
      f('B1')
    );
    const five: FixtureWorkbook = {
      sheets: [{ name: 'S', rows: `<row r="1"><c r="A1">${nested}</c></row>` }],
    };
    expect(refusal(storeOf(five), [op('insertRows', 'S', 20)])?.subject).toBe('S!A1');
  });

  test('a cell with wrapped copies in many shared formulas costs no more than its copies', () => {
    // Each copy has its own `@si`, so the cell is in each group. A check that read every copy of
    // the cell for each group would take minutes on a file of a few MB.
    const count = 20000;
    const many = (one: (i: number) => string) => {
      let list = '';
      for (let i = 0; i < count; i++) list += one(i);
      return `<c r="B2">${wrapper(list, '<v>1</v>')}</c>`;
    };
    for (const one of [
      (i: number) => `<f t="shared" si="${i}"/>`,
      (i: number) => `<f t="shared" si="${i}" ref="B2">A1</f>`,
    ]) {
      const s = storeOf(book(many(one)));
      const start = performance.now();
      expect(refusal(s, [op('insertRows', 'S', 10)])).toBeNull();
      expect(performance.now() - start).toBeLessThan(2000);
    }
  });

  test('a wrapped shared master is refused when the edit moves its range', () => {
    const master = wrappedCell('A1', 'B1', 'B1', '<f t="shared" ref="A1:A3" si="0">');
    const wb: FixtureWorkbook = { sheets: [{ name: 'S', rows: `<row r="1">${master}</row>` }] };
    expect(refusal(storeOf(wb), [op('insertRows', 'S', 2)])?.subject).toBe('S!A1');
    expect(refusal(storeOf(wb), [op('insertRows', 'S', 10)])).toBeNull();
  });
});

describe('a wrapped formula under sheet ops', () => {
  const sheets = (formula: string) => ({
    sheets: [
      { name: 'S', rows: row(1, [c('A1', 1)]) },
      { name: 'T', rows: `<row r="2">${wrappedCell('B2', formula)}</row>` },
      { name: 'U', rows: row(1, [c('A1', 1)]) },
      { name: 'V', rows: row(1, [c('A1', 1)]) },
    ],
  });

  test('a rename and a delete of the sheet that the formula names are refused', () => {
    for (const edit of [
      { op: 'renameSheet', sheet: 'S', name: 'New' },
      { op: 'deleteSheet', sheet: 'S' },
    ] as const) {
      expect(refusal(storeOf(sheets('S!A1')), [edit])).toEqual({
        code: 'structure-blocked',
        subject: 'T!B2',
        about: { kind: 'wrapped-formula', name: 'T!B2' },
      });
    }
  });

  test('a rename, a delete, and a move that keep the text go on', () => {
    for (const edit of [
      { op: 'renameSheet', sheet: 'S', name: 'New' },
      { op: 'deleteSheet', sheet: 'S' },
      { op: 'moveSheet', sheet: 'V', index: 0 },
    ] as const) {
      const s = storeOf(sheets('U!A1'));
      expect(refusal(s, [edit])).toBeNull();
      expect(wrappedTexts(reopen(s), 'T', 'B2')).toEqual(['U!A1', 'U!A1']);
    }
  });

  test('a move that changes the sheets of a wrapped sheet range is refused', () => {
    const s = storeOf(sheets('SUM(S:U!A1)'));
    const refused = {
      code: 'sheet-range-blocked',
      subject: 'T!B2',
      about: { kind: 'wrapped-formula', name: 'T!B2' },
    } as const;
    expect(refusal(s, [{ op: 'moveSheet', sheet: 'V', index: 1 }])).toEqual(refused);
    // A delete of an end of the range is refused too.
    const d = storeOf(sheets('SUM(S:U!A1)'));
    expect(refusal(d, [{ op: 'deleteSheet', sheet: 'U' }])).toEqual(refused);
  });
});

describe('a wrapped defined name', () => {
  const name = (text: string, local = '') => `<definedName name="N"${local}>${text}</definedName>`;
  const names = (text: string, local = '') => wrapper(name(text, local), name(text, local));
  const wb = (definedNames: string): FixtureWorkbook => ({
    sheets: [
      { name: 'S', rows: row(5, [c('A5', 5)]) },
      { name: 'T', rows: row(1, [c('A1', 1)]) },
    ],
    definedNames,
  });

  test('a row edit that moves its reference is refused and names it', () => {
    expect(refusal(storeOf(wb(names('S!$A$5'))), [op('insertRows', 'S', 3)])).toEqual({
      code: 'structure-blocked',
      subject: 'x:definedName N',
      about: { kind: 'defined-name', name: 'N' },
    });
    expect(refusal(storeOf(wb(names('S!$A$5'))), [op('insertRows', 'S', 10)])).toBeNull();
  });

  test('a rename or a delete of the sheet that it names is refused', () => {
    for (const edit of [
      { op: 'renameSheet', sheet: 'S', name: 'New' },
      { op: 'deleteSheet', sheet: 'S' },
    ] as const)
      expect(refusal(storeOf(wb(names('S!$A$5'))), [edit])?.subject).toBe('x:definedName N');
    expect(
      refusal(storeOf(wb(names('T!$A$1'))), [{ op: 'renameSheet', sheet: 'S', name: 'New' }])
    ).toBeNull();
  });

  test('a move or a delete that changes the sheet of its local scope is refused', () => {
    const scoped = names('T!$A$1', ' localSheetId="1"');
    for (const edit of [
      { op: 'moveSheet', sheet: 'T', index: 0 },
      { op: 'deleteSheet', sheet: 'S' },
    ] as const)
      expect(refusal(storeOf(wb(scoped)), [edit])?.subject).toBe('x:definedName N');
  });

  test('a sheet add that changes the position of its local scope sheet is refused', () => {
    const scoped = names('T!$A$1', ' localSheetId="1"');
    for (const index of [0, 1])
      expect(refusal(storeOf(wb(scoped)), [{ op: 'addSheet', name: 'New', index }])).toMatchObject({
        code: 'structure-blocked',
        subject: 'x:definedName N',
      });
    // An add after the scope sheet keeps its position.
    expect(refusal(storeOf(wb(scoped)), [{ op: 'addSheet', name: 'New', index: 2 }])).toBeNull();
  });

  test('a name list inside a wrapper of the workbook element is read too', () => {
    const list = (local: string) =>
      `<definedNames><definedName name="N"${local}>T!$A$1</definedName></definedNames>`;
    const outer = (local: string): FixtureWorkbook => ({
      ...wb(''),
      afterCalcPr: wrapper(list(local), list(local)),
    });
    for (const edit of [
      { op: 'addSheet', name: 'New', index: 0 },
      { op: 'moveSheet', sheet: 'T', index: 0 },
      { op: 'deleteSheet', sheet: 'S' },
    ] as const)
      expect(refusal(storeOf(outer(' localSheetId="1"')), [edit])?.subject).toBe('x:definedName N');
    // A rename or a row edit of the sheet that such a name names is refused by the scan of
    // workbook elements, which names the wrapper.
    const edits: WorkbookOp[] = [
      { op: 'renameSheet', sheet: 'T', name: 'New' },
      op('insertRows', 'T', 1),
    ];
    for (const edit of edits)
      expect(refusal(storeOf(outer('')), [edit])).toMatchObject({
        code: 'structure-blocked',
        subject: 'mc:AlternateContent (/xl/workbook.xml)',
      });
    // A local scope that keeps its position goes on.
    expect(
      refusal(storeOf(outer(' localSheetId="1"')), [{ op: 'addSheet', name: 'New', index: 2 }])
    ).toBeNull();
  });

  test('a row edit checks a name list in a wrapper of the workbook element', () => {
    const outer = (text: string): FixtureWorkbook => {
      const list = `<definedNames>${name(text)}</definedNames>`;
      return { ...wb(''), afterCalcPr: wrapper(list, list) };
    };
    // A reference with no sheet names no sheet, so only the check of each copy refuses it.
    expect(refusal(storeOf(outer('$A$1')), [op('insertRows', 'S', 1)])).toMatchObject({
      code: 'structure-blocked',
      subject: 'x:definedName N',
    });
    // A copy that the edit keeps goes on, and a save and reopen keeps it.
    const s = storeOf(outer('T!$A$1'));
    expect(refusal(s, [op('insertRows', 'S', 1)])).toBeNull();
    const back = reopen(s);
    const kept = (xmlPartOf(back, '/xl/workbook.xml')?.root.children ?? []).filter(
      (n): n is Element => n.type === 'element' && n.local === 'AlternateContent'
    );
    expect(kept.map((n) => textOf(n))).toEqual(['T!$A$1T!$A$1']);
  });

  test('a name in a wrapper of the workbook element marks a sheet add, rename, or move for a recalculation', () => {
    // A cell reads N, so the saved values depend on it. The cell itself calls no such function.
    const outer = (text: string): FixtureWorkbook => {
      const list = `<definedNames>${name(text)}</definedNames>`;
      return {
        sheets: [
          { name: 'S', rows: '<row r="5"><c r="A5"><f>N</f><v>1</v></c></row>' },
          { name: 'T', rows: row(1, [c('A1', 1)]) },
        ],
        afterCalcPr: wrapper(list, list),
      };
    };
    const after = (text: string, edit: WorkbookOp) => {
      const s = storeOf(outer(text));
      expect(refusal(s, [edit])).toBeNull();
      return workbookIndex(reopen(s))!.fullCalcOnLoad;
    };
    // SHEET reads the position of its sheet, which the move changes.
    expect(after('SHEET(T!$A$1)', { op: 'moveSheet', sheet: 'T', index: 0 })).toBe(true);
    // A name of the new sheet had no sheet to read before the add.
    expect(after('New!$A$1', { op: 'addSheet', name: 'New', index: 2 })).toBe(true);
    expect(after('New!$A$1', { op: 'renameSheet', sheet: 'T', name: 'New' })).toBe(true);
    // A name that reads neither keeps the cached values.
    expect(after('T!$A$1', { op: 'moveSheet', sheet: 'T', index: 0 })).toBe(false);
  });

  test('a wrapped name in the name list marks a sheet move by each branch', () => {
    // The scan reads each branch on its own: the text of the two branches together, `XSHEET(...)`,
    // calls no function of the list.
    const after = (fallback: string) => {
      const s = storeOf({
        sheets: [
          { name: 'S', rows: '<row r="5"><c r="A5"><f>N</f><v>1</v></c></row>' },
          { name: 'T', rows: row(1, [c('A1', 1)]) },
        ],
        definedNames: wrapper(name('X'), name(fallback)),
      });
      expect(refusal(s, [{ op: 'moveSheet', sheet: 'T', index: 0 }])).toBeNull();
      return workbookIndex(reopen(s))!.fullCalcOnLoad;
    };
    expect(after('SHEET(T!$A$1)')).toBe(true);
    // No branch calls such a function, so the cached values stay.
    expect(after('T!$A$1')).toBe(false);
  });

  test('a name in a nested wrapper is read too', () => {
    const nested = wrapper(wrapper(name('S!$A$5'), name('S!$A$5')), name('S!$A$5'));
    expect(refusal(storeOf(wb(nested)), [op('insertRows', 'S', 3)])?.subject).toBe(
      'x:definedName N'
    );
  });
});

describe('a merge over wrapped content', () => {
  const FOREIGN = '<e:x xmlns:e="urn:example:extension"><v>2</v></e:x>';

  test('counts a wrapped formula and a value in another namespace as content it would hide', () => {
    for (const b2 of [wrappedCell('B2', 'A1'), `<c r="B2">${FOREIGN}</c>`]) {
      const s = storeOf(book(b2));
      const part = xmlPartOf(s.package, '/xl/worksheets/sheet1.xml')!;
      expect(firstValueLost(part, [0, 0, 1, 1])).toBe('B2');
      expect(refusal(s, [{ op: 'mergeCells', sheet: 'S', ref: 'A1:B2' }])?.code).toBe(
        'merge-drops-values'
      );
      // The clear that a confirmed merge makes first is refused for the cell.
      expect(refusal(s, [{ op: 'clearValues', sheet: 'S', ref: 'B2' }])?.code).toBe(
        'cell-alternate-content'
      );
    }
  });
});

// The helper reads every branch, nested wrappers included.
test('wrappedFormulas reads every branch of nested wrappers', () => {
  const s = storeOf(
    book(`<c r="B2">${wrapper(`<f>A1</f>${wrapper('<f>A2</f>', '<f>A3</f>')}`, '<f>A4</f>')}</c>`)
  );
  const el = cell(s.package, 'S', 'B2') as Element;
  expect(wrappedFormulas(el).map((f) => textOf(f))).toEqual(['A4', 'A1', 'A3', 'A2']);
});
