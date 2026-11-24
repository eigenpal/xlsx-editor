import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import {
  buildXlsx,
  c,
  row,
  type FixtureWorkbook,
} from '../../../../../scripts/fixtures/builder.ts';
import { FIXTURES, fixtureEditOps, fixtureFile } from '../../../../../scripts/fixtures/catalog.ts';
import { packageFingerprints, readXlsx, xmlPartOf, type XlsxPackage } from '../package/package.ts';
import { writeXlsx } from '../package-write/write.ts';
import { attr, child, textOf } from '../package/query.ts';
import { serializeXmlPart } from '../package-write/serialize.ts';
import { sharedStringItems, sheetByName, workbookIndex } from '../workbook/workbook-index.ts';
import { cellKey, sheetIndex } from '../workbook/sheet-index.ts';
import { applySaveRules, saveXlsx, workbookOracles } from '../package-write/save.ts';
import { decodeXstring, encodeXstring } from '../workbook/xstring.ts';
import { applyHistory } from './history-helpers.ts';
import { touchedPartViolation } from '../store/part-checks.ts';
import { WorkbookDocumentStore } from '../store/document-store.ts';
import type { CellInput, WorkbookOp } from '../store/ops.ts';

const FIXTURE_DIR = `${import.meta.dir}/../../../../../fixtures`;

function open(wb: FixtureWorkbook): XlsxPackage {
  const read = readXlsx(buildXlsx(wb));
  if (!read.ok) throw new Error(JSON.stringify(read.error));
  return read.pkg;
}

function cellOf(pkg: XlsxPackage, sheet: string, row0: number, col0: number) {
  const index = workbookIndex(pkg)!;
  const part = xmlPartOf(pkg, sheetByName(index, sheet)!.partName)!;
  return sheetIndex(part).cells.get(cellKey(row0, col0))?.el;
}

const set = (ref: string, value: CellInput, sheet = 'S'): WorkbookOp => ({
  op: 'setCellValue',
  sheet,
  ref,
  value,
});

const base = (): FixtureWorkbook => ({
  strings: ['old'],
  sheets: [
    {
      name: 'S',
      rows: [row(1, [c('A1', 0, { t: 's' }), c('C1', 3)]), row(4, [c('B4', 1)])].join(''),
    },
  ],
});

describe('WorkbookDocumentStore.transact', () => {
  test('sets numbers, booleans, and strings, and keeps rows and cells in order', () => {
    const store = new WorkbookDocumentStore(open(base()));
    const r = store.transact([
      set('B1', { type: 'number', value: 2.5 }),
      set('A3', { type: 'boolean', value: true }),
      set('D4', { type: 'string', value: 'new text' }),
    ]);
    expect(r.ok).toBe(true);
    const pkg = store.package;
    expect(textOf(child(cellOf(pkg, 'S', 0, 1), 'x:v'))).toBe('2.5');
    expect(attr(cellOf(pkg, 'S', 2, 0), 't')).toBe('b');
    const d4 = cellOf(pkg, 'S', 3, 3)!;
    expect(attr(d4, 't')).toBe('s');
    const strings = sharedStringItems(pkg, workbookIndex(pkg)!);
    expect(textOf(strings[Number(textOf(child(d4, 'x:v')))])).toBe('new text');
    const index = sheetIndex(xmlPartOf(pkg, '/xl/worksheets/sheet1.xml')!);
    expect(index.problems).toEqual([]);
    expect(index.order).toEqual([
      cellKey(0, 0),
      cellKey(0, 1),
      cellKey(0, 2),
      cellKey(2, 0),
      cellKey(3, 1),
      cellKey(3, 3),
    ]);
  });

  test('an existing shared string is reused, and the table is never compacted', () => {
    const store = new WorkbookDocumentStore(open(base()));
    store.transact([
      set('B1', { type: 'string', value: 'old' }),
      set('A1', { type: 'number', value: 1 }),
    ]);
    const pkg = store.package;
    expect(textOf(child(cellOf(pkg, 'S', 0, 1), 'x:v'))).toBe('0');
    expect(sharedStringItems(pkg, workbookIndex(pkg)!)).toHaveLength(1);
  });

  test('a string edit creates the shared string table when the workbook has none', () => {
    const store = new WorkbookDocumentStore(open({ sheets: [{ name: 'S', rows: '' }] }));
    expect(store.transact([set('A1', { type: 'string', value: '  padded  ' })]).ok).toBe(true);
    const saved = readXlsx(saveXlsx(store.package));
    if (!saved.ok) throw new Error('reopen');
    const items = sharedStringItems(saved.pkg, workbookIndex(saved.pkg)!);
    expect(items.map((i) => textOf(i))).toEqual(['  padded  ']);
    expect(attr(xmlPartOf(saved.pkg, '/xl/sharedStrings.xml')!.root, 'count')).toBeUndefined();
  });

  test('undo of a string edit that created the table removes the part, relationship, and override', () => {
    for (const ops of [
      [set('A1', { type: 'string', value: 'one' })],
      [set('A1', { type: 'string', value: 'one' }), set('B1', { type: 'string', value: 'two' })],
    ]) {
      const before = open({ sheets: [{ name: 'S', rows: '' }] });
      const store = new WorkbookDocumentStore(before);
      expect(store.transact(ops).ok).toBe(true);
      expect(workbookIndex(store.package)!.sharedStringsPart).toBe('/xl/sharedStrings.xml');
      expect(store.undo().ok).toBe(true);
      const pkg = store.package;
      expect(pkg.parts.has('/xl/sharedStrings.xml')).toBe(false);
      expect(workbookIndex(pkg)!.sharedStringsPart).toBeUndefined();
      // The relationships and the content types are the parts of before the edit.
      expect(pkg.parts.get('/xl/_rels/workbook.xml.rels')).toBe(
        before.parts.get('/xl/_rels/workbook.xml.rels')
      );
      expect(pkg.parts.get('/[Content_Types].xml')).toBe(before.parts.get('/[Content_Types].xml'));
      // Redo puts the table back with the string.
      expect(store.redo().ok).toBe(true);
      expect(sharedStringItems(store.package, workbookIndex(store.package)!)[0]).toBeDefined();
    }
  });

  test('after an undo, a save keeps added text in a table that the workbook had, and drops a table that the edit created', () => {
    const fixture = (name: string) => {
      const read = readXlsx(new Uint8Array(readFileSync(`${FIXTURE_DIR}/${fixtureFile(name)}`)));
      if (!read.ok) throw new Error(name);
      return read.pkg;
    };
    const reopened = (store: WorkbookDocumentStore) => {
      const read = readXlsx(saveXlsx(store.package));
      if (!read.ok) throw new Error('reopen');
      return read.pkg;
    };
    // formulas.xlsx has no shared string table: the undo removes the table that the edit made.
    const created = new WorkbookDocumentStore(fixture('formulas'));
    expect(created.transact([set('H1', { type: 'string', value: 'Added' }, 'Formulas')]).ok).toBe(
      true
    );
    expect(created.undo().ok).toBe(true);
    const plain = reopened(created);
    expect(workbookIndex(plain)!.sharedStringsPart).toBeUndefined();
    expect([...plain.parts.keys()].some((p) => p.includes('sharedStrings'))).toBe(false);
    // basic-values.xlsx has a table: the added text stays in it after the undo.
    const grown = new WorkbookDocumentStore(fixture('basic-values'));
    expect(grown.transact([set('H1', { type: 'string', value: 'Added' }, 'Values')]).ok).toBe(true);
    expect(grown.undo().ok).toBe(true);
    const kept = reopened(grown);
    const items = sharedStringItems(kept, workbookIndex(kept)!).map((i) => textOf(i));
    expect(items).toContain('Added');
  });

  test('a text longer than 32,767 UTF-16 code units is refused with text-too-long, alone or in a batch', () => {
    const store = new WorkbookDocumentStore(open({ sheets: [{ name: 'S', rows: '' }] }));
    const long = set('A1', { type: 'string', value: 'x'.repeat(40_000) });
    for (const ops of [[long], [set('B1', { type: 'string', value: 'ok' }), long]]) {
      const r = store.transact(ops);
      expect(r.ok ? null : r.rejection.code).toBe('text-too-long');
    }
    expect(store.revision).toBe(0);
    expect(store.transact([set('A1', { type: 'string', value: 'x'.repeat(32_767) })]).ok).toBe(
      true
    );
  });

  test('the part check reads each new shared string item, also one before x:extLst', () => {
    const before = open({
      sheets: [{ name: 'S', rows: '' }],
      strings: ['a'],
      rawStrings: ['<extLst><ext uri="urn:x"/></extLst>'],
    });
    const name = workbookIndex(before)!.sharedStringsPart!;
    const part = xmlPartOf(before, name)!;
    const [first, ext] = part.root.children.filter((n) => n.type === 'element');
    if (first?.type !== 'element' || ext?.type !== 'element') throw new Error('no items');
    const si = first.children[0];
    if (si?.type !== 'element') throw new Error('no text');
    const text = { type: 'text' as const, id: `${si.id}-bad` as typeof si.id, value: '\u0001' };
    const bad = { ...first, children: [{ ...si, children: [text] }] };
    const parts = new Map(before.parts);
    parts.set(name, {
      kind: 'xml',
      partName: name,
      xml: { ...part, root: { ...part.root, children: [first, bad, ext] } },
    });
    const after = { ...before, parts } as XlsxPackage;
    expect(touchedPartViolation(before, after, [name])).toContain('XML cannot hold');
  });

  test('an undo or a redo that leaves the content as it was publishes no revision', () => {
    const store = new WorkbookDocumentStore(
      open({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] })
    );
    // Two writes that end where they started, kept apart so the batch path does not merge them.
    const r = store.transact([
      set('A1', { type: 'number', value: 9 }),
      { op: 'clearValues', sheet: 'S', ref: 'Z99' },
      set('A1', { type: 'number', value: 1 }),
    ]);
    expect(r.ok).toBe(true);
    const revision = store.revision;
    let published = 0;
    store.subscribe(() => (published += 1));
    expect(store.undo().ok).toBe(true);
    expect(store.redo().ok).toBe(true);
    expect(store.revision).toBe(revision);
    expect(published).toBe(0);
    expect(store.canUndo).toBe(true);
  });

  test('undo and redo with an empty history name it', () => {
    const store = new WorkbookDocumentStore(open({ sheets: [{ name: 'S', rows: '' }] }));
    const undo = store.undo();
    const redo = store.redo();
    expect(undo.ok ? null : undo.rejection.code).toBe('nothing-to-undo');
    expect(redo.ok ? null : redo.rejection.code).toBe('nothing-to-redo');
  });

  test('a clear that leaves a bare row drops it, alone, in a batch, or as a range', () => {
    const rows =
      row(1, [c('A1', 1)]) +
      row(2, [c('A2', 2)]) +
      row(3, [c('A3', 3)], 'ht="30" customHeight="1"');
    const store = new WorkbookDocumentStore(open({ sheets: [{ name: 'S', rows }] }));
    const rowsOf = () => {
      const part = xmlPartOf(store.package, '/xl/worksheets/sheet1.xml')!;
      return [...sheetIndex(part).rows.keys()];
    };
    expect(store.transact([set('A2', { type: 'empty' })]).ok).toBe(true);
    expect(rowsOf()).toEqual([0, 2]);
    expect(store.undo().ok).toBe(true);
    expect(rowsOf()).toEqual([0, 1, 2]);
    expect(store.transact([set('A1', { type: 'empty' }), set('A2', { type: 'empty' })]).ok).toBe(
      true
    );
    expect(rowsOf()).toEqual([2]);
    expect(store.undo().ok).toBe(true);
    // A row with a height set by hand keeps its element.
    expect(store.transact([{ op: 'clearValues', sheet: 'S', ref: 'A1:A3' }]).ok).toBe(true);
    expect(rowsOf()).toEqual([2]);
  });

  test('a bad index the file held does not refuse an edit of another cell', () => {
    const store = new WorkbookDocumentStore(
      open({
        sheets: [
          {
            name: 'S',
            rows:
              row(1, [c('A1', 77, { t: 's' })]) + row(2, [c('A2', 1), c('B2', 2, { s: 99999 })]),
          },
        ],
        strings: ['a'],
      })
    );
    const ops: WorkbookOp[] = [
      set('A2', { type: 'number', value: 5 }),
      { op: 'clearValues', sheet: 'S', ref: 'A2' },
      { op: 'insertRows', sheet: 'S', start: 1, count: 1 },
      set('C2', { type: 'number', value: 5 }),
    ];
    for (const op of ops) {
      const r = store.transact([op]);
      expect(r.ok ? 'ok' : r.rejection).toBe('ok');
    }
    expect(store.undo().ok).toBe(true);
  });

  test('undo puts back a bad index that an edit replaced, and redo applies the edit again', () => {
    // A format index and a shared string index past their lists, each replaced by an edit.
    const store = new WorkbookDocumentStore(
      open({
        sheets: [
          { name: 'S', rows: row(1, [c('A1', 1, { s: 99 }), c('B1', 77, { t: 's' }), c('C1', 2)]) },
        ],
        strings: ['a'],
      })
    );
    const cell = (ref: string) => cellOf(store.package, 'S', 0, ref.charCodeAt(0) - 65);
    const edits: WorkbookOp[] = [
      set('C1', { type: 'number', value: 5 }),
      { op: 'clearFormats', sheet: 'S', ref: 'A1' },
      set('B1', { type: 'number', value: 6 }),
    ];
    for (const op of edits) expect(store.transact([op]).ok).toBe(true);
    expect(attr(cell('A1')!, 's')).toBeUndefined();
    // Each undo puts back the node of its own revision, with its bad index, and keeps redo.
    const undo = store.undo();
    expect(undo.ok ? 'ok' : undo.rejection).toBe('ok');
    expect(attr(cell('B1')!, 't')).toBe('s');
    expect(store.undo().ok).toBe(true);
    expect(attr(cell('A1')!, 's')).toBe('99');
    expect(store.canRedo).toBe(true);
    expect(store.undo().ok).toBe(true);
    for (let i = 0; i < 3; i += 1) expect(store.redo().ok).toBe(true);
    expect(attr(cell('A1')!, 's')).toBeUndefined();
    expect(attr(cell('B1')!, 't')).toBeUndefined();
  });

  test('a history entry keeps frozen copies of the selections that the caller supplies', () => {
    const store = new WorkbookDocumentStore(
      open({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] })
    );
    const before = {
      sheet: 'S',
      range: [0, 0, 1, 1] as [number, number, number, number],
      active: { row: 0, col: 0 },
      anchor: { row: 0, col: 0 },
      focus: { row: 1, col: 1 },
    };
    const c2 = { row: 2, col: 2 };
    const after = { sheet: 'S', range: [2, 2, 2, 2] as const, active: c2, anchor: c2, focus: c2 };
    expect(
      store.transact([set('A1', { type: 'number', value: 2 })], {
        selectionBefore: before,
        selectionAfter: after,
      }).ok
    ).toBe(true);
    // A later change of the caller's objects does not reach the history.
    before.range[0] = 9;
    before.active.row = 9;
    before.focus.col = 9;
    const forUndo = store.selectionForUndo()!;
    expect(forUndo).toEqual({
      sheet: 'S',
      range: [0, 0, 1, 1],
      active: { row: 0, col: 0 },
      anchor: { row: 0, col: 0 },
      focus: { row: 1, col: 1 },
    });
    expect(
      [forUndo, forUndo.range, forUndo.active, forUndo.anchor, forUndo.focus].every((o) =>
        Object.isFrozen(o)
      )
    ).toBe(true);
    // Each read returns its own copy.
    expect(store.selectionForUndo()).not.toBe(forUndo);
    expect(store.selectionForRedo()).toBeNull();
    expect(store.undo().ok).toBe(true);
    expect(store.selectionForRedo()).toEqual({
      sheet: 'S',
      range: [2, 2, 2, 2],
      active: c2,
      anchor: c2,
      focus: c2,
    });
    expect(store.selectionForUndo()).toBeNull();
    // An entry without selections names none, and keeps its sheet.
    expect(store.transact([set('B1', { type: 'number', value: 3 })], { sheet: 'S' }).ok).toBe(true);
    expect(store.selectionForUndo()).toBeNull();
    expect(store.sheetForUndo()).toBe('S');
    // A sheet name that is not a worksheet of the workbook is refused, and nothing commits.
    const revision = store.revision;
    for (const options of [{ sheet: 'Nope' }, { selectionBefore: { ...before, sheet: 'Nope' } }]) {
      const r = store.transact([set('C1', { type: 'number', value: 4 })], options);
      expect(r.ok ? null : r.rejection.code).toBe('unknown-sheet');
    }
    expect(store.revision).toBe(revision);
    // The selection after names a sheet after the ops, the selection before one before them,
    // and `sheet` one before or after: a rename gives the new name to the selection after.
    const rename = { op: 'renameSheet', sheet: 'S', name: 'T' } as const;
    const refused = store.transact([rename], { selectionBefore: before, selectionAfter: after });
    expect(refused.ok ? null : refused.rejection.code).toBe('unknown-sheet');
    expect(store.revision).toBe(revision);
    const renamed = { ...after, sheet: 'T' };
    expect(
      store.transact([rename], { sheet: 'S', selectionBefore: before, selectionAfter: renamed }).ok
    ).toBe(true);
    expect(store.labelForUndo()).toBeNull();
    expect(store.undo().ok).toBe(true);
    expect(store.selectionForRedo()?.sheet).toBe('T');
    // The label of an entry comes back for the undo and the redo.
    const add = { op: 'addSheet', name: 'New', index: 1 } as const;
    expect(store.transact([add], { label: 'addSheet', sheet: 'New' }).ok).toBe(true);
    expect(store.labelForUndo()).toBe('addSheet');
    expect(store.undo().ok).toBe(true);
    expect([store.labelForRedo(), store.sheetForRedo()]).toEqual(['addSheet', 'New']);
  });

  test('clearing keeps the cell format and drops an unformatted cell', () => {
    const store = new WorkbookDocumentStore(
      open({
        sheets: [{ name: 'S', rows: row(1, [c('A1', 5, { s: 1 }), c('B1', 6)]) }],
        styles: {
          cellXfs: [
            '<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
          ],
        },
      })
    );
    store.transact([set('A1', { type: 'empty' }), set('B1', { type: 'empty' })]);
    const a1 = cellOf(store.package, 'S', 0, 0)!;
    expect(attr(a1, 's')).toBe('1');
    expect(child(a1, 'x:v')).toBeUndefined();
    expect(cellOf(store.package, 'S', 0, 1)).toBeUndefined();
  });

  test('a refused op has no effect: the revision does not change', () => {
    const store = new WorkbookDocumentStore(open(base()));
    const before = store.package;
    for (const op of [
      set('XFE1', { type: 'number', value: 1 }),
      set('A1', { type: 'number', value: Number.NaN }),
      set('A1', { type: 'number', value: 1 }, 'Missing'),
    ]) {
      const r = store.transact([set('B2', { type: 'number', value: 1 }), op]);
      expect(r.ok).toBe(false);
    }
    expect(store.revision).toBe(0);
    expect(store.package).toBe(before);
  });

  test('an op that leaves a part invalid is refused with a typed reason', () => {
    const pkg = open(base());
    const bad = open({ sheets: [{ name: 'S', rows: row(2, [c('B2', 1), c('A2', 2)]) }] });
    const badRow = child(xmlPartOf(bad, '/xl/worksheets/sheet1.xml')!.root, 'x:sheetData')!
      .children[0];
    const r = applyHistory(pkg, [{ op: 'restoreRow', sheet: 'S', row: 2, node: badRow as never }]);
    expect(r.ok ? null : r.rejection.code).toBe('invalid-part');
  });

  test('a sheet with invalid cell order refuses edits and names the problem', () => {
    const store = new WorkbookDocumentStore(
      open({ sheets: [{ name: 'S', rows: row(1, [c('B1', 1), c('A1', 2)]) }] })
    );
    const r = store.transact([set('C1', { type: 'number', value: 2 })]);
    expect(r.ok ? null : r.rejection.code).toBe('invalid-sheet');
  });
});

describe('undo and redo', () => {
  test('undo restores the prior row content and redo reapplies it', () => {
    const store = new WorkbookDocumentStore(open(base()));
    const original = serializeXmlPart(xmlPartOf(store.package, '/xl/worksheets/sheet1.xml')!);
    store.transact([
      set('A1', { type: 'number', value: 9 }),
      set('Z9', { type: 'boolean', value: false }),
    ]);
    expect(store.canUndo).toBe(true);
    expect(store.undo().ok).toBe(true);
    expect(serializeXmlPart(xmlPartOf(store.package, '/xl/worksheets/sheet1.xml')!)).toBe(original);
    expect(store.redo().ok).toBe(true);
    expect(textOf(child(cellOf(store.package, 'S', 0, 0), 'x:v'))).toBe('9');
    expect(cellOf(store.package, 'S', 8, 25)).toBeDefined();
    expect(store.revision).toBe(3);
  });

  test('a system transaction stays out of the history', () => {
    const store = new WorkbookDocumentStore(open(base()));
    store.transact([set('A1', { type: 'number', value: 1 })], { origin: 'system' });
    expect(store.canUndo).toBe(false);
  });

  test('a transaction outside the history clears both stacks, so no undo leaves a bad @s', () => {
    const store = new WorkbookDocumentStore(open(base()));
    const bold = (ref: string): WorkbookOp => ({
      op: 'applyFormat',
      sheet: 'S',
      ref,
      patch: { bold: true },
    });
    expect(store.transact([bold('A1')]).ok).toBe(true);
    expect(store.canUndo).toBe(true);
    // B4 reuses the new cell format, outside the history.
    expect(store.transact([bold('B4')], { origin: 'system' }).ok).toBe(true);
    expect(store.canUndo).toBe(false);
    expect(store.canRedo).toBe(false);
    expect(attr(cellOf(store.package, 'S', 3, 1)!, 's')).toBe(
      attr(cellOf(store.package, 'S', 0, 0)!, 's')
    );
    expect(store.transact([set('C1', { type: 'number', value: 4 })]).ok).toBe(true);
    expect(store.undo().ok).toBe(true);
    expect(
      store.transact([set('C1', { type: 'number', value: 5 })], { recordsHistory: false }).ok
    ).toBe(true);
    expect(store.canUndo).toBe(false);
    expect(store.canRedo).toBe(false);
  });

  test('a throwing listener does not fail the commit', () => {
    const store = new WorkbookDocumentStore(open(base()));
    const seen: number[] = [];
    const host = globalThis as { reportError?: (e: unknown) => void };
    const original = host.reportError;
    host.reportError = () => {};
    try {
      store.subscribe(() => {
        throw new Error('listener');
      });
      store.subscribe((change) => seen.push(change.toRevision));
      expect(store.transact([set('A1', { type: 'number', value: 1 })]).ok).toBe(true);
    } finally {
      host.reportError = original;
    }
    expect(seen).toEqual([1]);
  });
});

describe('save rules and fidelity oracles', () => {
  test('counts are written from the saved content', () => {
    const pkg = open(base());
    const store = new WorkbookDocumentStore(pkg);
    store.transact([
      set('B1', { type: 'string', value: 'x' }),
      set('B2', { type: 'string', value: 'x' }),
    ]);
    const saved = applySaveRules(store.package);
    const sst = xmlPartOf(saved, '/xl/sharedStrings.xml')!.root;
    expect([attr(sst, 'count'), attr(sst, 'uniqueCount')]).toEqual(['3', '2']);
    expect(applySaveRules(pkg)).toBe(pkg);
  });

  for (const fixture of FIXTURES) {
    test(`${fixture.name}: open and save keep the fingerprint and the digest`, () => {
      const read = readXlsx(
        new Uint8Array(readFileSync(`${FIXTURE_DIR}/${fixtureFile(fixture.name)}`))
      );
      if (!read.ok) throw new Error('read');
      const oracles = workbookOracles(read.pkg);
      expect(oracles.changedParts).toEqual([]);
      expect(oracles.digest).toBe(oracles.fingerprint);
    });

    test(`${fixture.name}: the edited copy keeps its generic content and its oracles`, () => {
      const read = readXlsx(
        new Uint8Array(readFileSync(`${FIXTURE_DIR}/${fixtureFile(fixture.name)}`))
      );
      if (!read.ok) throw new Error('read');
      const store = new WorkbookDocumentStore(read.pkg);
      const { sheet } = fixture.edits[0]!;
      const strings = fixture.edits.some((e) => 'value' in e && typeof e.value === 'string');
      expect(store.transact(fixtureEditOps(fixture, read.pkg)).ok).toBe(true);
      const oracles = workbookOracles(store.package);
      expect(oracles.changedParts).toEqual([]);
      expect(oracles.digest).toBe(oracles.fingerprint);
      // Every part that the edit and the save rules do not change keeps its fingerprint.
      const original = packageFingerprints(read.pkg);
      const edited = packageFingerprints(applySaveRules(store.package));
      const touched = [...original.keys()].filter((n) => original.get(n) !== edited.get(n));
      const index = workbookIndex(read.pkg)!;
      const editedPart = sheetByName(index, sheet)!.partName!;
      // A string edit in a workbook without a shared string table also adds the table, its
      // content type, and its workbook relationship. An edit in a workbook with formulas sets
      // the stale flag in the workbook part (design D3).
      const allowed = [editedPart, index.sharedStringsPart, index.stylesPart, index.mainPart];
      if (index.sharedStringsPart === undefined && strings)
        allowed.push('/[Content_Types].xml', '/xl/_rels/workbook.xml.rels');
      // An edit that removes or changes a formula drops the calculation chain, its content
      // type, and its workbook relationship (design D5).
      if (read.pkg.parts.has('/xl/calcChain.xml'))
        allowed.push('/xl/calcChain.xml', '/[Content_Types].xml', '/xl/_rels/workbook.xml.rels');
      // A structural edit moves references on every sheet, and the anchors of the drawing of
      // the edited sheet (design D4).
      if (fixture.edits.some((e) => 'structure' in e))
        for (const name of original.keys())
          if (
            index.sheets.some((s) => s.partName === name) ||
            name.startsWith('/xl/drawings/drawing')
          )
            allowed.push(name);
      // A sheet op rewrites references to its sheet on every sheet, and adds or removes a
      // worksheet with its relationships and content type (design D1 of the change
      // `workbook-structure-editing`). It never touches `docProps/app.xml`.
      if (fixture.edits.some((e) => 'sheetOp' in e)) {
        allowed.push('/[Content_Types].xml', '/xl/_rels/workbook.xml.rels');
        for (const name of original.keys())
          if (/^\/xl\/worksheets\/(_rels\/)?sheet\d+\.xml(\.rels)?$/.test(name)) allowed.push(name);
      }
      for (const name of touched) expect(allowed).toContain(name);
      const partName = sheetByName(workbookIndex(read.pkg)!, sheet)!.partName!;
      const generic = (pkg: XlsxPackage) =>
        xmlPartOf(pkg, partName)!.root.children.flatMap((n) =>
          n.type === 'element' && n.kind === 'generic' ? [n] : []
        );
      expect(generic(store.package)).toEqual(generic(read.pkg));
    });
  }
});

describe('string escapes', () => {
  test('decode and encode _xHHHH_ escapes', () => {
    expect(decodeXstring('a_x0009_b')).toBe('a\tb');
    expect(decodeXstring('_x005F_x0041_')).toBe('_x0041_');
    expect(encodeXstring('_x0041_')).toBe('_x005F_x0041_');
    expect(encodeXstring('a\rb\u0001')).toBe('a_x000D_b_x0001_');
    for (const s of ['plain', '_x00', 'x_x0041__x', '😀', '\uD800 lone', 'tab\tline\n'])
      expect(decodeXstring(encodeXstring(s))).toBe(s);
  });
});

describe('review regressions', () => {
  test('undo restores a row that has no row number', () => {
    const store = new WorkbookDocumentStore(
      open({ sheets: [{ name: 'S', rows: '<row><c r="A1"><v>1</v></c></row>' }] })
    );
    const original = serializeXmlPart(xmlPartOf(store.package, '/xl/worksheets/sheet1.xml')!);
    store.transact([set('B1', { type: 'number', value: 2 })]);
    expect(store.undo().ok).toBe(true);
    expect(serializeXmlPart(xmlPartOf(store.package, '/xl/worksheets/sheet1.xml')!)).toBe(original);
    expect(store.canRedo).toBe(true);
  });

  test('an edit inside an array formula range is refused', () => {
    const store = new WorkbookDocumentStore(
      open({
        sheets: [
          {
            name: 'S',
            rows:
              row(1, [c('A1', 1, { f: 'B1:B2' })]).replace('<f>', '<f t="array" ref="A1:A2">') +
              row(2, [c('A2', 2)]),
          },
        ],
      })
    );
    const r = store.transact([set('A2', { type: 'number', value: 5 })]);
    expect(r.ok ? null : r.rejection.detail).toContain('A1:A2');
  });

  test('a new shared string table needs the workbook relationships', () => {
    const pkg = open({ sheets: [{ name: 'S', rows: '' }] });
    const parts = new Map(pkg.parts);
    parts.delete('/[Content_Types].xml');
    const store = new WorkbookDocumentStore({ parts });
    const r = store.transact([set('A1', { type: 'string', value: 'x' })]);
    expect(r.ok ? null : r.rejection.code).toBe('unsupported-package');
  });

  test('the oracles fail and name the part when a save loses content', () => {
    const read = readXlsx(new Uint8Array(readFileSync(`${FIXTURE_DIR}/basic-values.xlsx`)));
    if (!read.ok) throw new Error('read');
    const lossy = (pkg: XlsxPackage) => {
      const parts = new Map(pkg.parts);
      const sheet = xmlPartOf(pkg, '/xl/worksheets/sheet1.xml')!;
      const root = {
        ...sheet.root,
        attrs: [...sheet.root.attrs, { ns: '', local: 'x', value: '1' }],
      };
      parts.set(sheet.partName, { kind: 'xml', partName: sheet.partName, xml: { ...sheet, root } });
      return writeXlsx({ parts });
    };
    const oracles = workbookOracles(read.pkg, lossy);
    expect(oracles.changedParts).toEqual(['/xl/worksheets/sheet1.xml']);
    expect(oracles.digest).not.toBe(oracles.fingerprint);
  });

  const dataTable = (): FixtureWorkbook => ({
    sheets: [
      {
        name: 'S',
        rows: row(1, [c('A1', 1, { f: 'X' }), c('B1', 2)]).replace(
          '<f>X</f>',
          '<f t="dataTable" ref="A1:B2" dt2D="0" dtr="1" r1="D1"/>'
        ),
      },
    ],
  });

  test('the master of a data table formula is refused, so its definition stays', () => {
    const store = new WorkbookDocumentStore(open(dataTable()));
    const r = store.transact([set('A1', { type: 'number', value: 3 })]);
    expect(r.ok ? null : r.rejection.code).toBe('shared-formula');
  });

  test('an edit inside a data table range is refused', () => {
    const store = new WorkbookDocumentStore(open(dataTable()));
    const r = store.transact([set('B2', { type: 'number', value: 3 })]);
    expect(r.ok ? null : r.rejection.detail).toContain('A1:B2');
  });

  test('restoreRow refuses an element that is not an x:row', () => {
    const pkg = open(base());
    const rowEl = child(xmlPartOf(pkg, '/xl/worksheets/sheet1.xml')!.root, 'x:sheetData')!
      .children[0]!;
    const foreign = { ...rowEl, ns: 'urn:example:foreign', children: [] };
    const r = applyHistory(pkg, [{ op: 'restoreRow', sheet: 'S', row: 1, node: foreign as never }]);
    expect(r.ok).toBe(false);
  });

  test('a shared string table that did not parse refuses a string edit early', () => {
    const pkg = open(base());
    const parts = new Map(pkg.parts);
    const sst = '/xl/sharedStrings.xml';
    parts.set(sst, { kind: 'binary', partName: sst, bytes: new Uint8Array([60]) });
    const store = new WorkbookDocumentStore({ parts });
    const r = store.transact([set('B2', { type: 'string', value: 'x' })]);
    expect(r.ok ? null : r.rejection.code).toBe('unsupported-package');
    expect(store.transact([set('B2', { type: 'number', value: 1 })]).ok).toBe(true);
  });
});
