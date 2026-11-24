// Structural edits and the code that the format edits added: the calculation side effects of
// designs D3 and D5 (a structural edit marks a workbook with formulas stale and drops the
// calculation chain; its undo and redo set neither again), and the inherited formats of rows
// and columns, which move with their rows and columns.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { c, row } from '../../../../../scripts/fixtures/builder.ts';
import { fixtureFile } from '../../../../../scripts/fixtures/catalog.ts';
import { partOf, readXlsx, xmlPartOf, type XlsxPackage } from '../package/package.ts';
import { attr, child } from '../package/query.ts';
import { sheetByName, workbookIndex } from '../workbook/workbook-index.ts';
import { sheetIndex } from '../workbook/sheet-index.ts';
import { inheritedFormat } from '../workbook/inherited-format.ts';
import { WorkbookDocumentStore } from '../store/document-store.ts';
import { op, storeOf } from './structure-helpers.ts';
import { applyHistory } from './history-helpers.ts';

const FIXTURES = join(import.meta.dir, '../../../../../fixtures');

function openFixture(name: string): XlsxPackage {
  const read = readXlsx(readFileSync(join(FIXTURES, fixtureFile(name))));
  if (!read.ok) throw new Error(JSON.stringify(read.error));
  return read.pkg;
}

const fullCalc = (pkg: XlsxPackage) =>
  attr(child(xmlPartOf(pkg, '/xl/workbook.xml')!.root, 'x:calcPr'), 'fullCalcOnLoad');

describe('structural edits and the calculation state', () => {
  test('an insertion in a workbook with formulas marks it stale and drops the chain', () => {
    const store = new WorkbookDocumentStore(openFixture('formulas'));
    expect(partOf(store.package, '/xl/calcChain.xml')).toBeDefined();
    expect(fullCalc(store.package)).toBeUndefined();
    const r = store.transact([op('insertRows', 'Formulas', 1)]);
    expect(r.ok).toBe(true);
    expect(fullCalc(store.package)).toBe('1');
    expect(partOf(store.package, '/xl/calcChain.xml')).toBeUndefined();
  });

  test('a deletion of columns does the same', () => {
    const store = new WorkbookDocumentStore(openFixture('formulas'));
    expect(store.transact([op('deleteColumns', 'Formulas', 1)]).ok).toBe(true);
    expect(fullCalc(store.package)).toBe('1');
    expect(partOf(store.package, '/xl/calcChain.xml')).toBeUndefined();
  });

  test('undo and redo keep the stale flag and set nothing more', () => {
    const store = new WorkbookDocumentStore(openFixture('formulas'));
    expect(store.transact([op('insertColumns', 'Formulas', 1)]).ok).toBe(true);
    expect(store.undo().ok).toBe(true);
    expect(fullCalc(store.package)).toBe('1');
    expect(store.redo().ok).toBe(true);
    expect(fullCalc(store.package)).toBe('1');
  });

  test('restoreParts reports no value or formula change', () => {
    // A restore of one worksheet part in a workbook with formulas and a chain: neither the
    // stale flag nor the chain changes, because only the edit that it undoes sets them.
    const pkg = openFixture('formulas');
    const partName = sheetByName(workbookIndex(pkg)!, 'Formulas')!.partName!;
    const part = pkg.parts.get(partName)!;
    const r = applyHistory(pkg, [{ op: 'restoreParts', parts: [{ partName, part }] }]);
    if (!r.ok) throw new Error(r.rejection.detail);
    expect(fullCalc(r.pkg)).toBeUndefined();
    expect(partOf(r.pkg, '/xl/calcChain.xml')).toBeDefined();
  });
});

describe('structural edits and inherited formats', () => {
  const wb = () => ({
    sheets: [
      {
        name: 'S',
        cols: '<cols><col min="2" max="2" width="12" style="2" customWidth="1"/></cols>',
        rows:
          row(1, [c('A1', 1)]) + '<row r="2" s="1" customFormat="1"><c r="A2"><v>2</v></c></row>',
      },
    ],
  });
  const formatAt = (pkg: XlsxPackage, r: number, col: number) =>
    inheritedFormat(
      sheetIndex(xmlPartOf(pkg, sheetByName(workbookIndex(pkg)!, 'S')!.partName)!),
      r,
      col
    );

  test('a row format moves with its row', () => {
    const s = storeOf(wb());
    expect(formatAt(s.package, 1, 3)).toBe(1);
    expect(s.transact([op('insertRows', 'S', 1, 2)]).ok).toBe(true);
    expect(formatAt(s.package, 1, 3)).toBe(0);
    expect(formatAt(s.package, 3, 3)).toBe(1);
    expect(s.undo().ok).toBe(true);
    expect(formatAt(s.package, 1, 3)).toBe(1);
  });

  test('a column format moves with its column, and goes with it', () => {
    const s = storeOf(wb());
    expect(formatAt(s.package, 5, 1)).toBe(2);
    expect(s.transact([op('insertColumns', 'S', 1)]).ok).toBe(true);
    expect(formatAt(s.package, 5, 1)).toBe(0);
    expect(formatAt(s.package, 5, 2)).toBe(2);
    expect(s.transact([op('deleteColumns', 'S', 3)]).ok).toBe(true);
    expect(formatAt(s.package, 5, 2)).toBe(0);
  });
});
