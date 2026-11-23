// Sheet names that compare without case (design D1 of the change `workbook-structure-editing`):
// the fold keeps the length of a name, so `Maße` and `MASSE` are two names, and a name that a
// file gives to two sheets without case refuses the ops that cannot tell them apart.

import { describe, expect, test } from 'bun:test';
import { c, row } from '../../../../../scripts/fixtures/builder.ts';
import { sheetByName, sheetNameKey, workbookIndex } from '../workbook/workbook-index.ts';
import type { XlsxPackage } from '../package/package.ts';
import type { WorkbookDocumentStore } from '../store/document-store.ts';
import type { WorkbookOp } from '../store/ops.ts';
import { sheetNameProblem } from '../structure/sheet-refs.ts';
import { formulaOf, storeOf } from './structure-helpers.ts';

const names = (pkg: XlsxPackage) => workbookIndex(pkg)!.sheets.map((s) => s.name);

const rejection = (s: WorkbookDocumentStore, ops: readonly WorkbookOp[]) => {
  const before = s.package;
  const r = s.transact(ops);
  expect(s.package).toBe(before);
  return r.ok ? null : { code: r.rejection.code, subject: r.rejection.subject };
};

describe('the sheet name fold', () => {
  test('keeps the length of a name', () => {
    expect(sheetNameKey('Maße')).toBe('MAßE');
    expect(sheetNameKey('MASSE')).toBe('MASSE');
    expect(sheetNameKey('Données')).toBe('DONNÉES');
    expect(sheetNameKey('ŉ')).toBe('ŉ');
  });

  test('takes Maße and MASSE as two names', () => {
    const s = storeOf({ sheets: [{ name: 'Maße', rows: '' }] });
    const index = workbookIndex(s.package)!;
    expect(sheetNameProblem('MASSE', index)).toBeUndefined();
    expect(sheetNameProblem('MAßE', index)?.code).toBe('sheet-name-taken');
    expect(s.transact([{ op: 'addSheet', name: 'MASSE', index: 1 }]).ok).toBe(true);
  });

  test('refuses a new name that is equal to another after NFC', () => {
    const cafe = storeOf({
      sheets: [
        { name: 'Caf\u00e9', rows: '' },
        { name: 'B', rows: '' },
      ],
    });
    // The subject is the name of the sheet that has the name, as its tab shows it.
    expect(rejection(cafe, [{ op: 'addSheet', name: 'Cafe\u0301', index: 2 }])).toEqual({
      code: 'sheet-name-taken',
      subject: 'Caf\u00e9',
    });
    expect(rejection(cafe, [{ op: 'renameSheet', sheet: 'B', name: 'CAFE\u0301' }])?.code).toBe(
      'sheet-name-taken'
    );
    const k = storeOf({ sheets: [{ name: 'k', rows: '' }] });
    expect(rejection(k, [{ op: 'addSheet', name: '\u212a', index: 1 }])?.code).toBe(
      'sheet-name-taken'
    );
    // A rename of a sheet to a canonical equivalent of its own name is applied.
    expect(cafe.transact([{ op: 'renameSheet', sheet: 'Caf\u00e9', name: 'Cafe\u0301' }]).ok).toBe(
      true
    );
    expect(names(cafe.package)).toEqual(['Cafe\u0301', 'B']);
  });

  const masse = (formula: string) =>
    storeOf({
      sheets: [
        { name: 'Maße', rows: row(1, [c('A1', 1)]) },
        { name: 'MASSE', rows: row(1, [c('A1', 2)]) },
        { name: 'T', rows: row(1, [c('A1', 2, { f: formula })]) },
      ],
    });

  test('a rename of Maße rewrites only its own prefixes', () => {
    const s = masse('MASSE!A1+Maße!A1');
    expect(s.transact([{ op: 'renameSheet', sheet: 'Maße', name: 'X' }]).ok).toBe(true);
    expect(names(s.package)).toEqual(['X', 'MASSE', 'T']);
    expect(formulaOf(s.package, 'T', 'A1')).toBe('MASSE!A1+X!A1');
  });

  test('a rename of MASSE acts on MASSE', () => {
    const s = masse('MASSE!A1+Maße!A1');
    expect(s.transact([{ op: 'renameSheet', sheet: 'MASSE', name: 'Y' }]).ok).toBe(true);
    expect(names(s.package)).toEqual(['Maße', 'Y', 'T']);
    expect(formulaOf(s.package, 'T', 'A1')).toBe('Y!A1+Maße!A1');
  });

  test('a delete of MASSE keeps the references to Maße', () => {
    const s = masse('Maße!A1');
    expect(s.transact([{ op: 'deleteSheet', sheet: 'MASSE' }]).ok).toBe(true);
    expect(names(s.package)).toEqual(['Maße', 'T']);
    expect(formulaOf(s.package, 'T', 'A1')).toBe('Maße!A1');
  });
});

describe('two sheets of a file with the same name without case', () => {
  const twins = () =>
    storeOf({
      sheets: [
        { name: 'Data', rows: row(1, [c('A1', 1)]) },
        { name: 'DATA', rows: row(1, [c('A1', 2)]) },
        { name: 'T', rows: row(1, [c('A1', 2, { f: 'DATA!A1' })]) },
      ],
    });

  test('an exact name finds its sheet, and a folded one finds none', () => {
    const index = workbookIndex(twins().package)!;
    expect(sheetByName(index, 'Data')?.index).toBe(0);
    expect(sheetByName(index, 'DATA')?.index).toBe(1);
    expect(sheetByName(index, 'data')).toBeUndefined();
  });

  test('a rename and a delete are refused, with the other sheet as the subject', () => {
    const s = twins();
    expect(rejection(s, [{ op: 'renameSheet', sheet: 'Data', name: 'X' }])).toEqual({
      code: 'sheet-name-ambiguous',
      subject: 'DATA',
    });
    expect(rejection(s, [{ op: 'deleteSheet', sheet: 'DATA' }])).toEqual({
      code: 'sheet-name-ambiguous',
      subject: 'Data',
    });
    expect(rejection(s, [{ op: 'renameSheet', sheet: 'data', name: 'X' }])).toEqual({
      code: 'unknown-sheet',
      subject: undefined,
    });
  });

  test('a row insert is refused, and a move and a cell edit are not', () => {
    const s = twins();
    expect(rejection(s, [{ op: 'insertRows', sheet: 'Data', start: 1, count: 1 }])).toEqual({
      code: 'sheet-name-ambiguous',
      subject: 'DATA',
    });
    const value = { type: 'number', value: 5 } as const;
    expect(s.transact([{ op: 'setCellValue', sheet: 'DATA', ref: 'B1', value }]).ok).toBe(true);
    expect(s.transact([{ op: 'moveSheet', sheet: 'T', index: 0 }]).ok).toBe(true);
    expect(names(s.package)).toEqual(['T', 'Data', 'DATA']);
  });

  test('a rename of a twin that no reference names is applied and ends the ambiguity', () => {
    const s = storeOf({
      sheets: [
        { name: 'Data', rows: row(1, [c('A1', 1)]) },
        { name: 'DATA', rows: row(1, [c('A1', 2)]) },
        { name: 'T', rows: row(1, [c('A1', 2, { f: 'Other!A1' })]) },
        { name: 'Other', rows: '' },
      ],
      definedNames: '<definedName name="n">Other!$A$1</definedName>',
    });
    expect(s.transact([{ op: 'renameSheet', sheet: 'Data', name: 'X' }]).ok).toBe(true);
    expect(names(s.package)).toEqual(['X', 'DATA', 'T', 'Other']);
    expect(s.transact([{ op: 'insertRows', sheet: 'DATA', start: 1, count: 1 }]).ok).toBe(true);
  });

  test('a defined name that names a twin refuses the rename', () => {
    const s = storeOf({
      sheets: [
        { name: 'Data', rows: row(1, [c('A1', 1)]) },
        { name: 'DATA', rows: row(1, [c('A1', 2)]) },
      ],
      definedNames: '<definedName name="n">DATA!$A$1</definedName>',
    });
    expect(rejection(s, [{ op: 'renameSheet', sheet: 'DATA', name: 'X' }])).toEqual({
      code: 'sheet-name-ambiguous',
      subject: 'Data',
    });
  });
});
