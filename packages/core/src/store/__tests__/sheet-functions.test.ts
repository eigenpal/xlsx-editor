// The scan for calls to the functions that read sheet facts (`sheet-functions.ts`): it finds a
// call in a cell formula, and it keeps its result for each package and set of names, so a
// second format op on the same package does not scan every formula again.

import { describe, expect, test } from 'bun:test';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx, type PackagePart, type XlsxPackage } from '../package/package.ts';
import { workbookIndex } from '../workbook/workbook-index.ts';
import {
  CELL_STATE_FUNCTIONS,
  ROW_HEIGHT_FUNCTIONS,
  workbookCallsSheetFunction,
} from '../structure/sheet-functions.ts';

function open(): XlsxPackage {
  const read = readXlsx(
    buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1, { f: 'CELL("format",B1)' })]) }] })
  );
  if (!read.ok) throw new Error('the fixture does not read');
  return read.pkg;
}

describe('workbookCallsSheetFunction', () => {
  test('finds a call, and a second call on the same package does not scan again', () => {
    const pkg = open();
    const index = workbookIndex(pkg)!;
    expect(workbookCallsSheetFunction(pkg, index, CELL_STATE_FUNCTIONS)).toBe(true);
    expect(workbookCallsSheetFunction(pkg, index, ROW_HEIGHT_FUNCTIONS)).toBe(false);
    // Remove the sheet part under the package. A scan would find no call now, so the kept
    // result shows that the second call did not scan. Only this test changes a package in
    // place: in production a package never changes, and each op returns a new package.
    const parts = pkg.parts as Map<string, PackagePart>;
    const sheet = index.sheets[0]!.partName!;
    const kept = parts.get(sheet)!;
    parts.delete(sheet);
    expect(workbookCallsSheetFunction(pkg, index, CELL_STATE_FUNCTIONS)).toBe(true);
    // A new package scans again.
    const other: XlsxPackage = { parts: new Map(parts) };
    expect(workbookCallsSheetFunction(other, index, CELL_STATE_FUNCTIONS)).toBe(false);
    parts.set(sheet, kept);
  });
});
