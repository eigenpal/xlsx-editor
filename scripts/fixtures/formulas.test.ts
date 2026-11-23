// The formulas fixture: the save check's saved copy keeps the calculation chain, and its edited
// copy (B2 = 100, the shared formula master) unshares the children (design D4), drops the chain
// (design D5), and marks the workbook stale (design D3).

import { describe, expect, test } from 'bun:test';
import {
  attr,
  cellKey,
  child,
  contentTypeOf,
  parseCellRef,
  readXlsx,
  relationshipsOf,
  sheetIndex,
  textOf,
  workbookIndex,
  xmlPartOf,
  type XlsxPackage,
} from '../../packages/core/src/store/index.ts';
import { fixtureCopies } from '../save-check/records.ts';
import { FIXTURES } from './catalog.ts';

const CHAIN = '/xl/calcChain.xml';
const CHAIN_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml';
const CHAIN_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain';

function open(bytes: Uint8Array): XlsxPackage {
  const read = readXlsx(bytes);
  if (!read.ok) throw new Error(JSON.stringify(read.error));
  return read.pkg;
}

const chain = (pkg: XlsxPackage) => ({
  part: pkg.parts.has(CHAIN),
  rel: relationshipsOf(pkg, '/xl/workbook.xml').some((r) => r.type === CHAIN_REL),
  type: contentTypeOf(pkg, CHAIN) === CHAIN_TYPE,
});

function cell(pkg: XlsxPackage, ref: string) {
  const a = parseCellRef(ref)!;
  return sheetIndex(xmlPartOf(pkg, '/xl/worksheets/sheet1.xml')!).cells.get(cellKey(a.row, a.col))
    ?.el;
}

const fixture = FIXTURES.find((f) => f.name === 'formulas')!;
const copies = fixtureCopies(fixture);

describe('the wrapped-only-formula fixture', () => {
  const only = fixtureCopies(FIXTURES.find((f) => f.name === 'wrapped-only-formula')!);

  test('its only formula is wrapped, so the sheet index counts no formula', () => {
    const part = xmlPartOf(open(only.saved), '/xl/worksheets/sheet1.xml')!;
    expect(sheetIndex(part).formulas).toBe(0);
    expect(workbookIndex(open(only.saved))!.fullCalcOnLoad).toBe(false);
  });

  test('the edited copy is marked stale, which only the formula scan can do', () => {
    expect(workbookIndex(open(only.edited))!.fullCalcOnLoad).toBe(true);
  });
});

describe('the formulas fixture', () => {
  test('the saved copy keeps the calculation chain and the shared formula', () => {
    const pkg = open(copies.saved);
    expect(chain(pkg)).toEqual({ part: true, rel: true, type: true });
    expect(workbookIndex(pkg)!.fullCalcOnLoad).toBe(false);
    const b3 = child(cell(pkg, 'B3')!, 'x:f')!;
    expect([attr(b3, 't'), attr(b3, 'si')]).toEqual(['shared', '0']);
  });

  test('the edited copy drops the chain, marks the workbook stale, and unshares B3:B5', () => {
    const pkg = open(copies.edited);
    expect(chain(pkg)).toEqual({ part: false, rel: false, type: false });
    expect(workbookIndex(pkg)!.fullCalcOnLoad).toBe(true);
    const b2 = cell(pkg, 'B2')!;
    expect(child(b2, 'x:f')).toBeUndefined();
    expect(textOf(child(b2, 'x:v'))).toBe('100');
    for (const [ref, text, cached] of [
      ['B3', 'A3*2', '10'],
      ['B4', 'A4*2', '14'],
      ['B5', 'A5*2', '22'],
    ] as const) {
      const f = child(cell(pkg, ref)!, 'x:f')!;
      expect([ref, textOf(f), attr(f, 't'), attr(f, 'si')]).toEqual([
        ref,
        text,
        undefined,
        undefined,
      ]);
      expect(textOf(child(cell(pkg, ref)!, 'x:v'))).toBe(cached);
    }
  });
});
