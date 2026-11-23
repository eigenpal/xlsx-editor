// The store lowers a number format code to its built-in id of ECMA-376 Part 1, 18.8.30, only
// when the workbook gives that id no other code: an id never shows a code other than the one
// that the command wrote.

import { describe, expect, test } from 'bun:test';
import {
  buildXlsx,
  c,
  row,
  type FixtureWorkbook,
} from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx, xmlPartOf, type XlsxPackage } from '../package/package.ts';
import { attr, children, child } from '../package/query.ts';
import { sheetByName, styleLists, workbookIndex } from '../workbook/workbook-index.ts';
import { cellKey, sheetIndex } from '../workbook/sheet-index.ts';
import { WorkbookDocumentStore } from '../store/document-store.ts';
import type { FormatPatch } from '../store/format-patch.ts';

function open(wb: FixtureWorkbook): XlsxPackage {
  const read = readXlsx(buildXlsx(wb));
  if (!read.ok) throw new Error(JSON.stringify(read.error));
  return read.pkg;
}

const sheet = (numFmts?: readonly string[]): FixtureWorkbook => ({
  sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }],
  ...(numFmts ? { styles: { numFmts } } : {}),
});

/** The `numFmtId` of A1's cell format, and the code that the styles part gives that id. */
function formatOfA1(pkg: XlsxPackage): { id: string | undefined; code: string | undefined } {
  const index = workbookIndex(pkg)!;
  const part = xmlPartOf(pkg, sheetByName(index, 'S')!.partName)!;
  const cell = sheetIndex(part).cells.get(cellKey(0, 0))!.el;
  const xf = styleLists(pkg, index).cellXfs[Number(attr(cell, 's') ?? '0')]!;
  const id = attr(xf, 'numFmtId');
  const root = xmlPartOf(pkg, index.stylesPart)!.root;
  const record = children(child(root, 'x:numFmts'), 'x:numFmt').find(
    (el) => attr(el, 'numFmtId') === id
  );
  return { id, code: record ? attr(record, 'formatCode') : undefined };
}

function apply(pkg: XlsxPackage, patch: FormatPatch): XlsxPackage {
  const store = new WorkbookDocumentStore(pkg);
  const r = store.transact([{ op: 'applyFormat', sheet: 'S', ref: 'A1', patch }]);
  if (!r.ok) throw new Error(JSON.stringify(r.rejection));
  return store.package;
}

describe('number format lowering', () => {
  test('a code equal to a built-in code is written as its id', () => {
    expect(formatOfA1(apply(open(sheet()), { numberFormat: { code: '0.00E+00' } }))).toEqual({
      id: '11',
      code: undefined,
    });
  });

  test('a code keeps its code when the workbook gives the built-in id another code', () => {
    const pkg = open(sheet(['<numFmt numFmtId="2" formatCode="0.000"/>']));
    const out = formatOfA1(apply(pkg, { numberFormat: { code: '0.00' } }));
    expect(out.id).not.toBe('2');
    expect(out.code).toBe('0.00');
  });

  test.each(['" 2"', '"02"'])('a record id %s that the cascade reads as 2 redefines it', (id) => {
    const pkg = open(sheet([`<numFmt numFmtId=${id} formatCode="0.000"/>`]));
    const out = formatOfA1(apply(pkg, { numberFormat: { code: '0.00' } }));
    expect(out.id).not.toBe('2');
    expect(out.code).toBe('0.00');
  });

  test('a patch with both an id and a code is refused, not read as the id', () => {
    const store = new WorkbookDocumentStore(open(sheet()));
    // The type refuses a patch with both keys; a caller without types can still send one, so the
    // store checks it too.
    // @ts-expect-error -- a number format is an id or a code, not both
    const patch: FormatPatch = { numberFormat: { id: 2, code: '0.000' } };
    const r = store.transact([{ op: 'applyFormat', sheet: 'S', ref: 'A1', patch }]);
    expect(r.ok ? null : r.rejection.code).toBe('bad-format');
  });

  test('a number format with neither an id nor a code is refused', () => {
    const store = new WorkbookDocumentStore(open(sheet()));
    const patch = { numberFormat: {} } as FormatPatch;
    const r = store.transact([{ op: 'applyFormat', sheet: 'S', ref: 'A1', patch }]);
    expect(r.ok ? null : r.rejection).toEqual({
      code: 'bad-format',
      detail: 'a number format is an id or a code',
    });
  });

  test('a key that is present but undefined does not count as given', () => {
    for (const numberFormat of [
      { id: 2, code: undefined },
      { code: '0.0', id: undefined },
    ] as FormatPatch['numberFormat'][]) {
      const store = new WorkbookDocumentStore(open(sheet()));
      const r = store.transact([
        { op: 'applyFormat', sheet: 'S', ref: 'A1', patch: { numberFormat } },
      ]);
      expect({ numberFormat, ok: r.ok }).toEqual({ numberFormat, ok: true });
    }
  });

  test('a code equal to the record of a built-in id is written as that id', () => {
    const pkg = open(sheet(['<numFmt numFmtId="2" formatCode="0.00"/>']));
    expect(formatOfA1(apply(pkg, { numberFormat: { code: '0.00' } })).id).toBe('2');
  });
});
