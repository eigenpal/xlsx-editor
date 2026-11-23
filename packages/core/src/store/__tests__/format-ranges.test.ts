// A run of `applyFormat` ops on one sheet applies in one pass (`applyFormatRanges`), each op
// with its own patch: the package equals the one that the ops give one by one, also the order
// of new style records and the refusal; ranges that overlap fall back to one op at a time; and
// one undo puts the package back.

import { describe, expect, test } from 'bun:test';
import {
  buildXlsx,
  c,
  row,
  type FixtureWorkbook,
} from '../../../../../scripts/fixtures/builder.ts';
import {
  canonicalPackageFingerprint,
  readXlsx,
  xmlPartOf,
  type XlsxPackage,
} from '../package/package.ts';
import { attr } from '../package/query.ts';
import { sheetByName, workbookIndex } from '../workbook/workbook-index.ts';
import { cellKey, sheetIndex } from '../workbook/sheet-index.ts';
import { WorkbookDocumentStore } from '../store/document-store.ts';
import type { FormatPatch } from '../store/format-patch.ts';
import type { WorkbookOp } from '../store/ops.ts';
import { MAX_STYLE_RECORDS } from '../store/style-records.ts';

function open(wb: FixtureWorkbook): XlsxPackage {
  const read = readXlsx(buildXlsx(wb));
  if (!read.ok) throw new Error(JSON.stringify(read.error));
  return read.pkg;
}

/** Cells inside and outside the ranges, a row format, a column format, and a bold format. */
const book = (): FixtureWorkbook => ({
  sheets: [
    {
      name: 'S',
      dimension: 'A1:F3',
      rows: [
        row(1, [c('A1', 1), c('B1', 2, { s: 1 }), c('D1', 'x'), c('F1', 3)], ' ht="30"'),
        row(2, [c('B2', 4), c('E2', 5)]),
        row(3, [c('C3', 6)], ' s="1" customFormat="1"'),
      ].join(''),
      cols: '<cols><col min="4" max="4" width="9" style="1" customWidth="1"/></cols>',
    },
  ],
  styles: {
    fonts: ['<font><b/><sz val="11"/><name val="Calibri"/></font>'],
    cellXfs: ['<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'],
  },
});

const format = (ref: string, patch: FormatPatch): WorkbookOp => ({
  op: 'applyFormat',
  sheet: 'S',
  ref,
  patch,
});

/** The fingerprint after the ops in one transaction, and after one transaction for each op. */
function both(ops: readonly WorkbookOp[]) {
  const batched = new WorkbookDocumentStore(open(book()));
  const before = canonicalPackageFingerprint(batched.package);
  expect(batched.transact(ops).ok).toBe(true);
  const single = new WorkbookDocumentStore(open(book()));
  for (const op of ops) expect(single.transact([op]).ok).toBe(true);
  return { batched, single, before };
}

/** `x:row/@ht` of the 0-based row of sheet S. */
function rowHeight(pkg: XlsxPackage, r: number): string | undefined {
  const part = xmlPartOf(pkg, sheetByName(workbookIndex(pkg)!, 'S')!.partName)!;
  const el = sheetIndex(part).rows.get(r);
  return el ? attr(el, 'ht') : undefined;
}

const RANGES = ['A1', 'C1:D1', 'F1:G2', 'B2', 'A3:D4', 'H8'];

describe('applyFormat runs', () => {
  test('a run with one number format equals the ops one by one, and one undo restores it', () => {
    const { batched, single, before } = both(
      RANGES.map((r) => format(r, { numberFormat: { id: 9 } }))
    );
    expect(canonicalPackageFingerprint(batched.package)).toBe(
      canonicalPackageFingerprint(single.package)
    );
    expect(batched.undo().ok).toBe(true);
    expect(canonicalPackageFingerprint(batched.package)).toBe(before);
  });

  test('outline borders follow the edges of each range', () => {
    const patch: FormatPatch = { borders: { edges: 'outline' } };
    const { batched, single } = both(RANGES.map((r) => format(r, patch)));
    expect(canonicalPackageFingerprint(batched.package)).toBe(
      canonicalPackageFingerprint(single.package)
    );
  });

  test('ranges that overlap apply one by one', () => {
    const ops = [
      format('A1:C2', { borders: { edges: 'outline' } }),
      format('B2:D3', { borders: { edges: 'outline' } }),
      format('A1', { numberFormat: { code: '"$"#,##0.00' } }),
      format('E5', { numberFormat: { code: '"$"#,##0.00' } }),
    ];
    const { batched, single } = both(ops);
    expect(canonicalPackageFingerprint(batched.package)).toBe(
      canonicalPackageFingerprint(single.package)
    );
  });

  test('a checkerboard of two formats on a large sheet costs the sheet once', () => {
    // 1000 rows of 50 cells. One op for each range costs the sheet for each of 5000 ranges,
    // which takes tens of seconds; one pass for the run takes well under a second.
    const cols = Array.from({ length: 50 }, (_, k) =>
      k < 26 ? String.fromCharCode(65 + k) : `A${String.fromCharCode(65 + k - 26)}`
    );
    const rows: string[] = [];
    for (let r = 1; r <= 1000; r += 1)
      rows.push(
        row(
          r,
          cols.map((col) => c(`${col}${r}`, r))
        )
      );
    const read = readXlsx(buildXlsx({ sheets: [{ name: 'S', rows: rows.join('') }] }));
    if (!read.ok) throw new Error(JSON.stringify(read.error));
    const ops: WorkbookOp[] = [];
    for (const id of [9, 4])
      for (let r = 1; r <= 100; r += 1)
        for (const [k, col] of cols.entries())
          if ((r + k) % 2 === (id === 9 ? 0 : 1))
            ops.push(format(`${col}${r}`, { numberFormat: { id } }));
    const store = new WorkbookDocumentStore(read.pkg);
    const start = performance.now();
    expect(store.transact(ops).ok).toBe(true);
    // A wide bound, so a busy test pool does not fail it.
    expect(performance.now() - start).toBeLessThan(5000);
  });

  test('ranges in any order give the same style records as the ops one by one', () => {
    const reversed = [...RANGES].reverse();
    for (const patch of [
      { numberFormat: { id: 9 } },
      { borders: { edges: 'outline' } },
      { bold: true },
    ] as FormatPatch[]) {
      const { batched, single } = both(reversed.map((r) => format(r, patch)));
      expect(canonicalPackageFingerprint(batched.package)).toBe(
        canonicalPackageFingerprint(single.package)
      );
    }
  });

  test('the run gives the refusal of the first op that the ops one by one refuse', () => {
    const store = new WorkbookDocumentStore(open(book()));
    const result = store.transact([
      format('A5:XFD6', { numberFormat: { id: 9 } }),
      format('ZZZZ1', { numberFormat: { id: 9 } }),
    ]);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.rejection.code).toBe('range-too-large');
  });

  test('an op whose patch is not valid is refused as it is alone', () => {
    const store = new WorkbookDocumentStore(open(book()));
    const junk = { numberFormat: { id: 9 }, junk: undefined } as unknown as FormatPatch;
    const result = store.transact([format('A1', { numberFormat: { id: 9 } }), format('B1', junk)]);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.rejection.code).toBe('bad-format');
  });

  test('a run with a patch for each range equals the ops one by one, in any order', () => {
    const patches: FormatPatch[] = [
      { numberFormat: { id: 9 } },
      { fontSize: 14 },
      { borders: { edges: 'outline' } },
      { numberFormat: { code: '"$"#,##0.00' } },
      { bold: true },
      { wrapText: true },
    ];
    for (const ranges of [RANGES, [...RANGES].reverse()]) {
      const { batched, single, before } = both(ranges.map((r, i) => format(r, patches[i]!)));
      // The font size of `C1:D1`, or in reverse the wrap of `A1`, removes the fitted height.
      expect(rowHeight(batched.package, 0)).toBeUndefined();
      expect(rowHeight(single.package, 0)).toBeUndefined();
      expect(canonicalPackageFingerprint(batched.package)).toBe(
        canonicalPackageFingerprint(single.package)
      );
      expect(batched.undo().ok).toBe(true);
      expect(canonicalPackageFingerprint(batched.package)).toBe(before);
    }
  });

  test('overlapping column formats give the run the result of the ops one by one', () => {
    const wb: FixtureWorkbook = {
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 1)]),
          cols: '<cols><col min="3" max="3" style="1"/><col min="1" max="5" style="2"/></cols>',
        },
      ],
      styles: {
        fonts: ['<font><i/><sz val="11"/><name val="Calibri"/></font>'],
        cellXfs: [
          '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
          '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
        ],
      },
    };
    const ops = [format('A5', { bold: true }), format('C5', { bold: true })];
    const batched = new WorkbookDocumentStore(open(wb));
    expect(batched.transact(ops).ok).toBe(true);
    const single = new WorkbookDocumentStore(open(wb));
    for (const op of ops) expect(single.transact([op]).ok).toBe(true);
    expect(canonicalPackageFingerprint(batched.package)).toBe(
      canonicalPackageFingerprint(single.package)
    );
  });

  test('a column format past the cell formats counts as 0, and the new cell keeps its format', () => {
    // Two cell formats, and column B names a third. A new record takes index 2, so a check
    // against the raw column index would skip the new cell.
    const wb: FixtureWorkbook = {
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 1)]),
          cols: '<cols><col min="2" max="2" style="2"/></cols>',
        },
      ],
      styles: {
        fonts: ['<font><i/><sz val="11"/><name val="Calibri"/></font>'],
        cellXfs: ['<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'],
      },
    };
    const ops = [format('B3', { bold: true }), format('A3', { numberFormat: { id: 9 } })];
    const batched = new WorkbookDocumentStore(open(wb));
    expect(batched.transact(ops).ok).toBe(true);
    const single = new WorkbookDocumentStore(open(wb));
    for (const op of ops) expect(single.transact([op]).ok).toBe(true);
    expect(canonicalPackageFingerprint(batched.package)).toBe(
      canonicalPackageFingerprint(single.package)
    );
    const part = xmlPartOf(
      batched.package,
      sheetByName(workbookIndex(batched.package)!, 'S')!.partName
    )!;
    const b3 = sheetIndex(part).cells.get(cellKey(2, 1));
    expect(b3).toBeDefined();
    expect(attr(b3!.el, 's')).toBeDefined();
  });

  test('a full style list refuses before a later range that is not valid', () => {
    const xf = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';
    const full: FixtureWorkbook = {
      sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }],
      styles: { cellXfs: Array.from({ length: MAX_STYLE_RECORDS - 1 }, () => xf) },
    };
    const store = new WorkbookDocumentStore(open(full));
    const result = store.transact([
      format('A1', { numberFormat: { id: 9 } }),
      format('ZZZZ1', { numberFormat: { id: 9 } }),
    ]);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.rejection.code).toBe('style-limit');
  });

  test('a refused range refuses the run', () => {
    const store = new WorkbookDocumentStore(open(book()));
    const result = store.transact([
      format('A1', { numberFormat: { id: 9 } }),
      format('A2:XFD2', { numberFormat: { id: 9 } }),
      format('A5:XFD6', { numberFormat: { id: 9 } }),
    ]);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.rejection.code).toBe('range-too-large');
  });
});
