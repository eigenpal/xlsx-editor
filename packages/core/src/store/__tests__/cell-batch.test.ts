// A run of setCellValue ops in one transaction (`cell-batch.ts`): the package equals the one
// that the ops give one transaction at a time, the inverse restores the sheet, a run that one
// op at a time refuses is refused the same way, and the cost grows with the writes, not with
// their square.

import { describe, expect, test } from 'bun:test';
import { c, row, type FixtureWorkbook } from '../../../../../scripts/fixtures/builder.ts';
import { xmlPartOf, type XlsxPackage } from '../package/package.ts';
import { serializeXmlPart } from '../package-write/serialize.ts';
import { columnLetters } from '../cell-ref.ts';
import { workbookIndex } from '../workbook/workbook-index.ts';
import { WorkbookDocumentStore } from '../store/document-store.ts';
import type { CellInput, WorkbookOp } from '../store/ops.ts';
import { open } from './structure-helpers.ts';

const SHEET = '/xl/worksheets/sheet1.xml';

const wb = (dimension?: string): FixtureWorkbook => ({
  sheets: [
    {
      name: 'S',
      ...(dimension ? { dimension } : {}),
      rows:
        row(1, [c('A1', 1, { s: 1 }), c('C1', 0, { t: 's' })]) +
        row(2, [c('B2', 2), c('D2', 0, { f: 'B2*2' })], 'ht="20" customHeight="1"') +
        '<row r="4"><c r="A4"><v>4</v></c><extLst><ext uri="{0}"/></extLst></row>',
    },
  ],
  strings: ['x', 'y'],
  styles: { cellXfs: ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'] },
});

const set = (ref: string, value: CellInput): WorkbookOp => ({
  op: 'setCellValue',
  sheet: 'S',
  ref,
  value,
});

const OPS: readonly WorkbookOp[] = [
  set('B1', { type: 'number', value: 7 }),
  set('A1', { type: 'empty' }),
  set('C1', { type: 'string', value: 'x' }),
  set('B2', { type: 'empty' }),
  set('A2', { type: 'formula', text: 'SUM(A1:A3)' }),
  set('E3', { type: 'string', value: 'new text' }),
  set('B4', { type: 'string', value: 'y' }),
  set('Z4', { type: 'boolean', value: true }),
  set('A3', { type: 'date', iso: '2024-03-01' }),
  set('A3', { type: 'number', value: 3 }),
  set('D2', { type: 'number', value: 9 }),
  set('C9', { type: 'string', value: 'new text' }),
  set('B1', { type: 'number', value: 8 }),
];

const xml = (pkg: XlsxPackage) => {
  const index = workbookIndex(pkg)!;
  return [SHEET, index.sharedStringsPart!, index.mainPart].map((p) =>
    serializeXmlPart(xmlPartOf(pkg, p)!)
  );
};

describe('a run of cell writes in one transaction', () => {
  for (const dimension of [undefined, 'A1:D4'])
    test(`equals the writes one at a time${dimension ? ', and widens the dimension' : ''}`, () => {
      const one = new WorkbookDocumentStore(open(wb(dimension)));
      for (const o of OPS) expect(one.transact([o]).ok).toBe(true);
      const batch = new WorkbookDocumentStore(open(wb(dimension)));
      const before = xml(batch.package);
      expect(batch.transact(OPS).ok).toBe(true);
      expect(xml(batch.package)).toEqual(xml(one.package));
      expect(batch.undo().ok).toBe(true);
      expect(xml(batch.package)[0]).toBe(before[0]);
      expect(batch.redo().ok).toBe(true);
      expect(xml(batch.package)).toEqual(xml(one.package));
    });

  test('a run that changes nothing commits nothing', () => {
    const s = new WorkbookDocumentStore(open(wb()));
    const r = s.transact([
      set('B2', { type: 'number', value: 2 }),
      set('C1', { type: 'string', value: 'x' }),
    ]);
    expect(r.ok).toBe(true);
    expect(s.revision).toBe(0);
  });

  test('a write that one op refuses refuses the run with the same reason', () => {
    const s = new WorkbookDocumentStore(open(wb()));
    const before = s.package;
    const r = s.transact([
      set('B1', { type: 'number', value: 1 }),
      set('B1', { type: 'formula', text: 'SUM(' }),
    ]);
    expect(r.ok ? null : r.rejection.code).toBe('unreadable-formula');
    expect(s.package).toBe(before);
  });

  test('50,000 writes and 2,000 formulas commit, undo, and redo in time', () => {
    const s = new WorkbookDocumentStore(open({ sheets: [{ name: 'S', rows: '' }] }));
    const ops: WorkbookOp[] = [];
    for (let i = 0; i < 50_000; i += 1) {
      const ref = `${columnLetters((i % 25) + 1)}${Math.floor(i / 25) + 1}`;
      ops.push(
        set(
          ref,
          i < 2_000
            ? { type: 'formula', text: `A1+${i}` }
            : i % 2
              ? { type: 'string', value: `s${i}` }
              : { type: 'number', value: i }
        )
      );
    }
    const t = performance.now();
    expect(s.transact(ops).ok).toBe(true);
    expect(s.undo().ok).toBe(true);
    expect(s.redo().ok).toBe(true);
    // About 0.1 s on a laptop; one op at a time took more than a minute.
    expect(performance.now() - t).toBeLessThan(3_000);
  });
});
