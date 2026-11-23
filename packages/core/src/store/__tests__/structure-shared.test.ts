// The shared formula safeguard of structural edits: the store never writes a new shared formula,
// and a master that stays shared never gains a sheet-qualified reference. A shared formula with
// one fails to open in the save check (scripts/fixtures/structure-case.ts).

import { describe, expect, test } from 'bun:test';
import { c, row } from '../../../../../scripts/fixtures/builder.ts';
import { attr, child, textOf } from '../package/query.ts';
import { xmlPartOf, type XlsxPackage } from '../package/package.ts';
import { formatCellRef } from '../cell-ref.ts';
import { sheetIndex } from '../workbook/sheet-index.ts';
import { workbookIndex } from '../workbook/workbook-index.ts';
import { gainsSheetPrefix } from '../structure/sheet-cells.ts';
import type { StructuralOpName } from '../store/ops.ts';
import { op, storeOf } from './structure-helpers.ts';

/** Every shared formula cell of the workbook: `sheet!ref` to its `si` and its master text. */
function sharedCells(pkg: XlsxPackage): Map<string, { si: string; text: string | undefined }> {
  const out = new Map<string, { si: string; text: string | undefined }>();
  for (const sheet of workbookIndex(pkg)!.sheets) {
    const part = xmlPartOf(pkg, sheet.partName);
    if (!part) continue;
    for (const entry of sheetIndex(part).cells.values()) {
      const f = child(entry.el, 'x:f');
      if (attr(f, 't') !== 'shared') continue;
      const text = attr(f, 'ref') !== undefined ? textOf(f) : undefined;
      out.set(`${sheet.name}!${formatCellRef(entry)}`, { si: attr(f, 'si') ?? '', text });
    }
  }
  return out;
}

const wb = () => ({
  sheets: [
    {
      name: 'S',
      rows:
        row(1, [
          c('A1', 1),
          c('B1', 2),
          '<c r="C1"><f t="shared" ref="C1:C4" si="0">A2*2</f><v>0</v></c>',
        ]) +
        row(2, [c('A2', 3), '<c r="C2"><f t="shared" si="0"/><v>0</v></c>']) +
        row(3, [c('A3', 4), '<c r="C3"><f t="shared" si="0"/><v>0</v></c>']) +
        row(4, [
          c('A4', 5),
          '<c r="C4"><f t="shared" si="0"/><v>0</v></c>',
          '<c r="D4"><f t="shared" ref="D4:F4" si="1">$A$1+A4</f><v>0</v></c>',
          '<c r="E4"><f t="shared" si="1"/><v>0</v></c>',
          '<c r="F4"><f t="shared" si="1"/><v>0</v></c>',
        ]),
    },
    { name: 'T', rows: row(1, [c('A1', 0, { f: 'S!A2*2' }), c('B1', 0, { f: 'SUM(S!A1:A4)' })]) },
  ],
});

describe('the shared formula safeguard', () => {
  test('a move never adds a sheet prefix; a new one is caught', () => {
    expect(gainsSheetPrefix('A1*2', 'A2*2')).toBe(false);
    expect(gainsSheetPrefix('S!A1*2', 'S!#REF!*2')).toBe(false);
    expect(gainsSheetPrefix('A1*2', 'S!A1*2')).toBe(true);
    expect(gainsSheetPrefix('A1*2', '"')).toBe(true);
  });

  const edits: [StructuralOpName, number, number][] = [];
  for (const name of ['insertRows', 'deleteRows', 'insertColumns', 'deleteColumns'] as const)
    for (const start of [1, 2, 3, 4, 5])
      for (const count of [1, 2]) edits.push([name, start, count]);

  test.each(edits)('%s at %i, count %i, writes no new shared formula', (name, start, count) => {
    const s = storeOf(wb());
    const before = sharedCells(s.package);
    const masters = new Map(
      [...before.values()].filter((v) => v.text !== undefined).map((v) => [v.si, v.text!])
    );
    const r = s.transact([op(name, 'S', start, count)]);
    expect(r.ok).toBe(true);
    const after = sharedCells(s.package);
    expect(after.size).toBeLessThanOrEqual(before.size);
    for (const [at, { si, text }] of after) {
      expect(masters.has(si)).toBe(true);
      if (text !== undefined) expect(gainsSheetPrefix(masters.get(si)!, text)).toBe(false);
      expect(at.startsWith('S!')).toBe(true);
    }
  });
});
