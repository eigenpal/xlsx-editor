// The format commands lower to store ops: a built-in code becomes its id, an unreadable code is
// refused before the store, a clear of both kinds is one transaction, and a toggle reads the
// resolved format.

import { describe, expect, test } from 'bun:test';
import { clearOps, formatOps, togglePatch } from '../index.ts';
import type { ResolvedCellFormat } from '../../cascade/index.ts';

describe('format commands', () => {
  test('the op keeps a code, and the store lowers it with the workbook records', () => {
    // The store writes a code that equals a built-in code as its id, unless the workbook
    // gives that id another code (number-format-entry.test.ts, number-format-lowering.test.ts).
    const r = formatOps('S', 'A1', { numberFormat: { code: '0%' } });
    expect(r.ok && r.ops).toEqual([
      { op: 'applyFormat', sheet: 'S', ref: 'A1', patch: { numberFormat: { code: '0%' } } },
    ]);
    const custom = formatOps('S', 'A1', { numberFormat: { code: '0.0%' }, bold: true });
    expect(custom.ok && custom.ops[0]).toMatchObject({
      patch: { numberFormat: { code: '0.0%' }, bold: true },
    });
  });

  test('a code that the number format parser cannot read is refused', () => {
    for (const code of ['"open', '[Red', '0;0;0;0;0', 'x'.repeat(300)]) {
      const r = formatOps('S', 'A1', { numberFormat: { code } });
      expect(r.ok ? null : r.rejection.code).toBe('unreadable-number-format');
    }
  });

  test('clears and toggles', () => {
    expect(clearOps('S', 'A1:B2', 'all').map((o) => o.op)).toEqual(['clearValues', 'clearFormats']);
    expect(clearOps('S', 'A1', 'formats').map((o) => o.op)).toEqual(['clearFormats']);
    const format = {
      font: { bold: true, italic: false, underline: 'double' },
    } as unknown as ResolvedCellFormat;
    expect(togglePatch(format, 'bold')).toEqual({ bold: false });
    expect(togglePatch(format, 'italic')).toEqual({ italic: true });
    expect(togglePatch(format, 'underline')).toEqual({ underline: false });
  });
});
