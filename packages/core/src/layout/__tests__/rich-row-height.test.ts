// The cost of the row heights of one rich shared string that many rows show.

import { beforeAll, describe, expect, test } from 'bun:test';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx } from '../../store/package/package.ts';
import { PRINT_PROFILE } from '../metrics.ts';
import { layoutSheet } from '../sheet-layout.ts';
import type { TextMeasurer } from '../text/measurer.ts';
import { substituteMeasurer } from './test-fonts.ts';

const RUNS = 20_000;
const ROWS = 20_000;
/** Far under the time of one walk of the runs for each row, and far over one walk in all. */
const LAYOUT_LIMIT_MS = 10_000;

let measurer: TextMeasurer;
beforeAll(async () => {
  measurer = await substituteMeasurer();
});

describe('a rich shared string in many rows', () => {
  test(
    'the layout cost does not grow with the rows that repeat the string',
    () => {
      const runs = Array.from({ length: RUNS }, (_, i) => `<r><t>r${i % 10}</t></r>`).join('');
      const rows = Array.from({ length: ROWS }, (_, r) =>
        row(r + 1, [c(`A${r + 1}`, 0, { t: 's' })])
      ).join('');
      const read = readXlsx(
        buildXlsx({ sheets: [{ name: 'S', rows }], rawStrings: [`<si>${runs}</si>`] })
      );
      if (!read.ok) throw new Error('read');
      const view = { width: 540, height: 720, scrollX: 0, scrollY: 0 };
      const start = performance.now();
      const laid = layoutSheet(read.pkg, 'S', view, { measurer, profile: PRINT_PROFILE });
      expect(laid).not.toBeNull();
      expect(performance.now() - start).toBeLessThan(LAYOUT_LIMIT_MS);
    },
    { timeout: 60_000 }
  );

  test(
    'empty runs in wrapped rows spend the measure budget',
    () => {
      const runs = '<r><t></t></r>'.repeat(RUNS);
      const rows = Array.from({ length: ROWS }, (_, r) =>
        row(r + 1, [c(`A${r + 1}`, 0, { t: 's', s: 1 })])
      ).join('');
      const wrap =
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf>';
      const read = readXlsx(
        buildXlsx({
          sheets: [{ name: 'S', rows }],
          rawStrings: [`<si>${runs}</si>`],
          styles: { cellXfs: [wrap] },
        })
      );
      if (!read.ok) throw new Error('read');
      const view = { width: 540, height: 720, scrollX: 0, scrollY: 0 };
      const start = performance.now();
      const laid = layoutSheet(read.pkg, 'S', view, { measurer, profile: PRINT_PROFILE });
      expect(laid).not.toBeNull();
      expect(performance.now() - start).toBeLessThan(LAYOUT_LIMIT_MS);
    },
    { timeout: 120_000 }
  );
});
