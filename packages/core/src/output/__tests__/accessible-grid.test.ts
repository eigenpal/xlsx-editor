// The accessible rows of a formula cell next to overflowing text, in a happy-dom window of its
// own (the engine tests run with no global DOM).
import { beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { Window } from 'happy-dom';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx, type XlsxPackage } from '../../store/index.ts';
import {
  layoutSheet,
  PRINT_PROFILE,
  type SheetLayout,
  type TextMeasurer,
} from '../../layout/index.ts';
import { substituteMeasurer } from '../../layout/__tests__/test-fonts.ts';
import { accessibleLabels } from '../../editor/editor-model.ts';
import { englishText } from '../../editor/english.ts';
import { accessibleGrid } from '../index.ts';

let measurer: TextMeasurer;
beforeAll(async () => {
  measurer = await substituteMeasurer();
});

const open = (bytes: Uint8Array): XlsxPackage => {
  const r = readXlsx(bytes);
  if (!r.ok) throw new Error('read');
  return r.pkg;
};
const VIEW = { width: 540, height: 720, scrollX: 0, scrollY: 0 };
const layout = (pkg: XlsxPackage, sheet: string): SheetLayout => {
  const out = layoutSheet(pkg, sheet, VIEW, { measurer, profile: PRINT_PROFILE });
  if (!out) throw new Error('no sheet');
  return out;
};
const rowsOf = (pkg: XlsxPackage, sheet: string) =>
  accessibleGrid(
    new Window().document as unknown as Document,
    layout(pkg, sheet),
    accessibleLabels(englishText)
  );

describe('accessibleGrid next to overflowing text', () => {
  test('formula-overflow: each formula cell is a grid cell with its cached value', () => {
    const bytes = readFileSync(`${import.meta.dir}/../../../../../fixtures/formula-overflow.xlsx`);
    const rows = rowsOf(open(new Uint8Array(bytes)), 'Overflow');
    expect(rows.querySelector('[data-cell="A1"]')!.textContent).toBe(
      'This text is longer than its cell'
    );
    expect(rows.querySelector('[data-cell="B1"]')!.textContent).toBe('2');
  });

  test('a formula with no cached value is a stale grid cell that reads as having no result', () => {
    const pkg = open(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            rows: row(1, [c('A1', 'A long text that overflows'), c('B1', undefined, { f: '1+1' })]),
          },
        ],
      })
    );
    const b1 = rowsOf(pkg, 'S').querySelector('[data-cell="B1"]')!;
    expect(b1.getAttribute('role')).toBe('gridcell');
    expect(b1.getAttribute('data-stale')).toBe('1');
    expect(b1.textContent).toBe('No result, not recalculated');
  });
});
