// The cells of one pane that show the same very long list of runs share one text layout.

import { beforeAll, describe, expect, test } from 'bun:test';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { cellText } from '../../output/accessible-grid.ts';
import { readXlsx } from '../../store/package/package.ts';
import { PRINT_PROFILE } from '../metrics.ts';
import { MANY_RUNS } from '../shared-cell-text.ts';
import { layoutSheet, type CellTextBox } from '../sheet-layout.ts';
import type { TextMeasurer } from '../text/measurer.ts';
import { substituteMeasurer } from './test-fonts.ts';

const VIEW = { width: 540, height: 720, scrollX: 0, scrollY: 0 };

let measurer: TextMeasurer;
beforeAll(async () => {
  measurer = await substituteMeasurer();
});

/** A workbook whose first `rows` rows all show one shared string of `runs` runs. */
function workbook(runs: number, rows: number) {
  const items = Array.from({ length: runs }, (_, i) => `<r><t>r${i % 10}</t></r>`).join('');
  const cells = Array.from({ length: rows }, (_, r) => row(r + 1, [c(`A${r + 1}`, 0, { t: 's' })]));
  const read = readXlsx(
    buildXlsx({ sheets: [{ name: 'S', rows: cells.join('') }], rawStrings: [`<si>${items}</si>`] })
  );
  if (!read.ok) throw new Error('read');
  return read.pkg;
}

/** The text boxes of a layout with `measure`. */
function boxes(runs: number, rows: number, measure = measurer): readonly CellTextBox[] {
  const laid = layoutSheet(workbook(runs, rows), 'S', VIEW, {
    measurer: measure,
    profile: PRINT_PROFILE,
  });
  if (!laid) throw new Error('layout');
  return laid.panes.flatMap((p) => p.texts);
}

/** The calls that one layout makes to measure text. */
function measureCalls(runs: number, rows: number): number {
  let calls = 0;
  const counting = Object.create(measurer) as TextMeasurer;
  counting.advances = (face, text) => {
    calls += 1;
    return measurer.advances(face, text);
  };
  boxes(runs, rows, counting);
  return calls;
}

describe('the cells that show one long list of runs', () => {
  test('measure the runs once, for any count of visible rows', () => {
    const runs = MANY_RUNS * 4;
    expect(measureCalls(runs, 30) - measureCalls(runs, 1)).toBeLessThan(runs);
  });

  test('keep only the spans that can reach the clip', () => {
    const runs = MANY_RUNS * 4;
    const [first, second] = boxes(runs, 2);
    for (const box of [first!, second!]) {
      const { spans } = box.text.lines[0]!;
      expect(spans.length).toBeGreaterThan(0);
      expect(spans.length).toBeLessThan(runs);
      const right = box.clip.x + box.clip.width;
      for (const s of spans) expect(s.x - s.style.sizePt).toBeLessThanOrEqual(right);
    }
  });

  test('place each cell at its own row', () => {
    const [first, second] = boxes(MANY_RUNS * 4, 2);
    const a = first!.text.lines[0]!;
    const b = second!.text.lines[0]!;
    expect(b.baseline - a.baseline).toBeCloseTo(second!.clip.y - first!.clip.y, 6);
    expect(b.spans[0]!.x).toBeCloseTo(a.spans[0]!.x, 6);
    expect(b.spans[0]!.positions).toEqual(a.spans[0]!.positions);
  });

  test('give assistive technology the complete text', () => {
    const runs = MANY_RUNS * 4;
    const [box] = boxes(runs, 1);
    const whole = Array.from({ length: runs }, (_, i) => `r${i % 10}`).join('');
    expect(cellText(box!)).toBe(whole);
  });

  test('leave a cell with fewer runs as it is', () => {
    const runs = MANY_RUNS - 1;
    const [box] = boxes(runs, 1);
    expect(box!.shownText).toBeUndefined();
    expect(box!.text.lines[0]!.spans).toHaveLength(runs);
  });
});
