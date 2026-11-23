// Layout after a structural edit equals the layout of the saved and reopened file (task 7.5):
// the store revision that the editor paints holds nothing that a save drops, so a user sees
// after the edit what a reopen shows. The edit is the structure fixture's (scripts/fixtures).

import { beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { readXlsx, type XlsxPackage } from '../../store/package/package.ts';
import { saveXlsx } from '../../store/package-write/save.ts';
import { WorkbookDocumentStore } from '../../store/store/document-store.ts';
import type { WorkbookOp } from '../../store/store/ops.ts';
import { layoutSheet, type SheetLayout } from '../sheet-layout.ts';
import { PRINT_PROFILE, SCREEN_PROFILE } from '../metrics.ts';
import type { TextMeasurer } from '../text/measurer.ts';
import { substituteMeasurer } from './test-fonts.ts';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';

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

/** The facts of a layout that a viewer paints, as plain data. */
function painted(pkg: XlsxPackage, sheet: string, screen: boolean): string {
  const l: SheetLayout | undefined = layoutSheet(pkg, sheet, VIEW, {
    measurer,
    profile: screen ? SCREEN_PROFILE : PRINT_PROFILE,
  });
  if (!l) throw new Error('no sheet');
  const offsets = (n: number, at: (i: number) => number) =>
    Array.from({ length: n }, (_, i) => at(i));
  return JSON.stringify({
    panes: l.panes,
    frozen: l.frozen,
    edges: l.frozenEdges,
    problems: l.problems,
    rows: offsets(40, (i) => l.geometry.rows.offsetOf(i)),
    cols: offsets(20, (i) => l.geometry.columns.offsetOf(i)),
  });
}

const edits: readonly (readonly WorkbookOp[])[] = [
  [
    { op: 'deleteRows', sheet: 'Data', start: 4, count: 1 },
    { op: 'insertColumns', sheet: 'Data', start: 3, count: 1 },
  ],
  [
    { op: 'insertRows', sheet: 'Data', start: 2, count: 3 },
    { op: 'deleteColumns', sheet: 'Data', start: 2, count: 2 },
  ],
];

describe('layout after a structural edit', () => {
  for (const ops of edits)
    test(`equals the layout of the saved file: ${ops.map((o) => o.op).join(', ')}`, () => {
      const pkg = open(
        new Uint8Array(readFileSync(`${import.meta.dir}/../../../../../fixtures/structure.xlsx`))
      );
      const store = new WorkbookDocumentStore(pkg);
      const r = store.transact(ops);
      expect(r.ok).toBe(true);
      const reopened = open(saveXlsx(store.package));
      for (const sheet of ['Data', 'Refs'])
        for (const screen of [false, true])
          expect(painted(store.package, sheet, screen)).toBe(painted(reopened, sheet, screen));
      expect(painted(store.package, 'Data', true)).not.toBe(painted(pkg, 'Data', true));
    });
});

describe('layout of a frozen sheet after a structural edit', () => {
  const frozen = () =>
    open(
      buildXlsx({
        sheets: [
          {
            name: 'F',
            rows: row(1, [c('A1', 'Head'), c('B1', 'Side')]) + row(2, [c('A2', 1)]),
            viewChildren:
              '<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>' +
              '<selection pane="bottomRight" activeCell="B2" sqref="B2"/>',
          },
        ],
      })
    );
  const cases: readonly (readonly [WorkbookOp, number, number])[] = [
    [{ op: 'insertRows', sheet: 'F', start: 1, count: 2 }, 3, 1],
    [{ op: 'deleteRows', sheet: 'F', start: 1, count: 1 }, 0, 1],
    [{ op: 'insertColumns', sheet: 'F', start: 1, count: 1 }, 1, 2],
    [{ op: 'deleteColumns', sheet: 'F', start: 1, count: 1 }, 1, 0],
  ];
  for (const [edit, rows, cols] of cases)
    test(`${edit.op} ${'start' in edit ? edit.start : ''}: the band follows, and a reopen agrees`, () => {
      const store = new WorkbookDocumentStore(frozen());
      expect(store.transact([edit]).ok).toBe(true);
      const live = layoutSheet(store.package, 'F', VIEW, { measurer, profile: SCREEN_PROFILE })!;
      expect(live.frozen).toEqual({ rows, cols });
      const reopened = open(saveXlsx(store.package));
      for (const screen of [false, true])
        expect(painted(store.package, 'F', screen)).toBe(painted(reopened, 'F', screen));
    });
});
