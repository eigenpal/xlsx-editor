// The demo's sample workbook (examples/vite/public/sample.xlsx): the committed file holds what
// the builder writes, it opens, every sheet lays out with no problem, a save loses nothing, the
// cached formula results are right, and it uses only features that the support matrix lists as
// supported or partial.
//
// Same role as pptx-editor/scripts/sample-deck/showcase.test.ts at 79a5d22d (EigenPal).

import { beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { unzipSync } from 'fflate';
import { supportFeatures } from '../../docs/site/data/support-matrix.ts';
import { layoutSheet } from '../../packages/core/src/layout/sheet-layout.ts';
import { PRINT_PROFILE } from '../../packages/core/src/layout/metrics.ts';
import type { TextMeasurer } from '../../packages/core/src/layout/text/measurer.ts';
import { substituteMeasurer } from '../../packages/core/src/layout/__tests__/test-fonts.ts';
import { cellFormat } from '../../packages/core/src/cascade/index.ts';
import {
  attr,
  cellKey,
  child,
  parseCellRef,
  readXlsx,
  sheetByName,
  sheetIndex,
  sharedStringItems,
  textOf,
  workbookIndex,
  workbookOracles,
  xmlPartOf,
  type XlsxPackage,
} from '../../packages/core/src/store/index.ts';
import { sampleRecordProblems } from './save-check.ts';
import {
  GUIDE_FIRST_ROW,
  SALES_FIRST_ROW,
  SALES_LAST_ROW,
  SHEET_GUIDE,
} from './showcase-sheets.ts';
import { buildSample } from './showcase.ts';

const FILE = join(import.meta.dir, '../../examples/vite/public/sample.xlsx');
const bytes = new Uint8Array(readFileSync(FILE));

function open(data: Uint8Array): XlsxPackage {
  const read = readXlsx(data);
  if (!read.ok) throw new Error(JSON.stringify(read.error));
  return read.pkg;
}

/** The `x:c` element of a cell, when the sheet has one. */
function cellAt(pkg: XlsxPackage, sheet: string, ref: string) {
  const part = xmlPartOf(pkg, sheetByName(workbookIndex(pkg)!, sheet)!.partName)!;
  const at = parseCellRef(ref)!;
  return sheetIndex(part).cells.get(cellKey(at.row, at.col))?.el;
}

/** The text of a shared string cell, or undefined. */
function text(pkg: XlsxPackage, sheet: string, ref: string): string | undefined {
  const el = cellAt(pkg, sheet, ref);
  if (attr(el, 't') !== 's') return undefined;
  const item = sharedStringItems(pkg, workbookIndex(pkg)!)[Number(textOf(child(el, 'x:v')))];
  return textOf(child(item, 'x:t'));
}

/**
 * The support matrix rows that the sample shows. Each must be supported or partial. Keep the
 * list in step with the sheets when a sheet gains a feature; the checks of parts and elements
 * below catch the unsupported features that the list cannot name.
 */
const USED_FEATURES = [
  'values.scalars',
  'values.strings',
  'values.rich-text',
  'values.cached-formulas',
  'values.number-formats',
  'formatting.colors',
  'formatting.fonts',
  'formatting.borders',
  'formatting.fills',
  'formatting.alignment',
  'formatting.wrap',
  'formatting.overflow',
  'formatting.rotation',
  'grid.column-widths',
  'grid.row-heights',
  'grid.hidden',
  'grid.merged-cells',
  'drawings.pictures',
  'drawings.shapes',
  'viewer.frozen-panes',
  'viewer.gridlines',
  'viewer.selection',
  'viewer.sheets',
];

let measurer: TextMeasurer;
beforeAll(async () => {
  measurer = await substituteMeasurer();
});

describe('sample workbook', () => {
  test('the committed file holds what the builder writes', () => {
    // Zip entry times depend on the time zone, so compare the parts, not the zip bytes.
    const committed = unzipSync(bytes);
    const built = unzipSync(buildSample());
    const names = [...new Set([...Object.keys(committed), ...Object.keys(built)])].sort();
    const stale = names.filter(
      (name) =>
        !committed[name] || !built[name] || !Buffer.from(built[name]).equals(committed[name])
    );
    // The remedy is part of the comparison, so a failure prints it.
    const remedy = 'Stale sample.xlsx: run `bun run sample:build`, then commit.';
    expect({ stale, remedy }).toEqual({ stale: [], remedy });
  });

  test('the welcome sheet lists every sheet, in tab order', () => {
    const pkg = open(bytes);
    const index = workbookIndex(pkg)!;
    expect(index.sheets.map((s) => s.name)).toEqual(SHEET_GUIDE.map(([name]) => name));
    expect(index.activeTab).toBe(0);
    expect(index.fullCalcOnLoad).toBe(false);
    const listed = SHEET_GUIDE.map((_, i) => [
      text(pkg, 'Welcome', `B${GUIDE_FIRST_ROW + i}`),
      text(pkg, 'Welcome', `C${GUIDE_FIRST_ROW + i}`),
    ]);
    expect(listed).toEqual(SHEET_GUIDE.map(([name, what]) => [name, what]));
  });

  test('each format code on the Formats sheet is the code of the value beside it', () => {
    const pkg = open(bytes);
    const index = workbookIndex(pkg)!;
    let rows = 0;
    for (let r = 2; r <= 60; r += 1) {
      const code = text(pkg, 'Formats', `C${r}`);
      const value = cellAt(pkg, 'Formats', `B${r}`);
      if (!code || !value) continue;
      const format = cellFormat(pkg, index, attr(value, 's')).format;
      const applied =
        format.numFmtId === 0
          ? 'General'
          : (format.formatCode ?? `Built-in format ${format.numFmtId}`);
      expect({ row: r, code }).toEqual({ row: r, code: applied });
      rows += 1;
    }
    expect(rows).toBe(29);
  });

  test('every sheet lays out with no problem, and the drawing has no unsupported object', () => {
    const pkg = open(bytes);
    const viewport = { width: 1200, height: 900, scrollX: 0, scrollY: 0 };
    for (const [name] of SHEET_GUIDE) {
      const layout = layoutSheet(pkg, name, viewport, { measurer, profile: PRINT_PROFILE });
      expect(layout).toBeDefined();
      expect({ name, problems: layout!.problems }).toEqual({ name, problems: [] });
      const drawings = layout!.panes.flatMap((p) => p.drawings);
      expect(drawings.filter((d) => d.kind === 'unsupported')).toEqual([]);
      if (name === 'Layout')
        expect(drawings.map((d) => d.kind).sort()).toEqual(['picture', 'shape', 'shape', 'shape']);
    }
  });

  test('a save reopens to the same content', () => {
    const oracles = workbookOracles(open(bytes));
    expect(oracles.changedParts).toEqual([]);
    expect(oracles.digest).toBe(oracles.fingerprint);
  });

  test('the cached result of each formula is the result of its formula', () => {
    const pkg = open(bytes);
    const valueOn = (sheet: string) => (ref: string) =>
      Number(textOf(child(cellAt(pkg, sheet, ref), 'x:v')));
    // The formula text of a cell, so that the checks below compare each cache with its formula.
    const formulaOn = (sheet: string) => (ref: string) =>
      textOf(child(cellAt(pkg, sheet, ref), 'x:f'));
    const value = valueOn('Sales');
    const formula = formulaOn('Sales');
    const rows = Array.from(
      { length: SALES_LAST_ROW - SALES_FIRST_ROW + 1 },
      (_, i) => i + SALES_FIRST_ROW
    );
    const total = SALES_LAST_ROW + 1;
    for (const col of ['B', 'C', 'D', 'E', 'F', 'G']) {
      const sum = rows.reduce((acc, r) => acc + value(`${col}${r}`), 0);
      expect({
        ref: `${col}${total}`,
        formula: formula(`${col}${total}`),
        value: value(`${col}${total}`),
      }).toEqual({
        ref: `${col}${total}`,
        formula: `SUM(${col}${SALES_FIRST_ROW}:${col}${SALES_LAST_ROW})`,
        value: sum,
      });
    }
    for (const r of [...rows, total]) {
      if (r !== total)
        expect({ r, formula: formula(`F${r}`), total: value(`F${r}`) }).toEqual({
          r,
          formula: `SUM(B${r}:E${r})`,
          total: value(`B${r}`) + value(`C${r}`) + value(`D${r}`) + value(`E${r}`),
        });
      if (r !== total)
        expect({ r, formula: formula(`G${r}`), share: value(`G${r}`) }).toEqual({
          r,
          formula: `F${r}/$F$${total}`,
          share: value(`F${r}`) / value(`F${total}`),
        });
      expect({ r, formula: formula(`H${r}`), change: value(`H${r}`) }).toEqual({
        r,
        formula: `(E${r}-B${r})/B${r}`,
        change: (value(`E${r}`) - value(`B${r}`)) / value(`B${r}`),
      });
    }
    // Schedule: each End is its Start plus its Days.
    const schedule = valueOn('Schedule');
    const scheduleFormula = formulaOn('Schedule');
    let checked = 0;
    for (let r = 2; cellAt(pkg, 'Schedule', `E${r}`); r += 1, checked += 1)
      expect({ r, formula: scheduleFormula(`E${r}`), end: schedule(`E${r}`) }).toEqual({
        r,
        formula: `C${r}+D${r}`,
        end: schedule(`C${r}`) + schedule(`D${r}`),
      });
    expect(checked).toBeGreaterThan(0);
  });

  test('each task of the Schedule sheet starts and ends on a weekday', () => {
    const pkg = open(bytes);
    const value = (ref: string) => Number(textOf(child(cellAt(pkg, 'Schedule', ref), 'x:v')));
    // 0 is Sunday, 6 is Saturday, for a serial of the 1900 date system after February 1900.
    const weekend = (serial: number) => [0, 6].includes((serial + 6) % 7);
    const days: { r: number; start: boolean; end: boolean }[] = [];
    for (let r = 2; cellAt(pkg, 'Schedule', `C${r}`); r += 1)
      days.push({ r, start: weekend(value(`C${r}`)), end: weekend(value(`E${r}`)) });
    expect(days.length).toBe(40);
    expect(days.filter((d) => d.start || d.end)).toEqual([]);
  });

  test('a saved and an edited copy open with no repair prompt (local save check)', () => {
    // Each problem names its remedy: `bun run sample:save-check` (CONTRIBUTING.md, "Save check").
    expect(sampleRecordProblems()).toEqual([]);
  });

  test('it uses only features that the support matrix lists as supported or partial', () => {
    const status = new Map(supportFeatures.map((f) => [f.id, f.status]));
    for (const id of USED_FEATURES)
      expect({ id, usable: ['supported', 'partial'].includes(status.get(id) ?? '') }).toEqual({
        id,
        usable: true,
      });
    const parts = unzipSync(bytes);
    const names = Object.keys(parts);
    const unsupportedParts =
      /charts\/|tables\/|pivot|externalLinks|connections|vbaProject|calcChain/;
    expect(names.filter((n) => unsupportedParts.test(n))).toEqual([]);
    const decoder = new TextDecoder();
    for (const [name, data] of Object.entries(parts))
      if (name.startsWith('xl/worksheets/'))
        expect({
          name,
          found: /<(conditionalFormatting|dataValidations|tableParts|hyperlinks)\b/.test(
            decoder.decode(data)
          ),
        }).toEqual({
          name,
          found: false,
        });
  });
});
