// The demo's heading switch (`headings.ts`): on, every worksheet view shows
// its headings; off, each view gets back its own `showRowColHeaders`, also a view with no value,
// a sheet with two views, and a sheet with no `x:sheetViews`.

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildXlsx } from '../../../scripts/fixtures/builder.ts';
import { readXlsx } from '../../../packages/core/src/store/index.ts';
import { headingFlags as flagsOf, withHeadingsSteps, type HeadingFlags } from './headings.ts';

const headingFlags = (bytes: Uint8Array) => flagsOf(readXlsx(bytes));

/** The steps of the headings switch, run to their end. */
function withHeadings(bytes: Uint8Array, show: boolean, restore: HeadingFlags): Uint8Array {
  const steps = withHeadingsSteps(bytes, show, restore);
  for (;;) {
    const next = steps.next();
    if (!next.done) continue;
    if (!next.value) throw new Error('the bytes did not read');
    return next.value;
  }
}

const VIEW = /<sheetViews>[\s\S]*?<\/sheetViews>/;

/** A workbook of three sheets, with the `x:sheetViews` of each sheet replaced by `views`. */
function workbook(views: readonly string[]): Uint8Array {
  const bytes = buildXlsx({
    sheets: views.map((_, i) => ({ name: `S${i + 1}`, rows: '' })),
  });
  const parts = unzipSync(bytes);
  views.forEach((v, i) => {
    const name = `xl/worksheets/sheet${i + 1}.xml`;
    parts[name] = strToU8(strFromU8(parts[name]!).replace(VIEW, v));
  });
  return zipSync(parts);
}

/** The `showRowColHeaders` values of each view, by sheet, in document order. */
const valuesOf = (bytes: Uint8Array) => [...headingFlags(bytes).values()];

const NONE = '<sheetViews><sheetView workbookViewId="0"/></sheetViews>';
const TWO =
  '<sheetViews><sheetView showRowColHeaders="0" workbookViewId="0"/>' +
  '<sheetView showRowColHeaders="1" workbookViewId="1"/></sheetViews>';

describe('the heading switch', () => {
  const file = workbook([NONE, TWO, '']);

  test('reads the flag of every view, and none for a sheet with no views', () => {
    expect(valuesOf(file)).toEqual([[null], ['0', '1'], []]);
  });

  test('on, every view shows its headings', () => {
    expect(valuesOf(withHeadings(file, true, headingFlags(file)))).toEqual([['1'], ['1', '1'], []]);
  });

  test('off, each view gets back its own value, also one with no value', () => {
    const flags = headingFlags(file);
    const shown = withHeadings(file, true, flags);
    expect(valuesOf(withHeadings(shown, false, flags))).toEqual(valuesOf(file));
  });

  test('bytes that do not read give no bytes, so their flags cannot reach a file', () => {
    const bad = new Uint8Array([1, 2, 3]);
    expect(() => withHeadings(bad, false, new Map())).toThrow('the bytes did not read');
    expect(headingFlags(bad).size).toBe(0);
  });
});
