// The row height cache: an edit measures the rows that it changed, and a row measures again when
// something that its height reads changes. After each kind of edit, the cached heights must equal
// the heights of the saved bytes read again, which no cache holds.

import { beforeAll, beforeEach, describe, expect, test } from 'bun:test';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx, type XlsxPackage } from '../../store/package/package.ts';
import { saveXlsx } from '../../store/package-write/save.ts';
import { WorkbookDocumentStore } from '../../store/store/document-store.ts';
import { sharedStringItems, workbookIndex } from '../../store/workbook/workbook-index.ts';
import { PRINT_PROFILE, SCREEN_PROFILE, type GridProfile } from '../metrics.ts';
import { rowHeightTestRecorder } from '../row-height.ts';
import { layoutSheet } from '../sheet-layout.ts';
import type { TextMeasurer } from '../text/measurer.ts';
import { substituteMeasurer } from './test-fonts.ts';

let measurer: TextMeasurer;
beforeAll(async () => {
  measurer = await substituteMeasurer();
});

const work = rowHeightTestRecorder();
beforeEach(() => work.reset());

const ROWS = 50;
/** The row whose shared string cell names the first item past the end of the table. */
const MISSING = 40;
const WRAP =
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">' +
  '<alignment wrapText="1"/></xf>';
/** Cell format 2: font 1, at 30 pt. No cell uses it when the workbook opens. */
const LARGE_FONT =
  '<font><sz val="30"/><color theme="1"/><name val="Carlito"/><family val="2"/>' +
  '<scheme val="minor"/></font>';
const LARGE = '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>';

/** `ROWS` rows of a number and a shared string; row `wrapped` (1-based) wraps a long text. */
function workbook(wrapped?: number): XlsxPackage {
  const rows = Array.from({ length: ROWS }, (_, i) => {
    const n = i + 1;
    const item = n === MISSING ? 2 : n % 2;
    const cells = [c(`A${n}`, n), c(`B${n}`, item, { t: 's' })];
    if (n === wrapped) cells.push(c(`C${n}`, 'word '.repeat(40), { s: 1 }));
    return row(n, cells);
  });
  const read = readXlsx(
    buildXlsx({
      sheets: [{ name: 'S', rows: rows.join('') }],
      strings: ['a', 'b'],
      styles: { fonts: [LARGE_FONT], cellXfs: [WRAP, LARGE] },
    })
  );
  if (!read.ok) throw new Error('read');
  return read.pkg;
}

const VIEW = { width: 540, height: 720, scrollX: 0, scrollY: 0 };

/** The height of each row after one layout. */
function heights(
  pkg: XlsxPackage,
  profile: GridProfile = PRINT_PROFILE,
  by: TextMeasurer = measurer
): number[] {
  const out = layoutSheet(pkg, 'S', VIEW, { measurer: by, profile });
  if (!out) throw new Error('no sheet');
  return Array.from({ length: ROWS }, (_, r) => out.geometry.rows.sizeOf(r));
}

function edit(store: WorkbookDocumentStore, ...ops: Parameters<typeof store.transact>[0]) {
  const result = store.transact(ops);
  if (!result.ok) throw new Error(`the edit was refused: ${JSON.stringify(result)}`);
}

/**
 * Commits `count` new texts, lays out after each one, and returns the store with a weak
 * reference to each item list. The caller keeps the store, so its row elements and their cache
 * entries stay alive, and the undo history keeps each old row. With `wrapped`, each text goes to
 * the wrapped cell of row 3; otherwise text `n` goes to cell `E<n>`.
 */
function commitTexts(
  count: number,
  wrapped = false
): { store: WorkbookDocumentStore; lists: WeakRef<object>[] } {
  const store = new WorkbookDocumentStore(workbook(wrapped ? 3 : undefined));
  heights(store.package);
  const lists: WeakRef<object>[] = [];
  for (let n = 1; n <= count; n += 1) {
    work.reset();
    edit(store, {
      op: 'setCellValue',
      sheet: 'S',
      ref: wrapped ? 'C3' : `E${n}`,
      value: { type: 'string', value: wrapped ? `${'word '.repeat(40)}${n}` : `text ${n}` },
    });
    heights(store.package);
    const pkg = store.package;
    lists.push(new WeakRef(sharedStringItems(pkg, workbookIndex(pkg)!)));
    // Row `n` or the wrapped row, and row 40 while the table grows to its item 2.
    if (!wrapped) expect(work.measured).toBe(n === 1 ? 2 : 1);
  }
  return { store, lists };
}

/** The lists of `lists` that are still in memory after garbage collection. */
function alive(lists: readonly WeakRef<object>[]): number {
  Bun.gc(true);
  Bun.gc(true);
  return lists.filter((l) => l.deref() !== undefined).length;
}

describe('the row height cache', () => {
  test('an edit measures only the row that it changed', () => {
    const store = new WorkbookDocumentStore(workbook());
    const before = heights(store.package);
    expect(work.measured).toBe(ROWS);
    work.reset();
    edit(store, { op: 'setCellValue', sheet: 'S', ref: 'A5', value: { type: 'number', value: 7 } });
    const after = heights(store.package);
    expect(after).toEqual(before);
    expect(work.measured).toBe(1);
    expect(work.cached).toBe(ROWS - 1);
    expect(after).toEqual(heights(readXlsxAgain(store.package)));
  });

  test('undo and redo of a new text and of a font size give the heights of a fresh read', () => {
    const store = new WorkbookDocumentStore(workbook());
    heights(store.package);
    edit(store, {
      op: 'setCellValue',
      sheet: 'S',
      ref: 'D9',
      value: { type: 'string', value: 'new' },
    });
    edit(store, { op: 'applyFormat', sheet: 'S', ref: 'A5', patch: { fontSize: 24 } });
    heights(store.package);
    for (const step of ['undo', 'undo', 'redo', 'redo'] as const) {
      const result = step === 'undo' ? store.undo() : store.redo();
      if (!result.ok) throw new Error(`${step} was refused`);
      expect(heights(store.package)).toEqual(heights(readXlsxAgain(store.package)));
    }
  });

  test('undo and redo of a new text in a wrapped cell give the heights of a fresh read', () => {
    const store = new WorkbookDocumentStore(workbook(3));
    const before = heights(store.package);
    edit(store, {
      op: 'setCellValue',
      sheet: 'S',
      ref: 'C3',
      value: { type: 'string', value: 'short' },
    });
    const after = heights(store.package);
    expect(after[2]).toBeLessThan(before[2]!);
    // Undo gives back the old cell element, which the cache of wrapped cells can still hold.
    for (const [step, want] of [
      ['undo', before],
      ['redo', after],
      ['undo', before],
    ] as const) {
      const result = step === 'undo' ? store.undo() : store.redo();
      if (!result.ok) throw new Error(`${step} was refused`);
      const got = heights(store.package);
      expect(got).toEqual(want);
      expect(got).toEqual(heights(readXlsxAgain(store.package)));
    }
  });

  test('a new font size measures every row again, and the row takes the font', () => {
    const store = new WorkbookDocumentStore(workbook());
    const before = heights(store.package);
    work.reset();
    edit(store, { op: 'applyFormat', sheet: 'S', ref: 'A5', patch: { fontSize: 24 } });
    const after = heights(store.package);
    // The styles part is a new revision, so no row can take its cached height.
    expect(work.measured).toBe(ROWS);
    expect(after[4]).toBeGreaterThan(before[4]!);
    expect(after.filter((h, r) => h !== before[r])).toEqual([after[4]!]);
    expect(after).toEqual(heights(readXlsxAgain(store.package)));
  });

  test('a cell format of a larger font measures only its row, and the row takes the font', () => {
    const store = new WorkbookDocumentStore(workbook());
    const before = heights(store.package);
    work.reset();
    // The cell format exists, so the styles part stays the same and only the cell changes.
    edit(store, { op: 'setCellFormatIndex', sheet: 'S', ref: 'A7', index: 2 });
    const after = heights(store.package);
    expect(work.measured).toBe(1);
    expect(work.cached).toBe(ROWS - 1);
    expect(after[6]).toBeGreaterThan(before[6]!);
    expect(after.filter((h, r) => h !== before[r])).toEqual([after[6]!]);
    expect(after).toEqual(heights(readXlsxAgain(store.package)));
  });

  test('a new shared string keeps the rows that read earlier items', () => {
    const store = new WorkbookDocumentStore(workbook());
    heights(store.package);
    work.reset();
    // The table only grows, so the new item takes index 2: the item that row 40 names.
    edit(store, {
      op: 'setCellValue',
      sheet: 'S',
      ref: 'D9',
      value: { type: 'string', value: 'new' },
    });
    const after = heights(store.package);
    // Row 9 changed, and row 40 named an item that did not exist.
    expect(work.measured).toBe(2);
    expect(after).toEqual(heights(readXlsxAgain(store.package)));
  });

  test('each new text measures only its row, and the old item lists leave memory', () => {
    const commits = 40;
    const { store, lists } = commitTexts(commits);
    // The row height cache keeps no list. The collector scans the stack conservatively, so a few
    // lists can stay; a cache that keeps each list keeps all of them.
    expect(alive(lists)).toBeLessThan(commits / 4);
    // The rows of the store are still cached: the next layout measures none of them.
    work.reset();
    const cached = heights(store.package);
    expect(work.measured).toBe(0);
    expect(cached).toEqual(heights(readXlsxAgain(store.package)));
  });

  test('each new text in a wrapped cell lets the old item lists leave memory', () => {
    const commits = 40;
    const { store, lists } = commitTexts(commits, true);
    // The undo history keeps each old wrapped cell, and the cache of wrapped cells keeps no list.
    expect(alive(lists)).toBeLessThan(commits / 4);
    expect(heights(store.package)).toEqual(heights(readXlsxAgain(store.package)));
  });

  test('the cached wrapped cells let an old measurer leave memory', async () => {
    // Each workbook stays in memory with its wrapped cell, laid out once with its own measurer,
    // as a sheet that is not laid out again after a font change.
    const count = 8;
    const workbooks: XlsxPackage[] = [];
    const measurers: WeakRef<object>[] = [];
    for (let i = 0; i < count; i += 1) {
      const pkg = workbook(3);
      const own = await substituteMeasurer();
      measurers.push(new WeakRef(own));
      layoutSheet(pkg, 'S', VIEW, { measurer: own, profile: PRINT_PROFILE });
      workbooks.push(pkg);
    }
    expect(alive(measurers)).toBeLessThan(count / 2);
    expect(workbooks.length).toBe(count);
  });

  test('a new measurer measures every row and each wrapped cell again', () => {
    // A font change gives the editor a new measurer (`editor-fonts.ts`).
    const pkg = workbook(3);
    const before = heights(pkg);
    work.reset();
    // Letters a quarter as wide, digits the same: the columns keep their width, so the wrapped
    // text takes fewer lines (the long text is at the row height limit), and a stale fitted
    // height shows.
    const narrow: TextMeasurer = {
      resolve: (request) => measurer.resolve(request),
      advances: (face, text, options) => {
        // A copy: the measurer can return the same array again.
        const out = Float64Array.from(measurer.advances(face, text, options));
        for (let i = 0; i < text.length; i += 1) if (/[a-z]/.test(text[i]!)) out[i]! /= 4;
        return out;
      },
    };
    const after = heights(pkg, PRINT_PROFILE, narrow);
    expect(work.measured).toBe(ROWS);
    expect(after[2]).toBeLessThan(before[2]!);
    expect(after).toEqual(heights(readXlsxAgain(pkg), PRINT_PROFILE, narrow));
  });

  test('a row with a wrapped cell is measured at each layout', () => {
    const store = new WorkbookDocumentStore(workbook(3));
    const before = heights(store.package);
    expect(before[2]).toBeGreaterThan(before[0]!);
    work.reset();
    edit(store, { op: 'setCellValue', sheet: 'S', ref: 'A9', value: { type: 'number', value: 1 } });
    const after = heights(store.package);
    expect(after).toEqual(before);
    expect(work.measured).toBe(2);
    expect(after).toEqual(heights(readXlsxAgain(store.package)));
  });

  test('a profile with other values measures again, and a copy of the same values does not', () => {
    const pkg = workbook();
    const print = heights(pkg);
    work.reset();
    expect(heights(pkg, { ...PRINT_PROFILE })).toEqual(print);
    expect(work.measured).toBe(0);
    const screen = heights(pkg, SCREEN_PROFILE);
    expect(work.measured).toBe(ROWS);
    expect(screen).not.toEqual(print);
  });
});

/** The same package read again from its saved bytes, so no element is in the cache. */
function readXlsxAgain(pkg: XlsxPackage): XlsxPackage {
  const read = readXlsx(saveXlsx(pkg));
  if (!read.ok) throw new Error('read');
  return read.pkg;
}
