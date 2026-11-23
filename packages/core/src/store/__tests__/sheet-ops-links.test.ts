// The hyperlink relationship targets of the sheet ops of design D1 of the change
// `workbook-structure-editing`: a target that is a fragment (`#Data!A1`) names a place in the
// workbook, as `x:hyperlink/@location` does, so a rename rewrites it, a twin rename refuses,
// and a delete gives the result of a location. A row edit moves it too (design D4). Also the
// cost of the mention scan of a move and a delete.

import { describe, expect, test } from 'bun:test';
import { c, row, type FixtureWorkbook } from '../../../../../scripts/fixtures/builder.ts';
import { attr, child, children } from '../package/query.ts';
import { readXlsx, xmlPartOf, type XlsxPackage } from '../package/package.ts';
import { saveXlsx } from '../package-write/save.ts';
import type { WorkbookOp } from '../store/ops.ts';
import { mentionScanWork } from '../structure/blockers.ts';
import { sheetRoot, storeOf } from './structure-helpers.ts';

const H = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
const RELS = '/xl/worksheets/_rels/sheet2.xml.rels';

/** `Data`, and `Other` with a hyperlink whose relationship target is `target`. */
const linked = (target: string, extra: FixtureWorkbook['sheets'] = []): FixtureWorkbook => ({
  sheets: [
    { name: 'Data', rows: row(1, [c('A1', 1)]) },
    {
      name: 'Other',
      rows: row(1, [c('A1', 9)]),
      afterMerges: '<hyperlinks><hyperlink ref="A1" r:id="rIdH"/></hyperlinks>',
      rels: [['rIdH', H, target, true]],
    },
    ...extra,
  ],
});

/** The `Target` of each relationship of `Other`, by id. */
function targets(pkg: XlsxPackage): Record<string, string> {
  const out: Record<string, string> = {};
  for (const el of xmlPartOf(pkg, RELS)?.root.children ?? [])
    if (el.type === 'element') out[attr(el, 'Id') ?? ''] = attr(el, 'Target') ?? '';
  return out;
}

const link = (pkg: XlsxPackage) =>
  children(child(sheetRoot(pkg, 'Other'), 'x:hyperlinks'), 'x:hyperlink')[0]!;

function commit(wb: FixtureWorkbook, ops: readonly WorkbookOp[]) {
  const s = storeOf(wb);
  const before = s.package;
  const r = s.transact(ops);
  return { s, before, r };
}

describe('hyperlink relationship targets', () => {
  test('a rename rewrites a target that names the sheet, and undo puts it back', () => {
    const { s, r } = commit(linked('#Data!A1'), [
      { op: 'renameSheet', sheet: 'Data', name: "Mine's" },
    ]);
    expect(r.ok).toBe(true);
    expect(targets(s.package)).toEqual({ rIdH: "#'Mine''s'!A1" });
    const reopened = readXlsx(saveXlsx(s.package));
    expect(reopened.ok && targets(reopened.pkg)).toEqual({ rIdH: "#'Mine''s'!A1" });
    expect(s.undo().ok).toBe(true);
    expect(targets(s.package)).toEqual({ rIdH: '#Data!A1' });
  });

  // A fragment holds only the characters of RFC 3986, section 3.5. A space or a character
  // outside ASCII is not one of them, so the place goes to `@location`, as `#` and `%` do.
  test('a rename to a name with a space or a character outside ASCII gives the location', () => {
    for (const name of ['Q 1', 'Zäh']) {
      const { s, before, r } = commit(linked('#Data!A1'), [
        { op: 'renameSheet', sheet: 'Data', name },
      ]);
      expect(r.ok).toBe(true);
      expect(targets(s.package)).toEqual({});
      expect(attr(link(s.package), 'location')).toBe(`'${name}'!A1`);
      expect(attr(link(s.package), 'r:id')).toBeUndefined();
      const reopened = readXlsx(saveXlsx(s.package));
      expect(reopened.ok && attr(link(reopened.pkg), 'location')).toBe(`'${name}'!A1`);
      expect(s.undo().ok).toBe(true);
      expect(xmlPartOf(s.package, RELS)).toBe(xmlPartOf(before, RELS)!);
    }
  });

  test('a rename keeps a target that names another sheet or an outside address', () => {
    const { s, before, r } = commit(linked('#Other!B2'), [
      { op: 'renameSheet', sheet: 'Data', name: 'X' },
    ]);
    expect(r.ok).toBe(true);
    expect(xmlPartOf(s.package, RELS)).toBe(xmlPartOf(before, RELS)!);
    const web = commit(linked('https://example.com/#Data!A1'), [
      { op: 'renameSheet', sheet: 'Data', name: 'X' },
    ]);
    expect(targets(web.s.package)).toEqual({ rIdH: 'https://example.com/#Data!A1' });
  });

  test('a rename of one of two sheets with the same name without case refuses', () => {
    const twin = linked('#Data!A1', [{ name: 'DATA', rows: row(1, [c('A1', 2)]) }]);
    const { s, before, r } = commit(twin, [{ op: 'renameSheet', sheet: 'Data', name: 'X' }]);
    expect(r.ok ? null : [r.rejection.code, r.rejection.subject]).toEqual([
      'sheet-name-ambiguous',
      'DATA',
    ]);
    expect(s.package).toBe(before);
  });

  test('a delete gives the hyperlink the location #REF!, and its relationship goes', () => {
    const { s, r } = commit(linked('#Data!A1:B2'), [{ op: 'deleteSheet', sheet: 'Data' }]);
    expect(r.ok).toBe(true);
    expect(attr(link(s.package), 'location')).toBe('#REF!');
    expect(attr(link(s.package), 'r:id')).toBeUndefined();
    expect(targets(s.package)).toEqual({});
    const reopened = readXlsx(saveXlsx(s.package));
    expect(reopened.ok && attr(link(reopened.pkg), 'location')).toBe('#REF!');
    expect(s.undo().ok).toBe(true);
    expect(targets(s.package)).toEqual({ rIdH: '#Data!A1:B2' });
    expect(attr(link(s.package), 'r:id')).toBe('rIdH');
  });

  test('a delete refuses a hyperlink that also has a location', () => {
    const base = linked('#Data!A1');
    const wb: FixtureWorkbook = {
      sheets: base.sheets.map((sheet, i) =>
        i === 1
          ? {
              ...sheet,
              afterMerges:
                '<hyperlinks><hyperlink ref="A1" r:id="rIdH" location="A2"/></hyperlinks>',
            }
          : sheet
      ),
    };
    const { r } = commit(wb, [{ op: 'deleteSheet', sheet: 'Data' }]);
    expect(r.ok ? null : [r.rejection.code, r.rejection.subject]).toEqual([
      'structure-blocked',
      'Other!x:hyperlink',
    ]);
  });

  test('a target with a % escape that can name the sheet refuses a rename', () => {
    const { r } = commit(linked('#%44ata!A1'), [{ op: 'renameSheet', sheet: 'Data', name: 'X' }]);
    expect(r.ok ? null : r.rejection.code).toBe('structure-blocked');
    const other = commit(linked('#Other!A1%20'), [{ op: 'renameSheet', sheet: 'Data', name: 'X' }]);
    expect(other.r.ok).toBe(true);
  });
});

describe('hyperlink relationship targets of a row edit', () => {
  test('an insertion moves a target that names a moved cell', () => {
    const { s, r } = commit(linked('#Data!A5'), [
      { op: 'insertRows', sheet: 'Data', start: 1, count: 2 },
    ]);
    expect(r.ok).toBe(true);
    expect(targets(s.package)).toEqual({ rIdH: '#Data!A7' });
    expect(s.undo().ok).toBe(true);
    expect(targets(s.package)).toEqual({ rIdH: '#Data!A5' });
  });

  test('a deletion of the named cell gives the hyperlink the location', () => {
    const { s, r } = commit(linked('#Data!A5'), [
      { op: 'deleteRows', sheet: 'Data', start: 5, count: 1 },
    ]);
    expect(r.ok).toBe(true);
    expect(attr(link(s.package), 'location')).toBe('Data!#REF!');
    expect(targets(s.package)).toEqual({});
  });
});

describe('the cost of the mention scan', () => {
  // The cost is a count of work, never a time: `chars` is the folded text that the scan reads,
  // and `steps` the characters that the sheet range walks read. Each op reads each scanned text
  // once, so `chars` stays within the XML text of the sheets. A walk reads at most a few
  // characters for each character of the text, whatever the number of sheets and spellings.
  const STEPS_PER_CHAR = 8;
  const xmlText = (sheets: FixtureWorkbook['sheets']) =>
    sheets.reduce((sum, sh) => sum + (sh.afterMerges ?? '').length, 0);
  const measure = (run: () => void) => {
    mentionScanWork.chars = 0;
    mentionScanWork.steps = 0;
    run();
    return { ...mentionScanWork };
  };

  // One formula with many `:` characters, which the sheet range scan reads at each `:`.
  const colons = (count: number) =>
    `<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1" dxfId="0"><formula>SUM(${Array.from(
      { length: count },
      (_, i) => `B${i + 1}:B${i + 2}`
    ).join(',')})</formula></cfRule></conditionalFormatting>`;

  // Every sheet holds conditional formatting, which the scan reads. A scan that searches each
  // text once for each sheet reads each text 2000 times here.
  test('a move and a delete of 1000 sheets with 100 formula elements each read each text once', () => {
    const cf = Array.from(
      { length: 100 },
      (_, i) =>
        `<conditionalFormatting sqref="A${i + 1}"><cfRule type="expression" priority="${i + 1}" dxfId="0"><formula>A1&gt;${i}</formula></cfRule></conditionalFormatting>`
    ).join('');
    const sheets = Array.from({ length: 1000 }, (_, i) => ({
      name: `Sheet${i}`,
      rows: row(1, [c('A1', 1)]),
      afterMerges: cf,
    }));
    const s = storeOf({ sheets });
    const work = measure(() => {
      expect(s.transact([{ op: 'moveSheet', sheet: 'Sheet0', index: 999 }]).ok).toBe(true);
      expect(s.transact([{ op: 'deleteSheet', sheet: 'Sheet500' }]).ok).toBe(true);
    });
    expect(work.chars).toBeGreaterThan(0);
    expect(work.chars).toBeLessThanOrEqual(2 * xmlText(sheets));
    expect(work.steps).toBeLessThanOrEqual(STEPS_PER_CHAR * work.chars);
  }, 60_000);

  // Sheet names of every length up to 31, with apostrophes that double in a quoted spelling,
  // and a formula with 50,000 `:` characters: the walk at each `:` reads at most 64 characters.
  test('many distinct name lengths and many colons keep the walk linear in the text', () => {
    const sheets = Array.from({ length: 600 }, (_, i) => ({
      name: `${i}`.padEnd(1 + (i % 31), "a'"),
      rows: row(1, [c('A1', 1)]),
      afterMerges: i === 1 ? colons(50_000) : '',
    }));
    const s = storeOf({ sheets });
    const work = measure(() => {
      expect(s.transact([{ op: 'moveSheet', sheet: '0', index: 599 }]).ok).toBe(true);
      expect(s.transact([{ op: 'deleteSheet', sheet: sheets[300]!.name }]).ok).toBe(true);
    });
    expect(work.chars).toBeGreaterThan(50_000);
    expect(work.chars).toBeLessThanOrEqual(2 * xmlText(sheets));
    expect(work.steps).toBeLessThanOrEqual(STEPS_PER_CHAR * work.chars);
  }, 60_000);

  // Sheet names of 30 lengths, so the spellings of each side have about 30 lengths, and a formula
  // text of 1,500,000 `:` characters. A look-up of each spelling length at each `:` reads about
  // 30 slices of up to 32 characters on each side. The walk back and forward from each `:`
  // stops at its second character, so the scan stays within the bound.
  test('many spelling lengths around 1,500,000 colons keep the walk linear in the text', () => {
    const last = 'ABCDEFGHIJKMNOPQRSTU';
    const name = (i: number) => `${'L'.repeat(1 + (i % 30))}${last[Math.floor(i / 30)]}`;
    const text = `<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1" dxfId="0"><formula>LEN("${':'.repeat(
      1_500_000
    )}")</formula></cfRule></conditionalFormatting>`;
    const sheets = Array.from({ length: 600 }, (_, i) => ({
      name: name(i),
      rows: row(1, [c('A1', 1)]),
      afterMerges: i === 599 ? text : '',
    }));
    const s = storeOf({ sheets });
    const work = measure(() => {
      expect(s.transact([{ op: 'deleteSheet', sheet: name(300) }]).ok).toBe(true);
      expect(s.transact([{ op: 'moveSheet', sheet: name(400), index: 100 }]).ok).toBe(true);
    });
    expect(work.chars).toBeGreaterThan(1_500_000);
    expect(work.chars).toBeLessThanOrEqual(2 * xmlText(sheets));
    expect(work.steps).toBeLessThanOrEqual(STEPS_PER_CHAR * work.chars);
  }, 60_000);

  // A file can store longer names. The search cost grows with the longest name, so a move, a
  // delete, and a row insertion refuse before the search, and a rename can shorten the name.
  test('a stored sheet name longer than 31 characters refuses the ops that search sheet ranges', () => {
    const sheets = Array.from({ length: 600 }, (_, i) => ({
      name: 'x'.repeat(i + 1),
      rows: row(1, [c('A1', 1)]),
      afterMerges: i === 0 ? colons(50_000) : '',
    }));
    const s = storeOf({ sheets });
    const results: ReturnType<typeof s.transact>[] = [];
    const work = measure(() => {
      results.push(s.transact([{ op: 'moveSheet', sheet: 'x', index: 599 }]));
      results.push(s.transact([{ op: 'deleteSheet', sheet: 'xx' }]));
      results.push(s.transact([{ op: 'insertRows', sheet: 'xx', start: 1, count: 1 }]));
    });
    const del = results[1]!;
    // Each op refuses before the scan reads any text.
    expect(work).toEqual({ chars: 0, steps: 0 });
    // Each names the long sheet by its full stored name, with its own code.
    expect(results).toHaveLength(3);
    for (const r of results) {
      expect(r.ok ? '' : r.rejection.code).toBe('sheet-name-too-long');
      expect(r.ok ? '' : r.rejection.subject).toBe('x'.repeat(32));
    }
    expect(del.ok ? '' : del.rejection.detail).toContain('more than 31 characters');
    expect(s.transact([{ op: 'renameSheet', sheet: 'x'.repeat(40), name: 'Short' }]).ok).toBe(true);
  }, 60_000);

  test('a long stored name with surrogate pairs is the whole subject, never cut inside a pair', () => {
    const long = '\u{1F600}'.repeat(20);
    const s = storeOf({
      sheets: [
        { name: 'A', rows: row(1, [c('A1', 1)]) },
        { name: long, rows: row(1, [c('A1', 1)]) },
      ],
    });
    const r = s.transact([{ op: 'deleteSheet', sheet: 'A' }]);
    expect(r.ok ? '' : r.rejection.code).toBe('sheet-name-too-long');
    expect(r.ok ? '' : r.rejection.subject).toBe(long);
  });
});
