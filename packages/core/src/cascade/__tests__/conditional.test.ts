import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx, type XlsxPackage } from '../../store/package/package.ts';
import { xmlPartOf } from '../../store/package/package.ts';
import { sheetByName, workbookIndex } from '../../store/workbook/workbook-index.ts';
import { cellKey, sheetIndex } from '../../store/workbook/sheet-index.ts';
import { attr } from '../../store/package/query.ts';
import { parseCellRef } from '../../store/cell-ref.ts';
import { cellFormat, type ResolvedFill } from '../cell-format.ts';
import { cellValue } from '../values.ts';
import {
  conditionalFormat,
  literalOperand,
  rowCacheSizeForTest,
  rowEntriesForTest,
  MAX_ENTRIES_PER_ROW,
  MAX_PAIRS,
  MAX_RANGES,
  MAX_RULES,
  MAX_RULES_PER_CELL,
  sheetHasHiddenRules,
  sheetRules,
  sqrefRanges,
} from '../conditional.ts';

function open(bytes: Uint8Array): XlsxPackage {
  const r = readXlsx(bytes);
  if (!r.ok) throw new Error(JSON.stringify(r.error));
  return r.pkg;
}

/** The shown format of one cell, with its conditional formatting. */
function shownAt(pkg: XlsxPackage, sheet: string, ref: string) {
  const index = workbookIndex(pkg)!;
  const part = xmlPartOf(pkg, sheetByName(index, sheet)!.partName)!;
  const a = parseCellRef(ref)!;
  const si = sheetIndex(part);
  const el = si.cells.get(cellKey(a.row, a.col))!.el;
  const base = cellFormat(pkg, index, attr(el, 's')).format;
  const value = cellValue(pkg, index, el, false).value;
  return conditionalFormat(pkg, index, si.root, a.row, a.col, value, base);
}

function rootOf(pkg: XlsxPackage, sheet = 'S') {
  const index = workbookIndex(pkg)!;
  return sheetIndex(xmlPartOf(pkg, sheetByName(index, sheet)!.partName)!).root;
}

const cf = (sqref: string, rules: string) =>
  `<conditionalFormatting sqref="${sqref}">${rules}</conditionalFormatting>`;
const cellIs = (op: string, dxf: number, priority: number, ...formulas: string[]) =>
  `<cfRule type="cellIs" dxfId="${dxf}" priority="${priority}" operator="${op}">` +
  formulas.map((f) => `<formula>${f}</formula>`).join('') +
  '</cfRule>';

const DXFS = [
  '<dxf><font><b/></font></dxf>',
  '<dxf><font><i/><color rgb="FFFF0000"/></font></dxf>',
  '<dxf><fill><patternFill><bgColor rgb="FF00FF00"/></patternFill></fill></dxf>',
  '<dxf><fill><patternFill><fgColor rgb="FF0000FF"/></patternFill></fill></dxf>',
  '<dxf><numFmt numFmtId="200" formatCode="0.00"/></dxf>',
  '<dxf><font><b val="0"/><strike/></font></dxf>',
  '<dxf><font><vertAlign val="superscript"/></font></dxf>',
];

function book(rules: string, cells = [c('A1', 5), c('A2', 'abc'), c('A3', true)]) {
  return open(
    buildXlsx({
      sheets: [
        { name: 'S', rows: cells.map((x, i) => row(i + 1, [x])).join(''), afterMerges: rules },
      ],
      styles: { dxfs: DXFS },
    })
  );
}

describe('conditional formatting rules', () => {
  test('the preserved-parts rule makes the value 5 bold, as the reference PDF shows', () => {
    const pkg = open(
      new Uint8Array(
        readFileSync(`${import.meta.dir}/../../../../../fixtures/preserved-parts.xlsx`)
      )
    );
    expect(shownAt(pkg, 'Kept', 'B1').font.bold).toBe(true);
    expect(shownAt(pkg, 'Kept', 'A1').font.bold).toBe(false);
  });

  test('cellIs compares numbers, strings without case, and booleans', () => {
    const pkg = book(
      cf('A1', cellIs('between', 0, 1, '10', '4')) +
        cf('A2', cellIs('equal', 0, 2, '"ABC"')) +
        cf('A3', cellIs('equal', 0, 3, 'TRUE'))
    );
    expect(shownAt(pkg, 'S', 'A1').font.bold).toBe(true);
    expect(shownAt(pkg, 'S', 'A2').font.bold).toBe(true);
    expect(shownAt(pkg, 'S', 'A3').font.bold).toBe(true);
  });

  test('numbers order before strings, and strings before booleans', () => {
    const pkg = book(
      cf('A1', cellIs('lessThan', 0, 1, '"a"')) + cf('A2', cellIs('lessThan', 0, 2, 'FALSE'))
    );
    expect(shownAt(pkg, 'S', 'A1').font.bold).toBe(true);
    expect(shownAt(pkg, 'S', 'A2').font.bold).toBe(true);
  });

  test('the first true rule by priority wins each property', () => {
    // Priority 1 sets italic and red; priority 2 sets bold. Both apply, red from priority 1.
    const pkg = book(cf('A1', cellIs('greaterThan', 0, 2, '1') + cellIs('greaterThan', 1, 1, '1')));
    const f = shownAt(pkg, 'S', 'A1').font;
    expect([f.bold, f.italic, f.color]).toEqual([true, true, 'FF0000']);
    // Priority 1 turns bold off; the bold of priority 2 does not show.
    const off = book(cf('A1', cellIs('greaterThan', 0, 2, '1') + cellIs('greaterThan', 5, 1, '1')));
    expect(shownAt(off, 'S', 'A1').font.bold).toBe(false);
    expect(shownAt(off, 'S', 'A1').font.strike).toBe(true);
  });

  test('a differential x:vertAlign sets the vertical alignment of the font (unmeasured)', () => {
    const pkg = book(cf('A1', cellIs('greaterThan', 6, 1, '1')));
    expect(shownAt(pkg, 'S', 'A1').font.vertAlign).toBe('superscript');
  });

  test('a true rule with stopIfTrue ends the walk', () => {
    const stop =
      '<cfRule type="cellIs" dxfId="1" priority="1" operator="greaterThan" stopIfTrue="1">' +
      '<formula>1</formula></cfRule>';
    const pkg = book(cf('A1', stop + cellIs('greaterThan', 0, 2, '1')));
    expect(shownAt(pkg, 'S', 'A1').font.bold).toBe(false);
    expect(shownAt(pkg, 'S', 'A1').font.italic).toBe(true);
  });

  test('a solid fill paints its bgColor; a fill with only fgColor paints white over no fill', () => {
    const pkg = book(
      cf('A1', cellIs('greaterThan', 2, 1, '1')) + cf('A2', cellIs('equal', 3, 1, '"abc"'))
    );
    expect(shownAt(pkg, 'S', 'A1').fill).toEqual({ kind: 'solid', color: '00FF00' });
    expect(shownAt(pkg, 'S', 'A2').fill).toEqual({ kind: 'solid', color: 'FFFFFF' });
  });

  test('an x:numFmt replaces the number format', () => {
    const f = shownAt(book(cf('A1', cellIs('greaterThan', 4, 1, '1'))), 'S', 'A1');
    expect([f.numFmtId, f.formatCode]).toEqual([200, '0.00']);
  });

  test('a false rule, a rule outside its range, and a missing dxf change nothing', () => {
    const pkg = book(
      cf('A1', cellIs('lessThan', 0, 1, '1')) +
        cf('B1', cellIs('greaterThan', 1, 2, '1')) +
        cf('A1', cellIs('greaterThan', 99, 3, '1'))
    );
    const f = shownAt(pkg, 'S', 'A1');
    expect([f.font.bold, f.font.italic]).toEqual([false, false]);
  });

  test('error and date values make every rule false; an empty value is 0 or the empty string', () => {
    const cells = [
      c('A1', '#DIV/0!', { t: 'e' }),
      c('A2', '2024-01-02T00:00:00', { t: 'd' }),
      '<c r="A3" s="0"/>',
    ];
    const rules =
      cf('A1:A3', cellIs('notEqual', 0, 1, '5')) +
      cf('A1:A3', cellIs('notBetween', 1, 2, '1', '2'));
    const pkg = book(rules, cells);
    for (const ref of ['A1', 'A2']) {
      const f = shownAt(pkg, 'S', ref).font;
      expect([ref, f.bold, f.italic]).toEqual([ref, false, false]);
    }
    // conditional-rules.pdf, B16 to B18: an empty cell is 0 against a number.
    const blank = shownAt(pkg, 'S', 'A3').font;
    expect([blank.bold, blank.italic]).toEqual([true, true]);
    const text = book(
      cf('A3', cellIs('equal', 0, 1, '""')) + cf('A3', cellIs('equal', 1, 2, 'FALSE')),
      cells
    );
    expect([shownAt(text, 'S', 'A3').font.bold, shownAt(text, 'S', 'A3').font.italic]).toEqual([
      true,
      false,
    ]);
  });

  test('rules in the extension content of a worksheet are not read', () => {
    const ext =
      '<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
      '<x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="cellIs" priority="1" operator="greaterThan"><xm:f xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">1</xm:f></x14:cfRule><xm:sqref xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">A1</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>';
    const pkg = open(
      buildXlsx({
        sheets: [{ name: 'S', rows: row(1, [c('A1', 5)]), ext: ext.slice(8, -9) }],
        styles: { dxfs: DXFS },
      })
    );
    expect(sheetRules(rootOf(pkg)).rules).toEqual([]);
    expect(shownAt(pkg, 'S', 'A1').font.bold).toBe(false);
  });

  test('rules that need a formula, other types, and malformed rules are not shown', () => {
    const pkg = book(
      cf('A1', cellIs('greaterThan', 0, 1, '$B$1')) +
        cf(
          'A1',
          '<cfRule type="expression" dxfId="0" priority="2"><formula>TRUE</formula></cfRule>'
        ) +
        cf('A1', cellIs('between', 0, 3, '1')) +
        cf(
          'A1',
          '<cfRule type="cellIs" dxfId="0" operator="greaterThan"><formula>1</formula></cfRule>'
        ) +
        cf('Z', cellIs('greaterThan', 0, 5, '1'))
    );
    expect(sheetRules(rootOf(pkg)).rules.map((r) => r.notShown)).toEqual([
      'needs-formula',
      'unsupported',
      'malformed',
      'malformed',
      'malformed',
    ]);
    expect(shownAt(pkg, 'S', 'A1').font.bold).toBe(false);
  });

  test('a rule that is not shown does not stop a lower rule', () => {
    const stop =
      '<cfRule type="expression" dxfId="1" priority="1" stopIfTrue="1"><formula>TRUE</formula></cfRule>';
    const pkg = book(cf('A1', stop + cellIs('greaterThan', 0, 2, '1')));
    expect(shownAt(pkg, 'S', 'A1').font.bold).toBe(true);
  });

  test('literal operands', () => {
    expect(literalOperand(' -1.5e2 ')).toEqual({ kind: 'number', value: -150 });
    expect(literalOperand('"A""b"')).toEqual({ kind: 'string', text: 'A"b', lower: 'a"b' });
    expect(literalOperand('"a"b"')).toBeUndefined();
    expect(literalOperand('true')).toEqual({ kind: 'boolean', value: true });
    expect(literalOperand('A1')).toBeUndefined();
    expect(literalOperand('\u00a01')).toBeUndefined();
    expect(literalOperand('1+1')).toBeUndefined();
  });
});

describe('conditional formatting limits', () => {
  test('an @sqref with too many ranges is not read', () => {
    const many = Array.from({ length: MAX_RANGES + 1 }, (_, i) => `A${i + 1}`).join(' ');
    expect(sqrefRanges(many)).toBeUndefined();
    expect(sqrefRanges(many.split(' ').slice(0, MAX_RANGES).join(' '))?.length).toBe(MAX_RANGES);
  });

  test('rules past the sheet limit are skipped and counted, never built', () => {
    const rules = Array.from({ length: MAX_RULES + 3 }, (_, i) =>
      cellIs('equal', 0, i + 1, '1')
    ).join('');
    const pkg = book(cf('A1', rules));
    const read = sheetRules(rootOf(pkg));
    expect(read.rules.length).toBe(MAX_RULES);
    expect(read.skipped).toBe(3);
    // The notice covers rules past the limits. More than 64 shown rules also give it, so this
    // cause never shows alone: the `skipped` term keeps the notice when a limit changes.
    expect(sheetHasHiddenRules(rootOf(pkg))).toBe(true);
  });

  test('whole-column, whole-row, absolute, and lowercase ranges are malformed, with the notice', () => {
    for (const sqref of ['B:B', '1:1', '$A$1', 'a1']) {
      const root = rootOf(book(cf(sqref, cellIs('greaterThan', 0, 1, '3'))));
      expect({ sqref, why: sheetRules(root).rules[0]!.notShown }).toEqual({
        sqref,
        why: 'malformed',
      });
      expect(sheetHasHiddenRules(root)).toBe(true);
    }
  });

  test('more shown ranges than one row checks give the notice', () => {
    // Shown rules of 512, 512, and n ranges: no rule is past a limit or not shown, and three
    // rules stay under the cap of one cell, so only the cap of one row can give the notice.
    const ranges = (from: number, n: number) =>
      Array.from({ length: n }, (_, i) => `A${from + i}`).join(' ');
    const sheetWith = (n: number) => {
      const third = n > 0 ? cf(ranges(1025, n), cellIs('equal', 0, 3, '1')) : '';
      return rootOf(
        book(
          cf(ranges(1, 512), cellIs('equal', 0, 1, '1')) +
            cf(ranges(513, 512), cellIs('equal', 0, 2, '1')) +
            third
        )
      );
    };
    expect(MAX_ENTRIES_PER_ROW).toBe(1_024);
    const over = sheetWith(1);
    expect(sheetRules(over).skipped).toBe(0);
    expect(sheetRules(over).rules.every((r) => r.test !== undefined)).toBe(true);
    expect(sheetHasHiddenRules(over)).toBe(true);
    expect(sheetHasHiddenRules(sheetWith(0))).toBe(false);
  });

  test('a huge range reference is stored as four bounds', () => {
    const pkg = book(cf('A1:XFD1048576', cellIs('greaterThan', 0, 1, '1')));
    expect(sheetRules(rootOf(pkg)).rules[0]!.ranges).toEqual([
      { firstRow: 0, firstCol: 0, lastRow: 1_048_575, lastCol: 16_383 },
    ]);
    expect(shownAt(pkg, 'S', 'A1').font.bold).toBe(true);
  });

  test('a rule whose @sqref holds too many ranges is malformed', () => {
    const many = Array.from({ length: MAX_RANGES + 1 }, (_, i) => `B${i + 1}`).join(' ');
    const pkg = book(cf(many, cellIs('greaterThan', 0, 1, '1')));
    expect(sheetRules(rootOf(pkg)).rules[0]!.notShown).toBe('malformed');
  });

  test('an @sqref longer than 16,384 characters is not read', () => {
    // Spaces between ranges count toward the length; one range keeps the range limit away.
    expect(sqrefRanges('A1'.padEnd(16_384, ' '))?.length).toBe(1);
    expect(sqrefRanges('A1'.padEnd(16_385, ' '))).toBeUndefined();
  });

  test('an operand longer than 8,192 characters is not a literal', () => {
    expect(literalOperand('1'.padStart(8_192, ' '))).toEqual({ kind: 'number', value: 1 });
    expect(literalOperand('1'.padStart(8_193, ' '))).toBeUndefined();
  });

  test('rules past the pair limit are skipped and counted', () => {
    // 128 rules of 512 ranges reach the pair limit; the next rule is skipped.
    const ranges = Array.from({ length: MAX_RANGES }, (_, i) => `C${i + 1}`).join(' ');
    const full = Array.from({ length: MAX_PAIRS / MAX_RANGES }, (_, i) =>
      cf(ranges, cellIs('equal', 0, i + 1, '1'))
    ).join('');
    const read = sheetRules(rootOf(book(full + cf('A1', cellIs('equal', 0, 999, '5')))));
    expect(read.rules.length).toBe(MAX_PAIRS / MAX_RANGES);
    expect(read.skipped).toBe(1);
  });

  test('at most 64 shown rules over one cell are walked', () => {
    // 64 false rules come first by priority; the true rule after them is not walked.
    const falses = Array.from({ length: MAX_RULES_PER_CELL }, (_, i) =>
      cellIs('lessThan', 0, i + 1, '0')
    ).join('');
    const pkg = book(cf('A1', falses + cellIs('greaterThan', 0, 100, '1')));
    expect(shownAt(pkg, 'S', 'A1').font.bold).toBe(false);
    const fewer = falses.slice(falses.indexOf('</cfRule>') + '</cfRule>'.length);
    const shown = book(cf('A1', fewer + cellIs('greaterThan', 0, 100, '1')));
    expect(shownAt(shown, 'S', 'A1').font.bold).toBe(true);
  });

  test('at most 1,024 rule ranges of a row are checked for one cell', () => {
    // 1,024 ranges of the row come first by priority and miss column A.
    const ranges = Array.from({ length: MAX_RANGES }, (_, i) => `B1:B${i + 2}`).join(' ');
    const misses = Array.from({ length: MAX_ENTRIES_PER_ROW / MAX_RANGES }, (_, i) =>
      cf(ranges, cellIs('equal', 0, i + 1, '1'))
    ).join('');
    const pkg = book(misses + cf('A1', cellIs('greaterThan', 0, 100, '1')));
    expect(shownAt(pkg, 'S', 'A1').font.bold).toBe(false);
    const one = misses.slice(misses.indexOf('</conditionalFormatting>') + 24);
    expect(
      shownAt(book(one + cf('A1', cellIs('greaterThan', 0, 100, '1'))), 'S', 'A1').font.bold
    ).toBe(true);
  });

  test('a rule with two ranges over one cell is walked once', () => {
    // 33 false rules with two ranges over A1 each, then one true rule. Counted once for each
    // range, the walk would stop at 64 before the true rule.
    const falses = Array.from({ length: 33 }, (_, i) => cellIs('lessThan', 0, i + 1, '0')).join('');
    const pkg = book(cf('A1 A1:A2', falses) + cf('A1', cellIs('greaterThan', 0, 100, '1')));
    expect(shownAt(pkg, 'S', 'A1').font.bold).toBe(true);
  });

  test('a row keeps at most 1,024 rule ranges in the cache', () => {
    const ranges = Array.from({ length: MAX_RANGES }, () => 'A1:XFD1048576').join(' ');
    const rules = Array.from({ length: 8 }, (_, i) =>
      cf(ranges, cellIs('equal', 0, i + 1, '1'))
    ).join('');
    const sheet = sheetRules(rootOf(book(rules)));
    expect(rowEntriesForTest(sheet, 0)).toBe(MAX_ENTRIES_PER_ROW);
    expect(rowEntriesForTest(sheet, 1_000_000)).toBe(MAX_ENTRIES_PER_ROW);
  });

  test('priority is an xsd:int, and ranges split on any XML white space', () => {
    const rule =
      '<cfRule type="cellIs" dxfId="0" priority="-2147483648" operator="greaterThan"><formula>1</formula></cfRule>';
    const pkg = book(cf('B9\tA1\nC3', rule));
    const read = sheetRules(rootOf(pkg)).rules[0]!;
    expect([read.priority, read.notShown, read.ranges.length]).toEqual([
      -2_147_483_648,
      undefined,
      3,
    ]);
    expect(shownAt(pkg, 'S', 'A1').font.bold).toBe(true);
  });

  test('dxfId and priority take a plus sign and leading zeros, as the schema types allow', () => {
    const rule =
      '<cfRule type="cellIs" dxfId="+00" priority="+00000000001" operator="greaterThan"><formula>1</formula></cfRule>';
    const pkg = book(cf('A1', rule));
    const read = sheetRules(rootOf(pkg)).rules[0]!;
    expect([read.priority, read.notShown]).toEqual([1, undefined]);
    expect(shownAt(pkg, 'S', 'A1').font.bold).toBe(true);
    // Past the range of the type, the rule is malformed.
    const big = rule.replace('dxfId="+00"', 'dxfId="04294967296"');
    expect(sheetRules(rootOf(book(cf('A1', big)))).rules[0]!.notShown).toBe('malformed');
  });

  test('a quoted string operand of 8,193 characters needs a formula', () => {
    const long = `"${'a'.repeat(8_191)}"`;
    expect(long.length).toBe(8_193);
    const pkg = book(cf('A2', cellIs('equal', 0, 1, long)));
    expect(sheetRules(rootOf(pkg)).rules[0]!.notShown).toBe('needs-formula');
  });

  test('the cache of a sheet clears past 262,144 rule ranges, and its results stay right', () => {
    const ranges = Array.from({ length: MAX_RANGES }, () => 'A1:XFD1048576').join(' ');
    const rules = Array.from({ length: 2 }, (_, i) =>
      cf(ranges, cellIs('equal', 0, i + 1, '1'))
    ).join('');
    const sheet = sheetRules(rootOf(book(rules)));
    // 256 rows of 1,024 ranges fill the cache; the next row clears it first.
    for (let r = 0; r < 256; r++) rowEntriesForTest(sheet, r);
    expect(rowCacheSizeForTest(sheet)).toEqual({ rows: 256, entries: 262_144 });
    rowEntriesForTest(sheet, 256);
    expect(rowCacheSizeForTest(sheet)).toEqual({ rows: 1, entries: 1_024 });
    expect(rowEntriesForTest(sheet, 0)).toBe(MAX_ENTRIES_PER_ROW);
  });

  test('white space around type, operator, priority, dxfId, and stopIfTrue is collapsed', () => {
    const rule =
      '<cfRule type=" cellIs " dxfId=" 0 " priority=" 1 " stopIfTrue=" 1 " operator=" greaterThan "><formula>1</formula></cfRule>';
    const pkg = book(cf('A1', rule + cellIs('greaterThan', 1, 2, '1')));
    const f = shownAt(pkg, 'S', 'A1').font;
    expect([f.bold, f.italic]).toEqual([true, false]);
  });

  test('a text operator on a cellIs rule is unsupported', () => {
    const pkg = book(cf('A1', cellIs('containsText', 0, 1, '1')));
    expect(sheetRules(rootOf(pkg)).rules[0]!.notShown).toBe('unsupported');
  });

  test('a fill with only fgColor keeps the fill of the cell format', () => {
    const pkg = open(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            rows: row(1, [c('A1', 5, { s: 1 })]),
            afterMerges: cf('A1', cellIs('greaterThan', 3, 1, '1')),
          },
        ],
        styles: {
          fills: [
            '<fill><patternFill patternType="solid"><fgColor rgb="FF112233"/></patternFill></fill>',
          ],
          cellXfs: ['<xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"/>'],
          dxfs: DXFS,
        },
      })
    );
    expect(shownAt(pkg, 'S', 'A1').fill).toEqual({ kind: 'solid', color: '112233' });
  });

  test('a none, a pattern, and a gradient differential fill keep the fill of the cell', () => {
    const fills = [
      '<dxf><fill><patternFill patternType="none"/></fill></dxf>',
      '<dxf><fill><patternFill patternType="gray125"><bgColor rgb="FFFF0000"/></patternFill></fill></dxf>',
      '<dxf><fill><gradientFill><stop position="0"><color rgb="FFFF0000"/></stop><stop position="1"><color rgb="FF0000FF"/></stop></gradientFill></fill></dxf>',
    ];
    for (const own of [true, false]) {
      const pkg = open(
        buildXlsx({
          sheets: [
            {
              name: 'S',
              rows: row(1, [
                c('A1', 5, own ? { s: 1 } : {}),
                c('B1', 5, own ? { s: 1 } : {}),
                c('C1', 5, own ? { s: 1 } : {}),
              ]),
              afterMerges:
                cf('A1', cellIs('greaterThan', 0, 1, '1')) +
                cf('B1', cellIs('greaterThan', 1, 2, '1')) +
                cf('C1', cellIs('greaterThan', 2, 3, '1')),
            },
          ],
          styles: {
            fills: [
              '<fill><patternFill patternType="solid"><fgColor rgb="FF112233"/></patternFill></fill>',
            ],
            cellXfs: ['<xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"/>'],
            dxfs: fills,
          },
        })
      );
      const want: ResolvedFill = own ? { kind: 'solid', color: '112233' } : { kind: 'none' };
      for (const ref of ['A1', 'B1', 'C1'])
        expect([ref, shownAt(pkg, 'S', ref).fill]).toEqual([ref, want]);
    }
  });

  test('a differential font sets the family and the size', () => {
    const pkg = open(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            rows: row(1, [c('A1', 5)]),
            afterMerges: cf('A1', cellIs('greaterThan', 0, 1, '1')),
          },
        ],
        styles: { dxfs: ['<dxf><font><sz val="20"/><name val="Cambria"/></font></dxf>'] },
      })
    );
    const f = shownAt(pkg, 'S', 'A1').font;
    expect([f.family, f.sizePt]).toEqual(['Cambria', 20]);
  });

  test('notBetween with a boolean bound is false for an empty value', () => {
    const pkg = book(cf('A1', cellIs('notBetween', 0, 1, '1', 'TRUE')), ['<c r="A1"/>']);
    expect(shownAt(pkg, 'S', 'A1').font.bold).toBe(false);
  });
});
