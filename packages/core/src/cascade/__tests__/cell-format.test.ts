import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { STYLE_CASES } from '../../../../../scripts/fixtures/cases.ts';
import { STYLE_RESOLUTION_CASES } from '../../../../../scripts/fixtures/fill-style-cases.ts';
import { measuredCells } from '../../../../../scripts/fixtures/measured.ts';
import { FIXTURES } from '../../../../../scripts/fixtures/catalog.ts';
import { fixtureCopies } from '../../../../../scripts/save-check/records.ts';
import { readXlsx, type XlsxPackage } from '../../store/package/package.ts';
import { sheetByName, workbookIndex } from '../../store/workbook/workbook-index.ts';
import { xmlPartOf } from '../../store/package/package.ts';
import { cellKey, sheetIndex } from '../../store/workbook/sheet-index.ts';
import { attr } from '../../store/package/query.ts';
import { cellFormat } from '../cell-format.ts';
import { applyTint, resolveColor } from '../color.ts';
import { cellValue } from '../values.ts';
import { defaultFont, richTextRuns } from '../cell-format.ts';
import { parseCellRef } from '../../store/cell-ref.ts';

function open(bytes: Uint8Array): XlsxPackage {
  const r = readXlsx(bytes);
  if (!r.ok) throw new Error(JSON.stringify(r.error));
  return r.pkg;
}

const fixture = (name: string) =>
  open(new Uint8Array(readFileSync(`${import.meta.dir}/../../../../../fixtures/${name}.xlsx`)));

/** The resolved format of one cell of a sheet. */
function formatAt(pkg: XlsxPackage, sheet: string, ref: string) {
  const index = workbookIndex(pkg)!;
  const part = xmlPartOf(pkg, sheetByName(index, sheet)!.partName)!;
  const a = parseCellRef(ref)!;
  const cell = sheetIndex(part).cells.get(cellKey(a.row, a.col))!.el;
  return cellFormat(pkg, index, attr(cell, 's'));
}

/** The largest channel difference of two `RRGGBB` colors. */
function channelDistance(a: string, b: string): number {
  let max = 0;
  for (let i = 0; i < 6; i += 2)
    max = Math.max(
      max,
      Math.abs(parseInt(a.slice(i, i + 2), 16) - parseInt(b.slice(i, i + 2), 16))
    );
  return max;
}

describe('colors and fonts against the measured reference', () => {
  const pkg = fixture('colors-fonts');
  const measured = measuredCells('colors-fonts');

  test.each(STYLE_CASES.map((k) => [k.label, k.ref, k.fill !== undefined] as const))(
    '%s',
    (label, ref, hasFill) => {
      const m = measured[ref]!;
      const { format, problems } = formatAt(pkg, 'Colors', ref);
      expect(problems).toEqual([]);
      // A cell with no fill resolves to `none`; only a cell with a fill compares its color.
      if (!hasFill) expect(format.fill).toEqual({ kind: 'none' });
      else {
        expect(format.fill.kind).toBe('solid');
        const fill = format.fill.kind === 'solid' ? format.fill.color : '';
        if (label.includes('tint')) expect(channelDistance(fill, m.fill!)).toBeLessThanOrEqual(2);
        else expect(fill).toBe(m.fill!);
      }
      expect(format.font.color).toBe(m.font!);
      expect(format.font.family).toBe(m.fontName!);
      expect(format.font.bold).toBe(m.bold!);
      if (m.bottom) {
        expect(format.border.bottom.style).toBe(m.bottom.style);
        expect(format.border.bottom.color).toBe(m.bottom.color);
      }
    }
  );
});

// Measured on fixture cell-styles: fixtures/measured.json for colors, fonts,
// borders, and number formats, and, for alignment, the text positions of the reference PDF, read by hand
// (`pdftotext -bbox`) into `shown` below: a cell format that
// omits a record shows record 0 (font, fill, border, number format) or the default alignment, for
// each `apply*` state, and never the record of its cell style format. An own `x:alignment` that
// sets only `wrapText` replaces the whole alignment of the style format.
describe('cell style formats against the measured reference', () => {
  const pkg = fixture('cell-styles');
  const measured = measuredCells('cell-styles');
  /** The horizontal alignment that the reference PDF shows, by the case label. */
  const shown = (label: string): string =>
    label.startsWith('alignment own apply')
      ? 'left'
      : label.startsWith('alignment') || label.includes('no xfId')
        ? 'general'
        : 'center';

  test.each(STYLE_RESOLUTION_CASES.map((k) => [k.label, k.ref] as const))('%s', (label, ref) => {
    const m = measured[ref]!;
    const { format, problems } = formatAt(pkg, 'Styles', ref);
    expect(problems).toEqual([]);
    if (m.fill === 'FFFFFF') expect(format.fill).toEqual({ kind: 'none' });
    else expect(format.fill).toEqual({ kind: 'solid', color: m.fill! });
    expect(format.font.color).toBe(m.font!);
    expect(format.font.bold).toBe(m.bold!);
    expect(format.border.bottom.style).toBe(m.bottom?.style ?? 'none');
    if (m.bottom) expect(format.border.bottom.color).toBe(m.bottom.color);
    expect(format.numFmtId).toBe(m.text === '0,5' ? 0 : 10);
    expect(format.alignment.horizontal).toBe(shown(label));
  });
});

describe('format edits under cell style formats', () => {
  test('the edited copy keeps record 0 under an edit, after save and reopen', () => {
    const fixture = FIXTURES.find((f) => f.name === 'cell-styles')!;
    const pkg = open(fixtureCopies(fixture).edited);
    const ref = (label: string) => STYLE_RESOLUTION_CASES.find((k) => k.label === label)!.ref;
    // Italic on a cell format with no font: font 0 with italic, not the bold red style font.
    const font = formatAt(pkg, 'Styles', ref('font omit apply absent')).format.font;
    expect([font.bold, font.italic, font.color]).toEqual([false, true, '000000']);
    // Wrap on a cell format with no alignment: the default alignment with wrap, not centered.
    const alignment = formatAt(pkg, 'Styles', ref('alignment omit apply absent')).format.alignment;
    expect([alignment.horizontal, alignment.wrapText]).toEqual(['general', true]);
  });
});

describe('path gradient rectangle', () => {
  const pathFill = (attrs: string) =>
    formatAt(
      open(
        buildXlsx({
          styles: {
            fills: [
              `<fill><gradientFill type="path" ${attrs}><stop position="0"><color rgb="FFFF0000"/></stop><stop position="1"><color rgb="FF0000FF"/></stop></gradientFill></fill>`,
            ],
            cellXfs: ['<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0"/>'],
          },
          sheets: [{ name: 'S', rows: row(1, [c('A1', 1, { s: 1 })]) }],
        })
      ),
      'S',
      'A1'
    );
  test('the edges are positions from the left and top edges', () => {
    const { format, problems } = pathFill('left="0.2" right="0.4" top="0.3" bottom="0.6"');
    expect(format.fill.kind === 'gradient' && format.fill.rect).toEqual({
      left: 0.2,
      right: 0.4,
      top: 0.3,
      bottom: 0.6,
    });
    expect(problems).toEqual([]);
  });
  test('an absent edge is 0, the schema default', () => {
    const { format, problems } = pathFill('');
    expect(format.fill.kind === 'gradient' && format.fill.rect).toEqual({
      left: 0,
      right: 0,
      top: 0,
      bottom: 0,
    });
    expect(problems).toEqual([]);
  });
  test('a pair in reverse order is put in order and recorded', () => {
    const { format, problems } = pathFill('left="0.8" right="0.2" top="0.5" bottom="0.5"');
    expect(format.fill.kind === 'gradient' && format.fill.rect).toEqual({
      left: 0.2,
      right: 0.8,
      top: 0.5,
      bottom: 0.5,
    });
    expect(problems).toEqual([{ code: 'gradient-rectangle-out-of-range', detail: 'left>right' }]);
  });
  test('a value outside 0 to 1, or not a number, is limited and recorded', () => {
    const { format, problems } = pathFill('left="-3" right="NaN" top="0" bottom="7"');
    expect(format.fill.kind === 'gradient' && format.fill.rect).toEqual({
      left: 0,
      right: 0,
      top: 0,
      bottom: 1,
    });
    expect(problems).toEqual([{ code: 'gradient-rectangle-out-of-range', detail: 'left=-3' }]);
  });
  test('a number in another base is not an xsd:double', () => {
    const { format, problems } = pathFill('left="0x1" right="0b1" top="0" bottom="1"');
    expect(format.fill.kind === 'gradient' && format.fill.rect).toEqual({
      left: 0,
      right: 0,
      top: 0,
      bottom: 1,
    });
    expect(problems).toEqual([{ code: 'gradient-rectangle-out-of-range', detail: 'left=0x1' }]);
  });
  test('INF and -INF limit to 1 and 0, and an empty value is not a number', () => {
    const { format, problems } = pathFill('left="-INF" right="INF" top="" bottom=" .5E0 "');
    expect(format.fill.kind === 'gradient' && format.fill.rect).toEqual({
      left: 0,
      right: 1,
      top: 0,
      bottom: 0.5,
    });
    expect(problems).toEqual([{ code: 'gradient-rectangle-out-of-range', detail: 'left=-INF' }]);
  });
});

describe('cell format cascade', () => {
  test('a bold red font resolves bold with #FF0000', () => {
    const pkg = open(
      buildXlsx({
        styles: {
          fonts: ['<font><b/><sz val="11"/><color rgb="FFFF0000"/><name val="Calibri"/></font>'],
          cellXfs: ['<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'],
        },
        sheets: [{ name: 'S', rows: row(1, [c('A1', 1, { s: 1 })]) }],
      })
    );
    const { format } = formatAt(pkg, 'S', 'A1');
    expect(format.font.bold).toBe(true);
    expect(format.font.color).toBe('FF0000');
  });

  test('a style index out of range falls back to the default and is reported', () => {
    const pkg = open(
      buildXlsx({
        sheets: [{ name: 'S', rows: '<row r="1"><c r="A1" s="999"><v>1</v></c></row>' }],
      })
    );
    const r = formatAt(pkg, 'S', 'A1');
    expect(r.format.xf).toBe(0);
    expect(r.problems).toEqual([{ code: 'cell-format-out-of-range', detail: '999' }]);
  });

  test('font, fill, and border indexes out of range fall back and are reported', () => {
    const pkg = open(
      buildXlsx({
        styles: { cellXfs: ['<xf numFmtId="0" fontId="40" fillId="40" borderId="40" xfId="0"/>'] },
        sheets: [{ name: 'S', rows: row(1, [c('A1', 1, { s: 1 })]) }],
      })
    );
    expect(formatAt(pkg, 'S', 'A1').problems.map((p) => p.code)).toEqual([
      'font-out-of-range',
      'fill-out-of-range',
      'border-out-of-range',
    ]);
  });

  test('a minor theme font resolves to the theme typeface', () => {
    const pkg = open(
      buildXlsx({ theme: { minor: 'Aptos' }, sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] })
    );
    expect(formatAt(pkg, 'S', 'A1').format.font.family).toBe('Aptos');
  });

  test('alignment and protection are read', () => {
    const pkg = open(
      buildXlsx({
        styles: {
          cellXfs: [
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment horizontal="center" vertical="top" wrapText="1" indent="2" textRotation="90"/><protection locked="0" hidden="1"/></xf>',
          ],
        },
        sheets: [{ name: 'S', rows: row(1, [c('A1', 1, { s: 1 })]) }],
      })
    );
    const { format } = formatAt(pkg, 'S', 'A1');
    expect(format.alignment).toEqual({
      horizontal: 'center',
      vertical: 'top',
      wrapText: true,
      shrinkToFit: false,
      indent: 2,
      textRotation: 90,
    });
    expect(format.protection).toEqual({ locked: false, hidden: true });
  });

  test('an indexed font color without a palette uses the default palette', () => {
    const el = {
      type: 'element',
      id: 'n1',
      kind: 'generic',
      ns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      local: 'color',
      nsDecls: [],
      attrs: [{ ns: '', local: 'indexed', value: '10' }],
      children: [],
    } as never;
    expect(resolveColor(el, { theme: [], indexed: undefined }, 'font')).toBe('FF0000');
  });

  test('values outside their enumerations fall back to the defaults', () => {
    const pkg = open(
      buildXlsx({
        styles: {
          fonts: [
            '<font><sz val="11"/><name val=""/></font>',
            `<font><name val="${'F'.repeat(300)}"/></font>`,
          ],
          fills: [
            '<fill><patternFill patternType="javascript"><fgColor rgb="FFFF0000"/></patternFill></fill>',
          ],
          borders: [
            '<border><left style="url(x)"/><right/><top/><bottom style="thick"/><diagonal/></border>',
          ],
          cellXfs: [
            '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0"><alignment horizontal="expression(x)" vertical="sideways"/></xf>',
            '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0"/>',
          ],
        },
        sheets: [{ name: 'S', rows: row(1, [c('A1', 1, { s: 1 }), c('B1', 1, { s: 2 })]) }],
      })
    );
    const { format, problems } = formatAt(pkg, 'S', 'A1');
    expect(problems).toEqual([]);
    expect(format.alignment.horizontal).toBe('general');
    expect(format.alignment.vertical).toBe('bottom');
    expect(format.border.left.style).toBe('none');
    expect(format.border.bottom.style).toBe('thick');
    expect(format.fill).toEqual({ kind: 'none' });
    // An empty `x:name/@val` is absent.
    expect(format.font.family).toBe('Calibri');
    expect(formatAt(pkg, 'S', 'B1').format.font.family).toHaveLength(255);
  });

  test('a malformed numFmtId is reported and reads as 0', () => {
    const pkg = open(
      buildXlsx({
        styles: { cellXfs: ['<xf numFmtId="0x0E" fontId="0" fillId="0" borderId="0" xfId="0"/>'] },
        sheets: [{ name: 'S', rows: row(1, [c('A1', 1, { s: 1 })]) }],
      })
    );
    const r = formatAt(pkg, 'S', 'A1');
    expect(r.format.numFmtId).toBe(0);
    expect(r.problems).toEqual([{ code: 'number-format-id-malformed', detail: '0x0E' }]);
  });

  test('a solid fill with an auto color is white, the system background', () => {
    const pkg = open(
      buildXlsx({
        styles: {
          fills: [
            '<fill><patternFill patternType="solid"><fgColor auto="1"/></patternFill></fill>',
          ],
          cellXfs: ['<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0"/>'],
        },
        sheets: [{ name: 'S', rows: row(1, [c('A1', 1, { s: 1 })]) }],
      })
    );
    expect(formatAt(pkg, 'S', 'A1').format.fill).toEqual({ kind: 'solid', color: 'FFFFFF' });
  });

  test('rich text runs resolve their fonts and inherit absent properties from the cell font', () => {
    const pkg = open(
      buildXlsx({
        rawStrings: [
          '<si><r><rPr><b/><color rgb="FFFF0000"/><rFont val="Arial"/></rPr><t>a</t></r><r><rPr><i/></rPr><t>b</t></r><r><t>c</t></r></si>',
        ],
        styles: {
          fonts: ['<font><sz val="14"/><color rgb="FF0000FF"/><name val="Courier New"/></font>'],
          cellXfs: ['<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"/>'],
        },
        sheets: [{ name: 'S', rows: row(1, [c('A1', 0, { t: 's', s: 1 })]) }],
      })
    );
    const index = workbookIndex(pkg)!;
    const part = xmlPartOf(pkg, sheetByName(index, 'S')!.partName)!;
    const a1 = parseCellRef('A1')!;
    const cell = sheetIndex(part).cells.get(cellKey(a1.row, a1.col))!.el;
    const value = cellValue(pkg, index, cell, false).value;
    const cellFont = cellFormat(pkg, index, attr(cell, 's')).format.font;
    const runs = richTextRuns(value.kind === 'string' ? value.runs! : [], cellFont, pkg, index);
    expect(runs.map((r) => r.text)).toEqual(['a', 'b', 'c']);
    expect(runs[0]!.font).toMatchObject({
      family: 'Arial',
      bold: true,
      color: 'FF0000',
      sizePt: 14,
    });
    expect(runs[1]!.font).toMatchObject({ family: 'Courier New', italic: true, color: '0000FF' });
    expect(runs[2]!.font).toBe(cellFont);
  });

  test('tint 0 keeps the color, and tints of 1 and -1 reach white and black', () => {
    expect(applyTint('4472C4', 1)).toBe('FFFFFF');
    expect(applyTint('4472C4', -1)).toBe('000000');
  });
});

describe('defaultFont', () => {
  test('is the font of the Normal cell style, through its theme typeface', () => {
    const pkg = open(
      buildXlsx({
        sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }],
        theme: { minor: 'Cambria' },
        defaultFont: '<sz val="20"/><name val="Calibri"/><scheme val="minor"/>',
      })
    );
    const font = defaultFont(pkg, workbookIndex(pkg)!);
    expect([font.family, font.sizePt]).toEqual(['Cambria', 20]);
  });
});
