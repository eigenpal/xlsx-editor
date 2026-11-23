// The fixtures of grid layout (task 5 of the active change): cell text layout, and the grid
// structure with drawings. Each case sits in its own cell so that a reference PDF measures it:
// see the quick reference, "Column widths and row heights" and "Text in cells".

import type { FixtureWorkbook } from './builder.ts';
import { c, row } from './builder.ts';

/** `x:alignment` of a cell format, on the default font, fill, and border. */
const xf = (alignment: string, extra = '') =>
  `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"${extra}><alignment ${alignment}/></xf>`;

const LONG = 'This text is longer than its cell';

/** Text layout: alignment, overflow, wrapping, indent, rotation, fitted numbers, and fonts. */
export function textLayoutWorkbook(): FixtureWorkbook {
  const cellXfs = [
    xf('horizontal="left"'), // 1
    xf('horizontal="center"'), // 2
    xf('horizontal="right"'), // 3
    xf('horizontal="fill"'), // 4
    xf('horizontal="justify" wrapText="1"'), // 5
    xf('horizontal="distributed" wrapText="1"'), // 6
    xf('horizontal="centerContinuous"'), // 7
    xf('wrapText="1"'), // 8
    xf('horizontal="left" indent="2"'), // 9
    xf('horizontal="right" indent="1"'), // 10
    xf('vertical="top"'), // 11
    xf('vertical="center"'), // 12
    xf('textRotation="45"'), // 13
    xf('textRotation="90"'), // 14
    xf('textRotation="135"'), // 15
    xf('textRotation="255"'), // 16
    xf('shrinkToFit="1"'), // 17
    '<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>', // 18
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>', // 19
    '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>', // 20
    xf('vertical="top" wrapText="1"'), // 21
  ];
  const rows = [
    row(1, [c('A1', LONG), c('C1', 7)]),
    row(2, [c('D2', LONG, { s: 3 })]),
    row(3, [c('C3', LONG, { s: 2 })]),
    row(4, [
      c('A4', 'left', { s: 1 }),
      c('B4', 'center', { s: 2 }),
      c('C4', 'right', { s: 3 }),
      c('D4', 'ab', { s: 4 }),
      c('E4', 'Indent 2', { s: 9 }),
      c('F4', 'Indent 1', { s: 10 }),
    ]),
    row(
      5,
      [
        c('A5', 'Justified text that wraps onto more lines', { s: 5 }),
        c('B5', 'Distributed text that wraps onto lines', { s: 6 }),
        c('C5', 'Wrapped text that breaks at spaces', { s: 8 }),
        c('D5', 'Top', { s: 11 }),
        c('E5', 'Middle', { s: 12 }),
        c('F5', 'Wrapped text at the top of its cell', { s: 21 }),
      ],
      'ht="60" customHeight="1"'
    ),
    row(6, [
      c('A6', 'Centered across', { s: 7 }),
      c('B6', undefined, { s: 7 }),
      c('C6', undefined, { s: 7 }),
    ]),
    row(
      7,
      [
        c('A7', 'Rotated 45', { s: 13 }),
        c('B7', 'Up 90', { s: 14 }),
        c('C7', 'Down 45', { s: 15 }),
        c('D7', 'ABC', { s: 16 }),
        c('E7', 'Shrink this long text to fit', { s: 17 }),
      ],
      'ht="72" customHeight="1"'
    ),
    row(8, [c('G8', 1234.5, { s: 18 }), c('H8', 1234.5678, { s: 18 })]),
    row(9, [c('G9', 0.000123456), c('H9', 123456789012)]),
    row(10, [c('G10', -12345.678), c('H10', 98765.4321)]),
    row(11, [c('G11', true), c('H11', '#N/A', { t: 'e' })]),
    row(12, [c('A12', 'Carlito 20 in a row without a height', { s: 19 })]),
    row(13, [c('A13', 'Carlito 8 in a row without a height', { s: 20 })]),
    row(14, [c('A14', 0, { t: 's' })]),
  ].join('');
  return {
    rawStrings: [
      '<si><r><rPr><b/><sz val="14"/><color rgb="FFC00000"/><rFont val="Carlito"/></rPr><t xml:space="preserve">Bold red 14 </t></r><r><t>and plain</t></r></si>',
    ],
    sheets: [
      {
        name: 'Text',
        cols:
          '<cols><col min="1" max="6" width="12" customWidth="1"/>' +
          '<col min="7" max="7" width="5" customWidth="1"/><col min="8" max="8" width="8" customWidth="1"/></cols>',
        rows,
      },
    ],
    styles: {
      fonts: [
        '<font><sz val="20"/><color theme="1"/><name val="Carlito"/><family val="2"/><scheme val="minor"/></font>',
        '<font><sz val="8"/><color theme="1"/><name val="Carlito"/><family val="2"/><scheme val="minor"/></font>',
      ],
      cellXfs,
    },
  };
}

/**
 * Text overflow next to formula cells: a long text stops at a formula with a cached number
 * (B1), and at a `t="str"` formula whose cached value is the empty string (B3), which shows
 * nothing. A5 shows overflow across empty cells; its text ends before the value in D5. Each
 * formula's cached value is its true result.
 */
export function formulaOverflowWorkbook(): FixtureWorkbook {
  return {
    sheets: [
      {
        name: 'Overflow',
        cols: '<cols><col min="1" max="4" width="12" customWidth="1"/></cols>',
        rows: [
          row(1, [c('A1', LONG), c('B1', 2, { f: '1+1' })]),
          row(3, [c('A3', LONG), c('B3', '', { f: '""', t: 'str' })]),
          row(5, [c('A5', LONG), c('D5', 7)]),
        ].join(''),
      },
    ],
  };
}

const NS_MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

/**
 * A formula cell whose `x:f` and cached `x:v` are only in an `mc:AlternateContent` child
 * (ECMA-376 Part 3, `AlternateContent`). The `mc:Choice` requires a namespace that no consumer
 * understands, so a consumer takes the `mc:Fallback`. Each branch caches its own true result:
 * the Choice 2, the Fallback 4.
 */
const wrappedCell = (ref: string, choice: [string, number], fallback: [string, number]) =>
  `<c r="${ref}"><mc:AlternateContent xmlns:mc="${NS_MC}" xmlns:x99="urn:example:x99">` +
  `<mc:Choice Requires="x99"><f>${choice[0]}</f><v>${choice[1]}</v></mc:Choice>` +
  `<mc:Fallback><f>${fallback[0]}</f><v>${fallback[1]}</v></mc:Fallback>` +
  '</mc:AlternateContent></c>';

/**
 * Wrapped formula values: the grid shows the cached value of a formula that only a branch of an
 * `mc:AlternateContent` child holds, from the `mc:Fallback` (B1). A direct formula (B2) is the
 * control.
 */
export function wrappedFormulaWorkbook(): FixtureWorkbook {
  return {
    sheets: [
      {
        name: 'Wrapped',
        cols: '<cols><col min="1" max="2" width="16" customWidth="1"/></cols>',
        rows: [
          row(1, [c('A1', 'Wrapped formula'), wrappedCell('B1', ['1+1', 2], ['2+2', 4])]),
          row(2, [c('A2', 'Direct formula'), c('B2', 6, { f: '3+3' })]),
          row(3, [c('A3', 'Edited cell')]),
        ].join(''),
      },
    ],
  };
}

/**
 * A workbook whose only formula is wrapped: B1 keeps its `x:f` only in an
 * `mc:AlternateContent` child, and no cell has a direct `x:f`. An edit sets `fullCalcOnLoad`
 * only if the formula scan finds the wrapped formula.
 */
export function wrappedOnlyFormulaWorkbook(): FixtureWorkbook {
  return {
    sheets: [
      {
        name: 'Wrapped',
        cols: '<cols><col min="1" max="2" width="16" customWidth="1"/></cols>',
        rows: [
          row(1, [c('A1', 'Wrapped formula'), wrappedCell('B1', ['1+1', 2], ['2+2', 4])]),
          row(2, [c('A2', 'Edited cell')]),
        ].join(''),
      },
    ],
  };
}

/** An 8 by 8 pixel PNG in four squares, built once by a script and kept as bytes. */
const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAHElEQVR42mO4YGAARwYJF+CIgYoSyBxkRVSUAAAKqFIBWUslBwAAAABJRU5ErkJggg=='
  ),
  (ch) => ch.charCodeAt(0)
);

const at = (tag: string, col: number, r: number, colOff = 0, rowOff = 0) =>
  `<xdr:${tag}><xdr:col>${col}</xdr:col><xdr:colOff>${colOff}</xdr:colOff><xdr:row>${r}</xdr:row><xdr:rowOff>${rowOff}</xdr:rowOff></xdr:${tag}>`;

const shape = (id: number, prst: string, fill: string) =>
  `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="Shape ${id}"/><xdr:cNvSpPr/></xdr:nvSpPr>` +
  `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>` +
  `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></xdr:spPr></xdr:sp>`;

const picture =
  '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Picture 2"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>' +
  '<xdr:blipFill><a:blip r:embed="rIdImage1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>' +
  '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>';

/** Grid structure: row heights, hidden rows and columns, merged ranges, borders, and drawings. */
export function gridStructureWorkbook(): FixtureWorkbook {
  const label = (r: number, text: string, attrs = '') => row(r, [c(`A${r}`, text)], attrs);
  const rows = [
    label(1, 'default row'),
    label(2, 'ht 10', 'ht="10" customHeight="1"'),
    label(3, 'ht 20', 'ht="20" customHeight="1"'),
    label(4, 'ht 33.75', 'ht="33.75" customHeight="1"'),
    label(5, 'hidden row', 'hidden="1"'),
    row(6, [c('A6', 'after hidden'), c('B6', 'B'), c('C6', 'hidden column'), c('D6', 'D')]),
    row(8, [c('B8', 'Merged B8:D10', { s: 1 })]),
    row(12, [c('F12', 'Merged F12:F14', { s: 2 })]),
  ].join('');
  const anchors =
    `<xdr:twoCellAnchor editAs="oneCell">${at('from', 1, 15)}${at('to', 3, 20)}${picture}<xdr:clientData/></xdr:twoCellAnchor>` +
    `<xdr:twoCellAnchor>${at('from', 4, 15, 127000, 63500)}${at('to', 5, 19, 254000, 127000)}${shape(3, 'ellipse', '4472C4')}<xdr:clientData/></xdr:twoCellAnchor>` +
    `<xdr:oneCellAnchor>${at('from', 1, 22)}<xdr:ext cx="1270000" cy="635000"/>${shape(4, 'roundRect', 'ED7D31')}<xdr:clientData/></xdr:oneCellAnchor>` +
    `<xdr:absoluteAnchor><xdr:pos x="3810000" y="5080000"/><xdr:ext cx="1270000" cy="1270000"/>${shape(5, 'rightArrow', '70AD47')}<xdr:clientData/></xdr:absoluteAnchor>`;
  return {
    sheets: [
      {
        name: 'Grid',
        cols:
          '<cols><col min="2" max="2" width="10" customWidth="1"/>' +
          '<col min="3" max="3" width="10" hidden="1" customWidth="1"/></cols>',
        rows,
        merges: ['B8:D10', 'F12:F14'],
        drawing: {
          anchors,
          images: [
            { relId: 'rIdImage1', name: 'image1.png', bytes: PNG, contentType: 'image/png' },
          ],
        },
      },
    ],
    styles: {
      fills: [
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>',
      ],
      borders: [
        '<border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border>',
      ],
      cellXfs: [
        '<xf numFmtId="0" fontId="0" fillId="2" borderId="1" xfId="0" applyFill="1" applyBorder="1"><alignment horizontal="center" vertical="center"/></xf>',
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"><alignment vertical="top" wrapText="1"/></xf>',
      ],
    },
  };
}
