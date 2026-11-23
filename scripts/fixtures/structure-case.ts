// The structure fixture (task 7.5, design D4): a sheet whose edit deletes row 4 and inserts
// column C, and a hidden second sheet that refers to it. Every cached value is the true result of its
// formula before the edit.
//
// No shared formula here holds a sheet-qualified reference. Measured by bisection with the
// save check: a workbook with a shared formula whose
// text holds a sheet-qualified reference (`Data!B2*2`, also `Data!$B$2*2`, also with another
// `si`) fails to open, and the same shared formula with a same-sheet reference (`A1*2`) opens.
// So the cross-sheet references of Refs are plain formulas, and the shared formula that the edit
// unshares is a same-sheet one on Data (column H). The store never writes a new shared formula
// (`store/structure/sheet-cells.ts`).

import { c, row, type FixtureWorkbook } from './builder.ts';

const at = (tag: string, col: number, r: number) =>
  `<xdr:${tag}><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${r}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:${tag}>`;

const shape = (id: number, prst: string, fill: string) =>
  `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="Shape ${id}"/><xdr:cNvSpPr/></xdr:nvSpPr>` +
  `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>` +
  `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></xdr:spPr></xdr:sp>`;

const sharedChild = (ref: string, v: number, si = 0) =>
  `<c r="${ref}"><f t="shared" si="${si}"/><v>${v}</v></c>`;

/** The workbook of the structure fixture. */
export function structureWorkbook(): FixtureWorkbook {
  const data = [
    row(1, [
      c('A1', 'Item'),
      c('B1', 'Q1'),
      c('C1', 'Q2'),
      c('D1', 'Total'),
      c('F1', 'Merged'),
      c('H1', 'Next Q1 x2'),
    ]),
    row(2, [
      c('A2', 'North'),
      c('B2', 10),
      c('C2', 20),
      '<c r="D2"><f t="shared" ref="D2:D5" si="0">B2+C2</f><v>30</v></c>',
      '<c r="H2"><f t="shared" ref="H2:H5" si="1">B3*2</f><v>60</v></c>',
    ]),
    row(3, [
      c('A3', 'South'),
      c('B3', 30),
      c('C3', 40),
      sharedChild('D3', 70),
      c('F3', 'F3:F5'),
      sharedChild('H3', 100, 1),
    ]),
    row(
      4,
      [
        c('A4', 'East'),
        c('B4', 50),
        c('C4', 60),
        sharedChild('D4', 110),
        sharedChild('H4', 140, 1),
      ],
      'ht="24" customHeight="1"'
    ),
    row(5, [
      c('A5', 'West'),
      c('B5', 70),
      c('C5', 80),
      sharedChild('D5', 150),
      sharedChild('H5', 320, 1),
    ]),
    row(6, [
      c('A6', 'Sum'),
      c('B6', 160, { f: 'SUM(B2:B5)' }),
      c('C6', 200, { f: 'SUM($C$2:$C$5)' }),
      c('D6', 50, { f: 'B4' }),
    ]),
  ].join('');
  const anchors =
    `<xdr:twoCellAnchor>${at('from', 1, 9)}${at('to', 3, 12)}${shape(2, 'rect', '4472C4')}<xdr:clientData/></xdr:twoCellAnchor>` +
    `<xdr:oneCellAnchor>${at('from', 4, 9)}<xdr:ext cx="635000" cy="635000"/>${shape(3, 'ellipse', 'ED7D31')}<xdr:clientData/></xdr:oneCellAnchor>` +
    `<xdr:absoluteAnchor><xdr:pos x="381000" y="3175000"/><xdr:ext cx="635000" cy="317500"/>${shape(4, 'rightArrow', '70AD47')}<xdr:clientData/></xdr:absoluteAnchor>`;
  const refs = [
    row(1, [c('A1', 'From Data'), c('B1', 30, { f: 'Data!D2' }), c('C1', 20, { f: 'Data!B2*2' })]),
    row(2, [c('B2', 360, { f: 'SUM(Data!B2:C5)' }), c('C2', 60, { f: 'Data!B3*2' })]),
    row(3, [c('B3', 50, { f: 'Data!B4' }), c('C3', 100, { f: 'Data!B4*2' })]),
    row(4, [c('B4', 'West', { t: 'str', f: 'Data!A5' }), c('C4', 140, { f: 'Data!B5*2' })]),
  ].join('');
  return {
    definedNames:
      '<definedName name="Gone">Data!$B$4</definedName>' +
      '<definedName name="Totals">Data!$D$2:$D$5</definedName>',
    sheets: [
      {
        name: 'Data',
        cols:
          '<cols><col min="1" max="1" width="12" customWidth="1"/>' +
          '<col min="2" max="4" width="9" customWidth="1"/></cols>',
        rows: data,
        merges: ['F3:F5'],
        afterMerges:
          '<hyperlinks><hyperlink ref="A5" location="Refs!A1" display="West"/></hyperlinks>',
        drawing: { anchors },
      },
      // Hidden, so that the reference PDF is one page of Data, the sheet that the harness paints.
      { name: 'Refs', state: 'hidden', rows: refs },
    ],
  };
}
