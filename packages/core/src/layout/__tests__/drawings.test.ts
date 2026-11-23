import { beforeAll, describe, expect, test } from 'bun:test';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx, type XlsxPackage } from '../../store/package/package.ts';
import { parseXmlPart } from '../../store/package/tree.ts';
import { layoutSheet } from '../sheet-layout.ts';
import type { DrawingLayout } from '../drawings.ts';
import { PRINT_PROFILE } from '../metrics.ts';
import type { TextMeasurer } from '../text/measurer.ts';
import { substituteMeasurer } from './test-fonts.ts';

let measurer: TextMeasurer;
beforeAll(async () => {
  measurer = await substituteMeasurer();
});

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
const pic = (id: number, embed: string) =>
  `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="Picture ${id}"/><xdr:cNvPicPr/></xdr:nvPicPr>` +
  `<xdr:blipFill><a:blip r:embed="${embed}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
  '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>';
const cell = (tag: string, col: number, row: number, colOff = 0, rowOff = 0) =>
  `<xdr:${tag}><xdr:col>${col}</xdr:col><xdr:colOff>${colOff}</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>${rowOff}</xdr:rowOff></xdr:${tag}>`;

function drawings(
  anchors: string,
  viewport = { width: 600, height: 400, scrollX: 0, scrollY: 0 },
  edit: (pkg: XlsxPackage) => XlsxPackage = (pkg) => pkg
) {
  const read = readXlsx(
    buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 1)]),
          drawing: {
            anchors,
            images: [
              { relId: 'rIdImg1', name: 'image1.png', bytes: PNG, contentType: 'image/png' },
            ],
          },
        },
      ],
    })
  );
  if (!read.ok) throw new Error('read');
  const l = layoutSheet(edit(read.pkg), 'S', viewport, { measurer, profile: PRINT_PROFILE })!;
  return { layout: l, all: l.panes.flatMap((p) => p.drawings) as DrawingLayout[] };
}

describe('drawing anchors', () => {
  test('the spec scenario: a picture from B2 to D6 with no offsets', () => {
    const { layout, all } = drawings(
      `<xdr:twoCellAnchor>${cell('from', 1, 1)}${cell('to', 3, 5)}${pic(2, 'rIdImg1')}<xdr:clientData/></xdr:twoCellAnchor>`
    );
    const g = layout.geometry;
    expect(all).toHaveLength(1);
    const p = all[0]!;
    expect(p.kind).toBe('picture');
    expect(p.rect).toEqual({
      x: g.columns.offsetOf(1),
      y: g.rows.offsetOf(1),
      width: g.columns.offsetOf(3) - g.columns.offsetOf(1),
      height: g.rows.offsetOf(5) - g.rows.offsetOf(1),
    });
    if (p.kind === 'picture') {
      expect(p.partName).toBe('/xl/media/image1.png');
      expect(p.contentType).toBe('image/png');
    }
  });

  test('a one-cell anchor takes its extent, and offsets are EMU', () => {
    const { all } = drawings(
      `<xdr:oneCellAnchor>${cell('from', 0, 0, 127000, 254000)}<xdr:ext cx="635000" cy="317500"/>${pic(2, 'rIdImg1')}<xdr:clientData/></xdr:oneCellAnchor>`
    );
    expect(all[0]!.rect).toEqual({ x: 10, y: 14, width: 50, height: 25 });
  });

  test('an absolute anchor sits at its position', () => {
    const { all } = drawings(
      `<xdr:absoluteAnchor><xdr:pos x="1270000" y="635000"/><xdr:ext cx="1270000" cy="1270000"/>${pic(2, 'rIdImg1')}<xdr:clientData/></xdr:absoluteAnchor>`
    );
    expect(all[0]!.rect).toEqual({ x: 100, y: 50, width: 100, height: 100 });
  });

  test('a drawing outside the viewport is not in any pane', () => {
    const { all } = drawings(
      `<xdr:absoluteAnchor><xdr:pos x="127000000" y="0"/><xdr:ext cx="1270000" cy="1270000"/>${pic(2, 'rIdImg1')}<xdr:clientData/></xdr:absoluteAnchor>`
    );
    expect(all).toHaveLength(0);
  });

  test('an anchor with a cell outside the grid is dropped', () => {
    const { all } = drawings(
      `<xdr:twoCellAnchor>${cell('from', 99999, 0)}${cell('to', 3, 5)}${pic(2, 'rIdImg1')}<xdr:clientData/></xdr:twoCellAnchor>`
    );
    expect(all).toHaveLength(0);
  });
});

describe('drawing objects', () => {
  test('a preset shape with a solid fill and a line, in points', () => {
    const shape =
      '<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="3" name="Oval"/><xdr:cNvSpPr/></xdr:nvSpPr>' +
      '<xdr:spPr><a:xfrm rot="5400000"><a:off x="0" y="0"/><a:ext cx="1270000" cy="635000"/></a:xfrm>' +
      '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent1"/></a:solidFill>' +
      '<a:ln w="25400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></xdr:spPr></xdr:sp>';
    const { all } = drawings(
      `<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="1270000" cy="635000"/>${shape}<xdr:clientData/></xdr:absoluteAnchor>`
    );
    const s = all[0]!;
    expect(s.kind).toBe('shape');
    expect(s.rotation).toBe(90);
    if (s.kind !== 'shape') return;
    expect(s.fill).toEqual({ r: 0x44, g: 0x72, b: 0xc4, a: 1 });
    expect(s.line).toEqual({ color: { r: 255, g: 0, b: 0, a: 1 }, width: 2 });
    const xs = s.geometry.paths[0]!.commands.flatMap((cmd) => ('x' in cmd ? [cmd.x] : []));
    expect(Math.max(...xs)).toBeCloseTo(100, 6);
  });

  test('a picture that links an external image is never loaded', () => {
    const external =
      '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="4" name="Linked"/><xdr:cNvPicPr/></xdr:nvPicPr>' +
      '<xdr:blipFill><a:blip r:link="rIdExternal"/></xdr:blipFill><xdr:spPr/></xdr:pic>';
    const { all } = drawings(
      `<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="1270000" cy="1270000"/>${external}<xdr:clientData/></xdr:absoluteAnchor>`
    );
    expect(all[0]).toMatchObject({ kind: 'unsupported', element: 'externalImage' });
  });

  test('a hidden object is not laid out, and a chart frame is unsupported', () => {
    const hidden = pic(5, 'rIdImg1').replace('name="Picture 5"', 'name="Picture 5" hidden="1"');
    const frame =
      '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="6" name="Chart"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>' +
      '<xdr:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/></a:graphic></xdr:graphicFrame>';
    const anchor = (body: string) =>
      `<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="1270000" cy="1270000"/>${body}<xdr:clientData/></xdr:absoluteAnchor>`;
    const { all } = drawings(anchor(hidden) + anchor(frame));
    expect(all.map((d) => d.kind)).toEqual(['unsupported']);
  });
});

const EMU = 12_700;
const absolute = (x: number, y: number, w: number, h: number, body: string) =>
  `<xdr:absoluteAnchor><xdr:pos x="${x * EMU}" y="${y * EMU}"/><xdr:ext cx="${w * EMU}" cy="${h * EMU}"/>${body}<xdr:clientData/></xdr:absoluteAnchor>`;
const connector = (xfrm = '') =>
  '<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="7" name="Line"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr>' +
  `<xdr:spPr>${xfrm}<a:prstGeom prst="line"><a:avLst/></a:prstGeom>` +
  '<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></xdr:spPr></xdr:cxnSp>';
const rect = (xfrm = '') =>
  '<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="8" name="Bar"/><xdr:cNvSpPr/></xdr:nvSpPr>' +
  `<xdr:spPr>${xfrm}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp>`;
const rotated = (rot: number) =>
  `<a:xfrm rot="${rot}"><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>`;

describe('drawing visibility and rotation', () => {
  test('a straight connector with no height or no width still shows', () => {
    const { all, layout } = drawings(
      absolute(10, 20, 72, 0, connector()) + absolute(100, 20, 0, 72, connector())
    );
    expect(all.map((d) => [d.kind, d.rect.width, d.rect.height])).toEqual([
      ['shape', 72, 0],
      ['shape', 0, 72],
    ]);
    expect(layout.problems).toEqual([]);
  });

  test('a two-cell connector along a row boundary shows', () => {
    const anchors = `<xdr:twoCellAnchor>${cell('from', 0, 2)}${cell('to', 3, 2)}${connector()}<xdr:clientData/></xdr:twoCellAnchor>`;
    const { all } = drawings(anchors);
    expect(all).toHaveLength(1);
  });

  test('rotation is normalized to [0, 360)', () => {
    const { all } = drawings(
      absolute(0, 0, 50, 50, rect(rotated(-5_400_000))) +
        absolute(0, 0, 50, 50, rect(rotated(27_000_000))) +
        absolute(0, 0, 50, 50, rect(rotated(21_600_000)))
    );
    expect(all.map((d) => d.rotation)).toEqual([270, 90, 0]);
  });

  test('the rotated bounding box decides visibility', () => {
    // 20 by 400 points, 10 points right of a 600 point viewport. A quarter turn makes it
    // 400 by 20 points around the same center, so it reaches back into the viewport.
    const flat = drawings(absolute(610, 0, 20, 400, rect()));
    expect(flat.all).toHaveLength(0);
    const turned = drawings(absolute(610, 0, 20, 400, rect(rotated(5_400_000))));
    expect(turned.all).toHaveLength(1);
    expect(turned.all[0]!.rect.x).toBe(610);
  });
});

describe('pictures without their image', () => {
  const picture = (blip: string) =>
    absolute(
      0,
      0,
      50,
      50,
      '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="9" name="P"/><xdr:cNvPicPr/></xdr:nvPicPr>' +
        `<xdr:blipFill><a:blip ${blip}/></xdr:blipFill><xdr:spPr/></xdr:pic>`
    );

  test('an embed relationship with an external target is an external image', () => {
    const { all } = drawings(picture('r:embed="rIdImg1"'), undefined, (pkg) => {
      const parts = new Map(pkg.parts);
      const name = [...parts.keys()].find((k) => k.endsWith('drawing1.xml.rels'))!;
      const xml = parseXmlPart(
        name,
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rIdImg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.com/a.png" TargetMode="External"/>' +
          '</Relationships>'
      );
      if (!xml.ok) throw new Error('rels');
      parts.set(name, { kind: 'xml', partName: name, xml: xml.part });
      return { parts };
    });
    expect(all[0]).toMatchObject({ kind: 'unsupported', element: 'externalImage' });
  });

  test('an embed that names no relationship, or a missing part, is a missing image', () => {
    const missingRel = drawings(picture('r:embed="rIdNope"'));
    expect(missingRel.all[0]).toMatchObject({ kind: 'unsupported', element: 'missingImage' });
    expect(missingRel.layout.problems).toContainEqual({ code: 'image-missing', detail: 'P' });
    const noEmbed = drawings(picture(''));
    expect(noEmbed.all[0]).toMatchObject({ kind: 'unsupported', element: 'missingImage' });
    const missingPart = drawings(picture('r:embed="rIdImg1"'), undefined, (pkg) => ({
      parts: new Map([...pkg.parts].filter(([k]) => !k.includes('/media/'))),
    }));
    expect(missingPart.all[0]).toMatchObject({ kind: 'unsupported', element: 'missingImage' });
  });
});
