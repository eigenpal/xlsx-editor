// A delete of a sheet with a drawing (design D1 of the change `workbook-structure-editing`):
// the Drawings part (ECMA-376 Part 1, 12.3.8), its relationship part, and the Image parts
// (15.2.14) that only the sheet uses go with the sheet. Content of the deleted sheet that names
// the sheet goes with it, and content of another sheet that names it still refuses the delete.

import { describe, expect, test } from 'bun:test';
import { c, row, type FixtureSheet } from '../../../../../scripts/fixtures/builder.ts';
import { attr, children } from '../package/query.ts';
import { CONTENT_TYPES_PART, readXlsx, xmlPartOf, type XlsxPackage } from '../package/package.ts';
import { saveXlsx } from '../package-write/save.ts';
import { serializeXmlPart } from '../package-write/serialize.ts';
import { parseXmlPart } from '../package/tree.ts';
import { WorkbookDocumentStore } from '../store/document-store.ts';
import { storeOf } from './structure-helpers.ts';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

const anchor = (col: number, body: string) =>
  `<xdr:oneCellAnchor><xdr:from><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="914400" cy="914400"/>${body}<xdr:clientData/></xdr:oneCellAnchor>`;

const picture = (id: number) =>
  anchor(
    1,
    `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="Picture"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdImage1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>`
  );

const shape = (id: number, textlink = '') =>
  anchor(
    4,
    `<xdr:sp macro="" textlink="${textlink}"><xdr:nvSpPr><xdr:cNvPr id="${id}" name="Shape"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp>`
  );

const image = {
  relId: 'rIdImage1',
  name: 'image1.png',
  bytes: PNG,
  contentType: 'image/png',
} as const;

const drawn = (name: string, anchors: string, shareImage = true): FixtureSheet => ({
  name,
  rows: row(1, [c('A1', 1)]),
  drawing: { anchors, images: shareImage ? [image] : [] },
});

const plain = (name: string): FixtureSheet => ({ name, rows: row(1, [c('A1', 2)]) });

const overrides = (pkg: XlsxPackage) =>
  children(xmlPartOf(pkg, CONTENT_TYPES_PART)!.root, 'ct:Override').map((o) => attr(o, 'PartName'));

describe('a delete of a sheet with a drawing', () => {
  test('removes the drawing, its relationships, and its image, and saves a package that reopens', () => {
    const s = storeOf({ sheets: [drawn('Layout', picture(2) + shape(3)), plain('Other')] });
    const before = s.package;
    expect(s.transact([{ op: 'deleteSheet', sheet: 'Layout' }]).ok).toBe(true);
    const names = [...s.package.parts.keys()];
    expect(names.filter((n) => /drawings|media|sheet1/.test(n))).toEqual([]);
    expect(overrides(s.package).some((n) => n?.includes('drawing'))).toBe(false);
    const reopened = readXlsx(saveXlsx(s.package));
    expect(reopened.ok).toBe(true);
    if (reopened.ok)
      expect([...reopened.pkg.parts.keys()].some((n) => n.includes('media'))).toBe(false);
    // Undo brings back every part, with the same content.
    expect(s.undo().ok).toBe(true);
    expect([...s.package.parts.keys()].sort()).toEqual([...before.parts.keys()].sort());
    for (const [name, part] of before.parts) expect(s.package.parts.get(name)).toBe(part);
  });

  test('a shape on the deleted sheet that links to that sheet goes with it', () => {
    const s = storeOf({
      sheets: [drawn('Layout', shape(2, 'Layout!$A$1'), false), plain('Other')],
    });
    expect(s.transact([{ op: 'deleteSheet', sheet: 'Layout' }]).ok).toBe(true);
  });

  test('a shape on another sheet that links to the deleted sheet still refuses the delete', () => {
    const s = storeOf({ sheets: [plain('Data'), drawn('Board', shape(2, 'Data!$A$1'), false)] });
    const before = s.package;
    const r = s.transact([{ op: 'deleteSheet', sheet: 'Data' }]);
    expect(r.ok ? null : r.rejection).toMatchObject({
      code: 'structure-blocked',
      subject: '/xl/drawings/drawing2.xml',
    });
    expect(s.package).toBe(before);
  });

  test('an image that the drawing of another sheet uses stays', () => {
    const s = storeOf({ sheets: [drawn('Layout', picture(2)), drawn('Other', picture(2))] });
    expect(s.transact([{ op: 'deleteSheet', sheet: 'Layout' }]).ok).toBe(true);
    const names = [...s.package.parts.keys()];
    expect(names).not.toContain('/xl/drawings/drawing1.xml');
    expect(names).toContain('/xl/drawings/drawing2.xml');
    expect(names).toContain('/xl/media/image1.png');
    expect(readXlsx(saveXlsx(s.package)).ok).toBe(true);
  });

  test('a drawing that relates to a part other than an image refuses the delete', () => {
    const diagram = 'application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml';
    const book = storeOf({
      sheets: [drawn('Layout', picture(2)), plain('Other')],
      extraParts: [
        [
          'xl/diagrams/data1.xml',
          diagram,
          new TextEncoder().encode(
            '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>'
          ),
        ],
      ],
    });
    // The builder relates extra parts to the workbook; relate the diagram to the drawing instead.
    const pkg = withDrawingRel(book.package, `${REL}/diagramData`, '../diagrams/data1.xml');
    const s = new WorkbookDocumentStore(pkg);
    const r = s.transact([{ op: 'deleteSheet', sheet: 'Layout' }]);
    expect(r.ok ? null : r.rejection.subject).toBe('/xl/drawings/drawing1.xml');
  });
});

/** The package with one more relationship in the relationship part of `drawing1.xml`. */
function withDrawingRel(pkg: XlsxPackage, type: string, target: string): XlsxPackage {
  const name = '/xl/drawings/_rels/drawing1.xml.rels';
  const text = serializeXmlPart(xmlPartOf(pkg, name)!).replace(
    '</Relationships>',
    `<Relationship Id="rIdX" Type="${type}" Target="${target}"/></Relationships>`
  );
  const parsed = parseXmlPart(name, text);
  if (!parsed.ok) throw new Error(parsed.reason);
  const parts = new Map(pkg.parts);
  parts.set(name, { kind: 'xml', partName: name, xml: parsed.part });
  return { parts };
}
