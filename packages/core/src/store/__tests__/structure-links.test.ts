// Structural edits and the links of other sheets (design D4): a shape (`xdr:sp/@textlink`), a
// control (`@fmlaLink`), or a legacy control (`x:FmlaLink`) on another sheet that names the
// edited sheet refuses the edit, and the refusal names its subject in words.

import { describe, expect, test } from 'bun:test';
import { c, row, type FixtureSheet } from '../../../../../scripts/fixtures/builder.ts';
import type { WorkbookDocumentStore } from '../store/document-store.ts';
import { op, storeOf } from './structure-helpers.ts';

const enc = (s: string) => new TextEncoder().encode(s);
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const CTRL = 'application/vnd.ms-excel.controlproperties+xml';
const VML = 'application/vnd.openxmlformats-officedocument.vmlDrawing';

const data: FixtureSheet = { name: 'Data', rows: row(1, [c('A1', 1)]) + row(5, [c('A5', 5)]) };

const point = (tag: string, col: number, r: number) =>
  `<xdr:${tag}><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff>` +
  `<xdr:row>${r}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:${tag}>`;

const shapeSheet = (textlink: string): FixtureSheet => ({
  name: 'Other',
  rows: row(1, [c('A1', 1)]),
  drawing: {
    anchors:
      `<xdr:oneCellAnchor>${point('from', 1, 1)}<xdr:ext cx="1" cy="1"/>` +
      `<xdr:sp macro="" textlink="${textlink}"><xdr:nvSpPr><xdr:cNvPr id="2" name="S2"/>` +
      '<xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
      '</xdr:spPr></xdr:sp><xdr:clientData/></xdr:oneCellAnchor>',
  },
});

const controlSheet = (rel: 'ctrlProp' | 'vmlDrawing'): FixtureSheet => ({
  name: 'Other',
  rows: row(1, [c('A1', 1)]),
  afterPageSetup:
    rel === 'vmlDrawing'
      ? '<legacyDrawing r:id="rIdL"/>'
      : '<controls><control shapeId="1025" r:id="rIdL" name="Check Box 1"/></controls>',
  rels: [
    [
      'rIdL',
      `${R}${rel}`,
      rel === 'vmlDrawing' ? '../drawings/vmlDrawing1.vml' : '../ctrlProps/ctrlProp1.xml',
    ],
  ],
});

const ctrlProp = (link: string) =>
  enc(
    '<?xml version="1.0"?><formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" ' +
      `objectType="CheckBox" fmlaLink="${link}"/>`
  );
const vml = (link: string) =>
  enc(
    '<xml xmlns:x="urn:schemas-microsoft-com:office:excel"><x:ClientData ObjectType="Checkbox">' +
      `<x:FmlaLink>${link}</x:FmlaLink></x:ClientData></xml>`
  );

function refused(store: WorkbookDocumentStore) {
  const before = store.package;
  const r = store.transact([op('insertRows', 'Data', 1)]);
  expect(store.package).toBe(before);
  return r.ok ? null : { code: r.rejection.code, about: r.rejection.about };
}

describe('links of other sheets to the edited sheet', () => {
  test('a shape on another sheet whose text comes from the edited sheet refuses', () => {
    const s = storeOf({ sheets: [data, shapeSheet('Data!$A$5')] });
    expect(refused(s)).toEqual({
      code: 'structure-blocked',
      about: { kind: 'linked-drawing', name: 'Other' },
    });
  });

  test('a shape linked to a cell of its own sheet does not refuse', () => {
    const s = storeOf({ sheets: [data, shapeSheet('$A$1')] });
    expect(s.transact([op('insertRows', 'Data', 1)]).ok).toBe(true);
  });

  test('a control property linked to the edited sheet refuses', () => {
    const s = storeOf({
      sheets: [data, controlSheet('ctrlProp')],
      extraParts: [['/xl/ctrlProps/ctrlProp1.xml', CTRL, ctrlProp('Data!$A$5')]],
    });
    expect(refused(s)).toEqual({
      code: 'structure-blocked',
      about: { kind: 'control', name: 'Other' },
    });
  });

  test('a legacy control linked to the edited sheet refuses', () => {
    const s = storeOf({
      sheets: [data, controlSheet('vmlDrawing')],
      extraParts: [['/xl/drawings/vmlDrawing1.vml', VML, vml('Data!$A$5')]],
    });
    expect(refused(s)).toEqual({
      code: 'structure-blocked',
      about: { kind: 'control', name: 'Other' },
    });
  });

  test('controls linked to another sheet do not refuse', () => {
    const s = storeOf({
      sheets: [data, controlSheet('vmlDrawing')],
      extraParts: [['/xl/drawings/vmlDrawing1.vml', VML, vml('Other!$A$1')]],
    });
    expect(s.transact([op('insertRows', 'Data', 1)]).ok).toBe(true);
  });
});
