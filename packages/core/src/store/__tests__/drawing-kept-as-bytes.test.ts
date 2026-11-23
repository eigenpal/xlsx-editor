// A drawing part that the edit cannot read: the reader keeps it as bytes, or its root
// is not a `xdr:wsDr`. A structural edit refuses with `structure-blocked` and a resize with
// `resize-anchor` and `causeUnreadable`, and neither changes the workbook, because the anchors
// would stay on their old cells. A drawing whose relationships part is kept as bytes refuses a
// structural edit, and so does a drawing that relates to a part that can refer to cells.
// `structureBlocker` holds these refusals, so the editor disables the commands.
// A drawing that the reader reads still moves and still allows a resize. Each workbook is built
// as bytes and read by the reader.

import { describe, expect, test } from 'bun:test';
import { strToU8, unzipSync, zipSync } from 'fflate';
import {
  buildXlsx,
  c,
  row,
  type FixtureImage,
  type FixtureSheet,
} from '../../../../../scripts/fixtures/builder.ts';
import { attr, child, textOf } from '../package/query.ts';
import { readXlsx, xmlPartOf, type XlsxPackage } from '../package/package.ts';
import { NS } from '../package/namespaces.ts';
import type { Element } from '../package/tree.ts';
import { structureBlocker } from '../structure/structure.ts';
import { WorkbookDocumentStore } from '../store/document-store.ts';
import type { WorkbookOp } from '../store/ops.ts';
import { op } from './structure-helpers.ts';

const DRAWING = '/xl/drawings/drawing1.xml';
const DRAWING_RELS = '/xl/drawings/_rels/drawing1.xml.rels';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

/** A one-cell anchor at 0-based column 1, row 3, with a shape named `Pic`. */
const ANCHOR =
  '<xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
  '<xdr:ext cx="100" cy="100"/><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="2" name="Pic"/><xdr:cNvSpPr/></xdr:nvSpPr>' +
  '<xdr:spPr/></xdr:sp><xdr:clientData/></xdr:oneCellAnchor>';

/** The same anchor with a prefix in `mc:Ignorable` that it does not declare. */
const UNDECLARED = ANCHOR.replace(
  '<xdr:oneCellAnchor>',
  `<xdr:oneCellAnchor xmlns:mc="${MC}" mc:Ignorable="zz">`
);

/** The same anchor with 300 nested elements, past the limit of 256 nested levels. */
const TOO_DEEP = ANCHOR.replace(
  '<xdr:clientData/>',
  `<xdr:clientData>${'<a:x>'.repeat(300)}${'</a:x>'.repeat(300)}</xdr:clientData>`
);

const IMAGE: FixtureImage = {
  relId: 'rIdImg1',
  name: 'p.png',
  bytes: new Uint8Array(8),
  contentType: 'image/png',
};

const sheet = (anchors: string): FixtureSheet => ({
  name: 'S',
  rows: row(1, [c('A1', 1)]) + row(5, [c('A5', 5)]),
  drawing: { anchors, images: [IMAGE] },
});

/** The workbook of `sheet(anchors)` as bytes, with zip entries replaced, read by the reader. */
function read(anchors: string, replace: Readonly<Record<string, string>> = {}): XlsxPackage {
  const parts = unzipSync(buildXlsx({ sheets: [sheet(anchors)] }));
  for (const [name, xml] of Object.entries(replace)) {
    if (!parts[name]) throw new Error(`no part ${name}`);
    parts[name] = strToU8(xml);
  }
  const result = readXlsx(zipSync(parts));
  if (!result.ok) throw new Error(result.error.reason);
  return result.pkg;
}

/** Edits, each with a name. */
type Edits = readonly (readonly [string, WorkbookOp])[];

const STRUCTURE: Edits = [
  ['insert a row', op('insertRows', 'S', 2)],
  ['delete a row', op('deleteRows', 'S', 2)],
  ['insert a column', op('insertColumns', 'S', 1)],
  ['delete a column', op('deleteColumns', 'S', 1)],
];

const RESIZE: Edits = [
  ['set a column width', { op: 'setColumnWidth', sheet: 'S', first: 2, last: 2, width: 30 }],
  ['set a row height', { op: 'setRowHeight', sheet: 'S', first: 4, last: 4, height: 40 }],
  ['set column widths', { op: 'setColumnWidths', sheet: 'S', columns: [{ column: 2, width: 30 }] }],
  ['fit row heights', { op: 'setFittedRowHeights', sheet: 'S', rows: [{ row: 5, height: 40 }] }],
];

/** The refusal of `edit`, after a check that the workbook did not change. */
function refusal(pkg: XlsxPackage, edit: WorkbookOp) {
  const store = new WorkbookDocumentStore(pkg);
  const before = store.package;
  const r = store.transact([edit]);
  expect(store.package).toBe(before);
  expect(store.canUndo).toBe(false);
  if (r.ok) return null;
  const { code, subject, about, causeUnreadable } = r.rejection;
  return { code, subject, about, causeUnreadable };
}

/** What `structureBlocker` gives, as `refusal` gives it. */
function blocker(pkg: XlsxPackage): Refusal | undefined {
  const b = structureBlocker(pkg, 'S');
  return b && { code: b.code, subject: b.subject, about: b.about, causeUnreadable: undefined };
}

type Refusal = ReturnType<typeof refusal>;

const blockedBy = (subject: string): Refusal => ({
  code: 'structure-blocked',
  subject,
  about: { kind: 'drawing' },
  causeUnreadable: undefined,
});

const RESIZE_UNREAD: Refusal = {
  code: 'resize-anchor',
  subject: undefined,
  about: undefined,
  causeUnreadable: true,
};

/** The refusal of each edit, by its name, so that a failure names the edit. */
const refusals = (pkg: XlsxPackage, edits: Edits): Record<string, Refusal> =>
  Object.fromEntries(edits.map(([name, edit]) => [name, refusal(pkg, edit)]));

/** The same expected refusal for each edit, by its name. */
const named = (edits: Edits, expected: Refusal): Record<string, Refusal> =>
  Object.fromEntries(edits.map(([name]) => [name, expected]));

/** Each structural op and each resize op refuses, and `structureBlocker` holds the refusal. */
function refusesEverything(pkg: XlsxPackage, subject: string) {
  expect(blocker(pkg)).toEqual(blockedBy(subject));
  expect(refusals(pkg, STRUCTURE)).toEqual(named(STRUCTURE, blockedBy(subject)));
  expect(refusals(pkg, RESIZE)).toEqual(named(RESIZE, RESIZE_UNREAD));
}

describe('a drawing part that the edit cannot read', () => {
  test('a prefix in mc:Ignorable that the part does not declare: kept as bytes', () => {
    const pkg = read(UNDECLARED);
    expect(pkg.parts.get(DRAWING)?.kind).toBe('binary');
    refusesEverything(pkg, DRAWING);
  });

  test('a part past the limit of nested levels: kept as bytes', () => {
    const pkg = read(TOO_DEEP);
    expect(pkg.parts.get(DRAWING)?.kind).toBe('binary');
    refusesEverything(pkg, DRAWING);
  });

  test('a part whose root is not a xdr:wsDr', () => {
    const pkg = read(ANCHOR, {
      'xl/drawings/drawing1.xml': '<other xmlns="urn:example:other"/>',
    });
    expect(pkg.parts.get(DRAWING)?.kind).toBe('xml');
    refusesEverything(pkg, DRAWING);
  });
});

describe('a drawing whose relationships part is kept as bytes', () => {
  const pkg = read(ANCHOR, { 'xl/drawings/_rels/drawing1.xml.rels': '<Relationships' });

  test('the reader keeps the relationships part as bytes', () => {
    expect(pkg.parts.get(DRAWING)?.kind).toBe('xml');
    expect(pkg.parts.get(DRAWING_RELS)?.kind).toBe('binary');
  });

  test('every structural edit is refused, and structureBlocker holds the refusal', () => {
    expect(blocker(pkg)).toEqual(blockedBy(DRAWING_RELS));
    expect(refusals(pkg, STRUCTURE)).toEqual(named(STRUCTURE, blockedBy(DRAWING_RELS)));
  });

  test('a resize reads only the anchors, so it is applied', () => {
    for (const [name, edit] of RESIZE) {
      const store = new WorkbookDocumentStore(pkg);
      expect([name, store.transact([edit]).ok]).toEqual([name, true]);
    }
  });
});

describe('a drawing that relates to a part that can refer to cells', () => {
  const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const pkg = read(ANCHOR, {
    'xl/drawings/_rels/drawing1.xml.rels':
      `<Relationships xmlns="${REL_NS}">` +
      `<Relationship Id="rIdImg1" Type="${R}/image" Target="../media/p.png"/>` +
      `<Relationship Id="rId9" Type="${R}/diagramData" Target="../diagrams/data1.xml"/>` +
      '</Relationships>',
  });
  const DIAGRAM = '/xl/diagrams/data1.xml';

  test('every structural edit is refused, and structureBlocker holds the refusal', () => {
    expect(blocker(pkg)).toEqual(blockedBy(DIAGRAM));
    expect(refusals(pkg, STRUCTURE)).toEqual(named(STRUCTURE, blockedBy(DIAGRAM)));
  });

  test('a resize reads only the anchors, so it is applied', () => {
    for (const [name, edit] of RESIZE) {
      const store = new WorkbookDocumentStore(pkg);
      expect([name, store.transact([edit]).ok]).toEqual([name, true]);
    }
  });
});

describe('a drawing part that the reader reads', () => {
  const first = (pkg: XlsxPackage) =>
    xmlPartOf(pkg, DRAWING)!.root.children.find((n): n is Element => n.type === 'element')!;
  const marker = (pkg: XlsxPackage, key: string) =>
    textOf(child(child(first(pkg), { ns: NS.xdr, local: 'from' }), { ns: NS.xdr, local: key }));

  test('moves with each structural edit, and no blocker holds the commands', () => {
    const pkg = read(ANCHOR);
    expect(pkg.parts.get(DRAWING)?.kind).toBe('xml');
    expect(structureBlocker(pkg, 'S')).toBeUndefined();
    const store = new WorkbookDocumentStore(pkg);
    expect(store.transact([op('insertRows', 'S', 2, 2)]).ok).toBe(true);
    expect(marker(store.package, 'row')).toBe('5');
    expect(store.transact([op('deleteRows', 'S', 2)]).ok).toBe(true);
    expect(marker(store.package, 'row')).toBe('4');
    expect(store.transact([op('insertColumns', 'S', 1)]).ok).toBe(true);
    expect(marker(store.package, 'col')).toBe('2');
    expect(store.transact([op('deleteColumns', 'S', 1)]).ok).toBe(true);
    expect(marker(store.package, 'col')).toBe('1');
  });

  test('allows each resize, and the one-cell anchor keeps its size', () => {
    for (const [name, edit] of RESIZE) {
      const store = new WorkbookDocumentStore(read(ANCHOR));
      expect([name, store.transact([edit]).ok]).toEqual([name, true]);
      expect(attr(child(first(store.package), { ns: NS.xdr, local: 'ext' }), 'cx')).toBe('100');
    }
  });
});
