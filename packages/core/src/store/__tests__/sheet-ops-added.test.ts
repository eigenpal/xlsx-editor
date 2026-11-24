// A rename and a delete of a sheet that an add op of the session made, in a workbook whose
// content refuses them for the sheets of the file (`sheetEditBlocker`, `blockers.ts`). The store
// keeps the part names of the package that it opened; a worksheet whose part is not among them
// is an added sheet. Content of the file can name such a sheet only by a dangling name, which
// the add made live: a chart formula (`c:f`, ECMA-376 Part 1, 21.2.2.65) or a pivot cache source
// (`worksheetSource`, 18.10.1.95). That content refuses the op as before.

import { describe, expect, test } from 'bun:test';
import {
  buildXlsx,
  c,
  row,
  type FixtureWorkbook,
} from '../../../../../scripts/fixtures/builder.ts';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { readXlsx, type XlsxPackage } from '../package/package.ts';
import { parseXmlPart } from '../package/tree.ts';
import { saveXlsx } from '../package-write/save.ts';
import { serializeXmlPart } from '../package-write/serialize.ts';
import { workbookIndex } from '../workbook/workbook-index.ts';
import { WorkbookDocumentStore } from '../store/document-store.ts';
import { applyOps, type WorkbookOp } from '../store/ops.ts';
import { open, storeOf } from './structure-helpers.ts';
import {
  addedSheetBlocker,
  partNamesOf,
  sameOpaqueContent,
  sheetEditBlocker,
} from '../structure/blockers.ts';
import { asciiFold } from '@eigenpal/ooxml-core';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CHART = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const enc = (text: string) => new TextEncoder().encode(text);

/** A chart part whose series reads `ref`, or no cells. */
const chart = (ref?: string): readonly [string, string, Uint8Array] => [
  'xl/charts/chart1.xml',
  CHART,
  enc(
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
      (ref
        ? `<c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numRef><c:f>${ref}</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart>`
        : '') +
      '</c:chartSpace>'
  ),
];

const book = (extraParts: FixtureWorkbook['extraParts']): FixtureWorkbook => ({
  sheets: [
    { name: 'Sheet1', rows: row(1, [c('A1', 1)]) },
    { name: 'Sheet2', rows: row(1, [c('A1', 2, { f: 'Sheet1!A1' })]) },
  ],
  extraParts,
});

const names = (s: WorkbookDocumentStore) => workbookIndex(s.package)!.sheets.map((x) => x.name);

/** The code and the subject of a refused transaction, or `ok`. */
function outcome(s: WorkbookDocumentStore, ops: readonly WorkbookOp[]): unknown {
  const before = s.package;
  const r = s.transact(ops);
  if (r.ok) return 'ok';
  expect(s.package).toBe(before);
  return { code: r.rejection.code, subject: r.rejection.subject };
}

const add = (name: string): WorkbookOp => ({ op: 'addSheet', name, index: 2 });
const rename = (sheet: string, name: string): WorkbookOp => ({ op: 'renameSheet', sheet, name });
const del = (sheet: string): WorkbookOp => ({ op: 'deleteSheet', sheet });

/** The package with an XML part set to `text`. */
function withXml(pkg: XlsxPackage, name: string, text: string): XlsxPackage {
  const parsed = parseXmlPart(name, text);
  if (!parsed.ok) throw new Error(parsed.reason);
  const parts = new Map(pkg.parts);
  parts.set(name, { kind: 'xml', partName: name, xml: parsed.part });
  return { parts };
}

/** The package with the text of an XML part changed by `edit`. */
function patched(pkg: XlsxPackage, name: string, edit: (xml: string) => string): XlsxPackage {
  const part = pkg.parts.get(name);
  if (part?.kind !== 'xml') throw new Error(`${name} is not XML`);
  return withXml(pkg, name, edit(serializeXmlPart(part.xml)));
}

/** The rename of an added `Sheet3` to `Notes`, in a store over `pkg`. */
function addedRename(pkg: XlsxPackage): unknown {
  const s = new WorkbookDocumentStore(pkg);
  expect(outcome(s, [add('Sheet3')])).toBe('ok');
  return outcome(s, [rename('Sheet3', 'Notes')]);
}

describe('a sheet that an add op of the session made', () => {
  test('takes a rename and a delete in a workbook with a chart; a sheet of the file does not', () => {
    const s = storeOf(book([chart('Sheet1!$A$1')]));
    const blocked = { code: 'structure-blocked', subject: '/xl/charts/chart1.xml' };
    expect(outcome(s, [rename('Sheet1', 'Data')])).toEqual(blocked);
    expect(outcome(s, [add('Sheet3')])).toBe('ok');
    expect(outcome(s, [rename('Sheet3', 'Notes')])).toBe('ok');
    expect(names(s)).toEqual(['Sheet1', 'Sheet2', 'Notes']);
    // The rename did not change the chart, and the file sheets stay refused.
    expect(outcome(s, [del('Sheet2')])).toEqual(blocked);
    expect(outcome(s, [del('Notes')])).toBe('ok');
    expect(names(s)).toEqual(['Sheet1', 'Sheet2']);
  });

  test('stays an added sheet through undo and redo', () => {
    const s = storeOf(book([chart()]));
    expect(outcome(s, [add('Sheet3')])).toBe('ok');
    expect(outcome(s, [del('Sheet3')])).toBe('ok');
    expect(s.undo().ok).toBe(true);
    expect(names(s)).toEqual(['Sheet1', 'Sheet2', 'Sheet3']);
    expect(outcome(s, [rename('Sheet3', 'Notes')])).toBe('ok');
    expect(s.undo().ok && s.undo().ok && s.redo().ok).toBe(true);
    expect(names(s)).toEqual(['Sheet1', 'Sheet2', 'Sheet3']);
    expect(outcome(s, [rename('Sheet3', 'Notes')])).toBe('ok');
  });

  test('is a sheet of the file after a save and a reopen', () => {
    const s = storeOf(book([chart()]));
    expect(outcome(s, [add('Sheet3')])).toBe('ok');
    const read = readXlsx(saveXlsx(s.package));
    if (!read.ok) throw new Error(read.error.reason);
    const reopened = new WorkbookDocumentStore(read.pkg);
    expect(outcome(reopened, [rename('Sheet3', 'Notes')])).toEqual({
      code: 'structure-blocked',
      subject: '/xl/charts/chart1.xml',
    });
  });

  test('a chart formula that names it, dangling before the add, refuses the rename and the delete', () => {
    for (const ref of ['Sheet3!$A$1', "'sheet3'!$A$1:$A$4", 'SHEET3!B2']) {
      const s = storeOf(book([chart(ref)]));
      expect(outcome(s, [add('Sheet3')])).toBe('ok');
      const blocked = { code: 'structure-blocked', subject: '/xl/charts/chart1.xml' };
      expect([ref, outcome(s, [rename('Sheet3', 'Notes')])]).toEqual([ref, blocked]);
      expect([ref, outcome(s, [del('Sheet3')])]).toEqual([ref, blocked]);
    }
    // A name with an apostrophe is found in its quoted spelling.
    const s = storeOf(book([chart("'Bob''s'!$A$1")]));
    expect(outcome(s, [add("Bob's")])).toBe('ok');
    expect(outcome(s, [rename("Bob's", 'Notes')])).toMatchObject({ code: 'structure-blocked' });
  });

  test('a pivot cache source that names it refuses the rename', () => {
    const cache = (source: string) =>
      [
        'xl/pivotCache/pivotCacheDefinition1.xml',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
        enc(
          `<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="worksheet"><worksheetSource ref="A1:B4" ${source}/></cacheSource></pivotCacheDefinition>`
        ),
        `${REL}/pivotCacheDefinition`,
      ] as const;
    const named = storeOf(book([cache('sheet="Sheet3"')]));
    expect(outcome(named, [add('Sheet3')])).toBe('ok');
    expect(outcome(named, [rename('Sheet3', 'Notes')])).toEqual({
      code: 'structure-blocked',
      subject: '/xl/pivotCache/pivotCacheDefinition1.xml',
    });
    // A reader reads a line break or a tab in an attribute value as a space.
    for (const name of ['My\nData', 'My\tData', 'My\r\nData']) {
      const split = storeOf(book([cache(`sheet="${name}"`)]));
      expect(outcome(split, [add('My Data')])).toBe('ok');
      expect([name, outcome(split, [rename('My Data', 'Notes')])]).toEqual([
        name,
        { code: 'structure-blocked', subject: '/xl/pivotCache/pivotCacheDefinition1.xml' },
      ]);
    }
    const other = storeOf(book([cache('sheet="Sheet1"')]));
    expect(outcome(other, [add('Sheet3')])).toBe('ok');
    expect(outcome(other, [rename('Sheet3', 'Notes')])).toBe('ok');
  });

  test('the pivot cache records hold values, not references, so they do not refuse', () => {
    const PIVOT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.';
    const pkg = open(
      book([
        [
          'xl/pivotCache/pivotCacheDefinition1.xml',
          `${PIVOT}pivotCacheDefinition+xml`,
          enc(
            '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="worksheet"><worksheetSource ref="A1:B4" sheet="Sheet1"/></cacheSource></pivotCacheDefinition>'
          ),
          `${REL}/pivotCacheDefinition`,
        ],
        [
          'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels',
          '',
          enc(
            `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/pivotCacheRecords" Target="pivotCacheRecords1.xml"/></Relationships>`
          ),
        ],
        [
          'xl/pivotCache/pivotCacheRecords1.xml',
          `${PIVOT}pivotCacheRecords+xml`,
          enc(
            '<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><r><s v="Sheet3"/></r></pivotCacheRecords>'
          ),
        ],
      ])
    );
    expect(addedRename(pkg)).toBe('ok');
  });

  test('a relationship of the content to the new sheet part refuses the rename', () => {
    // A chart that relates to a worksheet part that the file does not hold, as a pivot cache
    // source can (`worksheetSource/@r:id`, 18.10.1.95): the add gives that part to the new sheet.
    const s = storeOf(
      book([
        chart(),
        [
          'xl/charts/_rels/chart1.xml.rels',
          '',
          enc(
            `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="../worksheets/sheet3.xml"/></Relationships>`
          ),
        ],
      ])
    );
    expect(outcome(s, [add('Sheet3')])).toBe('ok');
    expect(workbookIndex(s.package)!.sheets[2]!.partName).toBe('/xl/worksheets/sheet3.xml');
    expect(outcome(s, [rename('Sheet3', 'Notes')])).toMatchObject({ code: 'structure-blocked' });
    // A target with a fragment names the same part.
    const fragment = open(
      book([
        chart(),
        [
          'xl/charts/_rels/chart1.xml.rels',
          '',
          enc(
            `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="../worksheets/sheet3.xml#x"/></Relationships>`
          ),
        ],
      ])
    );
    expect(addedRename(fragment)).toMatchObject({ code: 'structure-blocked' });
    // A relationship that the reader drops cannot be read, so it refuses.
    const dropped = open(
      book([
        chart(),
        [
          'xl/charts/_rels/chart1.xml.rels',
          '',
          enc(
            `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="../../../xl/worksheets/sheet3.xml"/></Relationships>`
          ),
        ],
      ])
    );
    expect(addedRename(dropped)).toMatchObject({ code: 'structure-blocked' });
  });

  test('a part that is not text refuses the rename: the scan cannot read it', () => {
    const s = storeOf(
      book([
        [
          'xl/opaque.bin',
          '',
          new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xff, 0xfe, 0x80]),
          `${REL}/x`,
        ],
      ])
    );
    expect(outcome(s, [add('Sheet3')])).toBe('ok');
    expect(outcome(s, [rename('Sheet3', 'Notes')])).toEqual({
      code: 'structure-blocked',
      subject: '/xl/opaque.bin',
    });
  });

  test('data connections that do not name it take the rename and the delete', () => {
    const s = storeOf(
      book([
        [
          'xl/connections.xml',
          'application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml',
          enc('<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>'),
          `${REL}/connections`,
        ],
      ])
    );
    expect(outcome(s, [add('Sheet3'), rename('Sheet3', 'Notes')])).toBe('ok');
    expect(outcome(s, [del('Notes')])).toBe('ok');
  });

  test('a delete and a new add give a sheet that is still an added sheet', () => {
    const s = storeOf(book([chart()]));
    expect(outcome(s, [add('Sheet3')])).toBe('ok');
    expect(outcome(s, [del('Sheet3')])).toBe('ok');
    expect(outcome(s, [add('Sheet3')])).toBe('ok');
    expect(outcome(s, [rename('Sheet3', 'Notes')])).toBe('ok');
  });

  test('a name split by markup or spelled with references in a chart kept as bytes refuses', () => {
    for (const ref of [
      'Sh<![CDATA[eet3]]>!$A$1',
      'Sheet<!--x-->3!$A$1',
      // Comment markers inside a CDATA section are text, not a comment.
      "('<![CDATA[<!--]]>'!A1,Sheet3!$A$1,'<![CDATA[-->]]>'!A1)",
      // A character reference can have any count of digits.
      '&#0000000083;heet3!$A$1',
      '&#x0000000053;heet3!$A$1',
    ])
      expect([ref, addedRename(open(book([chart(ref)])))]).toEqual([
        ref,
        { code: 'structure-blocked', subject: '/xl/charts/chart1.xml' },
      ]);
  });

  test('markup with no end costs linear time', () => {
    for (const opener of ['<?', '<!--', '<![CDATA[']) {
      const start = performance.now();
      const pkg = open(book([chart(opener.repeat(Math.floor(300_000 / opener.length)))]));
      expect([opener, addedRename(pkg)]).toEqual([opener, 'ok']);
      // A scan that searches to the end from each opener takes minutes on this part.
      expect([opener, performance.now() - start < 2000]).toEqual([opener, true]);
    }
  });

  test('a chart in another declared encoding is not read, so it refuses', () => {
    const utf7 = enc(
      '<?xml version="1.0" encoding="UTF-7"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>'
    );
    expect(addedRename(open(book([['xl/charts/chart1.xml', CHART, utf7]])))).toEqual({
      code: 'structure-blocked',
      subject: '/xl/charts/chart1.xml',
    });
  });

  test('a part in UTF-16 with no byte order mark is not text, so it refuses', () => {
    const text = 'Sheet3!A1';
    const bytes = new Uint8Array(text.length * 2);
    for (let i = 0; i < text.length; i += 1) bytes[i * 2] = text.charCodeAt(i);
    const pkg = open(book([['xl/opaque.bin', '', bytes, `${REL}/x`]]));
    expect(addedRename(pkg)).toEqual({ code: 'structure-blocked', subject: '/xl/opaque.bin' });
  });

  test('an external relationship refuses when its target names the sheet', () => {
    const rel = (target: string) =>
      `<Relationship Id="rIdE" Type="${REL}/x" Target="${target}" TargetMode="External"/>`;
    // A relationship of the workbook outside the allowlist.
    const top = (target: string) =>
      patched(open(book([])), '/xl/_rels/workbook.xml.rels', (xml) =>
        xml.replace('</Relationships>', `${rel(target)}</Relationships>`)
      );
    expect(addedRename(top('https://example.test/#Sheet3!A1'))).toEqual({
      code: 'structure-blocked',
      subject: `${REL}/x`,
    });
    expect(addedRename(top('https://example.test/#Sheet1!A1'))).toBe('ok');
    // A relationship of a chart.
    const rels = (target: string) =>
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel(target)}</Relationships>`;
    const nested = (target: string) =>
      open(book([chart(), ['xl/charts/_rels/chart1.xml.rels', '', enc(rels(target))]]));
    expect(addedRename(nested('https://example.test/#Sheet3!A1'))).toEqual({
      code: 'structure-blocked',
      subject: '/xl/charts/chart1.xml',
    });
    expect(addedRename(nested('https://example.test/#Sheet1!A1'))).toBe('ok');
  });

  test('a chartsheet refuses the sheets of the file, not an added sheet that it does not name', () => {
    const base = open(
      book([
        [
          'xl/chartsheets/sheet9.xml',
          'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml',
          enc(
            '<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"/></sheetViews></chartsheet>'
          ),
          `${REL}/chartsheet`,
        ],
      ])
    );
    const pkg = patched(base, '/xl/workbook.xml', (xml) =>
      xml.replace('</sheets>', '<sheet name="Chart1" sheetId="9" r:id="rIdX1"/></sheets>')
    );
    const s = new WorkbookDocumentStore(pkg);
    expect(outcome(s, [rename('Sheet1', 'Data')])).toEqual({
      code: 'structure-blocked',
      subject: 'Chart1',
    });
    expect(outcome(s, [add('Sheet3')])).toBe('ok');
    expect(outcome(s, [rename('Sheet3', 'Notes')])).toBe('ok');
  });

  test('a cell edit keeps the content that the scan read; a rename does not', () => {
    const s = storeOf(book([chart()]));
    const value = (n: number): WorkbookOp => ({
      op: 'setCellValue',
      sheet: 'Sheet1',
      ref: 'B2',
      value: { type: 'number', value: n },
    });
    // The first value edit also marks the cached values stale in the workbook part.
    expect(outcome(s, [add('Sheet3'), value(4)])).toBe('ok');
    const before = s.package;
    const index = workbookIndex(before)!;
    expect(addedSheetBlocker(before, index, index.sheets[2]!)).toBeUndefined();
    expect(outcome(s, [value(5)])).toBe('ok');
    expect(sameOpaqueContent(before, s.package, workbookIndex(s.package)!)).toBe(true);
    // Each next cell edit compares with the last revision, with no new scan.
    const second = s.package;
    expect(outcome(s, [value(6)])).toBe('ok');
    expect(sameOpaqueContent(second, s.package, workbookIndex(s.package)!)).toBe(true);
    expect(outcome(s, [rename('Sheet3', 'Notes')])).toBe('ok');
    expect(sameOpaqueContent(before, s.package, workbookIndex(s.package)!)).toBe(false);
  });

  test('without the part names of the opened package, each sheet is a sheet of the file', () => {
    const pkg = open(book([chart()]));
    const r = applyOps(pkg, [add('Sheet3'), rename('Sheet3', 'Notes')]);
    expect(r.ok ? 'ok' : r.rejection.code).toBe('structure-blocked');
  });
});

describe('macros', () => {
  test('block a sheet that counts as added, as they block every sheet', () => {
    const VBA = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject';
    const files = unzipSync(buildXlsx(book([])));
    const rels = 'xl/_rels/workbook.xml.rels';
    files[rels] = strToU8(
      strFromU8(files[rels]!).replace(
        '</Relationships>',
        `<Relationship Id="rIdMacros" Type="${VBA}" Target="vbaProject.bin"/></Relationships>`
      )
    );
    files['xl/vbaProject.bin'] = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0]);
    const read = readXlsx(zipSync(files));
    if (!read.ok) throw new Error('the package did not open');
    const pkg = read.pkg;
    const index = workbookIndex(pkg)!;
    const sheet2 = index.sheets[1]!;
    // The opened part names without the part of Sheet2: Sheet2 counts as a sheet of the session.
    const opened = new Set([...partNamesOf(pkg)].filter((n) => n !== asciiFold(sheet2.partName!)));
    expect(sheetEditBlocker(pkg, index, sheet2, opened)?.about?.kind).toBe('macros');
  });
});
