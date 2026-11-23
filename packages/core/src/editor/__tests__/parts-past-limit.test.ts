// The parts that stay as bytes past an XML part limit: the snapshot names each one with its kind
// and its limit, and an edit of another part keeps the list.

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import {
  CONTENT_TYPE,
  readXlsx,
  saveXlsx,
  workbookIndex,
  WorkbookDocumentStore,
  type XlsxPackage,
} from '../../store/index.ts';
import { en } from '@xlsx-editor.dev/i18n';
import type { WorkbookOp } from '../../store/index.ts';
import { NO_WORKBOOK_STATE, structureRefusalOf, workbookState } from '../editor-model.ts';
import { refusalOf, refusalText } from '../cell-input.ts';
import { englishText } from '../english.ts';

const NS_X = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
/** Elements nested past the 256 levels of the XML part limit. */
const deep = (tag: string) => `${`<${tag}>`.repeat(300)}${`</${tag}>`.repeat(300)}`;

const workbook = () =>
  buildXlsx({
    sheets: [
      {
        name: 'S',
        rows: row(1, [c('A1', 1)]),
        drawing: {
          anchors: '',
          images: [
            { relId: 'rIdImg', name: 'a.png', bytes: new Uint8Array(8), contentType: 'image/png' },
          ],
        },
      },
    ],
  });

/** The package of `bytes` with the part `name` replaced by `xml`. */
function readWith(name: string, xml: string): XlsxPackage {
  const parts = unzipSync(workbook());
  if (!parts[name]) throw new Error(`no part ${name}`);
  parts[name] = strToU8(xml);
  const read = readXlsx(zipSync(parts));
  if (!read.ok) throw new Error(read.error.reason);
  return read.pkg;
}

/**
 * The package with an extra part `name` of the content type `type` that nothing links to, and
 * with the part `used` replaced by `usedXml` when it is given.
 */
function readWithExtra(
  name: string,
  type: string,
  xml: string,
  used?: { name: string; xml: string }
): XlsxPackage {
  const parts = unzipSync(workbook());
  parts[name] = strToU8(xml);
  if (used) parts[used.name] = strToU8(used.xml);
  const types = strFromU8(parts['[Content_Types].xml']!);
  parts['[Content_Types].xml'] = strToU8(
    types.replace('</Types>', `<Override PartName="/${name}" ContentType="${type}"/></Types>`)
  );
  const read = readXlsx(zipSync(parts));
  if (!read.ok) throw new Error(read.error.reason);
  return read.pkg;
}

const deepStyles = `<x:styleSheet xmlns:x="${NS_X}">${deep('x:a')}</x:styleSheet>`;
const deepTheme = `<a:theme xmlns:a="${NS_A}">${deep('a:b')}</a:theme>`;

const pastLimit = (pkg: XlsxPackage) => workbookState(pkg, workbookIndex(pkg)!).partsPastLimit;

describe('the parts that stay as bytes past a limit', () => {
  test('a styles part past a limit is named with the kind styles', () => {
    const pkg = readWith(
      'xl/styles.xml',
      `<x:styleSheet xmlns:x="${NS_X}">${deep('x:a')}</x:styleSheet>`
    );
    expect(pastLimit(pkg)).toEqual([
      { part: '/xl/styles.xml', kind: 'styles', limit: 'xml.maxDepth' },
    ]);
  });

  test('a theme part past a limit is named with the kind theme', () => {
    const xml = `<a:theme xmlns:a="${NS_A}">${deep('a:b')}</a:theme>`;
    expect(pastLimit(readWith('xl/theme/theme1.xml', xml))).toEqual([
      { part: '/xl/theme/theme1.xml', kind: 'theme', limit: 'xml.maxDepth' },
    ]);
  });

  test('an unused styles part past a limit gets the kind other and locks no format', () => {
    const pkg = readWithExtra('xl/styles2.xml', CONTENT_TYPE.styles, deepStyles);
    expect(pastLimit(pkg)).toEqual([
      { part: '/xl/styles2.xml', kind: 'other', limit: 'xml.maxDepth' },
    ]);
    const r = new WorkbookDocumentStore(pkg).transact([
      { op: 'applyFormat', sheet: 'S', ref: 'A1', patch: { bold: true } },
    ]);
    expect(r.ok).toBe(true);
  });

  test('an unused theme part past a limit gets the kind other', () => {
    const pkg = readWithExtra('xl/theme/theme9.xml', CONTENT_TYPE.theme, deepTheme);
    expect(pastLimit(pkg)).toEqual([
      { part: '/xl/theme/theme9.xml', kind: 'other', limit: 'xml.maxDepth' },
    ]);
  });

  test('the used styles and theme parts keep their kinds beside unused parts of the same type', () => {
    const styles = readWithExtra('xl/styles2.xml', CONTENT_TYPE.styles, deepStyles, {
      name: 'xl/styles.xml',
      xml: deepStyles,
    });
    expect(pastLimit(styles)).toEqual([
      { part: '/xl/styles.xml', kind: 'styles', limit: 'xml.maxDepth' },
      { part: '/xl/styles2.xml', kind: 'other', limit: 'xml.maxDepth' },
    ]);
    const r = new WorkbookDocumentStore(styles).transact([
      { op: 'applyFormat', sheet: 'S', ref: 'A1', patch: { bold: true } },
    ]);
    expect(r.ok ? null : r.rejection.code).toBe('unsupported-package');
    const theme = readWithExtra('xl/theme/theme9.xml', CONTENT_TYPE.theme, deepTheme, {
      name: 'xl/theme/theme1.xml',
      xml: deepTheme,
    });
    expect(pastLimit(theme).map((p) => [p.part, p.kind])).toEqual([
      ['/xl/theme/theme1.xml', 'theme'],
      ['/xl/theme/theme9.xml', 'other'],
    ]);
  });

  test('a drawing past a limit is named with the kind drawing', () => {
    const xml = `<xdr:wsDr xmlns:xdr="${NS_XDR}">${deep('xdr:a')}</xdr:wsDr>`;
    expect(pastLimit(readWith('xl/drawings/drawing1.xml', xml))).toEqual([
      { part: '/xl/drawings/drawing1.xml', kind: 'drawing', limit: 'xml.maxDepth' },
    ]);
  });

  test('a relationship part of a drawing past a limit is named with the kind relationships', () => {
    const xml = `<Relationships xmlns="${NS_REL}">${deep('a')}</Relationships>`;
    expect(pastLimit(readWith('xl/drawings/_rels/drawing1.xml.rels', xml))).toEqual([
      {
        part: '/xl/drawings/_rels/drawing1.xml.rels',
        kind: 'relationships',
        limit: 'xml.maxDepth',
      },
    ]);
  });

  test('a drawing past a limit locks row and column edits of its sheet, with their messages', () => {
    const english = englishText;
    const drawing = `<xdr:wsDr xmlns:xdr="${NS_XDR}">${deep('xdr:a')}</xdr:wsDr>`;
    const pkg = readWith('xl/drawings/drawing1.xml', drawing);
    // The structural commands read this refusal as their disabled state.
    const blocked = structureRefusalOf(pkg, 'S');
    expect(blocked).toMatchObject({ code: 'structure-blocked', about: { kind: 'drawing' } });
    expect(refusalText(blocked!, english)).toBe(
      english('editing.refused.structure-blocked', { subject: english('structureSubject.drawing') })
    );
    const resize = new WorkbookDocumentStore(pkg).transact([
      { op: 'setColumnWidth', sheet: 'S', first: 1, last: 1, width: 30 },
    ]);
    expect(resize.ok).toBe(false);
    if (resize.ok) return;
    expect(refusalText(refusalOf(resize.rejection), english)).toBe(
      english('editing.resizeUnreadDrawing')
    );
    // The relationship part of the drawing past a limit locks them too.
    const rels = `<Relationships xmlns="${NS_REL}">${deep('a')}</Relationships>`;
    expect(
      structureRefusalOf(readWith('xl/drawings/_rels/drawing1.xml.rels', rels), 'S')
    ).toMatchObject({ code: 'structure-blocked', subject: '/xl/drawings/_rels/drawing1.xml.rels' });
  });

  test('a part that is not valid XML is not named: it passes no limit', () => {
    const pkg = readWith('xl/styles.xml', `<x:styleSheet xmlns:x="${NS_X}">`);
    expect(pkg.parts.get('/xl/styles.xml')?.kind).toBe('binary');
    expect(pastLimit(pkg)).toEqual([]);
  });

  test('a workbook with no part past a limit, and no workbook, give an empty list', () => {
    const read = readXlsx(workbook());
    if (!read.ok) throw new Error(read.error.reason);
    expect(pastLimit(read.pkg)).toEqual([]);
    expect(NO_WORKBOOK_STATE.partsPastLimit).toEqual([]);
  });

  test('the list keeps its identity, and an edit of a cell keeps the part in it', () => {
    const pkg = readWith(
      'xl/styles.xml',
      `<x:styleSheet xmlns:x="${NS_X}">${deep('x:a')}</x:styleSheet>`
    );
    const before = pastLimit(pkg);
    expect(pastLimit(pkg)).toBe(before);
    const store = new WorkbookDocumentStore(pkg);
    const edit = store.transact([
      { op: 'setColumnWidth', sheet: 'S', first: 1, last: 1, width: 20 },
    ]);
    expect(edit.ok).toBe(true);
    expect(store.package).not.toBe(pkg);
    expect(pastLimit(store.package)).toBe(before);
    expect(before).toEqual([{ part: '/xl/styles.xml', kind: 'styles', limit: 'xml.maxDepth' }]);
  });

  test('a change of the parts past a limit gives a new list', () => {
    const parts = unzipSync(workbook());
    parts['xl/styles.xml'] = strToU8(
      `<x:styleSheet xmlns:x="${NS_X}">${deep('x:a')}</x:styleSheet>`
    );
    parts['xl/theme/theme1.xml'] = strToU8(`<a:theme xmlns:a="${NS_A}">${deep('a:b')}</a:theme>`);
    const read = readXlsx(zipSync(parts));
    if (!read.ok) throw new Error(read.error.reason);
    const both = pastLimit(read.pkg);
    expect(both.map((p) => p.kind)).toEqual(['styles', 'theme']);
    const same: XlsxPackage = { parts: new Map(read.pkg.parts) };
    expect(pastLimit(same)).toBe(both);
    const fewer = new Map(read.pkg.parts);
    fewer.delete('/xl/theme/theme1.xml');
    const one = pastLimit({ parts: fewer });
    expect(one).not.toBe(both);
    expect(one).toEqual([{ part: '/xl/styles.xml', kind: 'styles', limit: 'xml.maxDepth' }]);
  });

  test('a save after an edit keeps the bytes of the part', () => {
    const styles = `<x:styleSheet xmlns:x="${NS_X}">${deep('x:a')}</x:styleSheet>`;
    const store = new WorkbookDocumentStore(readWith('xl/styles.xml', styles));
    expect(
      store.transact([{ op: 'setColumnWidth', sheet: 'S', first: 1, last: 1, width: 20 }]).ok
    ).toBe(true);
    expect(strFromU8(unzipSync(saveXlsx(store.package))['xl/styles.xml']!)).toBe(styles);
  });
});

describe('a format change while the styles part stays as bytes', () => {
  /** A store whose styles part passes a limit, with A1 in the cell format 1. */
  const store = () => {
    const parts = unzipSync(workbook());
    const sheet = strFromU8(parts['xl/worksheets/sheet1.xml']!);
    parts['xl/worksheets/sheet1.xml'] = strToU8(sheet.replace('<c r="A1"', '<c r="A1" s="1"'));
    parts['xl/styles.xml'] = strToU8(
      `<x:styleSheet xmlns:x="${NS_X}">${deep('x:a')}</x:styleSheet>`
    );
    const read = readXlsx(zipSync(parts));
    if (!read.ok) throw new Error(read.error.reason);
    expect(read.pkg.parts.get('/xl/styles.xml')?.kind).toBe('binary');
    return new WorkbookDocumentStore(read.pkg);
  };
  const refusal = (op: WorkbookOp) => {
    const r = store().transact([op]);
    return r.ok ? null : r.rejection.code;
  };

  test('applyFormat is refused', () => {
    expect(refusal({ op: 'applyFormat', sheet: 'S', ref: 'A1', patch: { bold: true } })).toBe(
      'unsupported-package'
    );
  });

  test('clearFormats gives the default format 0', () => {
    expect(refusal({ op: 'clearFormats', sheet: 'S', ref: 'A1' })).toBeNull();
  });

  test('setCellFormatIndex gives the format 0 and refuses every other index', () => {
    expect(refusal({ op: 'setCellFormatIndex', sheet: 'S', ref: 'A1', index: 0 })).toBeNull();
    expect(refusal({ op: 'setCellFormatIndex', sheet: 'S', ref: 'B1', index: 1 })).toBe(
      'unsupported-package'
    );
  });

  test('on a protected sheet, clearFormats is refused: the protection of each cell is unread', () => {
    const parts = unzipSync(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            rows: row(1, [c('A1', 1)]),
            beforeMerges: '<sheetProtection sheet="1" formatCells="0"/>',
          },
        ],
      })
    );
    parts['xl/styles.xml'] = strToU8(deepStyles);
    const read = readXlsx(zipSync(parts));
    if (!read.ok) throw new Error(read.error.reason);
    const r = new WorkbookDocumentStore(read.pkg).transact([
      { op: 'clearFormats', sheet: 'S', ref: 'A1' },
    ]);
    expect(r.ok ? null : r.rejection.code).toBe('unsupported-package');
  });

  test('the notice names what gives the default format and what is refused', () => {
    expect(en.status.pastLimit.styles).toContain(
      'On a sheet that is not protected, Paste, Clear formats, and Clear all give the default format. Every other format change is refused.'
    );
  });
});
