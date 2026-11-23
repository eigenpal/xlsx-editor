// The store's side of every rule in AGENTS.md, "Security". The bounded zip, name, and XML
// readers have their own tests in packages/ooxml-core; these prove that the workbook read goes
// through them and that the store adds no unbounded path of its own.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { strToU8, unzipSync, zipSync } from 'fflate';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx, xmlPartOf } from '../package/package.ts';
import { textOf } from '../package/query.ts';
import { sharedStringItems, sheetByName, workbookIndex } from '../workbook/workbook-index.ts';
import { sheetIndex } from '../workbook/sheet-index.ts';
import { sheetHyperlinks } from '../workbook/hyperlinks.ts';
import { saveXlsx } from '../package-write/save.ts';
import { WorkbookDocumentStore } from '../store/document-store.ts';

const FIXTURE = `${import.meta.dir}/../../../../../fixtures`;

function withPart(name: string, xml: string): Uint8Array {
  const entries = unzipSync(new Uint8Array(readFileSync(`${FIXTURE}/basic-values.xlsx`)));
  entries[name] = strToU8(xml);
  return zipSync(entries);
}

describe('XML: no DTD, no entity expansion, bounded depth', () => {
  test('a worksheet with a DTD is refused', () => {
    const read = readXlsx(
      withPart(
        'xl/worksheets/sheet1.xml',
        '<?xml version="1.0"?><!DOCTYPE w [<!ENTITY a "aaaa">]><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>'
      )
    );
    expect(read.ok ? null : read.error.reason).toBe('xml');
  });

  test('a worksheet nested past the depth cap is refused', () => {
    const deep = '<x>'.repeat(400) + '</x>'.repeat(400);
    const read = readXlsx(
      withPart(
        'xl/worksheets/sheet1.xml',
        `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><extLst>${deep}</extLst></worksheet>`
      )
    );
    expect(read.ok ? null : read.error.reason).toBe('xml');
  });

  test('an optional part that does not parse stays bytes and does not refuse the workbook', () => {
    const read = readXlsx(withPart('xl/styles.xml', '<styleSheet><unclosed></styleSheet>'));
    expect(read.ok).toBe(true);
    if (read.ok) expect(read.pkg.parts.get('/xl/styles.xml')?.kind).toBe('binary');
  });
});

describe('zip: names inside the package', () => {
  test('a zip entry name with .. is refused', () => {
    const entries = unzipSync(new Uint8Array(readFileSync(`${FIXTURE}/basic-values.xlsx`)));
    entries['xl/../../evil.xml'] = strToU8('<x/>');
    expect(readXlsx(zipSync(entries)).ok).toBe(false);
  });
});

describe('URLs reach a sink only through sanitizeHref', () => {
  test('a script hyperlink has no href, and its target stays inert in the part', () => {
    const read = readXlsx(new Uint8Array(readFileSync(`${FIXTURE}/preserved-parts.xlsx`)));
    if (!read.ok) throw new Error('read');
    const part = xmlPartOf(read.pkg, sheetByName(workbookIndex(read.pkg)!, 'Kept')!.partName)!;
    const links = sheetHyperlinks(read.pkg, part);
    expect(links).toEqual([{ ref: 'A1', href: 'https://example.com/report' }, { ref: 'B1' }]);
    const rels = new TextDecoder().decode(
      unzipSync(saveXlsx(read.pkg))['xl/worksheets/_rels/sheet1.xml.rels']
    );
    expect(rels).toContain('javascript:alert(1)');
  });
});

describe('escape on save', () => {
  test('markup in a cell string is escaped and reads back as text', () => {
    const read = readXlsx(buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] }));
    if (!read.ok) throw new Error('read');
    const store = new WorkbookDocumentStore(read.pkg);
    const hostile = '</t></si><si><t>injected</t></si>&amp;"\'<![CDATA[x]]>';
    store.transact([
      { op: 'setCellValue', sheet: 'S', ref: 'B1', value: { type: 'string', value: hostile } },
    ]);
    const reopened = readXlsx(saveXlsx(store.package));
    if (!reopened.ok) throw new Error('reopen');
    const items = sharedStringItems(reopened.pkg, workbookIndex(reopened.pkg)!);
    expect(items.map((i) => textOf(i))).toEqual([hostile]);
  });
});

describe('formulas are never evaluated on open', () => {
  test('a formula keeps its text and its cached value through open and save', () => {
    const bytes = buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 'stale', { t: 'str', f: 'HYPERLINK("http://example.com","x")' })]),
        },
      ],
    });
    const read = readXlsx(bytes);
    if (!read.ok) throw new Error('read');
    const sheet = new TextDecoder().decode(
      unzipSync(saveXlsx(read.pkg))['xl/worksheets/sheet1.xml']
    );
    expect(sheet).toContain(
      '<f>HYPERLINK(&quot;http://example.com&quot;,&quot;x&quot;)</f><v>stale</v>'
    );
  });
});

describe('resource limits: file numbers never size work', () => {
  const huge = buildXlsx({
    strings: ['a'],
    sheets: [
      {
        name: '__proto__',
        cols: '<cols><col min="1" max="16384" width="9"/></cols>',
        rows: row(1048576, [c('XFD1048576', 0, { t: 's' })]),
      },
    ],
  });
  // A dimension that names the whole grid, and counts that claim four billion strings.
  const entries = unzipSync(huge);
  const sheetXml = new TextDecoder().decode(entries['xl/worksheets/sheet1.xml']!);
  entries['xl/worksheets/sheet1.xml'] = strToU8(
    sheetXml.replace('<sheetData>', '<dimension ref="A1:XFD1048576"/><sheetData>')
  );
  const sstXml = new TextDecoder().decode(entries['xl/sharedStrings.xml']!);
  entries['xl/sharedStrings.xml'] = strToU8(
    sstXml.replace(/count="1" uniqueCount="1"/, 'count="4294967295" uniqueCount="4294967295"')
  );
  const bytes = zipSync(entries);

  test('a full-grid dimension, full-width columns, and huge counts cost only the cells that exist', () => {
    const started = performance.now();
    const read = readXlsx(bytes);
    if (!read.ok) throw new Error(JSON.stringify(read.error));
    const index = workbookIndex(read.pkg)!;
    expect(sharedStringItems(read.pkg, index)).toHaveLength(1);
    const sheet = sheetByName(index, '__proto__')!;
    expect(sheetIndex(xmlPartOf(read.pkg, sheet.partName)!).cells.size).toBe(1);
    const saved = saveXlsx(read.pkg);
    const sst = new TextDecoder().decode(unzipSync(saved)['xl/sharedStrings.xml']!);
    expect(sst).toContain('count="1"');
    expect(performance.now() - started).toBeLessThan(2000);
  });

  test('a sheet named __proto__ is found by name and pollutes no prototype', () => {
    const read = readXlsx(bytes);
    if (!read.ok) throw new Error('read');
    const index = workbookIndex(read.pkg)!;
    expect(sheetByName(index, '__proto__')?.name).toBe('__proto__');
    expect(sheetByName(index, 'constructor')).toBeUndefined();
    expect(({} as Record<string, unknown>).partName).toBeUndefined();
    const store = new WorkbookDocumentStore(read.pkg);
    expect(
      store.transact([
        { op: 'setCellValue', sheet: '__proto__', ref: 'A1', value: { type: 'number', value: 1 } },
      ]).ok
    ).toBe(true);
  });
});
