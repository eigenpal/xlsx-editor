// The read of a load that a newer load replaced, and the limit that a refused file names.

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx } from '../../store/index.ts';
import { limitOf, readForLoad, whileWanted } from '../load-steps.ts';

describe('a load that a newer load replaced', () => {
  test('stops its read at the next step', () => {
    let wanted = true;
    let steps = 0;
    let closed = false;
    function* read(): Generator<void, string> {
      try {
        for (;;) {
          steps += 1;
          yield;
        }
      } finally {
        closed = true;
      }
    }
    const reading = whileWanted(read(), () => wanted);
    reading.next();
    wanted = false;
    const end = reading.next();
    expect(end).toEqual({ done: true, value: null });
    expect(steps).toBe(2);
    expect(closed).toBe(true);
  });
});

const NS_X = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
/** Elements nested past the 256 levels of the XML part limit. */
const DEEP = `${'<x:r>'.repeat(300)}${'</x:r>'.repeat(300)}`;
const workbook = () =>
  buildXlsx({ sheets: [{ name: 'A', rows: row(1, [c('A1', 0, { t: 's' })]) }], strings: ['text'] });

/** `bytes` with the part `name` replaced by `xml`. */
function withPart(bytes: Uint8Array, name: string, xml: string): Uint8Array {
  const parts = unzipSync(bytes);
  parts[name] = strToU8(xml);
  return zipSync(parts);
}

const failureOf = async (bytes: Uint8Array) => {
  const read = await readForLoad(bytes, null, () => true);
  if (!read || read.ok) throw new Error('the load did not fail');
  return read.failure;
};

describe('the limit of a refused file', () => {
  test('a worksheet past the nested levels names xml.maxDepth', async () => {
    const deep = `<x:worksheet xmlns:x="${NS_X}"><x:sheetData/>${DEEP}</x:worksheet>`;
    const failure = await failureOf(withPart(workbook(), 'xl/worksheets/sheet1.xml', deep));
    expect(failure).toEqual({
      code: 'package',
      detail: 'xml: /xl/worksheets/sheet1.xml',
      part: '/xl/worksheets/sheet1.xml',
      limit: 'xml.maxDepth',
    });
  });

  test('a worksheet that is not valid XML names no limit', async () => {
    const bad = `<x:worksheet xmlns:x="${NS_X}"><x:sheetData>`;
    const failure = await failureOf(withPart(workbook(), 'xl/worksheets/sheet1.xml', bad));
    expect(failure).toEqual({
      code: 'package',
      detail: 'xml: /xl/worksheets/sheet1.xml',
      part: '/xl/worksheets/sheet1.xml',
    });
  });

  test('a package with too many entries names zip.maxEntries', async () => {
    const parts = unzipSync(workbook());
    for (let n = 0; n < 10_001; n += 1) parts[`extra/${n}.bin`] = new Uint8Array(1);
    const failure = await failureOf(zipSync(parts));
    expect(failure).toEqual({ code: 'package', detail: 'zip', limit: 'zip.maxEntries' });
  });

  test('a package past the compression ratio names zip.maxRatio', async () => {
    const parts = unzipSync(workbook());
    parts['extra/zeros.bin'] = new Uint8Array(9 * 1024 * 1024);
    const failure = await failureOf(zipSync(parts, { level: 9 }));
    expect(failure).toEqual({ code: 'package', detail: 'zip', limit: 'zip.maxRatio' });
  });

  test('each other limit of an XML part has its name', () => {
    const xml = (detail: 'too-large' | 'too-many-elements') =>
      limitOf({ reason: 'xml', partName: '/xl/worksheets/sheet1.xml', detail });
    expect(xml('too-large')).toBe('xml.maxBytes');
    expect(xml('too-many-elements')).toBe('xml.maxElements');
    expect(limitOf({ reason: 'zip', detail: 'too-large', limit: 'zip.maxTotalBytes' })).toBe(
      'zip.maxTotalBytes'
    );
  });

  test('a shared string table past a limit refuses the file', async () => {
    const deep = `<x:sst xmlns:x="${NS_X}"><x:si><x:t>text</x:t></x:si>${DEEP}</x:sst>`;
    const failure = await failureOf(withPart(workbook(), 'xl/sharedStrings.xml', deep));
    expect(failure).toEqual({
      code: 'package',
      detail: 'xml: /xl/sharedStrings.xml',
      part: '/xl/sharedStrings.xml',
      limit: 'xml.maxDepth',
    });
  });

  test('a shared string table that no relationship uses also refuses the file past a limit', async () => {
    // DOCX Editor reads each part by its content type, also a part that no relationship uses.
    const parts = unzipSync(workbook());
    const types = strFromU8(parts['[Content_Types].xml']!);
    const sst = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml';
    parts['[Content_Types].xml'] = strToU8(
      types.replace(
        '</Types>',
        `<Override PartName="/xl/unused.xml" ContentType="${sst}"/></Types>`
      )
    );
    parts['xl/unused.xml'] = strToU8(`<x:sst xmlns:x="${NS_X}">${DEEP}</x:sst>`);
    const failure = await failureOf(zipSync(parts));
    expect(failure).toEqual({
      code: 'package',
      detail: 'xml: /xl/unused.xml',
      part: '/xl/unused.xml',
      limit: 'xml.maxDepth',
    });
  });

  test('a shared string table that is not valid XML stays as bytes', () => {
    const bad = `<x:sst xmlns:x="${NS_X}"><x:si>`;
    const read = readXlsx(withPart(workbook(), 'xl/sharedStrings.xml', bad));
    if (!read.ok) throw new Error(read.error.reason);
    expect(read.pkg.parts.get('/xl/sharedStrings.xml')?.kind).toBe('binary');
  });
});
