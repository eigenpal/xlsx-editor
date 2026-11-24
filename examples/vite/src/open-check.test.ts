// The demo's open check (`open-check.ts`) gives the same answer as the
// editor's `openingSheet` for each workbook: visible, hidden, and very hidden sheets, an active
// tab on a hidden sheet, and no sheet. The limit values of its messages and of
// docs/site/content/workbook-limits.mdx stay equal to the constants of packages/ooxml-core.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  DEFAULT_ZIP_LIMITS,
  ZIP_HARD_LIMITS,
  ZIP_RATIO_EXEMPT_BYTES,
} from '../../../packages/ooxml-core/src/zip.ts';
import {
  MAX_XML_DEPTH,
  XML_DEFAULT_MAX_ELEMENTS,
  XML_HARD_MAX_BYTES,
} from '../../../packages/ooxml-core/src/xml-reader.ts';
import { buildXlsx, type FixtureSheet } from '../../../scripts/fixtures/builder.ts';
import { openingSheet } from '../../../packages/core/src/editor/editor-model.ts';
import { editorSheets } from '../../../packages/core/src/editor/editor-sheets.ts';
import {
  MAX_STORED_SHEET_NAME,
  readXlsx,
  workbookIndex,
} from '../../../packages/core/src/store/index.ts';
import { hasSheetToShow, readForOpen, readProblem, whyNotOpen } from './open-check.ts';

const sheet = (name: string, state?: FixtureSheet['state']): FixtureSheet => ({
  name,
  rows: '',
  ...(state ? { state } : {}),
});

/** The workbook of `bytes` with no `x:sheet` in its workbook part. */
function noSheets(bytes: Uint8Array): Uint8Array {
  const parts = unzipSync(bytes);
  const name = 'xl/workbook.xml';
  parts[name] = strToU8(strFromU8(parts[name]!).replace(/<sheets>[\s\S]*?<\/sheets>/, '<sheets/>'));
  return zipSync(parts);
}

const CASES: [string, Uint8Array][] = [
  ['one visible sheet', buildXlsx({ sheets: [sheet('A')] })],
  ['a hidden first sheet', buildXlsx({ sheets: [sheet('A', 'hidden'), sheet('B')] })],
  [
    'an active tab on a hidden sheet',
    buildXlsx({ sheets: [sheet('A'), sheet('B', 'hidden')], activeTab: 1 }),
  ],
  ['only hidden sheets', buildXlsx({ sheets: [sheet('A', 'hidden'), sheet('B', 'veryHidden')] })],
  ['no sheet', noSheets(buildXlsx({ sheets: [sheet('A')] }))],
];

describe('the demo open check', () => {
  for (const [name, bytes] of CASES)
    test(`gives the answer of openingSheet for ${name}`, () => {
      const read = readXlsx(bytes);
      if (!read.ok) throw new Error(read.error.reason);
      const index = workbookIndex(read.pkg)!;
      const editor = openingSheet(editorSheets(read.pkg, index), index.activeTab) >= 0;
      expect(hasSheetToShow(index)).toBe(editor);
    });

  test('the cases cover both answers', () => {
    const answers = CASES.map(([, bytes]) => {
      const read = readXlsx(bytes);
      return read.ok && hasSheetToShow(workbookIndex(read.pkg)!);
    });
    expect(new Set(answers)).toEqual(new Set([true, false]));
  });
});

describe('whyNotOpen', () => {
  test('names why a file does not open, and gives null for a workbook that opens', () => {
    expect(whyNotOpen(readXlsx(new TextEncoder().encode('not a workbook')))).toBe(
      'the file is not a ZIP package'
    );
    const one = buildXlsx({ sheets: [sheet('A')] });
    expect(whyNotOpen(readXlsx(one))).toBeNull();
    expect(whyNotOpen(readXlsx(noSheets(one)))).toBe('the file has no worksheet to show');
  });

  test('names the limit that a package or a part passes', () => {
    expect(readProblem({ reason: 'zip', detail: 'too-large', limit: 'zip.maxTotalBytes' })).toBe(
      'the package passes the limit of 512 MiB of content'
    );
    expect(readProblem({ reason: 'zip', detail: 'too-large', limit: 'zip.maxRatio' })).toBe(
      'the package passes the limit of a compression ratio of 200, which applies above 8 MiB of content'
    );
    expect(readProblem({ reason: 'zip', detail: 'too-many-entries' })).toBe(
      'the package passes the limit of 10,000 ZIP entries'
    );
    expect(readProblem({ reason: 'xml', partName: '/xl/a.xml', detail: 'too-large' })).toBe(
      'the part /xl/a.xml passes the limit of 64 MiB of XML'
    );
    expect(readProblem({ reason: 'xml', partName: '/xl/a.xml', detail: 'too-many-elements' })).toBe(
      'the part /xl/a.xml passes the limit of 10,000,000 elements'
    );
    expect(readProblem({ reason: 'xml', partName: '/xl/a.xml', detail: 'too-deep' })).toBe(
      'the part /xl/a.xml passes the limit of 256 nested levels'
    );
  });

  test('cuts a long part name to 64 characters, as each name from the file', () => {
    const name = `/xl/${'a'.repeat(300)}z.xml`;
    const text = readProblem({ reason: 'xml', partName: name, detail: 'too-deep' });
    const shown = text.slice('the part '.length, text.indexOf(' passes'));
    expect([...shown]).toHaveLength(64);
    expect(shown.startsWith('/xl/aaa')).toBe(true);
    expect(shown.endsWith('a\u2026')).toBe(true);
  });

  test('shows no bidirectional format control of a part name, raw or percent-encoded', () => {
    const raw = readProblem({ reason: 'xml', partName: '/xl/\u202Elmx.a', detail: 'too-deep' });
    expect(raw).toBe('the part /xl/lmx.a passes the limit of 256 nested levels');
    // The ZIP item name holds U+202E percent-encoded, and the read decodes it once.
    const ENCODED = '%E2%80%AEsharedStrings.xml';
    const parts = unzipSync(buildXlsx({ sheets: [{ name: 'S', rows: '' }], strings: ['x'] }));
    const deep = `${'<a>'.repeat(300)}${'</a>'.repeat(300)}`;
    const renamed: Record<string, Uint8Array> = {};
    for (const [name, data] of Object.entries(parts)) {
      if (name === 'xl/sharedStrings.xml') renamed[`xl/${ENCODED}`] = strToU8(`<sst>${deep}</sst>`);
      else if (name.endsWith('.rels') || name === '[Content_Types].xml')
        renamed[name] = strToU8(strFromU8(data).replaceAll('sharedStrings.xml', ENCODED));
      else renamed[name] = data;
    }
    const read = readXlsx(zipSync(renamed));
    expect(read.ok).toBe(false);
    const text = whyNotOpen(read);
    expect(text).not.toContain('\u202E');
    expect(text).toBe('the part /xl/sharedStrings.xml passes the limit of 256 nested levels');
  });

  test('names a size limit for a package that passes one', () => {
    const bomb = zipSync({ 'xl/zeros.bin': new Uint8Array(100 * 1024 * 1024) }, { level: 9 });
    expect(whyNotOpen(readXlsx(bomb))).toBe(
      'the package passes the limit of a compression ratio of 200, which applies above 8 MiB of content'
    );
  });
});

describe('the limit values of the messages and of docs/site/content/workbook-limits.mdx', () => {
  const MiB = 1024 * 1024;
  const count = (n: number) => n.toLocaleString('en-US');
  const size = (n: number) => (n >= 1024 * MiB ? `${n / (1024 * MiB)} GiB` : `${n / MiB} MiB`);
  const zip = { ...DEFAULT_ZIP_LIMITS, maxRatio: DEFAULT_ZIP_LIMITS.maxRatio ?? 0 };
  const doc = readFileSync(
    new URL('../../../docs/site/content/workbook-limits.mdx', import.meta.url),
    'utf8'
  );
  const part = (detail: 'too-large' | 'too-many-elements' | 'too-deep') =>
    readProblem({ reason: 'xml', partName: '/xl/a.xml', detail });

  test('each message gives the value of its constant', () => {
    const total = readProblem({ reason: 'zip', detail: 'too-large', limit: 'zip.maxTotalBytes' });
    expect(total).toEndWith(`${size(zip.maxTotalBytes)} of content`);
    expect(readProblem({ reason: 'zip', detail: 'too-large', limit: 'zip.maxRatio' })).toEndWith(
      `a compression ratio of ${zip.maxRatio}, which applies above ${size(ZIP_RATIO_EXEMPT_BYTES)} of content`
    );
    expect(readProblem({ reason: 'zip', detail: 'too-many-entries' })).toEndWith(
      `${count(zip.maxEntries)} ZIP entries`
    );
    expect(part('too-large')).toEndWith(`${size(XML_HARD_MAX_BYTES)} of XML`);
    expect(part('too-many-elements')).toEndWith(`${count(XML_DEFAULT_MAX_ELEMENTS)} elements`);
    expect(part('too-deep')).toEndWith(`${MAX_XML_DEPTH} nested levels`);
    expect(
      readProblem({
        reason: 'workbook',
        detail: 'sheet-name-too-long',
        limit: 'workbook.maxSheetNameLength',
      })
    ).toEndWith(`${count(MAX_STORED_SHEET_NAME)} UTF-16 code units in one sheet name`);
  });

  test('the page gives the value of each constant', () => {
    const row = (name: string, value: string) =>
      expect(doc).toMatch(new RegExp(`\\| ${name} +\\| ${value} +\\|`));
    row('ZIP entries', count(zip.maxEntries));
    row('Uncompressed content of all entries', size(zip.maxTotalBytes));
    row(
      `Compression ratio, once the uncompressed total passes ${size(ZIP_RATIO_EXEMPT_BYTES)}`,
      String(zip.maxRatio)
    );
    expect(doc).toContain(
      `hard caps of ${count(ZIP_HARD_LIMITS.maxEntries)} entries and ${size(ZIP_HARD_LIMITS.maxTotalBytes)}`
    );
    expect(doc).toContain(`- ${size(XML_HARD_MAX_BYTES)} of XML.`);
    expect(doc).toContain(`- ${count(XML_DEFAULT_MAX_ELEMENTS)} elements.`);
    expect(doc).toContain(`- ${MAX_XML_DEPTH} nested levels.`);
  });
});

describe('readForOpen', () => {
  const bytes = buildXlsx({ sheets: [sheet('Data')] });

  test('reads the file while the pick is the newest', async () => {
    expect((await readForOpen(bytes, () => true))?.ok).toBe(true);
  });

  test('stops when a newer pick replaces this one', async () => {
    expect(await readForOpen(bytes, () => false)).toBeNull();
  });
});
