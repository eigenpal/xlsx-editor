/** @jsxImportSource react */
// The messages of a file past a limit of the read: a refused file is too large to open, and the
// message names the part and the limit; a part that stays as bytes past a limit shows a notice
// that names it.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { cleanup, render, waitFor } from '@testing-library/react';
import { strToU8, unzipSync, zipSync } from 'fflate';
import {
  packageFailureText,
  pastLimitText,
  type PartPastLimit,
} from '@xlsx-editor.dev/core/editor';
import { MAX_STORED_SHEET_NAME } from '@xlsx-editor.dev/core/store';
import { buildXlsx, c, row } from '../../../../scripts/fixtures/builder.ts';
import { DEFAULT_ZIP_LIMITS, ZIP_RATIO_EXEMPT_BYTES } from '../../../ooxml-core/src/zip.ts';
import {
  MAX_XML_DEPTH,
  XML_DEFAULT_MAX_ELEMENTS,
  XML_HARD_MAX_BYTES,
} from '../../../ooxml-core/src/xml-reader.ts';
import { translatorFor } from '../context.ts';
import { XlsxEditor, type FontSource } from '../index.ts';

const t = translatorFor(undefined);
const fontsDir = join(
  dirname(
    Bun.resolveSync('@docx-editor.dev/fonts/package.json', join(import.meta.dir, '../../../core'))
  ),
  'assets'
);
const FONTS: FontSource[] = [
  {
    family: 'Carlito',
    bold: false,
    italic: false,
    bytes: new Uint8Array(readFileSync(join(fontsDir, 'Carlito-Regular.ttf'))),
    origin: 'substitute',
  },
];

afterEach(cleanup);

const NS_X = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const DEEP = `${'<x:a>'.repeat(300)}${'</x:a>'.repeat(300)}`;

/** A workbook with the part `name` replaced by `xml`. */
function workbookWith(name: string, xml: string): Uint8Array {
  const parts = unzipSync(
    buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 0, { t: 's' })]) }], strings: ['x'] })
  );
  parts[name] = strToU8(xml);
  return zipSync(parts);
}

describe('the message of a file past a limit', () => {
  test('names the part and the limit, as DOCX Editor says a file is too large', () => {
    expect(packageFailureText({ limit: 'xml.maxDepth', part: '/xl/sharedStrings.xml' }, t)).toBe(
      'This file is too large or too complex to open here. The part /xl/sharedStrings.xml passes the limit of 256 nested levels.'
    );
  });

  test('gives a next step only for a limit that a smaller copy can pass', () => {
    const SMALLER = ' Try a smaller copy, for example one with fewer rows or sheets.';
    const text = (limit: Parameters<typeof packageFailureText>[0]['limit']) =>
      packageFailureText({ limit, part: '/xl/a.xml' }, t);
    expect(text('zip.maxEntries')).toEndWith(SMALLER);
    expect(text('zip.maxTotalBytes')).toEndWith(SMALLER);
    expect(text('xml.maxBytes')).toEndWith(SMALLER);
    expect(text('xml.maxElements')).toEndWith(SMALLER);
    expect(text('zip.maxRatio')).not.toContain('Try');
    expect(text('xml.maxDepth')).not.toContain('Try');
    expect(text('workbook.maxSheetNameLength')).not.toContain('Try');
  });

  test('a very long part name is cut to 64 characters, as each name from the file', () => {
    const part = `/xl/${'a'.repeat(5000)}/end.xml`;
    const text = packageFailureText({ limit: 'xml.maxDepth', part }, t);
    const shown = text.slice(text.indexOf('/xl/'), text.indexOf(' passes'));
    expect([...shown]).toHaveLength(64);
    expect(shown.startsWith('/xl/aaa')).toBe(true);
    expect(shown.endsWith('a\u2026')).toBe(true);
    const notice = pastLimitText([{ part, kind: 'theme', limit: 'xml.maxDepth' }], t);
    expect(notice.length).toBeLessThan(64 + 200);
    const short = '/xl/sharedStrings.xml';
    expect(packageFailureText({ limit: 'xml.maxDepth', part: short }, t)).toContain(short);
  });

  test('a part name shows no bidirectional format control, which could reorder the message', () => {
    const part = '/xl/\u202Elmx.1teehs';
    const text = packageFailureText({ limit: 'xml.maxDepth', part }, t);
    expect(text).not.toContain('\u202E');
    expect(text).toContain('The part /xl/lmx.1teehs passes');
    const notice = pastLimitText([{ part, kind: 'theme', limit: 'xml.maxDepth' }], t);
    expect(notice).not.toContain('\u202E');
    expect(notice).toContain('/xl/lmx.1teehs');
  });

  test('a limit of the whole package names no part', () => {
    expect(packageFailureText({ limit: 'zip.maxEntries' }, t)).toBe(
      'This file is too large or too complex to open here. It passes the limit of 10,000 ZIP entries. Try a smaller copy, for example one with fewer rows or sheets.'
    );
  });

  test('a file with no limit keeps the message of a file that is not a workbook', () => {
    expect(packageFailureText({ part: '/xl/worksheets/sheet1.xml' }, t)).toBe(
      'This file is not a workbook that the editor can open.'
    );
  });

  test('each limit gives the value of its constant', () => {
    const MiB = 1024 * 1024;
    const count = (n: number) => n.toLocaleString('en-US');
    const limit = (name: Parameters<typeof packageFailureText>[0]['limit']) =>
      packageFailureText({ limit: name }, t);
    expect(limit('zip.maxEntries')).toContain(
      `${count(DEFAULT_ZIP_LIMITS.maxEntries)} ZIP entries`
    );
    expect(limit('zip.maxTotalBytes')).toContain(
      `${DEFAULT_ZIP_LIMITS.maxTotalBytes / MiB} MiB of content`
    );
    expect(limit('zip.maxRatio')).toContain(
      `a compression ratio of ${DEFAULT_ZIP_LIMITS.maxRatio}, which applies above ${ZIP_RATIO_EXEMPT_BYTES / MiB} MiB of content`
    );
    expect(limit('xml.maxBytes')).toContain(`${XML_HARD_MAX_BYTES / MiB} MiB of XML`);
    expect(limit('xml.maxElements')).toContain(`${count(XML_DEFAULT_MAX_ELEMENTS)} elements`);
    expect(limit('xml.maxDepth')).toContain(`${MAX_XML_DEPTH} nested levels`);
    expect(limit('workbook.maxSheetNameLength')).toContain(
      `${count(MAX_STORED_SHEET_NAME)} UTF-16 code units in one sheet name`
    );
  });

  test('the editor shows it when a shared string table passes a limit', async () => {
    const bytes = workbookWith('xl/sharedStrings.xml', `<x:sst xmlns:x="${NS_X}">${DEEP}</x:sst>`);
    const view = render(<XlsxEditor document={bytes} fonts={FONTS} />);
    await waitFor(() => expect(view.getByRole('alert').textContent).toContain('too large'), {
      timeout: 5000,
    });
    const alert = view.getByRole('alert').textContent ?? '';
    expect(alert).toContain(
      'The part /xl/sharedStrings.xml passes the limit of 256 nested levels.'
    );
  });

  test('the editor shows a percent-encoded bidirectional control in a part name as nothing', async () => {
    // The ZIP item name holds U+202E percent-encoded. The read decodes it once, so the part
    // name holds the control itself.
    const parts = unzipSync(
      workbookWith('xl/sharedStrings.xml', `<x:sst xmlns:x="${NS_X}">${DEEP}</x:sst>`)
    );
    const ENCODED = '%E2%80%AEsharedStrings.xml';
    const renamed: Record<string, Uint8Array> = {};
    for (const [name, data] of Object.entries(parts)) {
      const key = name === 'xl/sharedStrings.xml' ? `xl/${ENCODED}` : name;
      renamed[key] =
        name.endsWith('.rels') || name === '[Content_Types].xml'
          ? strToU8(new TextDecoder().decode(data).replaceAll('sharedStrings.xml', ENCODED))
          : data;
    }
    const view = render(<XlsxEditor document={zipSync(renamed)} fonts={FONTS} />);
    await waitFor(() => expect(view.getByRole('alert').textContent).toContain('too large'), {
      timeout: 5000,
    });
    const alert = view.getByRole('alert').textContent ?? '';
    expect(alert).not.toContain('\u202E');
    expect(alert).toContain(
      'The part /xl/sharedStrings.xml passes the limit of 256 nested levels.'
    );
  });
});

describe('the notice of a part that stays as bytes past a limit', () => {
  const part = (kind: PartPastLimit['kind'], name: string): PartPastLimit => ({
    part: name,
    kind,
    limit: 'xml.maxDepth',
  });

  test('names the part and what the workbook shows without it, for each kind', () => {
    const text = (p: PartPastLimit) => pastLimitText([p], t);
    expect(text(part('styles', '/xl/styles.xml'))).toBe(
      'The part /xl/styles.xml is too large or too complex to show, so every cell shows the default format. On a sheet that is not protected, Paste, Clear formats, and Clear all give the default format. Every other format change is refused. A save keeps the part as it is.'
    );
    expect(text(part('theme', '/xl/theme/theme1.xml'))).toContain(
      '/xl/theme/theme1.xml is too large or too complex to show, so the theme colors and fonts do not apply.'
    );
    expect(text(part('drawing', '/xl/drawings/drawing1.xml'))).toContain(
      '/xl/drawings/drawing1.xml is too large or too complex to show, so its pictures do not show.'
    );
    expect(text(part('relationships', '/xl/drawings/_rels/drawing1.xml.rels'))).toContain(
      'so the parts that it links to do not show.'
    );
    expect(text(part('other', '/xl/styles2.xml'))).toBe(
      'The part /xl/styles2.xml is too large or too complex to show. The workbook does not use it, so no cell depends on it. A save keeps the part as it is.'
    );
  });

  test('counts the other parts', () => {
    const one = part('styles', '/xl/styles.xml');
    const theme = part('theme', '/xl/theme/theme1.xml');
    const drawing = part('drawing', '/xl/drawings/drawing1.xml');
    expect(pastLimitText([one, theme], t)).toEndWith(
      ' 1 other part is also too large or too complex to show.'
    );
    expect(pastLimitText([one, theme, drawing], t)).toEndWith(
      ' 2 other parts are also too large or too complex to show.'
    );
    expect(pastLimitText([], t)).toBe('');
  });

  test('the editor opens the workbook and shows the notice when the styles part passes a limit', async () => {
    const bytes = workbookWith(
      'xl/styles.xml',
      `<x:styleSheet xmlns:x="${NS_X}">${DEEP}</x:styleSheet>`
    );
    const view = render(
      <div style={{ height: 400 }}>
        <XlsxEditor document={bytes} fonts={FONTS} />
      </div>
    );
    const notices = () =>
      [...view.container.querySelectorAll('.xlsx-notice')].map((n) => n.textContent ?? '');
    await waitFor(
      () => expect(notices().some((text) => text.includes('/xl/styles.xml'))).toBe(true),
      { timeout: 5000 }
    );
    const notice = [...view.container.querySelectorAll('.xlsx-notice')].find((n) =>
      n.textContent?.includes('/xl/styles.xml')
    )!;
    expect(notice.getAttribute('role')).toBe('status');
    expect(notice.querySelector('.xlsx-notice-full')!.textContent).toContain(
      'so every cell shows the default format'
    );
    expect(view.queryByRole('alert')).toBeNull();
  });
});
