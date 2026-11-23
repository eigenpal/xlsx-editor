// The copied text of a date: a shown date that reads back to the same day copies
// as shown; any other date copies its edit text, so a plain-text paste into another workbook
// stores the same value (`copiedDateText`).

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { strFromU8, unzipSync } from 'fflate';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { createXlsxEditor, type XlsxEditorInstance } from '../index.ts';
import { copiedDateText } from '../typed-forms.ts';
import { flushEditor } from './flush-editor.ts';
import type { FontSource } from '../../layout/index.ts';
import { EN_US, type NumberLocale } from '../../cascade/index.ts';

const fontsDir = join(
  dirname(Bun.resolveSync('@docx-editor.dev/fonts/package.json', import.meta.dir)),
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

/** A day-first locale with `/`. */
const GB: NumberLocale = { ...EN_US, builtins: { 14: 'dd/mm/yyyy' } };

let viewers: XlsxEditorInstance[] = [];
afterEach(() => {
  for (const v of viewers) v.destroy();
  viewers = [];
});

async function mount(bytes: Uint8Array, locale?: NumberLocale) {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.appendChild(host);
  const viewer = createXlsxEditor({ fonts: FONTS, translate: (k) => k });
  viewers.push(viewer);
  viewer.attach(host);
  const scroller = host.querySelector('.xlsx-scroller') as HTMLElement;
  Object.defineProperty(scroller, 'clientWidth', { get: () => 800 });
  Object.defineProperty(scroller, 'clientHeight', { get: () => 400 });
  await viewer.load(bytes);
  if (locale) viewer.setNumberLocale(locale);
  flushEditor(viewer);
  return viewer;
}

/** A1 holds 1; A2 holds serial 49461 (2035-06-01) with format 14. */
const book = () =>
  buildXlsx({
    sheets: [
      {
        name: 'S',
        rows: row(1, [c('A1', 1)]) + '<row r="2"><c r="A2" s="1"><v>49461</v></c></row>',
      },
    ],
    styles: {
      cellXfs: [
        '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
      ],
    },
  });

/** The saved `x:v` of `ref`, or undefined for a cell with no number. */
function savedValue(bytes: Uint8Array, ref: string): string | undefined {
  const sheet = strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']!);
  const cell = new RegExp(`<c r="${ref}"[^>]*?(?:/>|>.*?</c>)`).exec(sheet)?.[0] ?? '';
  return cell.includes('t="s"') ? undefined : /<v>([^<]*)<\/v>/.exec(cell)?.[1];
}

/** Type `entry` in B1, copy it, and paste the plain text into B1 of another workbook. */
async function roundTrip(entry: string, locale?: NumberLocale) {
  const from = await mount(book(), locale);
  from.select({ row: 0, col: 1 });
  expect(from.enterValue(entry)).toEqual({ ok: true });
  const out = from.copy();
  if (!out.ok) throw new Error(out.refusal.code);
  const to = await mount(book(), locale);
  to.select({ row: 0, col: 1 });
  expect(to.paste({ text: out.content.text }).ok).toBe(true);
  return {
    text: out.content.text,
    stored: savedValue(from.save()!, 'B1'),
    pasted: savedValue(to.save()!, 'B1'),
  };
}

describe('copied date text', () => {
  test('a date that reads back to the same day copies as shown', async () => {
    const out = await roundTrip('2026-11-12');
    expect(out.text).toBe('11-12-26');
    expect(out.pasted).toBe(out.stored);
    expect((await roundTrip('13:45')).text).toBe('13:45');
    expect((await roundTrip('36:00')).text).toBe('36:00');
  });

  test('a day and a month with no year copies its ISO 8601 date', async () => {
    for (const locale of [undefined, GB]) {
      const out = await roundTrip(locale ? '2/1' : '1/2', locale);
      expect(out.text).toMatch(/^\d{4}-01-02$/);
      expect(out.stored).toBeDefined();
      expect(out.pasted).toBe(out.stored);
    }
  });

  test('a two-digit year outside 1930 to 2029 copies its ISO 8601 date', async () => {
    for (const [entry, text] of [
      ['2035-06-01', '2035-06-01'],
      ['1925-06-01', '1925-06-01'],
      ['2045-01-01 9:00', '2045-01-01 9:00'],
    ] as const) {
      const out = await roundTrip(entry);
      expect(out.text).toBe(text);
      expect(out.pasted).toBe(out.stored);
    }
  });

  test('a date that the workbook holds copies the same way', async () => {
    const from = await mount(book());
    from.select({ row: 1, col: 0 });
    const out = from.copy();
    if (!out.ok) throw new Error(out.refusal.code);
    expect(out.content.text).toBe('2035-06-01');
  });

  test('copiedDateText keeps text that is not a date format', () => {
    expect(copiedDateText('1.5', 1.5, '0.0', EN_US, false)).toBe('1.5');
    expect(copiedDateText('12%', 0.12, '0%', EN_US, false)).toBe('12%');
    expect(copiedDateText('2-Jan', 46024, 'd-mmm', EN_US, false)).toBe('2026-01-02');
  });
});
