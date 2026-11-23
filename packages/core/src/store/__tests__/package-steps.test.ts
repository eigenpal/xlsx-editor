// The stepped read and save: the same tree, node ids, bytes, and refusals as the one-pass code,
// with yields on a long part.

import { describe, expect, test } from 'bun:test';
import { buildXlsx } from '../../../../../scripts/fixtures/builder.ts';
import {
  parseXmlPart,
  parseXmlPartSteps,
  readXlsx,
  readXlsxSteps,
  saveXlsx,
  saveXlsxSteps,
  serializeXmlPart,
  serializeXmlPartSteps,
  type XmlPart,
} from '../index.ts';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

/** Runs `steps` to the end, and counts the yields. */
function drain<T>(steps: Generator<void, T>): { value: T; yields: number } {
  let yields = 0;
  for (;;) {
    const next = steps.next();
    if (next.done) return { value: next.value, yields };
    yields += 1;
  }
}

/** A shared string table of more than 2000000 characters, with one item of many runs. */
function longTable(): string {
  const plain = Array.from({ length: 60_000 }, (_, i) => `<si><t>text number ${i}</t></si>`);
  const runs = Array.from({ length: 5_000 }, (_, i) => `<r><t xml:space="preserve">r${i} </t></r>`);
  return `<sst xmlns="${NS}">${plain.join('')}<si>${runs.join('')}</si></sst>`;
}

const part = (xml: string): XmlPart => {
  const read = parseXmlPart('/xl/sharedStrings.xml', xml);
  if (!read.ok) throw new Error(read.reason);
  return read.part;
};

describe('parseXmlPartSteps', () => {
  test('a long part gives the tree and node ids of the one-pass read, and it yields', () => {
    const xml = longTable();
    expect(xml.length).toBeGreaterThan(2_000_000);
    const { value, yields } = drain(parseXmlPartSteps('/xl/sharedStrings.xml', xml));
    expect(yields).toBeGreaterThan(100);
    expect(value).toEqual(parseXmlPart('/xl/sharedStrings.xml', xml));
  });

  test('a long part outside the plain subset reads as the bounded reader reads it', () => {
    const xml = longTable().replace('<si>', '<!-- a comment --><si>');
    const { value } = drain(parseXmlPartSteps('/xl/sharedStrings.xml', xml));
    expect(value).toEqual(parseXmlPart('/xl/sharedStrings.xml', xml));
    expect(value.ok).toBe(true);
  });

  test.each([
    ['a part over the size limit', longTable(), { maxBytes: 1000 }],
    ['a part that is not well formed', `${longTable()}<extra/>`, undefined],
    ['a part with a document type', `<!DOCTYPE x>${longTable()}`, undefined],
    ['a part with an unbound prefix', longTable().replace('<si>', '<q:si>'), undefined],
    [
      'a part past the nested levels',
      longTable().replace('</sst>', `${'<a>'.repeat(300)}${'</a>'.repeat(300)}</sst>`),
      undefined,
    ],
    [
      'a part past the element limit',
      longTable(),
      { maxBytes: 64 * 1024 * 1024, maxElements: 1000 },
    ],
  ])('%s has the refusal of the one-pass read', (_name, xml, limits) => {
    const { value } = drain(parseXmlPartSteps('/xl/sharedStrings.xml', xml, limits));
    expect(value.ok).toBe(false);
    expect(value).toEqual(parseXmlPart('/xl/sharedStrings.xml', xml, limits));
  });
});

describe('serializeXmlPartSteps', () => {
  test('a long part gives the text of the one-pass serializer, and it yields', () => {
    const tree = part(longTable());
    const { value, yields } = drain(serializeXmlPartSteps(tree));
    expect(yields).toBeGreaterThan(100);
    expect(value).toBe(serializeXmlPart(tree));
  });
});

describe('a part of few wide rows', () => {
  const cells = Array.from({ length: 7000 }, () => '<c><v>1</v></c>').join('');
  const xml = `<worksheet xmlns="${NS}"><sheetData>${Array.from(
    { length: 63 },
    () => `<row>${cells}</row>`
  ).join('')}</sheetData></worksheet>`;

  test('the read yields inside each row', () => {
    const { value, yields } = drain(parseXmlPartSteps('/xl/worksheets/sheet1.xml', xml));
    expect(yields).toBeGreaterThan(6000);
    expect(value).toEqual(parseXmlPart('/xl/worksheets/sheet1.xml', xml));
  });

  test('the serializer yields inside each row', () => {
    const tree = part(xml);
    const { value, yields } = drain(serializeXmlPartSteps(tree));
    expect(yields).toBeGreaterThan(6000);
    expect(value).toBe(serializeXmlPart(tree));
  });
});

describe('the package steps', () => {
  const file = buildXlsx({
    sheets: [{ name: 'Data', rows: '<row r="1"><c r="A1"><v>1</v></c></row>' }],
  });

  test('readXlsxSteps gives the package of readXlsx, and it yields between parts', () => {
    const { value, yields } = drain(readXlsxSteps(file));
    expect(yields).toBeGreaterThan(0);
    expect(value).toEqual(readXlsx(file));
  });

  test('readXlsxSteps refuses a file that readXlsx refuses', () => {
    const bad = new Uint8Array([1, 2, 3]);
    expect(drain(readXlsxSteps(bad)).value).toEqual(readXlsx(bad));
  });

  test('saveXlsxSteps gives the bytes of saveXlsx', () => {
    const read = readXlsx(file);
    if (!read.ok) throw new Error('read');
    expect(drain(saveXlsxSteps(read.pkg)).value).toEqual(saveXlsx(read.pkg));
  });
});
