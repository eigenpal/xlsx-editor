// A missing or stale reference is a gap, and a gap fails the gate (design D8): every fixture
// has a current reference PDF.

import { describe, expect, test } from 'bun:test';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { FIXTURES, fixtureFile } from '../fixtures/catalog.ts';
import { workbookFamilies } from './fonts.ts';
import {
  invalidField,
  pdfPages,
  readManifest,
  referenceProblems,
  type Manifest,
} from './references.ts';

describe('reference PDFs', () => {
  test('every fixture has a current one-page Letter reference', () => {
    expect(FIXTURES.flatMap((f) => referenceProblems(f.name))).toEqual([]);
  });

  test('no record names a fixture that does not exist', () => {
    const names = new Set(FIXTURES.map((f) => f.name));
    expect(Object.keys(readManifest().references).filter((n) => !names.has(n))).toEqual([]);
  });

  test('a missing or stale record is reported, never passed', () => {
    const empty: Manifest = { version: 1, references: {} };
    expect(referenceProblems('basic-values', empty)).toEqual([
      'basic-values: no reference in fixtures/references.json',
    ]);
    const real = readManifest().references['basic-values']!;
    const stale: Manifest = {
      version: 1,
      references: { 'basic-values': { ...real, workbook: '0' } },
    };
    expect(referenceProblems('basic-values', stale)[0]).toContain('stale');
  });

  test('a record names its missing workbook or PDF with the next step', () => {
    const real = readManifest().references['basic-values']!;
    const ghost: Manifest = { version: 1, references: { 'no-such-fixture': real } };
    expect(referenceProblems('no-such-fixture', ghost)).toEqual([
      'no-such-fixture: fixtures/no-such-fixture.xlsx is missing: build it with `bun run fixtures:build`.',
      'no-such-fixture: fixtures/no-such-fixture.pdf is missing: a maintainer makes it (CONTRIBUTING.md, "Reference PDFs").',
    ]);
  });

  test('a manifest record with a field of another type is refused, by fixture and field', () => {
    const real = readManifest().references['basic-values']!;
    expect(invalidField(real)).toBeUndefined();
    expect(invalidField({ ...real, pages: '1' })).toBe('pages');
    // A page count is a positive integer.
    for (const pages of [0, -1, 1.5, Number.NaN])
      expect(invalidField({ ...real, pages })).toBe('pages');
    expect(invalidField({ ...real, pageSizePt: [612] })).toBe('pageSizePt');
    expect(invalidField({ ...real, locale: { ...real.locale, decimal: 1 } })).toBe(
      'locale.decimal'
    );
    expect(invalidField({ ...real, locale: { decimal: '.' } })).toBe('locale.group');
    expect(invalidField({ ...real, locale: { ...real.locale, builtins: { 14: 1 } } })).toBe(
      'locale.builtins'
    );
    // Built-in number formats are keyed by their id, never a list.
    expect(invalidField({ ...real, locale: { ...real.locale, builtins: ['m/d/yyyy'] } })).toBe(
      'locale.builtins'
    );
    const dir = mkdtempSync(join(tmpdir(), 'xlsx-manifest-'));
    try {
      const file = join(dir, 'references.json');
      const bad = { ...real, locale: { decimal: '.', group: null } };
      writeFileSync(file, JSON.stringify({ version: 1, references: { 'basic-values': bad } }));
      expect(() => readManifest(file)).toThrow(
        `${file}: the record of basic-values has no valid locale.group`
      );
      // A JSON syntax error names the manifest too.
      writeFileSync(file, '{ "version": 1,');
      expect(() => readManifest(file)).toThrow(`${file}: `);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  test('page facts come from the PDF bytes', () => {
    const pdf = new TextEncoder().encode(
      '<< /Type /Pages >> << /Type /Page /MediaBox [0 0 612 792] >>'
    );
    expect(pdfPages(pdf)).toEqual({ pages: 1, size: [612, 792] });
  });
});

/** The open faces that the harness draws (`@docx-editor.dev/fonts`). */
const OPEN_FACES = /^(?:[A-Z]{6}\+)?(?:Carlito|Liberation(?:Sans|Serif|Mono))(?:-[A-Za-z]+)?$/;

describe('the faces of the reference PDFs', () => {
  // A reference PDF that embeds another face measures the outlines of that face, not the
  // engine. The font dictionaries of these PDFs are plain objects, never in an object stream,
  // so each `/Type /Font` dictionary has a `/BaseFont` name that this test can read.
  test('every reference PDF embeds only open faces that its fixture names', () => {
    const problems: string[] = [];
    for (const { name } of FIXTURES) {
      const pdf = readFileSync(
        join(import.meta.dir, '..', '..', 'fixtures', `${name}.pdf`),
        'latin1'
      );
      if (pdf.includes('/ObjStm')) problems.push(`${name}: an object stream can hide a font`);
      const fonts = pdf.match(/\/Type\s*\/Font\b/g)?.length ?? 0;
      const faces = [...pdf.matchAll(/\/BaseFont\s*\/([^\s/<>[\]()]+)/g)].map((m) => m[1]!);
      if (faces.length !== fonts) problems.push(`${name}: ${fonts} fonts, ${faces.length} names`);
      // A face is embedded when its descriptor holds a font program.
      const descriptors = pdf.match(/\/Type\s*\/FontDescriptor\b/g)?.length ?? 0;
      const programs = pdf.match(/\/FontFile[23]?\b/g)?.length ?? 0;
      if (programs !== descriptors)
        problems.push(`${name}: ${descriptors} font descriptors, ${programs} embedded programs`);
      // Each face is an open face, and a family that the fixture names: `LiberationSans-Bold`
      // is the family `liberation sans`.
      const xlsx = readFileSync(join(import.meta.dir, '..', '..', 'fixtures', fixtureFile(name)));
      const named = workbookFamilies(new Uint8Array(xlsx));
      for (const face of faces) {
        if (!OPEN_FACES.test(face)) problems.push(`${name}: embeds ${face}`);
        const family = face
          .replace(/^[A-Z]{6}\+/, '')
          .replace(/-[A-Za-z]+$/, '')
          .replace(/^Liberation(?=[A-Z])/, 'Liberation ')
          .toLowerCase();
        if (!named.has(family)) problems.push(`${name}: embeds ${face}, which it does not name`);
      }
    }
    expect(problems).toEqual([]);
  });
});
