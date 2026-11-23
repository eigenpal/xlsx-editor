// The fonts of one editor: a source change while the first open loads its faces wins over the
// load of the old source, and the packaged substitute faces that did not load are reported
// (`fontReport`), as PPTX Editor reports them. The editor loads only the packaged substitute
// families that layout draws with.

import { ALL_DEFAULT_FONT_FAMILIES, loadDefaultFonts } from '@docx-editor.dev/fonts';
import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { EditorFonts } from '../editor-fonts.ts';
import { createXlsxEditor } from '../index.ts';
import { loadSubstituteFonts, SUBSTITUTE_FAMILIES } from '../substitute-fonts.ts';
import { FONT_SUBSTITUTES, type FontSource } from '../../layout/index.ts';

const fontsDir = join(
  dirname(Bun.resolveSync('@docx-editor.dev/fonts/package.json', import.meta.dir)),
  'assets'
);
const bytes = new Uint8Array(readFileSync(join(fontsDir, 'Carlito-Regular.ttf')));
const face = (family: string): FontSource => ({
  family,
  bold: false,
  italic: false,
  bytes,
  origin: 'substitute',
});
const families = (fonts: EditorFonts) => fonts.measurer?.faces.map((f) => f.family);

describe('editor fonts', () => {
  test('a fonts change while the first open loads replaces the old faces', async () => {
    const slow = () =>
      new Promise<readonly FontSource[]>((r) => setTimeout(() => r([face('Old')]), 30));
    const editor = createXlsxEditor({ fonts: slow });
    const opening = editor.load(buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] }));
    expect(await editor.setFonts([face('New')])).toBeNull();
    expect((await opening).status).toBe('ready');
    expect(editor.fontReport().faces.map((f) => f.family)).toEqual(['New']);
    editor.destroy();
  });

  test('a load of a replaced source that fails does not fail the new one', async () => {
    let fail: (e: Error) => void = () => {};
    const fonts = new EditorFonts(() => new Promise((_, reject) => (fail = reject)));
    const ensure = fonts.ensure();
    await fonts.replace([face('New')]);
    fail(new Error('old source'));
    await ensure;
    expect(families(fonts)).toEqual(['New']);
  });

  test('packaged substitute faces that did not load are reported, and the rest load', async () => {
    const failures = [{ family: 'Cambria', file: 'Caladea-Regular.ttf', diagnostic: 'size' }];
    const fonts = new EditorFonts(undefined, async () => ({
      sources: [face('Carlito')],
      failures,
    }));
    expect(fonts.report.substituteFailures).toEqual([]);
    await fonts.ensure();
    expect(families(fonts)).toEqual(['Carlito']);
    expect(fonts.report.substituteFailures).toEqual(failures);
    expect(fonts.report.faces).toEqual([{ family: 'Carlito', origin: 'substitute' }]);
    // Host faces have no packaged substitute failures.
    await fonts.replace([face('Host')]);
    expect(fonts.report.substituteFailures).toEqual([]);
  });
});

describe('packaged substitute families', () => {
  test('the packaged families whose substitute layout draws with are the substitute families', async () => {
    // The stub fetch fails every face, so the package names each substitution and loads nothing.
    const packaged = await loadDefaultFonts({
      families: ALL_DEFAULT_FONT_FAMILIES,
      fetcher: Object.assign(async () => new Response(null, { status: 404 }), {
        preconnect: fetch.preconnect,
      }),
    });
    const drawn = new Set<string>();
    for (const s of packaged.substitutions) {
      const names = FONT_SUBSTITUTES[s.from.family.toLowerCase()] ?? [];
      if (names.includes(s.to.family)) drawn.add(s.from.family);
    }
    expect([...SUBSTITUTE_FAMILIES].sort() as string[]).toEqual([...drawn].sort());
    expect(SUBSTITUTE_FAMILIES as readonly string[]).not.toContain('Segoe UI');
  });

  test('loadSubstituteFonts fetches the faces of the substitute families only', async () => {
    // The package reads the global fetch at call time; the stub fails each face and records it.
    const fetched: string[] = [];
    const real = globalThis.fetch;
    globalThis.fetch = Object.assign(
      async (input: string | URL | Request) => {
        fetched.push(String(input instanceof Request ? input.url : input));
        return new Response(null, { status: 404 });
      },
      { preconnect: real.preconnect }
    );
    try {
      const { sources, failures } = await loadSubstituteFonts();
      expect(sources).toEqual([]);
      expect([...new Set(failures.map((f) => f.family))].sort()).toEqual(
        [...SUBSTITUTE_FAMILIES].sort()
      );
      expect(fetched.length).toBe(failures.length);
      expect(fetched.some((url) => url.includes('Selawik'))).toBe(false);
    } finally {
      globalThis.fetch = real;
    }
  });
});
