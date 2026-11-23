// Copied from docx-editor/scripts/__tests__/build-core-declarations.test.ts at 490a4c2d (EigenPal).
// Differences: the package scope is `@xlsx-editor.dev`; the sibling and path tests name this
// repository's packages (react, core, i18n); and one test pins the inlined private package.
import { expect, test } from 'bun:test';
import { fileURLToPath } from 'node:url';
import {
  declarationCandidates,
  packageName,
  INLINED,
  publishedEntries,
} from '../build-core-declarations.mjs';
import { siblingsOf } from '../check-built-siblings.mjs';
import { declarationCompilerOptions } from '../declaration-options.mjs';

const paths = {
  '@xlsx-editor.dev/core': ['./src/index.ts'],
  '@xlsx-editor.dev/core/store': ['./src/store/index.ts'],
};

test('each export with a top-level types condition becomes one declaration entry', () => {
  const entries = publishedEntries(
    {
      exports: {
        '.': { types: './dist/index.d.ts', import: './dist/index.mjs' },
        './store': { types: './dist/store.d.ts', import: './dist/store.mjs' },
        './editor.css': './dist/editor.css',
        './package.json': './package.json',
      },
    },
    paths
  );
  expect(entries.map(({ specifier, name }) => [specifier, name])).toEqual([
    ['@xlsx-editor.dev/core', 'index'],
    ['@xlsx-editor.dev/core/store', 'store'],
  ]);
  expect(entries[1]!.source.endsWith('/packages/core/src/store/index.ts')).toBe(true);
});

test('a subpath without a source path or with nested types fails instead of shipping no types', () => {
  expect(() =>
    publishedEntries({ exports: { './layout': { types: './dist/layout.d.ts' } } }, paths)
  ).toThrow('has no path for its source');
  expect(() =>
    publishedEntries(
      { exports: { './store': { import: { types: './dist/store.d.ts', default: './x.mjs' } } } },
      paths
    )
  ).toThrow('nests its types condition');
});

test('a bare import is checked against the package that owns it', () => {
  expect(packageName('prosemirror-model')).toBe('prosemirror-model');
  expect(packageName('fflate/browser')).toBe('fflate');
  expect(packageName('@xlsx-editor.dev/i18n/locales')).toBe('@xlsx-editor.dev/i18n');
});

test('an import keeps its module kind when it resolves to a declaration file', () => {
  expect(declarationCandidates('/o/foo.ts')).toEqual(['/o/foo.d.ts']);
  expect(declarationCandidates('/o/foo.js')).toEqual(['/o/foo.d.ts']);
  expect(declarationCandidates('/o/foo.mts')).toEqual(['/o/foo.d.mts']);
  expect(declarationCandidates('/o/foo.cjs')).toEqual(['/o/foo.d.cts']);
  expect(declarationCandidates('/o/foo')).toEqual(['/o/foo.d.ts', '/o/foo/index.d.ts']);
});

test('declaration builds drop sibling source paths and keep local aliases', () => {
  const react = declarationCompilerOptions(
    new URL('../../packages/react/tsup.config.ts', import.meta.url)
  );
  // React reads core and i18n from dist; the private OOXML layer keeps its source path.
  expect(Object.keys(react.paths)).toEqual(['@eigenpal/ooxml-core']);
  expect(react.pathsBase).toBe(fileURLToPath(new URL('../../packages/react', import.meta.url)));
});

test('a package checks every sibling it reads, including those behind core', () => {
  const siblings = (name: string) =>
    [...siblingsOf(new URL(`../../packages/${name}`, import.meta.url).pathname).keys()].sort();
  expect(siblings('react')).toEqual(['@xlsx-editor.dev/core', '@xlsx-editor.dev/i18n']);
  expect(siblings('core')).toEqual(['@xlsx-editor.dev/i18n']);
  expect(siblings('i18n')).toEqual([]);
});

test('core inlines the private OOXML layer, and its tsconfig names that source', () => {
  expect(INLINED).toEqual(['@eigenpal/ooxml-core']);
});
