import { expect, test } from 'bun:test';
import { resolveSpecifier } from './changed.mjs';
import { DOM_PRELOAD, bunArgumentsFor } from './xlsx-suite.mjs';

test('only the adapter tests and the demo flow tests preload a DOM', () => {
  expect(bunArgumentsFor('packages/react/src/__tests__/adapter.test.tsx')).toEqual([
    '--preload',
    DOM_PRELOAD,
  ]);
  expect(bunArgumentsFor('packages/vue/test/adapter.test.ts')).toEqual(['--preload', DOM_PRELOAD]);
  expect(bunArgumentsFor('examples/vite/src/__tests__/demo-save.test.tsx')).toEqual([
    '--preload',
    DOM_PRELOAD,
  ]);
  for (const file of [
    'examples/vite/src/save.test.ts',
    'packages/core/src/layout/__tests__/a.test.ts',
    'packages/ooxml-core/src/__tests__/zip.test.ts',
    'packages/editor-api/src/__tests__/a.test.ts',
    'scripts/test/xlsx-suite.test.mjs',
  ]) {
    expect(bunArgumentsFor(file)).toEqual([]);
  }
});

// This file runs in the DOM-free group, as every engine test does. A global `document` here
// means the runner preloaded happy-dom for a file that must not have it.
test('a DOM-free file runs with no global document', () => {
  expect(typeof document).toBe('undefined');
  expect(typeof window).toBe('undefined');
});

// Add the `@xlsx-editor.dev/core` subpath case of PPTX Editor's test when `packages/core` exists.
test('the shared package scope resolves to its package', () => {
  expect(resolveSpecifier('packages/core/src/a.ts', '@eigenpal/ooxml-core')).toBe(
    'packages/ooxml-core/src/index.ts'
  );
});
