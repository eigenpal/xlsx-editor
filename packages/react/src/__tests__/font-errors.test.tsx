// The root reports each font failure once: a failed open for its fonts, and each set of
// packaged substitute faces that did not all load, as PPTX Editor's `reportFontError` does.
// Without a handler, a failure prints one warning, so it is never silent.

import { afterEach, describe, expect, test } from 'bun:test';
import { cleanup } from '@testing-library/react';
import {
  EMPTY_SNAPSHOT,
  reportFontError,
  type EditorFailure,
  type EditorSnapshot,
  type FontReport,
} from '@xlsx-editor.dev/core/editor';
import { watchFontErrors } from '@xlsx-editor.dev/core/editor';

afterEach(cleanup);

function fakeEditor() {
  const listeners = new Set<() => void>();
  let snapshot: EditorSnapshot = EMPTY_SNAPSHOT;
  let report: FontReport = { faces: [], substituteFailures: [] };
  return {
    subscribe: (f: () => void) => (listeners.add(f), () => void listeners.delete(f)),
    snapshot: () => snapshot,
    fontReport: () => report,
    set(next: Partial<EditorSnapshot>, fonts?: FontReport) {
      snapshot = { ...snapshot, ...next };
      if (fonts) report = fonts;
      for (const f of listeners) f();
    },
  };
}

describe('font errors', () => {
  test('substitute faces that did not load are reported once for each set of faces', () => {
    const editor = fakeEditor();
    const seen: EditorFailure[] = [];
    const off = watchFontErrors(editor, (f) => seen.push(f));
    const failures = [{ family: 'Cambria', file: 'Caladea-Regular.ttf', diagnostic: 'size' }];
    const fonts = {
      faces: [{ family: 'Carlito', origin: 'substitute' }],
      substituteFailures: failures,
    };
    editor.set({ status: 'ready' }, fonts);
    editor.set({ zoom: 2 });
    expect(seen).toHaveLength(1);
    expect(seen[0]!.code).toBe('fonts');
    expect(seen[0]!.substituteFailures).toEqual(failures);
    expect(seen[0]!.detail).toContain('Caladea-Regular.ttf');
    // A failed open for its fonts reports once too.
    const failure: EditorFailure = { code: 'fonts', detail: 'no font sources' };
    editor.set({ status: 'failed', failure });
    editor.set({ zoom: 1 });
    expect(seen).toHaveLength(2);
    expect(seen[1]).toBe(failure);
    off();
  });

  test('without a handler, a font failure prints one warning', () => {
    const warnings: unknown[] = [];
    const original = console.warn;
    console.warn = (...a: unknown[]) => void warnings.push(a[0]);
    try {
      reportFontError({ code: 'fonts', detail: 'substitute faces did not load: X.ttf' }, undefined);
      const seen: EditorFailure[] = [];
      reportFontError({ code: 'fonts' }, (f) => seen.push(f));
      expect(seen).toHaveLength(1);
    } finally {
      console.warn = original;
    }
    expect(warnings).toEqual(['[xlsx-editor] fonts: substitute faces did not load: X.ttf']);
  });
});
