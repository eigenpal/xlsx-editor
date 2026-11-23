// XLSX Editor copy of pptx-editor/scripts/parity/parity.test.ts at d4502552 (EigenPal).
// Intended differences: the package and component names.
//
// The DOM-free parity checks. The contract check mounts Vue components, so its tests run in the
// DOM test run: packages/vue/test/parity.test.ts.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { composableErrors } from '../check-composable-parity.mjs';
import { cssThinErrors } from '../check-adapter-css-thin.mjs';
import { mirrorErrors } from '../check-adapter-mirror.mjs';
import { vueRuntimePropTypes } from './api-snapshot.mjs';

const root = join(import.meta.dir, '../..');
const read = (path: string): string => readFileSync(join(root, path), 'utf8');

const inputs = () => ({
  reactApi: read('docs/api/react/index.api.md'),
  vueApi: read('docs/api/vue/index.api.md'),
  contract: JSON.parse(read('scripts/parity/parity.contract.json')),
});

describe('check:composable-parity', () => {
  test('the committed Vue composables pass', () => {
    expect(composableErrors(inputs())).toEqual([]);
  });

  test('a composable without a named return type fails', () => {
    const input = inputs();
    input.vueApi = input.vueApi.replace(
      'export function useXlsxEditor(): UseXlsxEditorReturn;',
      'export function useXlsxEditor(): ShallowRef<XlsxEditorInstance | null>;'
    );
    expect(composableErrors(input)).toContain(
      'useXlsxEditor must declare its return type as UseXlsxEditorReturn (found ShallowRef).'
    );
  });

  test('a Vue composable with no React hook fails unless it is Vue-only', () => {
    const input = inputs();
    input.vueApi +=
      '\nexport function useExtra(): UseExtraReturn;\nexport type UseExtraReturn = number;\n';
    expect(composableErrors(input)).toContain(
      'useExtra has no React hook of the same name and is not vueExclusive.'
    );
  });
});

describe('check:adapter-css-thin', () => {
  test('an import-only stylesheet passes', () => {
    expect(
      cssThinErrors({ 'a.css': "/* note */\n@import '@xlsx-editor.dev/core/styles/editor.css';\n" })
    ).toEqual([]);
  });

  test('an adapter-local rule fails', () => {
    const errors = cssThinErrors({
      'a.css': "@import '@xlsx-editor.dev/core/styles/editor.css';\n.x { color: red; }",
    });
    expect(errors.length).toBe(1);
    expect(errors[0]).toContain('.x { color: red; }');
  });
});

describe('check:adapter-mirror', () => {
  const none = {};
  test('twins at the mirrored path pass, .ts and .tsx alike', () => {
    expect(
      mirrorErrors({
        reactFiles: ['a.tsx', 'b.ts'],
        vueFiles: ['a.tsx', 'b.tsx'],
        vueOnly: none,
        reactOnly: none,
      })
    ).toEqual([]);
  });

  test('a file with no twin fails unless it is listed with a reason', () => {
    const input = { reactFiles: ['a.tsx'], vueFiles: ['a.tsx', 'extra.ts'], reactOnly: none };
    expect(mirrorErrors({ ...input, vueOnly: none })).toEqual([
      'Vue file extra.ts has no React twin and no listed reason.',
    ]);
    expect(mirrorErrors({ ...input, vueOnly: { 'extra.ts': 'Vue-only glue.' } })).toEqual([]);
  });

  test('a listed file that is gone or has a twin is stale', () => {
    expect(
      mirrorErrors({
        reactFiles: ['a.tsx'],
        vueFiles: ['a.ts'],
        vueOnly: { 'a.ts': 'Was Vue-only.', 'gone.ts': 'Removed.' },
        reactOnly: none,
      })
    ).toEqual([
      'Vue-only entry a.ts now has a React twin (stale).',
      'Vue-only entry gone.ts names no Vue file (stale).',
    ]);
  });
});

describe('vueRuntimePropTypes', () => {
  test('reads the props object of a component that a namespace type names', () => {
    const snapshot = [
      'export const Editor: EditorNamespace;',
      '',
      'export type EditorProps = ExtractPublicPropTypes<typeof editorProps_2>;',
      '',
      'const editorProps_2: {',
      '    readonly mode: {',
      '        readonly type: PropType<EditorMode>;',
      '        readonly default: undefined;',
      '    };',
      '    readonly opening: {',
      '        readonly type: BooleanConstructor;',
      '        readonly default: undefined;',
      '    };',
      '};',
    ].join('\n');
    expect([...(vueRuntimePropTypes(snapshot, 'Editor') ?? [])]).toEqual([
      ['mode', 'EditorMode'],
      ['opening', 'boolean'],
    ]);
    expect(vueRuntimePropTypes('export const Other: OtherNamespace;', 'Other')).toBeNull();
  });

  test('reads the paired prop types of the composed Vue editor', () => {
    const types = vueRuntimePropTypes(inputs().vueApi, 'XlsxEditor');
    expect(types?.get('opening')).toBe('boolean');
    expect(types?.get('mode')).toBe('EditorMode');
  });
});
