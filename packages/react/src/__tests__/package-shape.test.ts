// The engine resolves to one copy (AGENTS.md, "Packages"): core is a peer of the adapter, never a
// regular dependency, and the adapter stylesheet only imports the core stylesheet.

import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';

const root = join(import.meta.dir, '..', '..');
const manifest = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8')) as {
  dependencies?: Record<string, string>;
  peerDependencies?: Record<string, string>;
};

test('core is a peer dependency, never a regular one', () => {
  expect(manifest.peerDependencies?.['@xlsx-editor.dev/core']).toBeDefined();
  expect(manifest.dependencies?.['@xlsx-editor.dev/core']).toBeUndefined();
});

test('the core peer range accepts only the minor of core that the adapter releases with', () => {
  const core = JSON.parse(readFileSync(join(root, '..', 'core', 'package.json'), 'utf8')) as {
    version: string;
  };
  const [major, minor] = core.version.split('.');
  const range = manifest.peerDependencies?.['@xlsx-editor.dev/core'] ?? '';
  // A prerelease of that minor (`~0.1.0-beta.0`) is the same minor: Changesets pre mode writes
  // the prerelease version of core into the range.
  expect(range).toMatch(new RegExp(`^~${major}\\.${minor}\\.\\d+(-[0-9A-Za-z.]+)?$`));
});

test('the adapter stylesheet holds no rules of its own', () => {
  const css = readFileSync(join(root, 'src/styles/editor.css'), 'utf8')
    .replace(/\/\*[\s\S]*?\*\//g, '')
    .trim();
  expect(css).toBe("@import '@xlsx-editor.dev/core/styles/editor.css';");
});

test('each export list of the entry is in alphabetical order', () => {
  const entry = readFileSync(join(root, 'src/index.ts'), 'utf8');
  for (const [, body] of entry.matchAll(/export (?:type )?\{([^}]*)\}/g)) {
    const names = body!
      .split(',')
      .map((n) => n.trim().replace(/^type /, ''))
      .filter(Boolean);
    expect(names).toEqual(
      [...names].sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()))
    );
  }
});

test('the entry gives the clipboard type and the resize limits of the editor', async () => {
  const entry = await import('../index.ts');
  const core = await import('@xlsx-editor.dev/core/editor');
  expect(entry.CLIPBOARD_TYPE).toBe(core.CLIPBOARD_TYPE);
  expect(entry.MAX_COLUMN_WIDTH).toBe(core.MAX_COLUMN_WIDTH);
  expect(entry.MAX_ROW_HEIGHT).toBe(core.MAX_ROW_HEIGHT);
  expect(entry.MAX_RESIZE_NEW_ROWS).toBe(core.MAX_RESIZE_NEW_ROWS);
});

test('the entry gives the state of the sheet commands, the one source of the sheet tabs', async () => {
  const entry = await import('../index.ts');
  const core = await import('@xlsx-editor.dev/core/editor');
  expect(entry.sheetCommandState).toBe(core.sheetCommandState);
});

test('the entry gives the number format functions of the engine', async () => {
  const entry = await import('../index.ts');
  const core = await import('@xlsx-editor.dev/core/editor');
  for (const name of [
    'MAX_LISTED_NUMBER_FORMATS',
    'NUMBER_FORMAT_CODE_PROBLEM_KEYS',
    'NUMBER_FORMAT_SAMPLE',
    'numberFormatDialogApplyStep',
    'numberFormatCodeProblem',
    'numberFormatDialogCurrent',
    'numberFormatDialogStart',
    'previewNumberFormat',
    'refusalText',
    'trapTabWithin',
    'workbookNumberFormatCodes',
  ] as const)
    expect(entry[name]).toBe(core[name]);
});
