// Copied from pptx-editor/scripts/harfbuzz-notices.test.ts at d4502552 (EigenPal). Differences:
// the Universal Shaping Engine notice is `HarfBuzz-USE-MIT.txt`, and the first test checks its
// MIT permission notice instead of quoting its copyright holder (AGENTS.md, "Naming"); the error
// tests also match the fix that each error names, and a test covers a missing binary.
import { afterEach, expect, test } from 'bun:test';
import { cpSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import {
  HARFBUZZ_MANIFEST,
  HARFBUZZ_NOTICE,
  validateHarfBuzzNotices,
} from './harfbuzz-notices.mjs';

const core = resolve(import.meta.dir, '../packages/core');
const temporary: string[] = [];

afterEach(() => {
  for (const directory of temporary.splice(0)) rmSync(directory, { recursive: true, force: true });
});

function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'harfbuzz-notices-'));
  temporary.push(root);
  const coreRoot = join(root, 'core');
  const dependencyRoot = join(root, 'dependency');
  mkdirSync(join(coreRoot, 'dist'), { recursive: true });
  mkdirSync(join(dependencyRoot, 'dist'), { recursive: true });
  cpSync(join(core, 'licenses'), join(coreRoot, 'licenses'), { recursive: true });
  cpSync(join(core, 'package.json'), join(coreRoot, 'package.json'));
  cpSync(join(core, 'node_modules/harfbuzzjs/package.json'), join(dependencyRoot, 'package.json'));
  for (const destination of [coreRoot, dependencyRoot]) {
    cpSync(
      join(core, 'node_modules/harfbuzzjs/dist/harfbuzz.wasm'),
      join(destination, 'dist/harfbuzz.wasm')
    );
  }
  return { coreRoot, dependencyRoot };
}

test('the installed binary has matching component notices', () => {
  expect(() => validateHarfBuzzNotices(fixture())).not.toThrow();
  const notices = HARFBUZZ_MANIFEST.notices
    .map((entry: { file: string }) => readFileSync(join(core, 'licenses', entry.file), 'utf8'))
    .join('\n');
  for (const copyright of ['Grigori Goronzy', 'Emscripten authors', 'Rich Felker', 'Unicode, Inc.'])
    expect(notices).toContain(copyright);
  const useData = readFileSync(join(core, 'licenses/HarfBuzz-USE-MIT.txt'), 'utf8');
  expect(useData).toMatch(/MIT License\s+Copyright \(c\) \S+/);
  expect(useData).toContain('The above copyright notice and this permission notice');
  for (const entry of HARFBUZZ_MANIFEST.notices) {
    expect(HARFBUZZ_NOTICE).toContain(`licenses/${entry.file}`);
  }
});

test('a dependency upgrade requires a new notice inventory', () => {
  const paths = fixture();
  const file = join(paths.dependencyRoot, 'package.json');
  const dependency = JSON.parse(readFileSync(file, 'utf8'));
  dependency.version = '99.0.0';
  writeFileSync(file, JSON.stringify(dependency));
  expect(() => validateHarfBuzzNotices(paths)).toThrow(
    /HarfBuzz version changed: installed 99\.0\.0.*bun install.*upgrade steps in packages\/core\/licenses\/README\.md/
  );
});

test('a replaced release binary cannot reuse the old notice inventory', () => {
  const paths = fixture();
  writeFileSync(join(paths.coreRoot, 'dist/harfbuzz.wasm'), 'different binary');
  expect(() => validateHarfBuzzNotices(paths)).toThrow(
    /HarfBuzz binary differs.*bun run build:packages.*upgrade steps/
  );
});

test('an altered permission notice fails the release check', () => {
  const paths = fixture();
  writeFileSync(join(paths.coreRoot, 'licenses/HarfBuzz-USE-MIT.txt'), 'MIT');
  expect(() => validateHarfBuzzNotices(paths)).toThrow(
    /HarfBuzz notice differs from the source manifest: \S*licenses\/HarfBuzz-USE-MIT\.txt\. Restore the file from git.*upgrade steps/
  );
});

test('a missing component notice fails the release check', () => {
  const paths = fixture();
  rmSync(join(paths.coreRoot, 'licenses/HarfBuzz-source-notices.txt'));
  expect(() => validateHarfBuzzNotices(paths)).toThrow(
    /HarfBuzz notice is missing: \S*licenses\/HarfBuzz-source-notices\.txt\. Restore the file from git/
  );
});

test('a missing release binary names the rebuild', () => {
  const paths = fixture();
  rmSync(join(paths.coreRoot, 'dist/harfbuzz.wasm'));
  expect(() => validateHarfBuzzNotices(paths)).toThrow(
    /HarfBuzz binary is missing: .*bun run build:packages/
  );
});
