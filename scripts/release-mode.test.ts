// Adapted from pptx-editor/scripts/release-mode.test.ts at 8e843068 (EigenPal). Differences: no
// canary cases; cases for prerelease mode and for the published package set of this
// repository.
import { afterEach, expect, test } from 'bun:test';
import { mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { pendingChangesets, prereleaseTag, publicPackages, releaseMode } from './release-mode.mjs';

const root = join(import.meta.dir, '..');
const directories: string[] = [];
afterEach(() => {
  for (const directory of directories.splice(0))
    rmSync(directory, { recursive: true, force: true });
});

function fixture(changesets: string[] = [], fixed = ['@scope/core', '@scope/react']): string {
  const directory = mkdtempSync(join(tmpdir(), 'xlsx-release-mode-'));
  directories.push(directory);
  const write = (path: string, value: unknown): void => {
    mkdirSync(dirname(join(directory, path)), { recursive: true });
    writeFileSync(join(directory, path), typeof value === 'string' ? value : JSON.stringify(value));
  };
  write('packages/core/package.json', { name: '@scope/core', version: '1.2.0' });
  write('packages/react/package.json', { name: '@scope/react', version: '1.2.0' });
  write('packages/internal/package.json', {
    name: '@scope/internal',
    version: '1.2.0',
    private: true,
  });
  write('.changeset/README.md', 'Not a changeset.');
  write('.changeset/config.json', { fixed: [fixed] });
  for (const name of changesets) write(`.changeset/${name}`, '---\n"@scope/core": patch\n---\n');
  return directory;
}

test('pending changesets take the release PR path', () => {
  expect(releaseMode(fixture(['fix-a.md']))).toEqual({ mode: 'pending' });
});

test('no pending changesets take the publish path', () => {
  expect(releaseMode(fixture())).toEqual({ mode: 'publish' });
});

test('the README is not a changeset', () => {
  expect(pendingChangesets(fixture())).toEqual([]);
});

// `changeset version` in prerelease mode moves each used changeset into `.changeset/pre/`.
test('a changeset that prerelease mode already used is not pending', () => {
  const directory = fixture();
  mkdirSync(join(directory, '.changeset/pre'));
  writeFileSync(join(directory, '.changeset/pre/fix-a.md'), '---\n"@scope/core": patch\n---\n');
  writeFileSync(join(directory, '.changeset/pre.json'), JSON.stringify({ mode: 'pre', tag: 'b' }));
  expect(releaseMode(directory)).toEqual({ mode: 'publish' });
});

test('the prerelease tag comes from .changeset/pre.json', () => {
  const directory = fixture();
  expect(prereleaseTag(directory)).toBeNull();
  writeFileSync(join(directory, '.changeset/pre.json'), JSON.stringify({ mode: 'pre', tag: 'b' }));
  expect(prereleaseTag(directory)).toBe('b');
  writeFileSync(join(directory, '.changeset/pre.json'), JSON.stringify({ mode: 'exit', tag: 'b' }));
  expect(prereleaseTag(directory)).toBeNull();
});

test('the public packages are the fixed group without private packages', () => {
  const directory = fixture([], ['@scope/core', '@scope/react', '@scope/internal']);
  expect(publicPackages(directory)).toEqual([
    { name: '@scope/core', version: '1.2.0' },
    { name: '@scope/react', version: '1.2.0' },
  ]);
});

test('a fixed group name without a manifest is an error', () => {
  expect(() => publicPackages(fixture([], ['@scope/missing']))).toThrow(/@scope\/missing/);
});

// The published set: every workspace under packages/ that is not private is in the fixed
// group, so one release moves every published package to one version.
test('the fixed group matches the public packages of this repository', () => {
  const manifests = readdirSync(join(root, 'packages'), { withFileTypes: true })
    .filter((entry) => entry.isDirectory())
    .map((entry) => join(root, 'packages', entry.name, 'package.json'))
    .map((path) => JSON.parse(readFileSync(path, 'utf8')) as { name: string; private?: boolean });
  const published = manifests.filter((manifest) => !manifest.private).map(({ name }) => name);
  expect(
    publicPackages(root)
      .map(({ name }) => name)
      .sort()
  ).toEqual(published.sort());
});
