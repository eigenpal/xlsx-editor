import { describe, expect, test } from 'bun:test';
import {
  cpSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  symlinkSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import {
  orphanReports,
  packageArgumentProblem,
  reportsBelow,
  unlistedEntries,
} from '../lib/api-entries.mjs';
import { PACKAGES } from '../lib-packages.mjs';

const repoRoot = resolve(import.meta.dir, '../..');

describe('unlistedEntries', () => {
  test('names a code subpath that has no entry', () => {
    const exportsMap = { '.': {}, './en': {}, './de': {} };
    expect(unlistedEntries({ entries: ['index', 'en'] }, exportsMap)).toEqual(['./de']);
  });

  test('names the root entry when the list leaves it out', () => {
    expect(unlistedEntries({ entries: [] }, { '.': {} })).toEqual(['.']);
  });

  test('skips stylesheets, the text shaper binary, and the manifest', () => {
    const exportsMap = {
      './styles.css': {},
      './styles/editor.css': {},
      './harfbuzz.wasm': {},
      './package.json': {},
    };
    expect(unlistedEntries({ entries: [] }, exportsMap)).toEqual([]);
  });

  test.each(PACKAGES.map((pkg) => [pkg.name, pkg] as const))(
    '%s lists each code subpath of its exports',
    (_name, pkg) => {
      const manifest = JSON.parse(
        readFileSync(resolve(repoRoot, pkg.root, 'package.json'), 'utf8')
      );
      expect(unlistedEntries(pkg, manifest.exports)).toEqual([]);
    }
  );
});

describe('orphanReports', () => {
  const packages = [{ slug: 'i18n', entries: ['index', 'en'] }];

  test('passes the report of each entry', () => {
    const reports = new Map([['i18n', ['en.api.md', 'index.api.md']]]);
    expect(orphanReports(packages, reports)).toEqual([]);
  });

  test('names a report with no entry and a folder with no package', () => {
    const reports = new Map([
      ['i18n', ['en.api.md', 'ghost.api.md']],
      ['gone', ['index.api.md']],
    ]);
    expect(orphanReports(packages, reports)).toEqual(['i18n/ghost.api.md', 'gone/index.api.md']);
  });
});

describe('orphanReports in the root of docs/api', () => {
  test('names a file directly in docs/api', () => {
    const packages = [{ slug: 'core', entries: ['index'] }];
    const reports = new Map([
      ['', ['stray.api.md']],
      ['core', ['index.api.md']],
    ]);
    expect(orphanReports(packages, reports)).toEqual(['stray.api.md']);
  });
});

describe('reportsBelow', () => {
  test('lists the files directly in docs/api under the empty key', () => {
    const apiRoot = mkdtempSync(join(tmpdir(), 'api-reports-'));
    try {
      mkdirSync(join(apiRoot, 'core'));
      writeFileSync(join(apiRoot, 'core/index.api.md'), '');
      writeFileSync(join(apiRoot, 'stray.api.md'), '');
      const packages = [{ slug: 'core', entries: ['index'] }];
      expect(orphanReports(packages, reportsBelow(apiRoot))).toEqual(['stray.api.md']);
    } finally {
      rmSync(apiRoot, { recursive: true, force: true });
    }
  });
});

describe('api-extractor with no build', () => {
  test('a check and a record run exit with 1, name only the build, and keep the baseline', () => {
    const copy = mkdtempSync(join(tmpdir(), 'api-record-'));
    try {
      mkdirSync(join(copy, 'docs/api'), { recursive: true });
      cpSync(join(repoRoot, 'scripts/lib'), join(copy, 'scripts/lib'), { recursive: true });
      for (const file of [
        'scripts/api-extractor.mjs',
        'scripts/lib-packages.mjs',
        'scripts/api-forgotten-exports.baseline.json',
        ...PACKAGES.map((pkg) => `${pkg.root}/package.json`),
      ])
        cpSync(join(repoRoot, file), join(copy, file));
      symlinkSync(join(repoRoot, 'node_modules'), join(copy, 'node_modules'));
      const baseline = join(copy, 'scripts/api-forgotten-exports.baseline.json');
      const before = readFileSync(baseline, 'utf8');
      const script = join(copy, 'scripts/api-extractor.mjs');
      for (const args of [[], ['--local', '--record-forgotten-exports']]) {
        const result = Bun.spawnSync(['node', script, ...args]);
        expect(result.exitCode).toBe(1);
        const lines = result.stderr.toString().trim().split('\n');
        expect(lines).toHaveLength(1);
        expect(lines[0]).toContain('Run bun run build:packages first.');
        expect(readFileSync(baseline, 'utf8')).toBe(before);
      }
    } finally {
      rmSync(copy, { recursive: true, force: true });
    }
  });
});

describe('packageArgumentProblem', () => {
  const packages = [{ name: '@a/one' }, { name: '@a/two' }];

  test('accepts a listed package and no flag', () => {
    expect(packageArgumentProblem(['--package', '@a/two'], packages)).toBeNull();
    expect(packageArgumentProblem(['--local'], packages)).toBeNull();
  });

  test('names the valid values for an unknown or missing value', () => {
    expect(packageArgumentProblem(['--package', '@a/typo'], packages)).toBe(
      "Unknown package '@a/typo'. Valid values: @a/one, @a/two."
    );
    expect(packageArgumentProblem(['--package', '--local'], packages)).toContain(
      'requires a value'
    );
  });

  test('the script exits with a non-zero code and names the valid values', () => {
    const script = resolve(repoRoot, 'scripts/api-extractor.mjs');
    const result = Bun.spawnSync(['node', script, '--package', '@xlsx-editor.dev/typo']);
    expect(result.exitCode).toBe(1);
    expect(result.stderr.toString()).toContain('Valid values: @xlsx-editor.dev/i18n');
  });
});
