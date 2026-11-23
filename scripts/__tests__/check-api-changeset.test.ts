// `scripts/check-api-changeset.mjs` passes an API change with no changeset while every published
// package is at 0.0.0 (its header records this difference from PPTX Editor). These tests run the
// script in a temporary git repository, so they hold both sides of that rule.
import { afterEach, describe, expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { copyFileSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { PACKAGES } from '../lib-packages.mjs';

const ROOT = resolve(import.meta.dirname, '..', '..');
let repo = '';

afterEach(() => {
  if (repo) rmSync(repo, { recursive: true, force: true });
  repo = '';
});

function git(...args: string[]): void {
  const result = spawnSync(
    'git',
    ['-c', 'user.name=test', '-c', 'user.email=test@example.invalid', ...args],
    { cwd: repo, encoding: 'utf8' }
  );
  if (result.status !== 0) throw new Error(result.stderr);
}

function put(file: string, text: string): void {
  mkdirSync(dirname(join(repo, file)), { recursive: true });
  writeFileSync(join(repo, file), text);
}

/** A repository whose second commit changes an API report, with core at `coreVersion`. */
function repository(
  coreVersion: string,
  changeset: boolean,
  report = 'docs/api/core/store.api.md'
): void {
  repo = mkdtempSync(join(tmpdir(), 'api-changeset-'));
  for (const file of ['scripts/check-api-changeset.mjs', 'scripts/lib-packages.mjs']) {
    mkdirSync(dirname(join(repo, file)), { recursive: true });
    copyFileSync(join(ROOT, file), join(repo, file));
  }
  put('.changeset/config.json', '{}');
  for (const { name, root } of PACKAGES) {
    put(`${root}/package.json`, JSON.stringify({ name, version: '0.0.0' }));
  }
  put(report, 'before\n');
  git('init', '-q', '-b', 'main');
  git('add', '.');
  git('commit', '-q', '-m', 'base');
  git('branch', 'base');
  put(report, 'after\n');
  put('packages/core/package.json', JSON.stringify({ name: 'x', version: coreVersion }));
  if (changeset) put('.changeset/api.md', '---\n---\n\nChange.\n');
  git('add', '.');
  git('commit', '-q', '-m', 'change');
}

function check(): { status: number | null; output: string } {
  const result = spawnSync('node', ['scripts/check-api-changeset.mjs', '--base', 'base'], {
    cwd: repo,
    encoding: 'utf8',
  });
  return { status: result.status, output: result.stdout + result.stderr };
}

describe('check-api-changeset', () => {
  test('passes an API change with no changeset while every package is at 0.0.0', () => {
    repository('0.0.0', false);
    const { status, output } = check();
    expect(status).toBe(0);
    expect(output).toContain('no package has been released');
  });

  test('fails an API change with no changeset after a package is released', () => {
    repository('0.1.0', false);
    const { status, output } = check();
    expect(status).toBe(1);
    expect(output).toContain('without a changeset');
  });

  test('passes an API change with a changeset after a package is released', () => {
    repository('0.1.0', true);
    const { status, output } = check();
    expect(status).toBe(0);
    expect(output).toContain('covered by 1 changeset');
  });

  test('passes an API change of a private package with no changeset after a release', () => {
    const slug = PACKAGES.find((pkg) => pkg.private === true)?.slug;
    expect(slug).toBe('editor-api');
    repository('0.1.0', false, `docs/api/${slug}/index.api.md`);
    const { status, output } = check();
    expect(status).toBe(0);
    expect(output).toContain('no versioned package API change');
  });
});
