import { afterEach, describe, expect, test } from 'bun:test';
import { cpSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';

const repoRoot = resolve(import.meta.dir, '../..');
let copy: string | undefined;

afterEach(() => {
  if (copy) rmSync(copy, { recursive: true, force: true });
  copy = undefined;
});

function copyRepository(): string {
  const directory = mkdtempSync(join(tmpdir(), 'validate-i18n-'));
  mkdirSync(join(directory, 'scripts'));
  mkdirSync(join(directory, 'packages'));
  cpSync(join(repoRoot, 'scripts/validate-i18n.mjs'), join(directory, 'scripts/validate-i18n.mjs'));
  cpSync(join(repoRoot, 'scripts/lib'), join(directory, 'scripts/lib'), { recursive: true });
  cpSync(join(repoRoot, 'packages/i18n'), join(directory, 'packages/i18n'), {
    recursive: true,
    filter: (file) => !file.includes('node_modules') && !file.includes('/dist'),
  });
  return directory;
}

const validate = (directory: string) =>
  Bun.spawnSync(['node', join(directory, 'scripts/validate-i18n.mjs'), 'validate']);

describe('i18n:validate with English as the only locale', () => {
  test('passes on the repository as it is', () => {
    copy = copyRepository();
    expect(validate(copy).exitCode).toBe(0);
  });

  test('fails when the exports map leaves out the English entry', () => {
    copy = copyRepository();
    const file = join(copy, 'packages/i18n/package.json');
    const manifest = JSON.parse(readFileSync(file, 'utf8'));
    delete manifest.exports['./en'];
    writeFileSync(file, JSON.stringify(manifest, null, 2));
    const result = validate(copy);
    expect(result.exitCode).toBe(1);
    expect(result.stderr.toString()).toContain('exports');
  });
});
