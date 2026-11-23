// Tests for scripts/check-max-lines-caps.mjs. Each case runs a copy of the script in a new folder.
import { expect, test } from 'bun:test';
import { copyFileSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';

const CACHE = join(import.meta.dir, '../.cache/test');

function caps(files: Record<string, string>) {
  // Inside the repository, so the copied script resolves `typescript` from node_modules.
  mkdirSync(CACHE, { recursive: true });
  const root = mkdtempSync(join(CACHE, 'max-lines-'));
  try {
    mkdirSync(join(root, 'scripts/lib'), { recursive: true });
    mkdirSync(join(root, 'packages/core/src'), { recursive: true });
    copyFileSync(
      join(import.meta.dir, 'check-max-lines-caps.mjs'),
      join(root, 'scripts/check-max-lines-caps.mjs')
    );
    copyFileSync(
      join(import.meta.dir, 'lib/max-lines-directives.mjs'),
      join(root, 'scripts/lib/max-lines-directives.mjs')
    );
    copyFileSync(join(import.meta.dir, '../.oxlintrc.json'), join(root, '.oxlintrc.json'));
    for (const [path, text] of Object.entries(files)) writeFileSync(join(root, path), text);
    return spawnSync('node', [join(root, 'scripts/check-max-lines-caps.mjs')], {
      cwd: root,
      encoding: 'utf8',
    });
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
}

const DISABLE = ['/* eslint-disable ', 'max-lines */'].join('');

test('passes a tree with no blanket disable', () => {
  expect(caps({ 'packages/core/src/a.ts': 'export {};\n' }).status).toBe(0);
});

test('refuses an undeclared blanket disable in a package', () => {
  const result = caps({ 'packages/core/src/a.ts': `${DISABLE}\nexport {};\n` });
  expect(result.status).toBe(1);
  expect(result.stderr + result.stdout).toContain('undeclared blanket disable');
});

test('refuses an undeclared blanket disable in a script', () => {
  const result = caps({ 'scripts/big.ts': `${DISABLE}\nexport {};\n` });
  expect(result.status).toBe(1);
  expect(result.stderr + result.stdout).toContain('scripts/big.ts');
});
