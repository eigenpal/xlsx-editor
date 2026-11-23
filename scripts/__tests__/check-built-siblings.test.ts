import { afterEach, beforeEach, expect, test } from 'bun:test';
import { mkdirSync, mkdtempSync, rmSync, utimesSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { staleness } from '../check-built-siblings.mjs';

const CORE = '@xlsx-editor.dev/core';
const built = new Date('2026-01-02T00:00:00Z');
const older = new Date('2026-01-01T00:00:00Z');
const newer = new Date('2026-01-03T00:00:00Z');
let root: string;
let packages: string;

function write(file: string, time: Date) {
  mkdirSync(join(file, '..'), { recursive: true });
  writeFileSync(file, '');
  utimesSync(file, time, time);
  utimesSync(join(file, '..'), time, time);
}

beforeEach(() => {
  root = mkdtempSync(join(tmpdir(), 'built-siblings-'));
  packages = join(root, 'packages');
  write(join(root, 'tsconfig.base.json'), older);
  write(join(packages, 'core/src/index.ts'), older);
  write(join(packages, 'core/package.json'), older);
  write(join(packages, 'core/dist/index.d.ts'), built);
  write(join(packages, 'ooxml-core/src/zip.ts'), older);
});

afterEach(() => rmSync(root, { recursive: true, force: true }));

test('a current build of core passes', () => {
  expect(staleness(CORE, join(packages, 'core'))).toBeNull();
});

test('a newer source of the inlined OOXML layer makes the build of core stale', () => {
  write(join(packages, 'ooxml-core/src/zip.ts'), newer);
  expect(staleness(CORE, join(packages, 'core'))).toContain('older than its source');
});

test('the OOXML layer does not make another package stale', () => {
  write(join(packages, 'ooxml-core/src/zip.ts'), newer);
  expect(staleness('@xlsx-editor.dev/i18n', join(packages, 'core'))).toBeNull();
});

test('a newer root tsconfig.base.json makes a build stale', () => {
  write(join(root, 'tsconfig.base.json'), newer);
  expect(staleness(CORE, join(packages, 'core'))).toContain('older than its source');
});
