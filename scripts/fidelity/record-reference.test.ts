// The reference recorder `bun run fidelity:record-reference` (`record-reference.ts`).

import { afterEach, describe, expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { missingFiles } from './references.ts';
import { RECORD_REFERENCE_USAGE } from './run-args.ts';

const dirs: string[] = [];
afterEach(() => {
  for (const d of dirs.splice(0)) rmSync(d, { recursive: true, force: true });
});

describe('the reference recorder', () => {
  test('names each missing file of the fixtures in one line, with its next step', () => {
    const dir = mkdtempSync(join(tmpdir(), 'xlsx-record-'));
    dirs.push(dir);
    writeFileSync(join(dir, 'whole.xlsx'), '');
    writeFileSync(join(dir, 'whole.pdf'), '');
    writeFileSync(join(dir, 'half.xlsx'), '');
    expect(missingFiles(['whole', 'half', 'none'], dir)).toEqual([
      'fixtures/half.pdf is missing: a maintainer makes it (CONTRIBUTING.md, "Reference PDFs").',
      'fixtures/none.xlsx is missing: build it with `bun run fixtures:build`.',
      'fixtures/none.pdf is missing: a maintainer makes it (CONTRIBUTING.md, "Reference PDFs").',
    ]);
    expect(missingFiles(['whole'], dir)).toEqual([]);
  });

  test('a usage error exits with code 2 before it reads a file', () => {
    const run = spawnSync(
      process.execPath,
      [join(import.meta.dir, 'cli.ts'), 'record-reference', '--decimal', ',', 'basic-values'],
      { encoding: 'utf8' }
    );
    expect(run.status).toBe(2);
    expect(run.stderr.trim().startsWith('fidelity: ')).toBe(true);
    expect(run.stderr.trim().endsWith(RECORD_REFERENCE_USAGE)).toBe(true);
  });

  test('reference fonts that cannot be read stop it in one line with code 2', () => {
    const dir = mkdtempSync(join(tmpdir(), 'xlsx-record-'));
    dirs.push(dir);
    const file = join(dir, 'not-a-folder');
    writeFileSync(file, 'text');
    const run = spawnSync(
      process.execPath,
      [
        join(import.meta.dir, 'cli.ts'),
        'record-reference',
        '--decimal',
        '.',
        '--group',
        ',',
        'basic-values',
      ],
      { encoding: 'utf8', env: { ...process.env, XLSX_REFERENCE_FONTS: file } }
    );
    expect(run.status).toBe(2);
    const lines = run.stderr.trim().split('\n');
    expect(lines).toHaveLength(1);
    expect(lines[0]).toStartWith(
      'fidelity: the reference fonts in XLSX_REFERENCE_FONTS cannot be read: ENOTDIR'
    );
    expect(run.stdout).toBe('');
  });
});
