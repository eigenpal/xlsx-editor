// The entry of the fidelity commands (`cli.ts`): a module that does not load stops it in one line.

import { afterEach, describe, expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { copyFileSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const dirs: string[] = [];
afterEach(() => {
  for (const d of dirs.splice(0)) rmSync(d, { recursive: true, force: true });
});

/** Run a copy of `cli.ts` in a temp folder that holds `files`. */
function runCopy(files: Record<string, string>, args: string[]) {
  const dir = mkdtempSync(join(tmpdir(), 'xlsx-fidelity-cli-'));
  dirs.push(dir);
  copyFileSync(join(import.meta.dir, 'cli.ts'), join(dir, 'cli.ts'));
  for (const [name, text] of Object.entries(files)) writeFileSync(join(dir, name), text);
  return spawnSync(process.execPath, [join(dir, 'cli.ts'), ...args], { encoding: 'utf8' });
}

describe('the entry of the fidelity commands', () => {
  test('a thresholds.ts that does not load exits with code 2 in one line', () => {
    const run = runCopy(
      {
        'check-known-gaps.ts':
          "import { THRESHOLDS } from './thresholds.ts';\nexport const command = async () => (THRESHOLDS ? 0 : 1);\n",
        'thresholds.ts': 'export const THRESHOLDS = {\n',
      },
      ['check-gaps', '--base', 'HEAD']
    );
    expect(run.status).toBe(2);
    const lines = run.stderr.trim().split('\n');
    expect(lines).toHaveLength(1);
    expect(lines[0]).toStartWith('fidelity: check-known-gaps.ts does not load: ');
    expect(run.stdout).toBe('');
  });

  test('a module that throws while it loads exits with code 2 in one line', () => {
    const run = runCopy({ 'run.ts': "throw new Error('broken module\\n    at stack');\n" }, [
      'run',
      'basic-values',
    ]);
    expect(run.status).toBe(2);
    expect(run.stderr.trim()).toBe('fidelity: run.ts does not load: broken module');
  });

  test('the command gets its own arguments and gives its exit code', () => {
    const run = runCopy(
      {
        'record-reference.ts':
          'export const command = async () => { console.log(JSON.stringify(process.argv.slice(2))); return 1; };\n',
      },
      ['record-reference', '--decimal', ',']
    );
    expect(run.status).toBe(1);
    expect(run.stdout.trim()).toBe('["--decimal",","]');
  });

  test('a missing command exits with code 2 in one line', () => {
    const run = runCopy({}, []);
    expect(run.status).toBe(2);
    expect(run.stderr.trim()).toBe(
      'fidelity: missing command: use run, check-gaps, record-reference'
    );
  });

  test('an unknown command exits with code 2 in one line', () => {
    const run = runCopy({}, ['nope']);
    expect(run.status).toBe(2);
    expect(run.stderr.trim()).toBe(
      'fidelity: unknown command "nope": use run, check-gaps, record-reference'
    );
  });
});
