// The arguments of the browser bench: each wrong value stops the command with a usage error
// (exit code 2) before it starts a server or a browser. The test runs the command, as the gate
// test runs the edit bench, so it starts no browser.

import { expect, test } from 'bun:test';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

function bench(...flags: string[]): { exitCode: number; stderr: string } {
  const out = Bun.spawnSync(
    [process.execPath, join(import.meta.dir, 'browser-bench.ts'), ...flags],
    { stdout: 'pipe', stderr: 'pipe' }
  );
  return { exitCode: out.exitCode, stderr: out.stderr.toString() };
}

test('the browser bench refuses too few rows of data for its edits', () => {
  for (const [flags, message] of [
    [['edit', '--workbooks', 'grid', '--rows', '40'], '--rows takes 41 or more in the edit mode'],
    [['edit', '--workbooks', 'strings', '--strings', '819'], '--strings takes 820 or more'],
    [['memory', '--workbooks', 'grid', '--rows', '12'], '--rows takes 13 or more in the memory'],
    [['memory', '--workbooks', 'strings', '--strings', '259'], '--strings takes 260 or more'],
  ] as const) {
    const run = bench(...flags);
    expect(run.exitCode, run.stderr).toBe(2);
    expect(run.stderr).toContain(message);
  }
}, 60_000);

test('the browser bench refuses an --out path that it cannot write', () => {
  const dir = mkdtempSync(join(tmpdir(), 'xlsx-browser-bench-'));
  try {
    const file = join(dir, 'file');
    writeFileSync(file, '');
    for (const [out, message] of [
      [dir, 'is a folder'],
      [join(dir, 'missing', 'out.json'), 'does not exist'],
      [join(file, 'out.json'), 'does not exist'],
    ] as const) {
      const run = bench('open', '--workbooks', 'grid', '--out', out);
      expect(run.exitCode, run.stderr).toBe(2);
      expect(run.stderr).toContain(message);
    }
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}, 60_000);
