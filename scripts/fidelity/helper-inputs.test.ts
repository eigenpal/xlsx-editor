import { afterEach, describe, expect, test } from 'bun:test';
import { mkdirSync, mkdtempSync, rmSync, utimesSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { helperInputs } from './helper-inputs.ts';
import { REPORT_DIR_ENV } from './report.ts';

const roots: string[] = [];
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true });
});

function setUp(): { fixtures: string; env: Record<string, string>; report: string } {
  const root = mkdtempSync(join(tmpdir(), 'xlsx-helper-inputs-'));
  roots.push(root);
  const fixtures = join(root, 'fixtures');
  const report = join(root, 'report');
  mkdirSync(fixtures);
  mkdirSync(join(report, 'sheet'), { recursive: true });
  return { fixtures, env: { [REPORT_DIR_ENV]: report }, report };
}

const tool = 'glyph-share.ts';

describe('the input check of a helper tool', () => {
  test('refuses a missing or unknown fixture name with the usage line', () => {
    const { fixtures, env } = setUp();
    const usage = ['usage: bun scripts/fidelity/glyph-share.ts <fixture>', 'fixtures: sheet'];
    expect(helperInputs(tool, undefined, ['diff.png'], ['sheet'], fixtures, env).errors).toEqual(
      usage
    );
    expect(helperInputs(tool, 'typo', ['diff.png'], ['sheet'], fixtures, env).errors).toEqual([
      'unknown fixture: typo',
      ...usage,
    ]);
  });

  test('refuses a missing reference PDF before it reads a raster', () => {
    const { fixtures, env, report } = setUp();
    writeFileSync(join(report, 'sheet', 'diff.png'), '');
    expect(helperInputs(tool, 'sheet', ['diff.png'], ['sheet'], fixtures, env).errors).toEqual([
      `no reference PDF for sheet in ${fixtures}`,
    ]);
  });

  test('refuses a missing raster, and names the report folder variable in the remedy', () => {
    const { fixtures, env, report } = setUp();
    writeFileSync(join(fixtures, 'sheet.pdf'), '');
    writeFileSync(join(report, 'sheet', 'ours.png'), '');
    const inputs = helperInputs(
      tool,
      'sheet',
      ['ours.png', 'reference.png'],
      ['sheet'],
      fixtures,
      env
    );
    expect(inputs.errors).toEqual([
      `no ${join(report, 'sheet', 'reference.png')}: run \`bun run fidelity\` with the same ${REPORT_DIR_ENV} first`,
    ]);
  });

  test('needs only the rasters that the run always writes', () => {
    const { fixtures, env, report } = setUp();
    writeFileSync(join(fixtures, 'sheet.pdf'), '');
    writeFileSync(join(report, 'sheet', 'ours.png'), '');
    // A run with 0 differing pixels writes `ours.png` and no `diff.png`.
    expect(helperInputs(tool, 'sheet', ['ours.png'], ['sheet'], fixtures, env).errors).toEqual([]);
  });

  test('warns when the raster of the run is older than the reference PDF', () => {
    const { fixtures, env, report } = setUp();
    writeFileSync(join(fixtures, 'sheet.pdf'), '');
    writeFileSync(join(report, 'sheet', 'diff.png'), '');
    const fresh = helperInputs(tool, 'sheet', ['diff.png'], ['sheet'], fixtures, env);
    expect(fresh.errors).toEqual([]);
    expect(fresh.warning).toBeUndefined();
    expect(fresh.report).toBe(join(report, 'sheet'));
    expect(fresh.name).toBe('sheet');
    utimesSync(join(report, 'sheet', 'diff.png'), 1, 1);
    expect(helperInputs(tool, 'sheet', ['diff.png'], ['sheet'], fixtures, env).warning).toBe(
      `sheet: diff.png is older than the reference PDF: run \`bun run fidelity\` with the same ${REPORT_DIR_ENV} again`
    );
  });
});

describe('glyph-share.ts', () => {
  const run = (args: string[], env: Record<string, string>) =>
    Bun.spawnSync(['bun', join(import.meta.dir, 'glyph-share.ts'), ...args], {
      env: { ...process.env, ...env },
    });

  test('exits with code 2 and the usage line on an unknown fixture', () => {
    const { env } = setUp();
    const result = run(['typo'], env);
    expect(result.exitCode).toBe(2);
    expect(result.stderr.toString()).toContain('unknown fixture: typo');
  });

  test('reports 0 differing pixels when the run wrote ours.png and no diff.png', () => {
    const { env, report } = setUp();
    mkdirSync(join(report, 'number-formats'));
    writeFileSync(join(report, 'number-formats', 'ours.png'), '');
    const result = run(['number-formats'], env);
    expect(result.exitCode).toBe(0);
    expect(result.stdout.toString()).toBe(
      'number-formats: 0 differing pixels (the run wrote no diff.png)\n'
    );
  });
});
