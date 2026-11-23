// The setup of `bun run fidelity` (`run.ts`) before the harness starts. Each case stops before
// the tools, so it needs neither Poppler nor Chromium.

import { afterEach, describe, expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import {
  chmodSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { chromium } from '@playwright/test';
import { CHROMIUM_MISSING, POPPLER_MISSING } from './chromium.ts';
import { popplerVersion } from './poppler.ts';

const dirs: string[] = [];
afterEach(() => {
  for (const d of dirs.splice(0)) rmSync(d, { recursive: true, force: true });
});

function tempDir(): string {
  const dir = mkdtempSync(join(tmpdir(), 'xlsx-fidelity-run-'));
  dirs.push(dir);
  return dir;
}

/** Run `bun run fidelity` of `names` with `env`, through the entry of the commands. */
function fidelity(env: Record<string, string>, names = ['basic-values']) {
  return spawnSync(process.execPath, [join(import.meta.dir, 'cli.ts'), 'run', ...names], {
    encoding: 'utf8',
    env: { ...process.env, ...env },
    timeout: 300_000,
  });
}

/** Whether this machine has the tools of a run: Poppler and Playwright's Chromium. */
const hasTools = popplerVersion() !== 'missing' && existsSync(chromium.executablePath());

/** A report folder with the report of an earlier run and a file that a person put there. */
function earlierReport(): string {
  const report = tempDir();
  writeFileSync(join(report, 'report.md'), 'the report of an earlier run');
  mkdirSync(join(report, 'src'));
  writeFileSync(join(report, 'src', 'a.ts'), 'source');
  writeFileSync(join(report, 'important.txt'), 'keep');
  return report;
}

/** Expect exit code 2 and one stderr line that starts with `start`, and no stdout. */
function expectOneLine(run: ReturnType<typeof fidelity>, start: string): void {
  expect(run.status).toBe(2);
  const lines = run.stderr.trim().split('\n');
  expect(lines).toHaveLength(1);
  expect(lines[0]).toStartWith(start);
  expect(run.stdout).toBe('');
}

describe('the setup of the fidelity run', () => {
  test('a reference font folder that is a file prints one line and exits with code 2', () => {
    const file = join(tempDir(), 'not-a-folder');
    writeFileSync(file, 'text');
    expectOneLine(
      fidelity({ XLSX_REFERENCE_FONTS: file, XLSX_FIDELITY_REPORT_DIR: tempDir() }),
      'fidelity: the reference fonts in XLSX_REFERENCE_FONTS cannot be read: ENOTDIR'
    );
  });

  test('a malformed reference manifest prints one line, exits with code 2, and keeps the report', () => {
    const report = earlierReport();
    const manifest = join(tempDir(), 'references.json');
    const badRecord = '{ "version": 1, "references": { "basic-values": { "workbook": 1 } } }';
    for (const text of ['{ "version": 1, "references": ', '[]', badRecord]) {
      writeFileSync(manifest, text);
      expectOneLine(
        fidelity({ XLSX_FIDELITY_REPORT_DIR: report, XLSX_FIDELITY_MANIFEST: manifest }),
        'fidelity: the reference manifest cannot be read: '
      );
      expect(readFileSync(join(report, 'report.md'), 'utf8')).toBe('the report of an earlier run');
    }
  });

  test('a lock that another live process holds prints one line and exits with code 2', () => {
    const report = tempDir();
    mkdirSync(join(report, '.run.lock'));
    // The test runner is alive, and it is not the run.
    writeFileSync(join(report, '.run.lock', 'pid'), String(process.pid));
    writeFileSync(join(report, 'report.md'), 'the report of an earlier run');
    expectOneLine(
      fidelity({ XLSX_FIDELITY_REPORT_DIR: report }),
      `fidelity: another \`bun run fidelity\` (process ${process.pid}) holds `
    );
    // The held run clears nothing.
    expect(existsSync(join(report, 'report.md'))).toBe(true);
  });

  test('a report folder that cannot be written prints one line and exits with code 2', () => {
    const file = join(tempDir(), 'a-file');
    writeFileSync(file, 'text');
    expectOneLine(
      fidelity({ XLSX_FIDELITY_REPORT_DIR: join(file, 'report') }),
      'fidelity: the run lock cannot be taken: ENOTDIR'
    );
  });

  test('a missing browser, or a missing Poppler, keeps the report of the earlier run', () => {
    const report = earlierReport();
    const run = fidelity({
      XLSX_FIDELITY_REPORT_DIR: report,
      PLAYWRIGHT_BROWSERS_PATH: join(tempDir(), 'no-browsers'),
    });
    // Without Poppler, the run stops at Poppler, before the browser; the report stays as well.
    const missing = popplerVersion() === 'missing' ? POPPLER_MISSING : CHROMIUM_MISSING;
    expect(run.status).toBe(2);
    expect(run.stderr.trim()).toBe(`fidelity: ${missing}`);
    expect(readFileSync(join(report, 'report.md'), 'utf8')).toBe('the report of an earlier run');
    expect(existsSync(join(report, 'important.txt'))).toBe(true);
  });

  // The run clears the report folder after it starts the browser, so this test needs the tools.
  // The test of `clearReport` (`run-lock.test.ts`) checks the throw without them.
  test.skipIf(!hasTools || process.getuid?.() === 0)(
    'a report entry that cannot be cleared prints one line and exits with code 2',
    () => {
      const report = tempDir();
      const fixture = join(report, 'basic-values');
      mkdirSync(fixture);
      writeFileSync(join(fixture, 'ours.png'), 'old');
      chmodSync(fixture, 0o500);
      try {
        const run = fidelity({ XLSX_FIDELITY_REPORT_DIR: report });
        expect(run.status).toBe(2);
        expect(run.stderr).toContain('fidelity: the report folder cannot be cleared: ');
      } finally {
        chmodSync(fixture, 0o700);
      }
    },
    120_000
  );

  // A whole run needs Poppler and Chromium. The unit job of CI has neither and skips this test;
  // the tests of `clearsEntry` and `clearReport` (`run-lock.test.ts`) check the rule there.
  test.skipIf(!hasTools)(
    'a run of every fixture keeps the entries of the report folder that it does not write',
    () => {
      const report = earlierReport();
      const run = fidelity({ XLSX_FIDELITY_REPORT_DIR: report }, []);
      expect(run.status === 0 || run.status === 1).toBe(true);
      expect(readFileSync(join(report, 'src', 'a.ts'), 'utf8')).toBe('source');
      expect(readFileSync(join(report, 'important.txt'), 'utf8')).toBe('keep');
      expect(readFileSync(join(report, 'report.md'), 'utf8')).not.toBe(
        'the report of an earlier run'
      );
    },
    300_000
  );
});
