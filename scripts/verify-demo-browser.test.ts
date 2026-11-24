import { describe, expect, spyOn, test } from 'bun:test';
import {
  CI_SKIP_REASON,
  DEFAULT_ENGINE,
  engineSummary,
  CLIPBOARD_READ_CAUSE,
  HEADLESS_SKIPS,
  launchProblem,
  launchRemedy,
  optionsOf,
  PLAIN_PASTE_CAUSE,
  pageAlerts,
  runEngines,
  runPasses,
  skipIssue,
  skipCounter,
  skipLine,
  stepResult,
  stopReason,
  TimeLimitError,
  unknownSkips,
  USAGE,
  withTimeLimit,
} from './verify-demo-browser.ts';

describe('the skip report', () => {
  const issue = HEADLESS_SKIPS.firefox!['cut and paste']!;

  test('the line of a skipped step names the engine and the issue', () => {
    expect(skipLine('firefox', 'cut and paste', issue)).toBe(
      `cut and paste: SKIPPED in firefox, ${issue}`
    );
  });

  test('the summary names the issue of a skip, and only the issues of this run', () => {
    expect(engineSummary('firefox', { passed: true, skipped: 1, skipIssues: [issue] }, false)).toBe(
      `firefox: passed, 1 skipped (${issue})`
    );
  });

  test('the summary names no issue when nothing was skipped', () => {
    expect(engineSummary('firefox', { passed: true, skipped: 0, skipIssues: [] })).toBe(
      'firefox: passed, 0 skipped'
    );
    expect(engineSummary('webkit', { passed: false, skipped: 0, skipIssues: [] })).toBe(
      'webkit: FAILED, 0 skipped'
    );
  });
});

describe('the count of skips', () => {
  test('each printed SKIPPED line counts in the summary', () => {
    const lines: string[] = [];
    const skips = skipCounter('firefox', (line) => lines.push(line));
    skips.skip('cut and paste', PLAIN_PASTE_CAUSE);
    skips.skip('copy and paste', PLAIN_PASTE_CAUSE);
    skips.skip('menu bar clipboard', CLIPBOARD_READ_CAUSE);
    expect(lines.filter((line) => line.includes(': SKIPPED in firefox, '))).toHaveLength(3);
    const result = { passed: true, skipped: skips.count, skipIssues: skips.issues };
    expect(engineSummary('firefox', result, false)).toBe(
      `firefox: passed, 3 skipped (${PLAIN_PASTE_CAUSE}, ${CLIPBOARD_READ_CAUSE})`
    );
  });
});

describe('the result of a run', () => {
  const skipped = [['firefox', { passed: true, skipped: 2, skipIssues: [] }]] as const;

  test('a skip fails a run in CI', () => {
    expect(runPasses(skipped, true)).toBe(false);
  });

  test('a skip that fails a run in CI gives the reason on its own line, and no pass', () => {
    const result = { passed: true, skipped: 2, skipIssues: [PLAIN_PASTE_CAUSE] };
    const summary = engineSummary('firefox', result, true);
    expect(summary.split('\n')).toEqual([
      `firefox: FAILED, 2 skipped (${PLAIN_PASTE_CAUSE})`,
      `firefox: FAILED, ${CI_SKIP_REASON}`,
    ]);
    expect(summary).not.toContain('passed');
  });

  test('the help text states the rule of a skip in CI', () => {
    expect(USAGE).toContain(CI_SKIP_REASON);
  });

  test('a skip keeps a local run green, and the summary names it', () => {
    expect(runPasses(skipped, false)).toBe(true);
  });

  test('a run in CI with no skip passes', () => {
    expect(runPasses([['webkit', { passed: true, skipped: 0, skipIssues: [] }]], true)).toBe(true);
  });

  test('the summary names the first failed step', () => {
    const result = { passed: false, skipped: 0, skipIssues: [], firstFailed: 'cut and paste' };
    expect(engineSummary('webkit', result)).toBe(
      'webkit: FAILED, 0 skipped, first failed step "cut and paste"'
    );
  });
});

describe('the stop reason', () => {
  test('a sample that does not open gives its reason in the summary', () => {
    const reason = stopReason('sample open', 'the sample did not open');
    expect(
      engineSummary('chromium', { passed: false, skipped: 0, skipIssues: [], ...reason })
    ).toBe('chromium: FAILED, 0 skipped, stopped at "sample open": the sample did not open');
  });

  test('a run that reached every step gives no reason', () => {
    expect(stopReason(null, 'the page was lost')).toEqual({});
  });
});

describe('the command line', () => {
  const run = (...args: string[]) =>
    Bun.spawnSync(['bun', 'scripts/verify-demo.ts', ...args], {
      cwd: `${import.meta.dir}/..`,
      stdout: 'pipe',
      stderr: 'pipe',
    });

  test('--help prints the usage and the step limit, and exits with code 0', () => {
    const out = run('--help');
    expect(out.exitCode).toBe(0);
    expect(out.stdout.toString()).toContain('usage: bun scripts/verify-demo.ts');
    expect(out.stdout.toString()).toContain('longer than 180 seconds');
  });

  test('an unknown engine prints the usage, and exits with code 2', () => {
    const out = run('--browser', 'nope');
    expect(out.exitCode).toBe(2);
    expect(out.stderr.toString()).toContain('usage: bun scripts/verify-demo.ts');
  });
});

describe('withTimeLimit', () => {
  test('gives the result of work that ends in time', async () => {
    expect(await withTimeLimit(Promise.resolve('done'), 1000, 'the step')).toBe('done');
  });

  test('fails work that hangs, and names the limit', async () => {
    const hung = withTimeLimit(new Promise<string>(() => {}), 5, 'the step');
    await expect(hung).rejects.toBeInstanceOf(TimeLimitError);
    await expect(withTimeLimit(new Promise(() => {}), 5, 'the step')).rejects.toThrow(
      'the step passed its time limit of 0.005 seconds'
    );
  });

  test('a page that does not answer is lost', async () => {
    const silent = {
      locator: () => ({ allInnerTexts: () => new Promise<string[]>(() => {}) }),
    } as unknown as Parameters<typeof pageAlerts>[0];
    const started = Date.now();
    const { lost } = await pageAlerts(silent);
    expect(lost).toContain('passed its time limit');
    expect(Date.now() - started).toBeLessThan(8000);
  }, 10_000);
});

describe('a run that loses a page or an engine', () => {
  const closed = {
    locator: () => ({
      allInnerTexts: async () => {
        throw new Error('allInnerTexts: Target page, context or browser has been closed');
      },
    }),
  } as unknown as Parameters<typeof pageAlerts>[0];

  test('a closed page gives the step a FAILED line in place of an error', async () => {
    const { shown, lost } = await pageAlerts(closed);
    expect(shown).toEqual([]);
    expect(lost).toContain('has been closed');
    const { line, passed } = stepResult('keyboard regions', lost, shown, 0, '');
    expect(passed).toBe(false);
    expect(line).toStartWith('keyboard regions: FAILED, the page is lost');
  });

  test('a page that answers has no problem', async () => {
    const open = {
      locator: () => ({ allInnerTexts: async () => ['a message'] }),
    } as unknown as Parameters<typeof pageAlerts>[0];
    expect(await pageAlerts(open)).toEqual({ shown: ['a message'], lost: null });
  });

  test('the engines after one that throws still run, and the run fails', async () => {
    const ran: string[] = [];
    // The failure of this test's engine is expected, so a green run prints no `FAILED` line.
    const printed = [spyOn(console, 'log'), spyOn(console, 'error')];
    for (const spy of printed) spy.mockImplementation(() => {});
    const results = await runEngines(['chromium', 'firefox', 'webkit'], async (engine) => {
      ran.push(engine);
      if (engine === 'firefox') throw new Error('the browser went away');
      return { passed: true, skipped: 0, skipIssues: [] };
    }).finally(() => {
      for (const spy of printed) spy.mockRestore();
    });
    expect(ran).toEqual(['chromium', 'firefox', 'webkit']);
    expect(results.map(([, r]) => r.passed)).toEqual([true, false, true]);
    expect(results.every(([, r]) => r.passed)).toBe(false);
    expect(engineSummary('firefox', results[1]![1])).toBe(
      'firefox: FAILED, 0 skipped, did not finish: the browser went away'
    );
  });

  test('the summary names an engine that stopped or did not run', () => {
    const stopped = 'stopped at "keyboard regions": the page was lost';
    expect(engineSummary('chromium', { passed: false, skipped: 0, skipIssues: [], stopped })).toBe(
      `chromium: FAILED, 0 skipped, ${stopped}`
    );
    const launch = 'did not run: the browser did not launch';
    expect(
      engineSummary('webkit', { passed: false, skipped: 0, skipIssues: [], stopped: launch })
    ).toBe(`webkit: FAILED, 0 skipped, ${launch}`);
  });
});

describe('stepResult', () => {
  test('passes a step with no problem, alert, or console error', () => {
    expect(stepResult('theme', null, [], 0, 'it switched')).toEqual({
      line: 'theme: it switched, console errors 0',
      passed: true,
    });
  });

  test('prints FAILED for a step that fails only on a console error', () => {
    expect(stepResult('theme', null, [], 2, 'it switched')).toEqual({
      line: 'theme: FAILED, 2 console or page errors, console errors 2',
      passed: false,
    });
  });

  test('prints FAILED with the alert text', () => {
    expect(stepResult('theme', null, ['One', 'Two'], 0, 'it switched')).toEqual({
      line: 'theme: FAILED, alert "One; Two", console errors 0',
      passed: false,
    });
  });

  test('names the problem of the step first', () => {
    expect(stepResult('theme', 'no sheet', ['One'], 1, '').line).toBe(
      'theme: FAILED, no sheet, console errors 1'
    );
  });
});

describe('unknownSkips', () => {
  test('names a listed step that the run does not have', () => {
    const run = new Set(['cut and paste', 'menu bar clipboard']);
    expect(unknownSkips('firefox', run)).toEqual([]);
    expect(unknownSkips('firefox', new Set(['copy and paste']))).toEqual([
      'cut and paste',
      'menu bar clipboard',
    ]);
  });

  test('an engine skip applies with and without a browser window', () => {
    expect(skipIssue('firefox', 'menu bar clipboard', false)).toBe(CLIPBOARD_READ_CAUSE);
    expect(skipIssue('firefox', 'menu bar clipboard', true)).toBe(CLIPBOARD_READ_CAUSE);
    expect(skipIssue('chromium', 'menu bar clipboard', false)).toBeUndefined();
  });

  test('names nothing for an engine with no skip', () => {
    expect(unknownSkips('chromium', new Set())).toEqual([]);
  });
});

describe('optionsOf', () => {
  test('runs Chromium alone with no option', () => {
    expect(DEFAULT_ENGINE).toBe('chromium');
    expect(optionsOf([])).toEqual({ engines: ['chromium'] });
  });

  test('takes each --browser once, in the order given', () => {
    expect(optionsOf(['--browser', 'webkit', '--browser=firefox', '--browser', 'webkit'])).toEqual({
      engines: ['webkit', 'firefox'],
    });
  });

  test('refuses an unknown engine, a missing engine, and an unknown option', () => {
    expect(() => optionsOf(['--browser', 'edge'])).toThrow('--browser takes');
    expect(() => optionsOf(['--browser'])).toThrow('--browser takes');
    expect(() => optionsOf(['--browser', '__proto__'])).toThrow('--browser takes');
    expect(() => optionsOf(['--headed'])).toThrow('unknown option');
    expect(() => optionsOf(['--fixtures'])).toThrow('unknown option');
  });
});

// The form of the error of `validateDependenciesLinux` in playwright-core 1.65.0, written out
// from its source. No Linux host produced this text.
const MISSING_LIBRARIES = [
  'browserType.launch: ',
  '╔══════════════════════════════════════════════════════╗',
  '║ Host system is missing dependencies to run browsers. ║',
  '║ Missing libraries:                                   ║',
  '║     libgtk-4.so.1                                    ║',
  '╚══════════════════════════════════════════════════════╝',
].join('\n');

describe('launchProblem', () => {
  test('keeps each line of an error whose first line has no cause', () => {
    const text = launchProblem(new Error(`\n${MISSING_LIBRARIES}\n`));
    expect(text).toStartWith('browserType.launch:');
    expect(text).toContain('Host system is missing dependencies to run browsers.');
    expect(text).toContain('libgtk-4.so.1');
  });

  test('reads a thrown value that is not an error', () => {
    expect(launchProblem('  no browser  ')).toBe('no browser');
  });
});

describe('launchRemedy', () => {
  test('names install-deps for missing system libraries', () => {
    const remedy = 'bunx playwright install-deps webkit';
    expect(launchRemedy('webkit', MISSING_LIBRARIES)).toContain(remedy);
    expect(launchRemedy('webkit', 'libgtk-4.so.1: cannot open shared object file')).toContain(
      remedy
    );
  });

  test('names the install only for a missing executable', () => {
    const missing = "browserType.launch: Executable doesn't exist at /cache/firefox";
    expect(launchRemedy('firefox', missing)).toContain('bunx playwright install firefox');
    const stopped = launchRemedy('firefox', 'browserType.launch: Could not find profile folder');
    expect(stopped).not.toContain('install');
    expect(stopped).toContain('firefox did not start');
    expect(stopped).toContain('default engine (chromium)');
  });
});

describe('HEADLESS_SKIPS', () => {
  test('names a cause for each skipped step', () => {
    for (const steps of Object.values(HEADLESS_SKIPS))
      for (const cause of Object.values(steps ?? {}))
        expect(cause.trim().length).toBeGreaterThan(10);
  });

  test('finds no skip under a key of Object.prototype', () => {
    for (const name of ['constructor', 'toString', '__proto__', 'hasOwnProperty'])
      expect(skipIssue('webkit', name, false)).toBeUndefined();
  });

  test('a headed run skips no step', () => {
    expect(skipIssue('firefox', 'cut and paste', false)).toContain('headless Firefox clipboard');
    expect(skipIssue('firefox', 'cut and paste', true)).toBeUndefined();
  });
});
