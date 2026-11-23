// The browser of `bun run fidelity` (`chromium.ts`).

import { describe, expect, test } from 'bun:test';
import {
  CHROMIUM_MISSING,
  checkPoppler,
  eachFixture,
  harnessStopped,
  launchBrowser,
  POPPLER_MISSING,
  TOOL_EXIT,
  withHarness,
  CLOSE_TIMEOUT_MS,
  chromiumPid,
} from './chromium.ts';

/** A browser whose pid CDP does not give. */
const noPid = (): Promise<undefined> => Promise.resolve(undefined);

describe('the browser of the fidelity harness', () => {
  test('a missing Chromium is one line with the install command', async () => {
    const missing = await launchBrowser(() =>
      Promise.reject(new Error("browserType.launch: Executable doesn't exist at /tmp/chrome"))
    );
    expect(missing).toEqual({ error: CHROMIUM_MISSING });
    expect(CHROMIUM_MISSING).toContain('bunx playwright install chromium');
    expect(CHROMIUM_MISSING).not.toContain('\n');
  });

  test('a launched browser passes, and another launch error is the line to print', async () => {
    expect(await launchBrowser(() => Promise.resolve('browser'))).toEqual({ browser: 'browser' });
    expect(await launchBrowser(() => Promise.reject(new Error('no display')))).toEqual({
      error: 'no display',
    });
    // A host without the libraries of Chromium needs another command: its error passes through
    // unchanged, and the run prints it.
    const deps = new Error(
      'browserType.launch: Host system is missing dependencies to run browsers. Please install them with: bunx playwright install-deps'
    );
    expect(await launchBrowser(() => Promise.reject(deps))).toEqual({ error: deps.message });
  });

  test('a missing Poppler is one line with the install commands and exit code 2', () => {
    expect(checkPoppler(() => 'missing')).toEqual({ error: POPPLER_MISSING, code: 2 });
    expect(POPPLER_MISSING).toContain('brew install poppler');
    expect(POPPLER_MISSING).toContain('sudo apt-get install poppler-utils');
    expect(checkPoppler(() => 'poppler 26.10.0')).toEqual({ poppler: 'poppler 26.10.0' });
  });

  test('a harness server that does not start gives TOOL_EXIT and one line', async () => {
    let launched = false;
    const done = await withHarness(
      () => Promise.reject(new Error('listen EADDRINUSE: address already in use\n    at stack')),
      () => {
        launched = true;
        return Promise.resolve({ close: () => Promise.resolve() });
      },
      noPid,
      () => Promise.resolve(0)
    );
    expect(done).toEqual({
      error: 'the harness server did not start: listen EADDRINUSE: address already in use',
      code: TOOL_EXIT,
    });
    expect(launched).toBe(false);
  });

  test('an error of the run closes the browser and the server, and stops in one line', async () => {
    const closed: string[] = [];
    const closable = (name: string) => ({
      close: () => {
        closed.push(name);
        return Promise.resolve();
      },
    });
    const run = withHarness(
      () => Promise.resolve(closable('server')),
      () => Promise.resolve(closable('browser')),
      noPid,
      () => {
        // As a manifest read that throws.
        throw new Error('ENOENT: no such file, open fixtures/references.json\n    at readManifest');
      }
    );
    const error = await run.then(
      () => undefined,
      (e: unknown) => e
    );
    expect(closed).toEqual(['browser', 'server']);
    expect(harnessStopped(error)).toBe(
      'fidelity: the harness stopped: ENOENT: no such file, open fixtures/references.json'
    );
    // A browser that does not start closes the server too.
    closed.length = 0;
    const noBrowser = await withHarness(
      () => Promise.resolve(closable('server')),
      () => Promise.reject(new Error('no display')),
      noPid,
      () => Promise.resolve(0)
    );
    expect(noBrowser).toEqual({ error: 'no display', code: TOOL_EXIT });
    expect(closed).toEqual(['server']);
  });

  test('a browser close that never ends is killed after the time limit', async () => {
    expect(CLOSE_TIMEOUT_MS).toBe(10_000);
    const hung = { close: () => new Promise<never>(() => undefined) };
    const kills: [number, string][] = [];
    const closed: string[] = [];
    const server = {
      close: () => {
        closed.push('server');
        return Promise.resolve();
      },
    };
    const done = await withHarness(
      () => Promise.resolve(server),
      () => Promise.resolve(hung),
      () => Promise.resolve(4242),
      () => Promise.resolve(0),
      {
        ms: 20,
        kill: (pid, signal) => {
          kills.push([pid, signal]);
          // No such process group: the kill goes to the pid.
          if (pid < 0) throw new Error('ESRCH');
        },
      }
    );
    expect(done).toEqual({ result: 0 });
    expect(kills).toEqual([
      [-4242, 'SIGKILL'],
      [4242, 'SIGKILL'],
    ]);
    expect(closed).toEqual(['server']);

    // Without a pid, the run warns in one line and ends.
    const warnings: string[] = [];
    const noKill = await withHarness(
      () => Promise.resolve(server),
      () => Promise.resolve(hung),
      () => Promise.reject(new Error('CDP is not available')),
      () => Promise.resolve(1),
      { ms: 20, warn: (line) => warnings.push(line), kill: () => kills.push([0, 'none']) }
    );
    expect(noKill).toEqual({ result: 1 });
    expect(kills).toHaveLength(2);
    expect(warnings).toEqual([
      'fidelity: a browser that does not answer keeps running until the run ends: CDP gave no pid to stop it',
    ]);
  });

  test('a browser close that fails with an error is killed too', async () => {
    const kills: number[] = [];
    const done = await withHarness(
      () => Promise.resolve({ close: () => Promise.resolve() }),
      () => Promise.resolve({ close: () => Promise.reject(new Error('Target closed')) }),
      () => Promise.resolve(4242),
      () => Promise.resolve(0),
      { ms: 1000, kill: (pid) => void kills.push(pid) }
    );
    expect(done).toEqual({ result: 0 });
    expect(kills).toEqual([-4242]);
  });

  test('chromiumPid reads the browser process from CDP', async () => {
    const browser = {
      newBrowserCDPSession: () =>
        Promise.resolve({
          send: () =>
            Promise.resolve({
              processInfo: [
                { type: 'renderer', id: 11 },
                { type: 'browser', id: 7 },
              ],
            }),
          detach: () => Promise.resolve(),
        }),
    };
    expect(await chromiumPid(browser)).toBe(7);
  });
});

describe('one fixture of the fidelity run', () => {
  // The run measures each fixture through `eachFixture`. The failures are injected, so the test
  // needs no browser, Poppler, or reference fonts.
  type Result = { name: string; status: 'compared' | 'error'; error?: string };
  const failed = (name: string, error: string): Result => ({ name, status: 'error', error });

  test('a workbook, folder, or page error fails that fixture only, and the next fixture runs', async () => {
    const errors: Record<string, Error> = {
      workbook: new Error("ENOENT: no such file or directory, open 'workbook.xlsx'"),
      folder: new Error(
        "EACCES: permission denied, mkdir 'fidelity-report/folder'\n    at mkdirSync"
      ),
      page: new Error('page.goto: net::ERR_CONNECTION_REFUSED'),
    };
    const ran: string[] = [];
    const results = await eachFixture(
      ['workbook', 'basic-values', 'folder', 'page', 'formulas'],
      async (name) => {
        ran.push(name);
        const error = errors[name];
        if (error) throw error;
        return { name, status: 'compared' } as Result;
      },
      failed
    );
    expect(ran).toEqual(['workbook', 'basic-values', 'folder', 'page', 'formulas']);
    expect(results).toEqual([
      failed('workbook', "ENOENT: no such file or directory, open 'workbook.xlsx'"),
      { name: 'basic-values', status: 'compared' },
      // The error is one line: the stack stays out of the report.
      failed('folder', "EACCES: permission denied, mkdir 'fidelity-report/folder'"),
      failed('page', 'page.goto: net::ERR_CONNECTION_REFUSED'),
      { name: 'formulas', status: 'compared' },
    ]);
  });
});
