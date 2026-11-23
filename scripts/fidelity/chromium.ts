// The tools of `bun run fidelity` (`run.ts`): a missing Poppler or Chromium, a harness server or
// a browser that does not start, and an error that stops the harness each print one line, never
// a stack trace.
//
// Same role as `startHarnessServer` and `withHarnessPage` of
// pptx-editor/scripts/fidelity/harness-page.ts at 1943d2c9 (EigenPal). As there, a browser close
// that takes longer than `CLOSE_TIMEOUT_MS` ends with a kill of the browser's process group, else
// of its pid, which CDP gives after the launch (`chromiumPid`). `run.ts` takes the server options
// of `startHarnessServer`. The `fidelity: the harness stopped: <message>` line and its exit code
// 1 follow `main` of pptx-editor/scripts/fidelity/run.ts at 1943d2c9.
//
// Differences:
// - A tool that is missing or does not start exits with code 2 (`TOOL_EXIT`), as `run.ts`
//   records.
// - The harness opens one browser and a new page for each fixture, and it has no page reset. So
//   `RESET_TIMEOUT_MS` and the browser restart of the source have no counterpart here.
// - `startServer` of `run.ts` keeps the default Vite cache, where the source gives each server a
//   cache folder of its own (`cacheDir`). The source runs servers in parallel shards; here the
//   run lock (`run-lock.ts`) allows one run in a checkout.
// - `startServer` keeps the dependency discovery of Vite, where the source turns it off
//   (`optimizeDeps.noDiscovery`). The harness page here loads the engine sources, which import
//   npm packages, so the server must discover them.
//
// Compared with 1943d2c9; later commits 67ae7309 and 06534363 checked: 67ae7309 does not change
// this file, and the font cache of 06534363 is not ported. That cache hashes each font file once
// a process, and one page of a lane keeps the fonts of all its fixture decks. Here each fixture
// gets a new page, and each workbook serves only the reference faces of the families that it
// names (`referenceFaces`), so no page loads the fonts of another fixture. A shared page would
// let the fonts of one fixture reach the next.

/** The line that `bun run fidelity` prints when Playwright has no Chromium. */
export const CHROMIUM_MISSING =
  'Chromium is missing: install it with `bunx playwright install chromium`.';

/** The line that `bun run fidelity` prints when Poppler (`pdftoppm`) is missing. */
export const POPPLER_MISSING =
  'Poppler is missing: install it with `brew install poppler` or `sudo apt-get install poppler-utils`.';

/** The exit code of `bun run fidelity` when a tool is missing or does not start. */
export const TOOL_EXIT = 2;

/**
 * A browser close that takes longer than this ends with a kill. The value and the reason are
 * those of the source: in CI, a browser close never ended, because Playwright waits for the
 * output pipes of the browser, and they stayed open.
 */
export const CLOSE_TIMEOUT_MS = 10_000;

/** The first line of the message of `error`. */
export function firstLine(error: unknown): string {
  const text = error instanceof Error ? error.message : String(error);
  return text.split('\n')[0] ?? '';
}

/**
 * Run `body` for each fixture in turn. An error of one fixture, such as a workbook that cannot be
 * read, a folder that cannot be made, or a page that fails, becomes the result of that fixture
 * (`failed`, with the first line of the error). The next fixture still runs.
 */
export async function eachFixture<R>(
  names: readonly string[],
  body: (name: string) => Promise<R>,
  failed: (name: string, error: string) => R
): Promise<R[]> {
  const out: R[] = [];
  for (const name of names) {
    try {
      out.push(await body(name));
    } catch (error) {
      out.push(failed(name, firstLine(error)));
    }
  }
  return out;
}

/** The line that `bun run fidelity` prints when an error stops the harness, as in the source. */
export function harnessStopped(error: unknown): string {
  return `fidelity: the harness stopped: ${firstLine(error)}`;
}

/** Whether `error` is the launch error of a browser that Playwright has not installed. */
function isMissing(error: unknown): boolean {
  const text = error instanceof Error ? error.message : String(error);
  return text.includes("Executable doesn't exist");
}

/**
 * The browser that `launch` starts, or the line to print: `CHROMIUM_MISSING` when Chromium is not
 * installed, or the first line of the launch error, such as missing host libraries.
 */
export async function launchBrowser<B>(
  launch: () => Promise<B>
): Promise<{ readonly browser: B } | { readonly error: string }> {
  try {
    return { browser: await launch() };
  } catch (error) {
    if (isMissing(error)) return { error: CHROMIUM_MISSING };
    return { error: firstLine(error) };
  }
}

/** A server or a browser that the harness closes. */
interface Closable {
  close(): Promise<unknown>;
}

/** The Chromium calls that `chromiumPid` uses. */
interface CdpBrowser {
  newBrowserCDPSession(): Promise<{
    send(method: 'SystemInfo.getProcessInfo'): Promise<unknown>;
    detach(): Promise<void>;
  }>;
}

/** The pid of the browser process of `browser`, as CDP gives it, as in the source. */
export async function chromiumPid(browser: CdpBrowser): Promise<number | undefined> {
  const cdp = await browser.newBrowserCDPSession();
  const info = (await cdp.send('SystemInfo.getProcessInfo')) as {
    processInfo: { type: string; id: number }[];
  };
  await cdp.detach();
  return info.processInfo.find((p) => p.type === 'browser')?.id;
}

/** How `withHarness` closes a browser that does not answer. Tests replace the parts. */
export interface BrowserClose {
  /** The time limit of the close. The default is `CLOSE_TIMEOUT_MS`. */
  readonly ms?: number;
  /** Send a signal to a process or, for a negative pid, a process group. */
  readonly kill?: (pid: number, signal: 'SIGKILL') => void;
  /** Print a warning line. */
  readonly warn?: (line: string) => void;
}

/** Whether `work` resolves before `ms` passes: false when it rejects or takes longer. */
async function resolvesWithin(work: Promise<unknown>, ms: number): Promise<boolean> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  const limit = new Promise<false>((done) => {
    timer = setTimeout(() => done(false), ms);
  });
  const resolved = work.then(
    () => true,
    () => false
  );
  try {
    return await Promise.race([resolved, limit]);
  } finally {
    clearTimeout(timer);
  }
}

/**
 * Close `browser` within the time limit. Playwright's close never ends for a browser that does
 * not answer, and that browser keeps running. So it is then killed: its process group
 * (Playwright starts the browser as a group leader), else its pid, as in the source. A close
 * that fails with an error ends with the same kill, as in the source.
 */
async function closeBrowser(
  browser: Closable,
  pid: number | undefined,
  options: BrowserClose
): Promise<void> {
  let closed: Promise<unknown>;
  try {
    closed = browser.close();
  } catch (error) {
    closed = Promise.reject(error);
  }
  if (await resolvesWithin(closed, options.ms ?? CLOSE_TIMEOUT_MS)) return;
  if (pid === undefined) {
    (options.warn ?? console.warn)(
      'fidelity: a browser that does not answer keeps running until the run ends: CDP gave no pid to stop it'
    );
    return;
  }
  const kill = options.kill ?? ((target, signal) => process.kill(target, signal));
  for (const target of [-pid, pid]) {
    try {
      kill(target, 'SIGKILL');
      return;
    } catch {
      // No such group or process: try the pid, or it exited in the meantime.
    }
  }
}

/** Close `item`; an error of the close does not hide the result of the run. */
async function closeQuietly(item: Closable): Promise<void> {
  try {
    await item.close();
  } catch {
    // The run already has its result or its error.
  }
}

/**
 * Start the server (`start`) and the browser (`launch`), run `body` with them, and close both,
 * also when `body` throws. `pidOf` gives the pid of the browser process after the launch, for a
 * close that does not end in time (`BrowserClose`). A server or a browser that does not start
 * gives the line to print and `TOOL_EXIT`. An error of `body` is thrown on, after both close.
 */
export async function withHarness<S extends Closable, B extends Closable, T>(
  start: () => Promise<S>,
  launch: () => Promise<B>,
  pidOf: (browser: B) => Promise<number | undefined>,
  body: (server: S, browser: B) => Promise<T>,
  close: BrowserClose = {}
): Promise<{ readonly result: T } | { readonly error: string; readonly code: number }> {
  let server: S;
  try {
    server = await start();
  } catch (error) {
    return { error: `the harness server did not start: ${firstLine(error)}`, code: TOOL_EXIT };
  }
  try {
    const launched = await launchBrowser(launch);
    if ('error' in launched) return { error: launched.error, code: TOOL_EXIT };
    // Without a pid, a browser that does not answer cannot be killed (`closeBrowser` warns).
    const pid = await pidOf(launched.browser).catch(() => undefined);
    try {
      return { result: await body(server, launched.browser) };
    } finally {
      await closeBrowser(launched.browser, pid, close);
    }
  } finally {
    await closeQuietly(server);
  }
}

/**
 * The Poppler version that `version` reads, or the line to print and the exit code when Poppler
 * is missing. `version` is `popplerVersion` of `poppler.ts`, replaced in tests.
 */
export function checkPoppler(
  version: () => string
): { readonly poppler: string } | { readonly error: string; readonly code: number } {
  const poppler = version();
  return poppler === 'missing' ? { error: POPPLER_MISSING, code: TOOL_EXIT } : { poppler };
}
