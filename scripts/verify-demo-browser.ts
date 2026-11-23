// The browser helpers of `verify-demo.ts`: the engines and options of a run, the stand-in save
// picker, and the waits and reports of a step. Kept apart so the runner stays under the line cap.
// A run takes Chromium by default; `--browser firefox` and `--browser webkit` add an engine.

import { chromium, firefox, webkit, type Page } from '@playwright/test';
import { mkdirSync } from 'node:fs';
import { join } from 'node:path';

const ROOT = join(import.meta.dir, '..');

/** The browser engines of the check, in the order that a run takes them. */
export const ENGINES = { chromium, firefox, webkit };
export type Engine = keyof typeof ENGINES;

/** The engine of a run with no `--browser`, and of the CI job. */
export const DEFAULT_ENGINE: Engine = 'chromium';

/** The cause of a headless Firefox clipboard skip. */
export const PLAIN_PASTE_CAUSE = 'the headless Firefox clipboard drops the private clipboard type';

/** Whether the run shows browser windows (`DEMO_HEADED=1`). */
export const HEADED = process.env.DEMO_HEADED === '1';

/**
 * The steps that an engine does not run with no browser window, each with its cause. A skipped
 * step prints the cause on its result line, and the summary names it. A headed run skips none of
 * them. Remove an entry when its cause is fixed.
 */
export const HEADLESS_SKIPS: Readonly<Partial<Record<Engine, Readonly<Record<string, string>>>>> = {
  // The headless clipboard keeps text, HTML, and images only. A paste gets only `text/plain`,
  // so a cut does not move the cell. A headed run (`DEMO_HEADED=1`) runs the step.
  firefox: { 'cut and paste': PLAIN_PASTE_CAUSE },
};

/** How long one step can run. A step that passes it fails, and its engine stops. */
export const STEP_LIMIT_MS = 180_000;
/** How long the page can take to give its alerts after a step. */
const ALERTS_LIMIT_MS = 5_000;

/** The error of work that passed its time limit. */
export class TimeLimitError extends Error {}

/** The result of `work`, or a `TimeLimitError` that names `what` after `ms` milliseconds. */
export async function withTimeLimit<T>(work: Promise<T>, ms: number, what: string): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  const limit = new Promise<never>((_, reject) => {
    timer = setTimeout(
      () => reject(new TimeLimitError(`${what} passed its time limit of ${ms / 1000} seconds`)),
      ms
    );
  });
  try {
    return await Promise.race([work, limit]);
  } finally {
    clearTimeout(timer);
  }
}

/** The result of one engine: whether every step that ran passed, and what it did not prove. */
export interface EngineResult {
  readonly passed: boolean;
  /** The steps that `HEADLESS_SKIPS` skipped. */
  readonly skipped: number;
  /** Why the engine did not run every step, for the summary. */
  readonly stopped?: string;
  /** The first step that failed: fix it first, because later failures can follow from it. */
  readonly firstFailed?: string;
}

/** The result line of a step that `HEADLESS_SKIPS` skips. */
export const skipLine = (engine: Engine, name: string, issue: string): string =>
  `${name}: SKIPPED in ${engine}, ${issue}`;

/**
 * Prints each skip of an engine and counts it, so the summary counts every `SKIPPED` line. A
 * skip is a step that the engine could run but did not, and it names its issue.
 */
export function skipCounter(engine: Engine, log: (line: string) => void = console.log) {
  let count = 0;
  return {
    skip(name: string, issue: string): void {
      log(skipLine(engine, name, issue));
      count += 1;
    },
    get count(): number {
      return count;
    },
  };
}

/** The stop reason of an engine for its summary, or nothing when every step could run. */
export const stopReason = (at: string | null, why: string): { readonly stopped?: string } =>
  at === null ? {} : { stopped: `stopped at "${at}": ${why}` };

/**
 * The result line of a step that ran, and whether the step passes. Its own problem, an alert on
 * the page, or a console error fails it, and the line then says `FAILED`.
 */
export function stepResult(
  name: string,
  problem: string | null,
  shownAlerts: readonly string[],
  consoleErrors: number,
  detail: string
): { readonly line: string; readonly passed: boolean } {
  const why =
    problem ??
    (shownAlerts.length ? `alert "${shownAlerts.join('; ')}"` : null) ??
    (consoleErrors ? `${consoleErrors} console or page errors` : null);
  return {
    line: `${name}: ${why ? `FAILED, ${why}` : detail}, console errors ${consoleErrors}`,
    passed: why === null,
  };
}

/** Why a skip fails a run in CI. */
export const CI_SKIP_REASON =
  'a skipped step fails a run in CI, so each engine must run every step';

/**
 * Whether the skips of an engine fail the run. In CI each engine must run every step: a job that
 * loses its browser window cannot stay green with less coverage.
 */
const skipFailsRun = (result: EngineResult, ci: boolean): boolean => ci && result.skipped > 0;

/**
 * The summary of an engine. It names the issue of each skip, so no skip stays hidden. When the
 * skips fail the run, the engine does not read as passed, and a second line gives the reason.
 */
export function engineSummary(
  engine: Engine,
  result: EngineResult,
  ci = Boolean(process.env.CI)
): string {
  const issues = [...new Set(Object.values(HEADLESS_SKIPS[engine] ?? {}))].join(', ');
  const skips = result.skipped ? ` (${issues})` : '';
  const stopped = result.stopped ? `, ${result.stopped}` : '';
  const first = result.firstFailed ? `, first failed step "${result.firstFailed}"` : '';
  const skipFails = skipFailsRun(result, ci);
  const state = result.passed && !skipFails ? 'passed' : 'FAILED';
  const line = `${engine}: ${state}, ${result.skipped} skipped${skips}${first}${stopped}`;
  return skipFails ? `${line}\n${engine}: FAILED, ${CI_SKIP_REASON}` : line;
}

/** Whether a run passes: each engine passed, and in CI no engine skipped a step. */
export const runPasses = (
  results: readonly (readonly [Engine, EngineResult])[],
  ci = Boolean(process.env.CI)
): boolean => results.every(([, r]) => r.passed && !skipFailsRun(r, ci));

/**
 * Runs each engine, one after another. An engine that throws gets a failed result that says why,
 * and the engines after it still run.
 */
export async function runEngines(
  engines: readonly Engine[],
  run: (engine: Engine) => Promise<EngineResult>
): Promise<[Engine, EngineResult][]> {
  const results: [Engine, EngineResult][] = [];
  for (const engine of engines) {
    try {
      results.push([engine, await run(engine)]);
    } catch (cause) {
      const stopped = `did not finish: ${problemOf(cause)}`;
      console.log(`${engine}: FAILED, ${stopped}`);
      results.push([engine, { passed: false, skipped: 0, stopped }]);
    }
  }
  return results;
}

/** The issue of the step `name` for `engine` in a run with no window, or `undefined`. */
export function skipIssue(engine: Engine, name: string, headed = HEADED): string | undefined {
  if (headed) return undefined;
  const steps = Object.hasOwn(HEADLESS_SKIPS, engine) ? HEADLESS_SKIPS[engine] : undefined;
  return steps && Object.hasOwn(steps, name) ? steps[name] : undefined;
}

/**
 * The whole message of an error from a browser launch. Playwright can start the message with a
 * line break and put the cause in a box on the lines after it, so the first line can be empty.
 */
export const launchProblem = (cause: unknown): string =>
  (cause instanceof Error ? cause.message : String(cause)).trim();

/**
 * What to do after an engine does not launch. A missing executable needs an install, and missing
 * system libraries on Linux need `install-deps`. For an installed engine that does not start,
 * report it, and run the default engine, which the CI job runs.
 */
export function launchRemedy(engine: Engine, problem: string): string {
  if (/executable doesn't exist/i.test(problem))
    return `install the engine with: bunx playwright install ${engine}`;
  if (/missing (system )?dependencies|cannot open shared object file/i.test(problem))
    return `install the system libraries of the engine with: bunx playwright install-deps ${engine}`;
  return `report that ${engine} did not start, and run the default engine (${DEFAULT_ENGINE}), which the CI job demo runs`;
}

/** The steps that `HEADLESS_SKIPS` lists for `engine` and that `steps` of a full run does not name. */
export const unknownSkips = (engine: Engine, steps: ReadonlySet<string>): string[] =>
  Object.keys(HEADLESS_SKIPS[engine] ?? {}).filter((name) => !steps.has(name));

/**
 * The options of the command line: `--browser <engine>`, once for each engine to run, and
 * Chromium alone when there is none.
 */
export function optionsOf(args: readonly string[]): { engines: Engine[] } {
  const engines: Engine[] = [];
  for (let i = 0; i < args.length; i++) {
    const arg = args[i]!;
    if (arg === '--browser' || arg.startsWith('--browser=')) {
      const name = arg === '--browser' ? args[++i] : arg.slice('--browser='.length);
      if (!name || !Object.hasOwn(ENGINES, name))
        throw new Error(`--browser takes ${Object.keys(ENGINES).join(', ')}, not "${name ?? ''}"`);
      if (!engines.includes(name as Engine)) engines.push(name as Engine);
    } else throw new Error(`unknown option "${arg}"`);
  }
  return { engines: engines.length ? engines : [DEFAULT_ENGINE] };
}

/** The options of `verify-demo.ts`, for `--help` and a usage error. */
export const USAGE = [
  'usage: bun scripts/verify-demo.ts [--browser <engine>]...',
  `  --browser <engine>  run one engine (${Object.keys(ENGINES).join(', ')}); repeat for more.`,
  `                      With no --browser, ${DEFAULT_ENGINE} runs alone.`,
  '  --help              print this text',
  `A step that runs longer than ${STEP_LIMIT_MS / 1000} seconds fails, and its engine stops.`,
  'environment:',
  `  CI                  when set, ${CI_SKIP_REASON}.`,
  '  DEMO_HEADED=1       show a browser window. Firefox needs it for the cut and paste step.',
].join('\n');

/**
 * The options of the command line, or exits: `--help` prints the usage, and a usage error
 * prints the error and the usage with exit code 2.
 */
export function optionsOrExit(args: readonly string[]): ReturnType<typeof optionsOf> {
  if (args.includes('--help')) {
    console.log(USAGE);
    process.exit(0);
  }
  try {
    return optionsOf(args);
  } catch (cause) {
    console.error(`${(cause as Error).message}\n${USAGE}`);
    process.exit(2);
  }
}

/** Writes a screenshot of the page for a failed step, to `screenshots/demo-<engine>-<step>.png`. */
export async function screenshotOf(page: Page, step: string): Promise<void> {
  const dir = join(ROOT, 'screenshots');
  const file = join(dir, `demo-${step.replace(/[^a-z0-9]+/gi, '-')}.png`);
  try {
    mkdirSync(dir, { recursive: true });
    await page.screenshot({ path: file, fullPage: true });
    console.log(`  screenshot: ${file}`);
  } catch (cause) {
    console.log(`  no screenshot: ${String(cause)}`);
  }
}

/** A stand-in for `showSaveFilePicker` that keeps the written bytes in `window.__saved`. */
export function stubSavePicker(): void {
  const w = window as unknown as {
    __saved: number[][];
    __cancelNext?: boolean;
    showSaveFilePicker: unknown;
  };
  w.__saved = [];
  // `__cancelNext` makes the next pick fail as a cancelled picker does.
  w.showSaveFilePicker = async () => {
    if (w.__cancelNext) {
      w.__cancelNext = false;
      throw new DOMException('The user cancelled the picker.', 'AbortError');
    }
    return {
      createWritable: async () => {
        const chunks: Blob[] = [];
        return {
          write: async (data: Blob) => void chunks.push(data),
          close: async () =>
            void w.__saved.push([...new Uint8Array(await new Blob(chunks).arrayBuffer())]),
        };
      },
    };
  };
}

/** Waits up to 60 seconds for the grid to show the sheet `name`. False on a timeout. */
export async function showsSheet(page: Page, name: string): Promise<boolean> {
  try {
    await page.waitForFunction(
      (n) => document.querySelector('[role=grid]')?.getAttribute('aria-label') === `Worksheet ${n}`,
      name,
      { timeout: 60_000 }
    );
    return true;
  } catch {
    return false;
  }
}

/** The text of every alert on the page: the demo's own and the editor's. */
export async function alerts(page: Page): Promise<string[]> {
  return page.locator('[role=alert]').allInnerTexts();
}

/**
 * The alerts on the page, and the problem when the page cannot answer because it closed or
 * crashed. A step then fails with that problem, in place of an error that ends the run.
 */
export async function pageAlerts(
  page: Pick<Page, 'locator'>
): Promise<{ readonly shown: string[]; readonly lost: string | null }> {
  try {
    const shown = await withTimeLimit(alerts(page as Page), ALERTS_LIMIT_MS, 'the page');
    return { shown, lost: null };
  } catch (cause) {
    return { shown: [], lost: `the page is lost: ${problemOf(cause)}` };
  }
}

/** How long a step waits for the result of a save or a headings switch. */
export const WAIT_MS = 10_000;

/** The alerts on the page, as a note for an error message, or nothing. */
export async function alertNote(page: Page): Promise<string> {
  const shown = await alerts(page);
  return shown.length ? `; the page shows "${shown.join('; ')}"` : '';
}

/** The first line of the message of a thrown value. */
export const problemOf = (cause: unknown): string => {
  const lines = (cause instanceof Error ? cause.message : String(cause)).split('\n');
  // A Playwright timeout names the locator that it waited for on a later line of its call log.
  const waited = lines.find((l) => /waiting for/.test(l))?.trim();
  return waited ? `${lines[0]} (${waited})` : lines[0]!;
};

/** Selects `ref` with the name box, and returns the text of the formula bar for it. */
export async function goTo(page: Page, ref: string): Promise<string> {
  const nameBox = page.getByRole('textbox', { name: 'Cell reference' });
  await nameBox.click();
  await nameBox.fill(ref);
  await page.keyboard.press('Enter');
  const at = await nameBox.inputValue();
  if (at !== ref) throw new Error(`the name box went to ${at}, not ${ref}`);
  return page.getByRole('textbox', { name: 'Cell content' }).inputValue();
}
