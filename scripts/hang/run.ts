// The browser hang test: no hostile or large workbook hangs the page.
//
//   bun run test:hang:browser
//   HANG_ONLY=sst-50mb,fonts-stalled bun run test:hang:browser
//   HANG_HEADED=1 HANG_ONLY=sst-50mb bun run test:hang:browser
//
// Each case of `cases.ts` runs in its own browser, one browser at a time. The test opens the
// workbook in the demo through its file input, with the demo's default options, unless the case
// sets `headingsOff`. It then moves in the grid, scrolls, saves, types in a cell, and opens a small
// workbook with the network fault lifted. The page must answer after each step. A watchdog outside
// the page (`watchdog.ts`) stops the browser when it passes the memory cap or the time of the case.
// A read of the watchdog that comes more than 5 s late (a sleeping machine or a stalled test
// process) does not count its delay as time of the case.
//
// A race case replaces the moves, the save, and the typing:
// - `headings-switch` and `save` open a small workbook during the long work. That workbook must
//   stay open, and the demo must say that another workbook opened.
// - `headings-second-press` presses the headings option and Save again during a switch. The demo
//   must refuse each press, write no file, and end the switch.
//
// Verdicts, worst first, as in PPTX Editor's hang check. Each verdict but `ok` and `SLOW` fails
// the run:
// - CRASH: the page or the browser closed, or the watchdog stopped the browser at the memory cap.
// - HANG: the workbook did not show and was not refused within the budget, the page stopped
//   answering, or the case passed its time.
// - BLOCK: the main thread ran no timer for longer than the block limit.
// - FETCH: the page asked for a URL outside the demo's origin.
// - ERROR: an uncaught page error, the wrong outcome, a failed save, a lost race, a picked file
//   with no write, or a page that opens no workbook after the case.
// - SLOW: the blocks of 100 ms or more add up to more than the mark for repeated blocks. The run
//   prints each SLOW case and lists it in the job summary, but the mark never fails the run.
//
// The report is `test-results/hang/report.json`, written after each case. A case that is not `ok`
// leaves a screenshot there when its page is still open. Each run starts with an empty directory. A rerun
// with `HANG_ONLY` builds the same workbook again.
//
// Same method as pptx-editor/e2e/hang.stress.ts at f448820c (EigenPal), and the same SLOW mark:
// "It never fails". pptx-editor/scripts/stress/hang-check.ts at c47f5cea also has a SLOW mark.
// Differences, each for a reason:
// - It runs without the Playwright test runner, so it can start one browser for each case and
//   read the process tree of that browser.
// - It runs Chromium only. The browser matrix of the repository is a separate check.
// - It has no FREEZE verdict. A block over the block limit is BLOCK, and it always fails.
// - Its SLOW mark is the sum of the blocks of 100 ms or more, not the time to open. The open
//   budget is a HANG. hang-check.ts counts SLOW in its exit code, but no CI job runs it. This run
//   is a CI gate, so it keeps SLOW out of the exit code, as hang.stress.ts does.
// - Its limits are constants and not environment variables, so a run cannot loosen them.
// - Its files are in `scripts/`, with the other browser checks of this repository.
// - It has no case for a font body that stops or trickles. The unit tests of `stallLimitedFetch`
//   cover these faults with exact timing.

import {
  chromium,
  type Browser,
  type BrowserServer,
  type Locator,
  type Page,
} from '@playwright/test';
import { appendFileSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, relative } from 'node:path';
import { createServer } from 'vite';
import {
  HANG_CASES,
  HOSTILE_SHEET,
  RECOVERY_SHEET,
  RACE_SHEET,
  recoveryWorkbook,
  type HangCase,
} from './cases.ts';
import { LAST_CELL, SAVE_KEEPS_CONTENT } from './first-cells.ts';
import { readZip } from '../../packages/ooxml-core/src/zip.ts';
import { HEADINGS_FROM_FILE } from '../../examples/vite/src/headings.ts';
import { SAVE_READY } from '../../examples/vite/src/save.ts';
import { fails, showsException, verdictOf, type Verdict } from './verdict.ts';
import { killTree, watchBrowser, type Watch } from './watchdog.ts';

const ROOT = join(import.meta.dir, '../..');
const DEMO = join(ROOT, 'examples/vite');
const REPORT_DIR = join(ROOT, 'test-results/hang');

/** The time for a workbook to show or to be refused, as in PPTX Editor's hang check. */
const OPEN_BUDGET_MS = 20_000;
/** The longest time that the main thread may run no timer, as in PPTX Editor's hang check. */
const BLOCK_LIMIT_MS = 5_000;
/** The SLOW mark: the sum of the blocks of 100 ms or more in one case. It never fails the run. */
const REPEATED_BLOCK_LIMIT_MS = 10_000;
/** The time for the demo to start, before the case begins. */
const START_BUDGET_MS = 60_000;
/** The time for one key press, one save, or one answer of the page. */
const STEP_BUDGET_MS = 20_000;
/** The time after which the watchdog stops the browser of a case. */
const CASE_DEADLINE_MS = 180_000;
/** The time for the work of a race case to end after the other workbook opened. */
const RACE_SETTLE_MS = 8_000;
/** The words of the demo when another workbook opened during long work. */
const RACE_MESSAGE = 'another workbook opened';

interface Row {
  id: string;
  covers: string;
  verdict: Verdict;
  why?: string;
  outcome: string;
  openMs: number;
  /** The longest main-thread block, or null when the case failed before the run read it. */
  longestBlockMs: number | null;
  /** The sum of the main-thread blocks of 100 ms or more. */
  repeatedBlockMs: number;
  peakMiB: number;
  /** The JS heap of the page after the open, in MB, from the DevTools protocol. It fails no case. */
  heapMB?: number;
  fetched: string[];
  errors: string[];
  screenshot?: string;
  /** The demo itself did not start, so no later case can run. */
  demoDown?: boolean;
}

type Probe = {
  __hang: {
    picked: number;
    saved: number;
    /** The bytes of XML and other parts that the last saved package holds when inflated. */
    savedContent: number;
    repeated: number;
    take(): number;
  };
};

/** A step that failed, with its verdict. */
type Problem = { verdict: Verdict; why: string };

/**
 * Records the longest gap of a 25 ms timer, and takes the bytes of a save without keeping them.
 * `repeated` is the sum of the gaps of 100 ms or more since the page started. `take` returns the
 * longest gap and starts a new measure. It counts the gap that is still open,
 * because a read that waited for a block runs before the timer does. The picker refuses a call
 * with no recent user action, as the browser's picker does. It counts each picked file, because
 * a picked file that gets no write can lose its earlier content.
 */
function probe(): void {
  /** The inflated size of the parts of the zip `file`, from its central directory, or 0. */
  const inflatedSize = async (file: Blob): Promise<number> => {
    const tailStart = Math.max(0, file.size - 66_000);
    const tail = new DataView(await file.slice(tailStart).arrayBuffer());
    let end = tail.byteLength - 22;
    while (end >= 0 && tail.getUint32(end, true) !== 0x06054b50) end -= 1;
    if (end < 0) return 0;
    const count = tail.getUint16(end + 10, true);
    const size = tail.getUint32(end + 12, true);
    const offset = tail.getUint32(end + 16, true);
    const list = new DataView(await file.slice(offset, offset + size).arrayBuffer());
    let total = 0;
    let at = 0;
    for (let i = 0; i < count && at + 46 <= list.byteLength; i += 1) {
      if (list.getUint32(at, true) !== 0x02014b50) return 0;
      total += list.getUint32(at + 24, true);
      at +=
        46 +
        list.getUint16(at + 28, true) +
        list.getUint16(at + 30, true) +
        list.getUint16(at + 32, true);
    }
    return total;
  };
  let longest = 0;
  let tick = performance.now();
  const hang = {
    picked: 0,
    saved: 0,
    savedContent: 0,
    repeated: 0,
    take: () => {
      const now = performance.now();
      const value = Math.max(longest, now - tick);
      longest = 0;
      tick = now;
      return value;
    },
  };
  (window as unknown as Probe).__hang = hang;
  setInterval(() => {
    const now = performance.now();
    longest = Math.max(longest, now - tick);
    if (now - tick >= 100) hang.repeated += now - tick;
    tick = now;
  }, 25);
  (window as unknown as { showSaveFilePicker: unknown }).showSaveFilePicker = async () => {
    if (!navigator.userActivation.isActive)
      throw new DOMException('The picker needs a recent user action.', 'SecurityError');
    hang.picked += 1;
    return {
      createWritable: async () => {
        const chunks: Blob[] = [];
        return {
          write: async (data: Blob) => void chunks.push(data),
          // A save that writes no valid package is not a save.
          close: async () => {
            const content = await inflatedSize(new Blob(chunks));
            hang.savedContent = content;
            hang.saved += content > 0 ? 1 : 0;
          },
        };
      },
    };
  };
}

/** The longest main-thread block since the last call, or null when the page does not answer. */
const takeBlock = (page: Page) =>
  within(
    page.evaluate(() => (window as unknown as Probe).__hang.take()),
    STEP_BUDGET_MS
  );

/** How a step ended: with its value, with the error of Playwright, or not within its time. */
type Settled<T> = { value: T } | { failed: string } | { late: true };

/** Waits for `work` for `ms`, and keeps its error. */
async function settled<T>(work: Promise<T>, ms: number): Promise<Settled<T>> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  const late = new Promise<Settled<T>>((resolve) => {
    timer = setTimeout(() => resolve({ late: true }), ms);
  });
  try {
    return await Promise.race([
      work.then(
        (value) => ({ value }),
        (cause: unknown) => ({ failed: String(cause) })
      ),
      late,
    ]);
  } finally {
    clearTimeout(timer);
  }
}

/** `work`, or `null` when it fails or does not end within `ms`. */
async function within<T>(work: Promise<T>, ms: number): Promise<T | null> {
  const end = await settled(work, ms);
  return 'value' in end ? end.value : null;
}

/** The sheet that the grid shows, and the text of each alert and notice on the page. */
const pageState = (page: Page) =>
  page.evaluate(() => {
    const texts = (selector: string) =>
      [...document.querySelectorAll(selector)]
        .map((a) => a.textContent ?? '')
        .filter((t) => t.trim().length > 0);
    return {
      grid: document.querySelector('[role=grid]')?.getAttribute('aria-label') ?? null,
      alerts: texts('[role=alert]'),
      notices: texts('.demo-message'),
      picked: (window as unknown as Probe).__hang.picked,
      saved: (window as unknown as Probe).__hang.saved,
      savedContent: (window as unknown as Probe).__hang.savedContent,
    };
  });

/** Waits for the sheet `name` or for a refusal. Null when the page shows neither in `budget`. */
async function settle(
  page: Page,
  name: string | null,
  budget: number,
  refusalCounts = true
): Promise<{ outcome: 'open' | 'refuse'; detail: string } | null> {
  const end = Date.now() + budget;
  while (Date.now() < end) {
    const state = await within(pageState(page), Math.max(1, end - Date.now()));
    if (state) {
      const refused = [...state.alerts, ...state.notices].find((a) => a.includes('Could not open'));
      if (refused && refusalCounts) return { outcome: 'refuse', detail: refused };
      if (name ? state.grid === `Worksheet ${name}` : state.grid?.startsWith('Worksheet '))
        return { outcome: 'open', detail: state.grid ?? '' };
    }
    await Bun.sleep(100);
  }
  return null;
}

/** Whether the page answers within the step budget. */
const answers = async (page: Page): Promise<boolean> =>
  (await within(
    page.evaluate(() => 1),
    STEP_BUDGET_MS
  )) === 1;

/** Gives the focus to the grid. */
const focusGrid = (page: Page) =>
  within(
    page.evaluate(() => (document.querySelector('[role=grid]') as HTMLElement | null)?.focus()),
    STEP_BUDGET_MS
  );

/** Presses `key`. The reason when the page does not answer after it, or null. */
async function press(page: Page, key: string): Promise<string | null> {
  const sent = await settled(page.keyboard.press(key), STEP_BUDGET_MS);
  if (!('value' in sent)) return `the key ${key} did not reach the page`;
  return (await answers(page)) ? null : `the page did not answer after ${key}`;
}

/** Scrolls the grid with the wheel, with the pointer over the grid. */
async function scroll(page: Page): Promise<string | null> {
  const box = await within(page.locator('[role=grid]').first().boundingBox(), STEP_BUDGET_MS);
  if (box)
    await within(page.mouse.move(box.x + box.width / 2, box.y + box.height / 2), STEP_BUDGET_MS);
  const sent = await settled(page.mouse.wheel(0, 4000), STEP_BUDGET_MS);
  if (!('value' in sent)) return 'the wheel did not reach the page';
  return (await answers(page)) ? null : 'the page did not answer after a scroll';
}

/** Clicks `target`. The problem when the click fails or does not end, or null. */
async function click(target: Locator, name: string): Promise<Problem | null> {
  const end = await settled(target.click({ timeout: STEP_BUDGET_MS }), STEP_BUDGET_MS + 1000);
  if ('late' in end) return { verdict: 'HANG', why: `${name} did not answer` };
  if ('failed' in end) return { verdict: 'ERROR', why: `${name} did not work: ${end.failed}` };
  return null;
}

/** The Save button of the top bar. */
const saveButton = (page: Page) =>
  page.locator('.demo-topbar').getByRole('button', { name: 'Save', exact: true });

/**
 * Saves from the top bar. When the browser refused its file dialog because the bytes took long,
 * the demo keeps them, and a second press writes them. The problem, or null.
 */
async function save(page: Page): Promise<Problem | null> {
  const late = await click(saveButton(page), 'the Save button');
  if (late) return late;
  const end = Date.now() + STEP_BUDGET_MS;
  let pressedAgain = false;
  while (Date.now() < end) {
    const state = await within(pageState(page), Math.max(1, end - Date.now()));
    if (state && state.saved > 0) return null;
    if (state && state.alerts.length > 0)
      return { verdict: 'ERROR', why: `the save failed: ${state.alerts.join('; ')}` };
    if (!pressedAgain && state?.notices.includes(SAVE_READY)) {
      pressedAgain = true;
      const again = await click(saveButton(page), 'the Save button');
      if (again) return again;
    }
    await Bun.sleep(100);
  }
  return { verdict: 'HANG', why: 'the save did not end' };
}

/** Whether the live region of the grid says `text` within the step budget. */
async function announces(page: Page, text: string): Promise<boolean> {
  const end = Date.now() + STEP_BUDGET_MS;
  while (Date.now() < end) {
    const said = await within(
      page.evaluate(() => document.querySelector('.xlsx-live')?.textContent ?? ''),
      Math.max(1, end - Date.now())
    );
    if (said?.includes(text)) return true;
    await Bun.sleep(100);
  }
  return false;
}

/** The inflated size of the parts of `bytes`. */
function inflatedBytes(bytes: Uint8Array | undefined): number {
  const read = bytes ? readZip(bytes) : null;
  if (!read?.ok) return 0;
  let total = 0;
  for (const data of read.entries.values()) total += data.length;
  return total;
}

/** Moves in the grid, scrolls, saves, and types in a cell. The first problem, or null. */
async function exercise(page: Page, testCase: HangCase): Promise<Problem | null> {
  await focusGrid(page);
  const lastCell = LAST_CELL.get(testCase.id);
  for (const key of ['ArrowRight', 'PageDown', 'ControlOrMeta+End', 'ControlOrMeta+Home']) {
    const stuck = await press(page, key);
    if (stuck) return { verdict: 'HANG', why: stuck };
    if (key === 'ControlOrMeta+End' && lastCell && !(await announces(page, lastCell)))
      return { verdict: 'ERROR', why: `the last cell does not show ${lastCell}` };
  }
  if (testCase.firstCell && !(await announces(page, testCase.firstCell)))
    return { verdict: 'ERROR', why: `cell A1 does not show ${testCase.firstCell}` };
  const stuck = await scroll(page);
  if (stuck) return { verdict: 'HANG', why: stuck };
  const saved = await save(page);
  if (saved) return saved;
  if (SAVE_KEEPS_CONTENT.has(testCase.id)) {
    const kept = (await within(pageState(page), STEP_BUDGET_MS))?.savedContent ?? 0;
    const opened = inflatedBytes(testCase.bytes?.());
    if (kept * 2 < opened)
      return { verdict: 'ERROR', why: `the save kept ${kept} of ${opened} bytes of content` };
  }
  // The edit comes last: a sheet that refuses edits shows an alert, which the save would read.
  await focusGrid(page);
  for (const key of ['ControlOrMeta+Home', 'x', 'Enter', 'Escape']) {
    const late = await press(page, key);
    if (late) return { verdict: 'HANG', why: late };
  }
  return null;
}

/** The option of the demo that shows the headings of every sheet. */
const headingsPill = (page: Page) =>
  page.locator('label.demo-pill', { hasText: 'Show all headings' });

/**
 * Starts the long work of a race case, and opens the small workbook while it runs. The work must
 * say that another workbook opened, which proves the overlap, and the small workbook must still
 * show after the work ended. A save must write nothing. The problem, or null.
 */
async function raceOpen(
  page: Page,
  testCase: HangCase,
  input: Locator,
  file: string
): Promise<Problem | null> {
  const start = testCase.race === 'headings-switch' ? headingsPill(page) : saveButton(page);
  const late = await click(start, `the control of the ${testCase.race}`);
  if (late) return late;
  writeFileSync(file, recoveryWorkbook(RACE_SHEET));
  await within(
    input.setInputFiles(file).then(() => true),
    STEP_BUDGET_MS
  );
  const wanted = `Worksheet ${RACE_SHEET}`;
  const alerts = new Set<string>();
  const end = Date.now() + OPEN_BUDGET_MS;
  let openedAt: number | null = null;
  let last: Awaited<ReturnType<typeof pageState>> | null = null;
  while (Date.now() < end) {
    last = (await within(pageState(page), STEP_BUDGET_MS)) ?? last;
    for (const alert of last?.alerts ?? []) alerts.add(alert);
    if (openedAt === null && last?.grid === wanted) openedAt = Date.now();
    if (openedAt !== null && Date.now() - openedAt >= RACE_SETTLE_MS) break;
    await Bun.sleep(50);
  }
  if (openedAt === null)
    return { verdict: 'ERROR', why: `no workbook opened during the ${testCase.race}` };
  if (last?.grid !== wanted)
    return {
      verdict: 'ERROR',
      why: `the ${testCase.race} replaced the workbook that opened during it: ${last?.grid}`,
    };
  if (![...alerts].some((alert) => alert.includes(RACE_MESSAGE)))
    return { verdict: 'ERROR', why: `the ${testCase.race} did not say "${RACE_MESSAGE}"` };
  if (testCase.race === 'save' && (last.picked > 0 || last.saved > 0))
    return { verdict: 'ERROR', why: 'the save wrote a workbook that another workbook replaced' };
  return null;
}

/** Clicks `target`, and waits for an alert that holds `text`. */
async function refusedPress(page: Page, target: Locator, name: string, text: string) {
  const late = await click(target, name);
  if (late) return late;
  const end = Date.now() + STEP_BUDGET_MS;
  while (Date.now() < end) {
    const state = await within(pageState(page), STEP_BUDGET_MS);
    if (state?.alerts.some((alert) => alert.includes(text))) return null;
    await Bun.sleep(25);
  }
  return { verdict: 'ERROR', why: `${name} during the switch did not say "${text}"` } as Problem;
}

/** Presses the headings option again and Save during one switch: the switch still ends well. */
async function raceSecondPress(page: Page): Promise<Problem | null> {
  const first = await click(headingsPill(page), 'the headings option');
  if (first) return first;
  const second = await refusedPress(
    page,
    headingsPill(page),
    'the second press',
    'a switch is running'
  );
  if (second) return second;
  const save = await refusedPress(page, saveButton(page), 'Save', 'the headings are switching');
  if (save) return save;
  const end = Date.now() + OPEN_BUDGET_MS;
  while (Date.now() < end) {
    const state = await within(pageState(page), STEP_BUDGET_MS);
    if (state?.notices.some((n) => n.includes(HEADINGS_FROM_FILE))) {
      if (state.picked > 0) return { verdict: 'ERROR', why: 'a refused Save wrote a file' };
      return null;
    }
    await Bun.sleep(50);
  }
  return { verdict: 'ERROR', why: 'the headings switch did not end after the refused presses' };
}

/** Turns off the demo's headings option, so the editor reads the bytes of the case. */
async function headingsOff(page: Page): Promise<Problem | null> {
  const late = await click(headingsPill(page), 'the headings option');
  if (late) return late;
  const end = Date.now() + STEP_BUDGET_MS;
  while (Date.now() < end) {
    const state = await within(pageState(page), Math.max(1, end - Date.now()));
    if (state?.notices.some((n) => n.includes(HEADINGS_FROM_FILE))) return null;
    await Bun.sleep(50);
  }
  return { verdict: 'ERROR', why: 'the demo did not switch its headings option' };
}

/** Applies the font fault of the case before the first request of the page. */
async function applyNetwork(page: Page, testCase: HangCase): Promise<void> {
  const fonts = /\.(ttf|otf|woff2?)(\?|$)/;
  if (testCase.network === 'fonts-stalled') await page.route(fonts, () => undefined);
  if (testCase.network === 'fonts-partly-stalled')
    await page.route(/Bold[^/]*\.(ttf|otf|woff2?)(\?|$)/, () => undefined);
  if (testCase.network === 'fonts-blocked')
    await page.route(fonts, (route) => route.abort('blockedbyclient'));
}

/**
 * Runs one case in `page`. It fills `row`, and it never throws for a page that stops. It keeps
 * in `shot` the page as the case left it, before the recovery workbook replaces it.
 */
async function runCase(
  page: Page,
  testCase: HangCase,
  dir: string,
  base: string,
  row: Row,
  shot: { image?: Uint8Array }
) {
  const fail = (verdict: Verdict, why: string) => {
    row.verdict = verdict;
    row.why = why;
  };
  await applyNetwork(page, testCase);
  const started = Date.now();
  if ((await within(page.goto(base, { waitUntil: 'commit' }), START_BUDGET_MS)) === null) {
    row.demoDown = true;
    return fail('HANG', 'the demo did not load');
  }
  const first = await settle(page, null, testCase.bytes ? START_BUDGET_MS : OPEN_BUDGET_MS);
  const input = page.locator('input[type=file]').first();
  let result = first;
  if (testCase.bytes) {
    if (first?.outcome !== 'open') {
      row.demoDown = true;
      const why = first ? `: ${first.detail}` : '';
      return fail('ERROR', `the demo did not open its sample${why}`);
    }
    if (testCase.network === 'offline-after-start') await page.context().setOffline(true);
    const file = join(dir, `${testCase.id}.xlsx`);
    writeFileSync(file, testCase.bytes());
    if (testCase.headingsOff) {
      const off = await headingsOff(page);
      if (off) return fail(off.verdict, off.why);
    }
    await takeBlock(page);
    const t0 = Date.now();
    const taken = await within(
      input.setInputFiles(file).then(
        () => 'ok',
        (cause: unknown) => String(cause)
      ),
      STEP_BUDGET_MS
    );
    if (taken === null) return fail('HANG', 'the file input did not take the workbook');
    if (taken !== 'ok') return fail('ERROR', `the file input refused the workbook: ${taken}`);
    result = await settle(page, HOSTILE_SHEET, OPEN_BUDGET_MS);
    row.openMs = Date.now() - t0;
  } else row.openMs = Date.now() - started;
  if (!result) {
    const shown = (await within(pageState(page), STEP_BUDGET_MS))?.grid ?? 'no sheet';
    const wanted = testCase.bytes ? `sheet ${HOSTILE_SHEET}` : 'sheet';
    return fail(
      'HANG',
      `no ${wanted} and no refusal in ${OPEN_BUDGET_MS} ms; the grid shows ${shown}`
    );
  }
  row.outcome = result.outcome;
  const heap = await jsHeapMB(page);
  if (heap !== undefined) row.heapMB = heap;
  if (testCase.expect !== result.outcome)
    return fail('ERROR', `expected ${testCase.expect}, got ${result.outcome}: ${result.detail}`);
  if (result.outcome === 'refuse' && showsException(result.detail))
    return fail('ERROR', `the refusal shows an exception: ${result.detail}`);

  if (testCase.network === 'fonts-partly-stalled') {
    const state = await within(pageState(page), STEP_BUDGET_MS);
    if (!state?.notices.some((n) => n.includes('did not load')))
      return fail('ERROR', 'the demo did not say that some fonts did not load');
  }
  if (result.outcome === 'open') {
    const problem =
      testCase.race === 'headings-second-press'
        ? await raceSecondPress(page)
        : testCase.race
          ? await raceOpen(page, testCase, input, join(dir, 'race.xlsx'))
          : await exercise(page, testCase);
    if (problem) return fail(problem.verdict, problem.why);
    const files = await within(pageState(page), STEP_BUDGET_MS);
    if (files && files.picked !== files.saved)
      return fail('ERROR', `${files.picked} files were picked and ${files.saved} were written`);
  }
  const blocked = await takeBlock(page);
  if (blocked === null) return fail('HANG', 'the page did not answer after the case');
  row.longestBlockMs = Math.round(blocked);

  const image = await within(page.screenshot(), STEP_BUDGET_MS);
  if (image) shot.image = image;

  // With the fault lifted, the same page opens a workbook: nothing stays stuck.
  await within(page.unrouteAll({ behavior: 'ignoreErrors' }), STEP_BUDGET_MS);
  await within(page.context().setOffline(false), STEP_BUDGET_MS);
  const next = join(dir, 'recovery.xlsx');
  writeFileSync(next, recoveryWorkbook());
  await within(
    input.setInputFiles(next).then(() => true),
    STEP_BUDGET_MS
  );
  // The message of a refused case stays on the page, so only the sheet counts here.
  const after = await settle(page, RECOVERY_SHEET, OPEN_BUDGET_MS, false);
  if (after?.outcome !== 'open')
    return fail('ERROR', 'the editor did not open a workbook after this one');
  // The open after the case counts too: a case must not leave the page slow.
  const recovery = await takeBlock(page);
  if (recovery === null) return fail('HANG', 'the page did not answer after the recovery open');
  row.longestBlockMs = Math.max(row.longestBlockMs ?? 0, Math.round(recovery));
  const repeated = await within(
    page.evaluate(() => (window as unknown as Probe).__hang.repeated),
    STEP_BUDGET_MS
  );
  row.repeatedBlockMs = Math.round(repeated ?? 0);
}

/** The JS heap of `page` in MB, as PPTX Editor's hang check reads it, or undefined. */
async function jsHeapMB(page: Page): Promise<number | undefined> {
  const cdp = await within(page.context().newCDPSession(page), STEP_BUDGET_MS);
  if (!cdp) return undefined;
  try {
    await within(cdp.send('Performance.enable'), STEP_BUDGET_MS);
    const read = await within(cdp.send('Performance.getMetrics'), STEP_BUDGET_MS);
    const heap = read?.metrics.find((m) => m.name === 'JSHeapUsedSize')?.value;
    return heap === undefined ? undefined : Math.round(heap / 1e6);
  } finally {
    await within(cdp.detach(), STEP_BUDGET_MS);
  }
}

/** Chromium did not start, so no case of the run can run. */
class NoBrowser extends Error {}

/** The process id of the browser of the running case, for a stop of the whole run. */
let runningBrowser: number | undefined;

/** Starts one browser, runs one case in it, and stops that browser in every path. */
async function inOwnBrowser(testCase: HangCase, dir: string, base: string): Promise<Row> {
  const row: Row = {
    id: testCase.id,
    covers: testCase.covers,
    verdict: 'ok',
    outcome: '',
    openMs: 0,
    longestBlockMs: null,
    repeatedBlockMs: 0,
    peakMiB: 0,
    fetched: [],
    errors: [],
  };
  let server: BrowserServer | undefined;
  let browser: Browser | undefined;
  let page: Page | undefined;
  let watch: Watch | undefined;
  let pid: number | undefined;
  let closed = false;
  const shot: { image?: Uint8Array } = {};
  try {
    server = await chromium
      .launchServer({ headless: process.env.HANG_HEADED !== '1' })
      .catch((cause: unknown) => {
        throw new NoBrowser(String(cause));
      });
    pid = server.process().pid;
    if (pid === undefined) throw new Error('the browser has no process id');
    runningBrowser = pid;
    watch = watchBrowser(pid, CASE_DEADLINE_MS);
    browser = await chromium.connect(server.wsEndpoint());
    browser.on('disconnected', () => (closed = true));
    const context = await browser.newContext({ viewport: { width: 1440, height: 900 } });
    page = await context.newPage();
    await page.addInitScript(probe);
    page.on('crash', () => (closed = true));
    page.on('pageerror', (e) => row.errors.push(e.message));
    page.on('dialog', (d) => void d.accept().catch(() => undefined));
    const origin = new URL(base).origin;
    page.on('request', (r) => {
      const url = r.url();
      if (/^https?:/.test(url) && new URL(url).origin !== origin) row.fetched.push(url);
    });
    page.on('websocket', (socket) => {
      if (new URL(socket.url()).host !== new URL(base).host) row.fetched.push(socket.url());
    });
    await runCase(page, testCase, dir, base, row, shot);
  } catch (cause) {
    if (cause instanceof NoBrowser) throw cause;
    row.verdict = 'ERROR';
    row.why = String(cause);
  } finally {
    watch?.stop();
    row.peakMiB = Math.round((watch?.peakBytes() ?? 0) / 1048576);
    const { verdict, why } = verdictOf({
      stepVerdict: row.verdict,
      stepWhy: row.why,
      tripped: watch?.tripped() ?? null,
      trippedMemory: watch?.trippedMemory() ?? false,
      closed,
      peakMiB: row.peakMiB,
      longestBlockMs: row.longestBlockMs ?? 0,
      blockLimitMs: BLOCK_LIMIT_MS,
      repeatedBlockMs: row.repeatedBlockMs,
      repeatedBlockLimitMs: REPEATED_BLOCK_LIMIT_MS,
      fetched: row.fetched,
      errors: row.errors,
    });
    row.verdict = verdict;
    if (why) row.why = why;
    else delete row.why;
    if (verdict !== 'ok') {
      const file = join(REPORT_DIR, `${testCase.id}.png`);
      const image = shot.image ?? (page && !closed ? await within(page.screenshot(), 5000) : null);
      if (image) {
        writeFileSync(file, image);
        row.screenshot = relative(ROOT, file);
      }
    }
    await within(browser?.close() ?? Promise.resolve(), 5000);
    await within(server?.close() ?? Promise.resolve(), 5000);
    // A browser that has exited no longer owns its process id, which the system can reuse.
    const process_ = server?.process();
    if (pid !== undefined && process_?.exitCode === null && process_.signalCode === null)
      await killTree(pid);
    runningBrowser = undefined;
  }
  return row;
}

/** The cases that `HANG_ONLY` selects. It stops the run with the valid ids for a wrong value. */
function selectedCases(): readonly HangCase[] {
  const raw = process.env.HANG_ONLY;
  if (raw === undefined) return HANG_CASES;
  const only = raw
    .split(',')
    .map((id) => id.trim())
    .filter(Boolean);
  const unknown = only.filter((id) => !HANG_CASES.some((c) => c.id === id));
  if (only.length > 0 && unknown.length === 0) return HANG_CASES.filter((c) => only.includes(c.id));
  console.error(
    `HANG_ONLY must name cases of this list${unknown.length ? `, not ${unknown.join(', ')}` : ''}:`
  );
  console.error(HANG_CASES.map((c) => c.id).join(', '));
  process.exit(2);
}

async function main(): Promise<void> {
  if (process.argv.length > 2) {
    console.error('Name cases with HANG_ONLY: HANG_ONLY=id,id bun run test:hang:browser');
    process.exit(2);
  }
  const cases = selectedCases();
  const dir = mkdtempSync(join(tmpdir(), 'xlsx-hang-'));
  rmSync(REPORT_DIR, { recursive: true, force: true });
  mkdirSync(REPORT_DIR, { recursive: true });
  const rows: Row[] = [];
  const report = () =>
    writeFileSync(join(REPORT_DIR, 'report.json'), `${JSON.stringify(rows, null, 2)}\n`);
  for (const signal of ['SIGINT', 'SIGTERM'] as const)
    process.on(signal, () => {
      const stop = runningBrowser === undefined ? Promise.resolve() : killTree(runningBrowser);
      void stop.finally(() => {
        report();
        rmSync(dir, { recursive: true, force: true });
        process.exit(130);
      });
    });
  const server = await createServer({
    configFile: join(DEMO, 'vite.config.ts'),
    root: DEMO,
    server: { host: '127.0.0.1', port: 5298, strictPort: false },
    logLevel: 'error',
  });
  try {
    await server.listen();
    const base = server.resolvedUrls?.local[0] ?? 'http://127.0.0.1:5298/';
    for (const testCase of cases) {
      console.log(`run    ${testCase.id}`);
      const row = await inOwnBrowser(testCase, dir, base);
      rows.push(row);
      report();
      const shot = row.screenshot ? `, screenshot ${row.screenshot}` : '';
      const note = row.why ? `, ${row.why} (case: ${row.covers})${shot}` : '';
      const block = row.longestBlockMs === null ? 'not measured' : `${row.longestBlockMs} ms`;
      console.log(
        `${row.verdict.padEnd(6)} ${row.id}: ${row.outcome || 'no outcome'} in ${row.openMs} ms, ` +
          `block ${block}, repeated ${row.repeatedBlockMs} ms, peak ${row.peakMiB} MiB, ` +
          `heap ${row.heapMB ?? 'not read'} MB${note}`
      );
      if (row.demoDown) {
        console.error('The demo did not start, so the run stops. Read the errors of the demo.');
        break;
      }
    }
  } catch (cause) {
    if (!(cause instanceof NoBrowser)) throw cause;
    console.error(`Chromium did not start: ${cause.message}`);
    console.error('Install it with: bunx playwright install chromium');
    process.exitCode = 2;
    return;
  } finally {
    await server.close();
    rmSync(dir, { recursive: true, force: true });
  }
  const failed = rows.filter((r) => fails(r.verdict));
  const slow = rows.filter((r) => r.verdict === 'SLOW');
  console.log(`${rows.length - failed.length} of ${rows.length} cases passed`);
  if (rows.length < cases.length) console.log(`${cases.length - rows.length} cases did not run`);
  if (slow.length > 0) console.log(`SLOW, not failed: ${slow.map((r) => r.id).join(', ')}`);
  summarize(rows, failed, slow);
  if (failed.length > 0) {
    const ids = failed.map((r) => r.id).join(',');
    console.log(`Report: ${join(REPORT_DIR, 'report.json')}`);
    console.log(`Run the failed cases again: HANG_ONLY=${ids} bun run test:hang:browser`);
  }
  process.exit(failed.length > 0 ? 1 : 0);
}

/** Writes the counts and the SLOW cases to the job summary, when the run is a CI job. */
function summarize(rows: readonly Row[], failed: readonly Row[], slow: readonly Row[]): void {
  const file = process.env.GITHUB_STEP_SUMMARY;
  if (!file) return;
  const lines = [
    '## Browser hang test',
    '',
    `${rows.length - failed.length} of ${rows.length} cases passed.`,
  ];
  if (failed.length > 0) lines.push('', `Failed: ${failed.map((r) => `\`${r.id}\``).join(', ')}.`);
  if (slow.length > 0) {
    lines.push('', 'SLOW cases. The mark does not fail the job:', '');
    for (const r of slow) lines.push(`- \`${r.id}\`: ${r.why ?? ''}`);
  }
  appendFileSync(file, `${lines.join('\n')}\n`);
}

await main();
