// The browser bench: open, edit, scroll, and memory in the demo, on the generated workbooks of
// `workbooks.ts`, in Playwright Chromium. It is a local tool, as PPTX Editor's browser benches: it
// prints a table, and no time fails it. It uses `BROWSER_SIZES`, which are smaller than the sizes
// of the edit bench, to keep a margin under the 3 GiB cap of the watchdog, a hang
// limit that stays. `scripts/bench/README.md` gives the usage.
//
//   bun run bench:<mode>:browser -- [options]  (`--help` lists the options)
//
// What each mode measures. Each time sample is on the page clock, and it ends at the start of the
// first task after the frame that paints the change, so it holds the style, layout, and paint of
// that frame:
// - open: from the `change` event of the file input to the end of the frame in which the grid
//   shows the first sheet, in a new browser for each run. It also reads the JS heap after a garbage
//   collection. The editor reads the whole workbook before it shows the first sheet (issue
//   #164), so this is also the time until the editor takes keys.
// - edit: from the `keydown` event to the end of the frame after the editor handled it, for the
//   first key of an edit (`begin`), for each key that types into the cell editor (`type`), and for
//   Enter, which writes the value (`commit`). Twenty edits run first as a warmup, in columns I and
//   J, as DOCX Editor's browser bench. Then 60 edits run in six groups of 10, one group in each
//   column from C to H. Each group starts 12 rows above the last row of the data and goes down. No
//   commit may scroll the grid, so each commit does the same work, and the formula bar must show
//   the last value of the group. The table gives the median of the first and of the last 10 samples
//   of each step: a step that grows slower with each edit shows a last median far above the first.
//   DOCX Editor types 180 keys; here the 60 edits give 180 `type` keys and 60 of each other step,
//   and an edit with its commit takes much longer than a typed key. Then 5 untimed and 30 timed
//   ArrowDown keys move the active cell inside the data (`move`), and the name box must show the
//   cell 30 rows lower. On the sheets workbook it also times Alt+PageDown, after 5 untimed keys,
//   from the key to the end of the frame in which the grid names the next sheet (`sheet-switch`),
//   and each key must show a new sheet.
// - scroll: from each `scroll` event of the grid to the end of the frame after it, for 40 wheel
//   steps, after 5 untimed steps. The grid must have scrolled at the end. The sheets workbook has
//   100 rows a sheet, so on it the mode only opens the workbook and records the peak memory.
// - memory: after a warmup of 200 commits, which fills the undo history of the store
//   (`MEMORY_WARMUP_COMMITS`), and after each of 20 groups of 5 commits, it collects garbage and
//   reads the JS heap and its array buffers, the DOM nodes, the listeners, the documents, and the
//   resources of the page probe (`memory-probe.ts`). The groups edit columns B to G in turn, and
//   the warmup edits the same cells of column H 40 times, with a garbage collection after each
//   group of 5 commits. Each commit writes a text that is new to the workbook, so each commit
//   adds a shared string item. The leak check follows PPTX
//   Editor's memory verdict (`leaksOf` in `measure.ts`): a bounded cache levels off, and a leak
//   keeps its slope. As in PPTX Editor's memory bench, `MEMORY_SNAPSHOTS=1` takes a heap
//   snapshot after the warmup and one at the end, and prints the constructors that grew most
//   (`heap-summary.ts`). A leak without it prints how to get them.
// A key adds a sample only when the grid or its cell editor has the focus, so a key that goes to
// another element stops the run at the step time limit. Each mode also prints the longest task
// (`longtask` entries) of the whole page, from the load of the demo and the open of the workbook
// to the last step, and the peak memory of the browser process tree
// (`scripts/hang/watchdog.ts`), which stops a browser past 3 GiB.
//
// The run fails only when a workbook stops (a page error, a step over its time limit, or an edit,
// a move, or a scroll that does not land) or when the memory mode finds a leak, as PPTX Editor's
// memory bench fails only on its verdict.
//
// Model: PPTX Editor's `e2e/open.bench.ts`, `e2e/edit.bench.ts`, and `e2e/memory.bench.ts` at
// d4502552 (EigenPal), and DOCX Editor's `e2e/edit-browser.bench.spec.ts` at
// 1c9b6a29b for the warmup and the first and last ten edits. Differences, each for a reason:
// - It runs without the Playwright test runner, as `scripts/hang/run.ts` does, so it reads the
//   memory of the browser process tree. It starts one browser for each run of open, and one
//   browser for each workbook in the other modes.
// - It measures on the page clock, as PPTX Editor's key timings do, not with the main-thread
//   time of the DevTools protocol: a grid paints in a frame, and the frame is what the user sees.
// - PPTX Editor installs its page probe only in its memory bench. This bench installs it in each
//   mode, so the open reads its heap with the same reading as the memory mode. The probe adds one
//   counter step to each `addEventListener` and `setInterval` call of the page. It also counts
//   the console errors and warnings in each mode, as PPTX Editor's memory bench counts them.
// - It adds a scroll mode, which neither sister has. The grid is virtualized: a wheel step lays
//   out and paints only the cells that come into view (AGENTS.md, "Layout"), so the scroll mode
//   times that work.
// - It reads the used JS heap from `Runtime.getHeapUsage`, the call that also gives the array
//   buffers, where PPTX Editor reads the `JSHeapUsedSize` metric. Both give the used JS heap, and
//   as in PPTX Editor, a browser that gives no array buffer size adds 0.
// - PPTX Editor's memory bench (f448820c) has a scenario for each kind of repeated work: opens,
//   mounts, typing, moves, slide switches, pictures, exports, find, slideshows, and dialogs. The
//   memory mode here repeats commits, the work of its typing scenario, because the perf fix of
//   this bench changes what a commit keeps. Repeated opens and sheet switches have no
//   scenario.
// - PPTX Editor's typing scenario types 5000 keys and reads memory 11 times: once before the keys
//   and after each tenth of them. The memory mode here commits 100 times, 500 keys, and reads
//   memory 21 times, so the verdict reads 50 commits over 11 readings: a commit takes many more
//   steps than a key. The browser's own edit history of the cell editor does not level off in
//   300 edits (`LEAK_LIMITS` in `measure.ts`), so the node limit allows that growth.
// - PPTX Editor's memory bench collects no garbage during its warmup. Here the warmup collects
//   garbage after each group of 5 commits, as the timed groups do before each reading. Without
//   it, the garbage of 200 commits piled up, and on the grid workbook the browser took 1934 to
//   3212 MiB in nine local runs and went above the 3 GiB cap of the watchdog in three of them.
//   With it, four runs of the grid workbook alone took 1395 to 1409 MiB.
// - DOCX Editor's browser bench at 1c9b6a29b has a self-test that adds a known delay to each
//   edit and checks that the bench measures it (`EDIT_BROWSER_BENCH_DELAY_MS`), so that its
//   timing gates respond to a known regression. PPTX Editor's benches have none. This bench has
//   no timing gate, so it leaves the self-test out; a time sample ends after the frame that
//   paints its change, which `measure.test.ts` does not check.
// - DOCX Editor's browser bench at 1c9b6a29b also pins the layout work of a browser edit to the
//   work of its headless bench, bounds the pages that it keeps in the DOM
//   (`e2e/edit-browser-bench-gates.ts`), and sends key bursts at 30 Hz with no wait between keys
//   (`e2e/edit-browser-burst.ts`). PPTX Editor's browser benches have none of these, and this
//   bench takes that simpler form: the work counters are pinned in `bun run test` on the editor in
//   happy-dom (`edit-bench-gates.test.ts`), and each timed key waits for its frame, so a backlog
//   of keys does not show.
// - DOCX Editor's browser bench fails on timing tails and on a last ten edits twice as slow as the
//   first ten. PPTX Editor's browser benches do not, and this bench takes the simpler form: a time
//   on a shared machine depends on its load, so the table only reports it.
// - DOCX Editor's browser bench at 1c9b6a29b also times copy and paste and checks the pasted text
//   (`e2e/edit-browser.bench.spec.ts`). PPTX Editor's browser benches have no copy and paste
//   case, and this bench takes that form. `bun run test:e2e` checks copy, cut, and paste in the
//   demo in Chromium, and in another engine with `--browser`.

import { chromium, type Browser, type BrowserServer, type Page } from '@playwright/test';
import {
  accessSync,
  constants,
  existsSync,
  mkdtempSync,
  rmSync,
  statSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { createServer } from 'vite';
import { killTree, watchBrowser, type Watch } from '../hang/watchdog.ts';
import { heapSummary, summaryDiff } from './heap-summary.ts';
import { installProbe, type Resources } from './memory-probe.ts';
import { names, positive, readFlags, stopOnUsageError, UsageError } from './args.ts';
import {
  growthOf,
  leaksOf,
  MEMORY_WARMUP_COMMITS,
  RESOURCES,
  valuesOf,
  type Mode,
  type Reading,
  type Row,
  type WorkbookName,
  WORKBOOKS,
} from './measure.ts';
import {
  BENCH_SHEET,
  BENCH_SIZES,
  BROWSER_SIZES,
  gridWorkbook,
  SHEET_COLS,
  SHEET_ROWS,
  sheetsWorkbook,
  stringsWorkbook,
} from './workbooks.ts';

const ROOT = join(import.meta.dir, '../..');
const DEMO = join(ROOT, 'examples/vite');
/** The time of one step: a step that takes longer is a hang, and the run stops. */
const STEP_MS = 120_000;
/** The time of one browser: the open of the largest workbook and all its steps. */
const BROWSER_MS = 600_000;
/** Heap snapshots in the memory mode, as PPTX Editor's `MEMORY_SNAPSHOTS`. */
const SNAPSHOTS = process.env.MEMORY_SNAPSHOTS === '1';

const MODES: readonly Mode[] = ['open', 'edit', 'scroll', 'memory'];

/** The edits of each group of the edit mode. */
const EDITS_PER_GROUP = 10;

/** The rows above the last row of the data where each group of edits starts. */
const GROUP_START_ROWS = 12;

/** The timed ArrowDown keys. */
const MOVES = 30;

/** The untimed keys before the timed arrow keys, wheel steps, and sheet switches. */
const STEP_WARMUP = 5;

/** The rows above the last row of the data where the timed arrow keys start (`moves`). */
const MOVES_START_ROWS = MOVES + STEP_WARMUP + 5;

/** The full rows of data that the edit groups and the arrow keys of the edit mode need. */
const EDIT_ROWS_NEEDED = Math.max(GROUP_START_ROWS, MOVES_START_ROWS) + 1;

interface Args {
  mode: Mode;
  workbooks: WorkbookName[];
  rows: number;
  strings: number;
  runs: number;
  out?: string;
  headed: boolean;
}

const USAGE = `Usage: bun run bench:<mode>:browser -- [options]
       bun scripts/bench/browser-bench.ts <mode> [options]

  <mode>               open, edit, scroll, or memory
  --workbooks <names>  grid, sheets, strings, or a list (default: all three)
  --rows <n>           rows of the grid workbook (default: ${BROWSER_SIZES.gridRows})
  --strings <n>        shared strings of the strings workbook
                       (default: ${BROWSER_SIZES.strings})
  --runs <n>           opens of each workbook in open mode (default: 3)
  --out <file>         also write the values of the table, the samples, the
                       browser version, and the SHA-256 of each workbook to
                       <file> as JSON
  --headed             show the browser
  --help               print this text

The edit mode needs ${EDIT_ROWS_NEEDED} or more rows of data, and the memory
mode ${GROUP_START_ROWS + 1} or more. The strings workbook has ${BENCH_SIZES.gridCols} strings to a row.
The sheets workbook always has ${BROWSER_SIZES.sheets} worksheets.

With MEMORY_SNAPSHOTS=1, the memory mode also prints the heap growth by
constructor of each workbook.

Each mode needs Playwright's Chromium: bunx playwright install chromium`;

/** The arguments, or undefined after `--help`. */
function parseArgs(argv: readonly string[]): Args | undefined {
  const [first, ...rest] = argv;
  if (first === '--help' || first === '-h') return void console.log(USAGE);
  const mode = first as Mode;
  if (!MODES.includes(mode))
    throw new UsageError(`the first argument is one of ${MODES.join(', ')}\n\n${USAGE}`);
  const flags = readFlags(
    rest,
    {
      values: ['--workbooks', '--rows', '--strings', '--runs', '--out'],
      switches: ['--headed'],
    },
    USAGE
  );
  if (!flags) return undefined;
  const out = flags.get('--out');
  // A file that cannot be written must stop the run before any work, not after it.
  if (typeof out === 'string') {
    const file = resolve(out);
    const folder = dirname(file);
    if (!existsSync(folder) || !statSync(folder).isDirectory())
      throw new UsageError(`--out: the folder of ${out} does not exist`);
    if (existsSync(file) && statSync(file).isDirectory())
      throw new UsageError(`--out: ${out} is a folder; give a file path`);
    try {
      accessSync(existsSync(file) ? file : folder, constants.W_OK);
    } catch {
      throw new UsageError(`--out: you cannot write to ${out}`);
    }
  }
  if (mode !== 'open' && flags.has('--runs'))
    throw new UsageError(`--runs applies to the open mode only\n\n${USAGE}`);
  const workbooks = names(flags, '--workbooks', WORKBOOKS);
  const rows = positive(flags, '--rows', BROWSER_SIZES.gridRows);
  const strings = positive(flags, '--strings', BROWSER_SIZES.strings);
  // The edit groups and the arrow keys start above the last row of the data (`goTo`), so each
  // selected workbook needs that many full rows of data in the modes that edit.
  const needed = mode === 'edit' ? EDIT_ROWS_NEEDED : mode === 'memory' ? GROUP_START_ROWS + 1 : 0;
  if (workbooks.includes('grid') && rows < needed)
    throw new UsageError(`--rows takes ${needed} or more in the ${mode} mode: its edits need them`);
  const perRow = BENCH_SIZES.gridCols;
  if (workbooks.includes('strings') && Math.floor(strings / perRow) < needed)
    throw new UsageError(
      `--strings takes ${needed * perRow} or more in the ${mode} mode: ${perRow} to a row, ` +
        `its edits need ${needed} rows`
    );
  return {
    mode,
    workbooks,
    rows,
    strings,
    runs: positive(flags, '--runs', 3),
    out: typeof out === 'string' ? out : undefined,
    headed: flags.has('--headed'),
  };
}

/** The page side of the bench, installed before the demo's scripts. */
interface Probe {
  /** The samples of each kind, in milliseconds. */
  samples: Record<string, number[]>;
  /**
   * The kind that the next key or `scroll` event adds a sample to, or undefined. A key adds a
   * sample to each kind but `scroll`, and a `scroll` event only to `scroll`.
   */
  kind: string | undefined;
  longest: number;
  /** The highest `scrollTop` of an element that a timed `scroll` event came from. */
  scrolledTo: number;
}

function probe(): void {
  const bench: Probe = { samples: {}, kind: undefined, longest: 0, scrolledTo: 0 };
  (window as unknown as { __bench: Probe }).__bench = bench;
  // The editor paints in an animation frame that it asks for while it handles the event. The
  // task after the next frame starts when that frame has done its style, layout, and paint, as
  // in PPTX Editor's `e2e/open.bench.ts`.
  const afterFrame = (done: () => void) => requestAnimationFrame(() => setTimeout(done));
  // The editor names the sheet of the grid in its own animation frame (`render`), before the
  // style, layout, and paint of that frame. The observer runs at the end of that frame callback,
  // so the next task starts when that frame has painted. A new animation frame would wait for the
  // frame after it.
  const afterThisFrame = (done: () => void) => setTimeout(done);
  const add = (kind: string, start: number) =>
    afterFrame(() => (bench.samples[kind] ??= []).push(performance.now() - start));
  let opening: number | undefined;
  // A sheet switch ends at the frame in which the grid names the new sheet, not after the key.
  let switching: number | undefined;
  let shown: string | null | undefined;
  document.addEventListener(
    'change',
    (e) => {
      if ((e.target as HTMLElement).matches?.('input[type=file]')) opening = e.timeStamp;
    },
    true
  );
  new MutationObserver(() => {
    const label = document.querySelector('[role=grid]')?.getAttribute('aria-label');
    const changed = label !== shown;
    shown = label;
    if (opening !== undefined && label === 'Worksheet Data') {
      const start = opening;
      opening = undefined;
      afterThisFrame(() => (bench.samples.open ??= []).push(performance.now() - start));
    }
    if (switching !== undefined && changed) {
      const start = switching;
      switching = undefined;
      afterThisFrame(() => (bench.samples['sheet-switch'] ??= []).push(performance.now() - start));
    }
  }).observe(document, { subtree: true, attributes: true, attributeFilter: ['aria-label'] });
  document.addEventListener(
    'keydown',
    (e) => {
      // A key that goes to another element than the grid adds no sample, and neither does a
      // modifier key: Alt+PageDown sends an Alt keydown before the PageDown keydown.
      const target = e.target as Element | null;
      if (
        bench.kind === undefined ||
        bench.kind === 'scroll' ||
        ['Alt', 'Control', 'Meta', 'Shift'].includes(e.key) ||
        !target?.closest?.('[role=grid], .xlsx-cell-editor')
      )
        return;
      if (bench.kind === 'sheet-switch') switching = e.timeStamp;
      else add(bench.kind, e.timeStamp);
    },
    true
  );
  document.addEventListener(
    'scroll',
    (e) => {
      // A switch or a commit can scroll the grid: only the scroll step times a scroll.
      if (bench.kind !== 'scroll') return;
      add('scroll', e.timeStamp);
      const top = (e.target as Element | null)?.scrollTop;
      if (typeof top === 'number') bench.scrolledTo = Math.max(bench.scrolledTo, top);
    },
    true
  );
  new PerformanceObserver((list) => {
    for (const entry of list.getEntries()) bench.longest = Math.max(bench.longest, entry.duration);
  }).observe({ type: 'longtask', buffered: true });
}

const samplesOf = (page: Page, kind: string): Promise<number> =>
  page.evaluate(
    (k) => (window as unknown as { __bench: Probe }).__bench.samples[k]?.length ?? 0,
    kind
  );

/**
 * Waits until the page has `count` samples of `kind`. A step that adds no sample, for example a
 * key that goes to another element than the grid or its cell editor, stops the run with the step
 * and the element that has the focus.
 */
async function waitSamples(page: Page, kind: string, count: number, key = ''): Promise<void> {
  try {
    await page.waitForFunction(
      ([k, n]) =>
        ((window as unknown as { __bench: Probe }).__bench.samples[k as string]?.length ?? 0) >=
        (n as number),
      [kind, count] as const,
      { timeout: STEP_MS, polling: 20 }
    );
  } catch (error) {
    const focus = await page
      .evaluate(() => {
        const el = document.activeElement;
        if (!el) return 'no element';
        const label = el.getAttribute('aria-label') ?? el.getAttribute('class') ?? '';
        return `${el.tagName.toLowerCase()} ${label}`.trim();
      })
      .catch(() => 'an unknown element');
    const what = key ? `${kind} key ${key}` : kind;
    throw new Error(`${what}: no sample in ${STEP_MS / 1000} s; the focus is on ${focus}`, {
      cause: error,
    });
  }
}

/** Presses `key` and waits for its sample of `kind`. */
async function timedPress(page: Page, kind: string, key: string): Promise<void> {
  const before = await samplesOf(page, kind);
  await page.evaluate((k) => ((window as unknown as { __bench: Probe }).__bench.kind = k), kind);
  await page.keyboard.press(key);
  await waitSamples(page, kind, before + 1, key);
  await page.evaluate(() => ((window as unknown as { __bench: Probe }).__bench.kind = undefined));
}

/**
 * A memory reading after garbage collection: the JS heap and the array buffers behind it, the DOM
 * nodes, the listeners, the documents, and the resources of the page probe (`memory-probe.ts`).
 * As in PPTX Editor's `e2e/memory.bench.ts`, pending timers and idle work run first, and three
 * collections let finalizers free more.
 */
async function readingAfterGc(page: Page): Promise<Reading> {
  const cdp = await page.context().newCDPSession(page);
  try {
    await cdp.send('Performance.enable');
    await page.waitForTimeout(300);
    for (let i = 0; i < 3; i += 1) {
      await cdp.send('HeapProfiler.collectGarbage');
      await page.waitForTimeout(100);
    }
    const usage = (await cdp.send('Runtime.getHeapUsage')) as {
      usedSize: number;
      backingStorageSize?: number;
    };
    const { metrics } = (await cdp.send('Performance.getMetrics')) as {
      metrics: { name: string; value: number }[];
    };
    const metric = (name: string): number => {
      const value = metrics.find((m) => m.name === name)?.value;
      if (value === undefined) throw new Error(`the browser gives no ${name} metric`);
      return value;
    };
    const resources = await page.evaluate(() =>
      (window as unknown as { __memProbe: () => Resources }).__memProbe()
    );
    return {
      ...resources,
      heapBytes: usage.usedSize + (usage.backingStorageSize ?? 0),
      nodes: metric('Nodes'),
      listeners: metric('JSEventListeners'),
      documents: metric('Documents'),
    };
  } finally {
    await cdp.detach();
  }
}

function workbookOf(name: WorkbookName, args: Args): { bytes: Uint8Array; cells: number } {
  switch (name) {
    case 'grid':
      return { bytes: gridWorkbook(args.rows), cells: args.rows * BENCH_SIZES.gridCols };
    case 'sheets':
      return {
        bytes: sheetsWorkbook(BROWSER_SIZES.sheets),
        cells: BROWSER_SIZES.sheets * SHEET_ROWS * SHEET_COLS,
      };
    case 'strings':
      return { bytes: stringsWorkbook(args.strings), cells: args.strings };
  }
}

/** The version of the last browser of the run, for the `--out` file. */
let browserVersion: string | undefined;

/** One browser with one page on the demo, stopped in every path. */
async function inBrowser<T>(
  base: string,
  args: Args,
  work: (page: Page, watch: Watch) => Promise<T>
): Promise<{ value: T; peakMiB: number }> {
  let server: BrowserServer | undefined;
  let browser: Browser | undefined;
  let watch: Watch | undefined;
  let pid: number | undefined;
  try {
    server = await chromium.launchServer({ headless: !args.headed });
    pid = server.process().pid;
    if (pid === undefined) throw new Error('the browser has no process id');
    watch = watchBrowser(pid, BROWSER_MS);
    browser = await chromium.connect(server.wsEndpoint());
    browserVersion = browser.version();
    const context = await browser.newContext({ viewport: { width: 1440, height: 900 } });
    const page = await context.newPage();
    const errors: string[] = [];
    page.on('pageerror', (e) => errors.push(e.message));
    // As PPTX Editor's memory bench counts them, so a warning that repeats on each step shows.
    const consoleCounts = new Map<string, number>();
    page.on('console', (m) => {
      if (m.type() !== 'error' && m.type() !== 'warning') return;
      const text = `${m.type()}: ${m.text().slice(0, 160)}`;
      consoleCounts.set(text, (consoleCounts.get(text) ?? 0) + 1);
    });
    page.on('dialog', (d) => void d.accept().catch(() => undefined));
    await page.addInitScript(installProbe);
    await page.addInitScript(probe);
    await page.goto(base);
    // The demo opens its sample workbook first.
    await page.waitForFunction(
      () =>
        document.querySelector('[role=grid]')?.getAttribute('aria-label')?.startsWith('Worksheet '),
      undefined,
      { timeout: STEP_MS }
    );
    // A browser that the watchdog stopped fails its step with "Browser closed": name the cause.
    const stoppedBy = (w: Watch): string | null => {
      const why = w.tripped();
      if (!why) return null;
      return w.trippedMemory()
        ? `${why}, over the 3 GiB cap of the watchdog`
        : `the browser passed the time limit of the bench (${why})`;
    };
    let value: T;
    try {
      value = await work(page, watch);
    } catch (error) {
      const why = stoppedBy(watch);
      throw why ? new Error(why, { cause: error }) : error;
    } finally {
      for (const [text, k] of consoleCounts) console.error(`  console x${k} ${text}`);
    }
    const tripped = stoppedBy(watch);
    if (tripped) throw new Error(tripped);
    if (errors.length > 0) throw new Error(`page errors: ${errors.join('; ')}`);
    return { value, peakMiB: Math.round(watch.peakBytes() / 1048576) };
  } finally {
    watch?.stop();
    await browser?.close().catch(() => undefined);
    await server?.close().catch(() => undefined);
    const child = server?.process();
    if (pid !== undefined && child?.exitCode === null && child.signalCode === null)
      await killTree(pid);
  }
}

/**
 * Opens `file` through the demo's file input and waits for its open sample. A workbook that the
 * demo refuses, for example with a part larger than the XML part limit, stops the run with the
 * error message of the demo (`role=alert`), not at the step time limit.
 */
async function open(page: Page, file: string): Promise<void> {
  const before = await samplesOf(page, 'open');
  await page.locator('input[type=file]').first().setInputFiles(file);
  const outcome = await page.waitForFunction(
    (n) => {
      const opens = (window as unknown as { __bench: Probe }).__bench.samples.open?.length ?? 0;
      if (opens >= n) return { refused: null };
      const alert = document.querySelector('[role=alert]')?.textContent?.trim();
      return alert ? { refused: alert } : false;
    },
    before + 1,
    { timeout: STEP_MS, polling: 20 }
  );
  const { refused } = (await outcome.jsonValue()) as { refused: string | null };
  if (refused !== null) throw new Error(`the demo did not open the workbook: ${refused}`);
  const label = await page.evaluate(() =>
    document.querySelector('[role=grid]')?.getAttribute('aria-label')
  );
  if (label !== `Worksheet ${BENCH_SHEET}`) throw new Error(`the grid shows ${label}`);
}

async function focusGrid(page: Page): Promise<void> {
  const focused = await page.evaluate(() => {
    const grid = document.querySelector<HTMLElement>('[role=grid]');
    grid?.focus();
    return grid !== null && document.activeElement === grid;
  });
  if (!focused) throw new Error('the grid did not take the focus');
}

/** The edits that run before the timed edits, as DOCX Editor's `SUSTAINED_WARMUP_EDITS`. */
const WARMUP_EDITS = 20;

/** The timed edits: the first and the last 10 of `windowsOf`, and more between them. */
const TIMED_EDITS = 60;

/** The first column of the timed groups: C. They use C to H, and the warmup I and J. */
const FIRST_COLUMN = 2;

/**
 * The groups of commits of the memory mode, after the warmup groups in column H. The groups edit
 * columns B to G in turn, so each column takes 3 or 4 groups, and each group writes new texts.
 */
const MEMORY_GROUPS = 20;
const MEMORY_COLUMNS = 6;
const COMMITS_PER_GROUP = 5;

/**
 * The text of commit `i` of memory group `group`: new to the workbook, so each commit adds a
 * shared string item.
 */
const memoryText = (group: number) => (i: number) => `x${100 + group * 10 + i}`;

/**
 * The text of commit `i` of warmup group `group`: new to the workbook too. The workbooks hold
 * `w0`, `w1`, and so on (`sharedText`), so the warmup texts start with another letter.
 */
const warmupText = (group: number) => (i: number) => `y${group * COMMITS_PER_GROUP + i}`;

/** The scroll position of the grid, as `top,left`. */
const scrollOf = (page: Page): Promise<string> =>
  page.evaluate(() => {
    const scroller = document.querySelector('.xlsx-scroller');
    return `${scroller?.scrollTop},${scroller?.scrollLeft}`;
  });

/**
 * One group of edits in column `column` (0 is A), down from `GROUP_START_ROWS` rows above the
 * last row of the data: for each edit, the first key of `value(i)` begins it, the other keys type
 * into the cell editor, and Enter commits it. Each group starts from the same place of the grid,
 * and no commit of a group may scroll the grid, so each commit does the same work. With `timed`,
 * each key adds a sample. The formula bar must then show the last value.
 */
async function editGroup(
  page: Page,
  column: number,
  count: number,
  timed: boolean,
  value: (i: number) => string = () => '7123'
): Promise<void> {
  await goTo(page, column, GROUP_START_ROWS);
  const start = await scrollOf(page);
  const press = (kind: string, key: string) =>
    timed ? timedPress(page, kind, key) : page.keyboard.press(key);
  for (let i = 0; i < count; i += 1) {
    const [first, ...rest] = value(i);
    await press('begin', first!);
    for (const key of rest) await press('type', key);
    await press('commit', 'Enter');
  }
  const end = await scrollOf(page);
  if (end !== start) throw new Error(`the commits scrolled the grid from ${start} to ${end}`);
  await page.keyboard.press('ArrowUp');
  const shown = await page.evaluate(
    () => document.querySelector<HTMLTextAreaElement>('textarea.xlsx-cell-content')?.value
  );
  const last = value(count - 1);
  if (shown !== last) throw new Error(`the last edit shows ${shown}, not ${last}`);
}

/** Waits for the frame after the next one, so that the work of earlier keys is painted. */
async function settle(page: Page): Promise<void> {
  await page.evaluate(
    () => new Promise<void>((r) => requestAnimationFrame(() => requestAnimationFrame(() => r())))
  );
}

/**
 * Makes the cell `up` rows above the last row of the data in column `column` (0 is A) active,
 * from the top of the sheet, and waits until the grid painted it.
 */
async function goTo(page: Page, column: number, up: number): Promise<void> {
  await focusGrid(page);
  await page.keyboard.press('Control+Home');
  for (let i = 0; i < column; i += 1) await page.keyboard.press('ArrowRight');
  await page.keyboard.press('Control+ArrowDown');
  for (let i = 0; i < up; i += 1) await page.keyboard.press('ArrowUp');
  await settle(page);
}

/** The warmup edits, then the timed edits in groups of `EDITS_PER_GROUP`, one column each. */
async function edits(page: Page): Promise<void> {
  const groups = TIMED_EDITS / EDITS_PER_GROUP;
  for (let i = 0; i < WARMUP_EDITS / EDITS_PER_GROUP; i += 1)
    await editGroup(page, FIRST_COLUMN + groups + i, EDITS_PER_GROUP, false);
  for (let group = 0; group < groups; group += 1)
    await editGroup(page, FIRST_COLUMN + group, EDITS_PER_GROUP, true);
}

/** The column letters and the row number that the name box shows. */
async function activeCell(page: Page): Promise<{ column: string; row: number }> {
  const ref = await page.evaluate(
    () => document.querySelector<HTMLInputElement>('input.xlsx-name-box')?.value ?? ''
  );
  const match = /^([A-Z]{1,3})([1-9][0-9]*)$/.exec(ref);
  if (!match) throw new Error(`the name box shows ${JSON.stringify(ref)}, not a cell`);
  return { column: match[1]!, row: Number(match[2]) };
}

/**
 * Times ArrowDown `MOVES` times in the data, after `STEP_WARMUP` untimed keys; the active cell
 * must then be `MOVES` rows lower.
 */
async function moves(page: Page): Promise<void> {
  await goTo(page, FIRST_COLUMN, MOVES_START_ROWS);
  for (let i = 0; i < STEP_WARMUP; i += 1) await page.keyboard.press('ArrowDown');
  await settle(page);
  const before = await activeCell(page);
  for (let i = 0; i < MOVES; i += 1) await timedPress(page, 'move', 'ArrowDown');
  const after = await activeCell(page);
  if (after.column !== before.column || after.row !== before.row + MOVES)
    throw new Error(
      `ArrowDown ${MOVES} times from ${before.column}${before.row} shows ` +
        `${after.column}${after.row}`
    );
}

/**
 * Times Alt+PageDown `count` times, from the key to the frame after the grid names the next
 * sheet. Each key must show another sheet.
 */
async function switches(page: Page, count: number): Promise<void> {
  const label = () =>
    page.evaluate(() => document.querySelector('[role=grid]')?.getAttribute('aria-label'));
  for (let i = 0; i < STEP_WARMUP; i += 1) await page.keyboard.press('Alt+PageDown');
  await settle(page);
  for (let i = 0; i < count; i += 1) {
    const before = await label();
    await timedPress(page, 'sheet-switch', 'Alt+PageDown');
    const after = await label();
    if (after === before) throw new Error(`Alt+PageDown kept ${before}`);
  }
}

async function runWorkbook(
  name: WorkbookName,
  args: Args,
  base: string,
  dir: string
): Promise<Row> {
  const { bytes, cells } = workbookOf(name, args);
  const file = join(dir, `${name}.xlsx`);
  writeFileSync(file, bytes);
  const workbookSha256 = new Bun.CryptoHasher('sha256').update(bytes).digest('hex');
  const row: Row = {
    workbook: name,
    cells,
    workbookSha256,
    samples: {},
    longestTaskMs: 0,
    peakMiB: 0,
  };
  const take = async (page: Page) => {
    const got = await page.evaluate(() => (window as unknown as { __bench: Probe }).__bench);
    for (const [kind, xs] of Object.entries(got.samples))
      if (kind !== 'open' || args.mode === 'open') (row.samples[kind] ??= []).push(...xs);
    row.longestTaskMs = Math.max(row.longestTaskMs, Math.round(got.longest));
  };
  const runs = args.mode === 'open' ? args.runs : 1;
  for (let run = 0; run < runs; run += 1) {
    const { peakMiB } = await inBrowser(base, args, async (page) => {
      await open(page, file);
      if (args.mode === 'open' && run === runs - 1)
        row.heapMiB = Math.round((await readingAfterGc(page)).heapBytes / 1048576);
      if (args.mode === 'edit') {
        await edits(page);
        await moves(page);
        if (name === 'sheets') await switches(page, 10);
      }
      if (args.mode === 'scroll' && name !== 'sheets') {
        const box = await page.locator('[role=grid]').first().boundingBox();
        if (!box) throw new Error('the grid has no box');
        await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
        for (let i = 0; i < STEP_WARMUP; i += 1) {
          await page.mouse.wheel(0, 600);
          await settle(page);
        }
        await page.evaluate(
          () => ((window as unknown as { __bench: Probe }).__bench.kind = 'scroll')
        );
        for (let i = 0; i < 40; i += 1) {
          const before = await samplesOf(page, 'scroll');
          await page.mouse.wheel(0, 600);
          await waitSamples(page, 'scroll', before + 1);
        }
        const top = await page.evaluate(
          () => (window as unknown as { __bench: Probe }).__bench.scrolledTo
        );
        if (top <= 0) throw new Error('the grid did not scroll');
      }
      if (args.mode === 'memory') {
        // As in PPTX Editor's memory bench, the first reading is after a warmup.
        // The warmup fills the undo history (`MEMORY_WARMUP_COMMITS`), in the same cells. It
        // collects garbage after each group, as the timed groups do before each reading: else the
        // garbage of 200 commits piles up, and the browser can pass the cap of the watchdog.
        const cdp = await page.context().newCDPSession(page);
        for (let group = 0; group < MEMORY_WARMUP_COMMITS / COMMITS_PER_GROUP; group += 1) {
          await editGroup(page, 1 + MEMORY_COLUMNS, COMMITS_PER_GROUP, false, warmupText(group));
          await cdp.send('HeapProfiler.collectGarbage');
        }
        const readings = [await readingAfterGc(page)];
        row.heapMiB = Math.round(readings[0]!.heapBytes / 1048576);
        const before = SNAPSHOTS ? await heapSummary(cdp) : null;
        for (let group = 0; group < MEMORY_GROUPS; group += 1) {
          const column = 1 + (group % MEMORY_COLUMNS);
          await editGroup(page, column, COMMITS_PER_GROUP, true, memoryText(group));
          readings.push(await readingAfterGc(page));
          // As PPTX Editor's memory bench prints each reading, so a long run shows its trend.
          const r = readings[readings.length - 1]!;
          console.error(
            `  memory ${name} ${(group + 1) * COMMITS_PER_GROUP}/` +
              `${MEMORY_GROUPS * COMMITS_PER_GROUP}: heap ` +
              `${(r.heapBytes / 1048576).toFixed(1)} MiB, nodes ${r.nodes}, listeners ` +
              `${r.listeners}, blobs ${r.blobs}, fonts ${r.fonts}, observers ${r.observers}`
          );
        }
        if (before || leaksOf(readings, COMMITS_PER_GROUP).length > 0) {
          const lines = before
            ? summaryDiff(before, await heapSummary(cdp))
            : ['  (no baseline snapshot: run with MEMORY_SNAPSHOTS=1)'];
          console.log(`  memory ${name}: heap growth by constructor, baseline to end`);
          console.log(lines.join('\n'));
        }
        await cdp.detach();
        const heap = readings.map((r) => r.heapBytes);
        row.growthKiBPerCommit = Math.round(growthOf(heap, COMMITS_PER_GROUP) / 1024);
        row.readings = readings;
      }
      await take(page);
    });
    row.peakMiB = Math.max(row.peakMiB, peakMiB);
  }
  return row;
}

/** The memory line and the leaks of each memory row. */
function leakChecks(rows: readonly Row[]): { lines: string[]; failures: string[] } {
  const lines: string[] = [];
  const failures: string[] = [];
  for (const row of rows) {
    if (!row.readings) continue;
    const first = row.readings[0]!;
    const last = row.readings[row.readings.length - 1]!;
    const change = RESOURCES.map((k) => `${k} ${first[k]} -> ${last[k]}`).join(', ');
    lines.push(
      `Memory on ${row.workbook}: DOM nodes ${first.nodes} -> ${last.nodes}, listeners ` +
        `${first.listeners} -> ${last.listeners}, ${change}.`
    );
    for (const leak of leaksOf(row.readings, COMMITS_PER_GROUP))
      failures.push(`memory ${row.workbook} leak: ${leak}`);
  }
  return { lines, failures };
}

function report(
  mode: Mode,
  rows: readonly Row[],
  leaks: ReturnType<typeof leakChecks>,
  stopped: readonly string[],
  about: string
): string[] {
  const lines = [
    `## Browser bench: ${mode}`,
    '',
    '| Workbook | Cells | Metric | Value |',
    '| --- | --: | --- | --: |',
  ];
  const round = (n: number) => Math.round(n * 10) / 10;
  for (const row of rows)
    for (const v of valuesOf(mode, row))
      lines.push(`| ${row.workbook} | ${row.cells} | ${v.metric} | ${round(v.value)} |`);
  lines.push('', about);
  for (const row of rows) lines.push(`Longest task on ${row.workbook}: ${row.longestTaskMs} ms.`);
  lines.push(...leaks.lines, '');
  if (stopped.length > 0) lines.push('Stopped:');
  for (const why of stopped) lines.push(`- ${why}`);
  if (leaks.failures.length > 0) lines.push('Leaks:');
  for (const f of leaks.failures) lines.push(`- ${f}`);
  if (leaks.failures.length === 0 && stopped.length === 0)
    lines.push(
      mode === 'memory' ? 'Each workbook ran, and no leak was found.' : 'Each workbook ran.'
    );
  else lines.push('The run fails.');
  return lines;
}

async function main(): Promise<void> {
  const args = parseArgs(process.argv.slice(2));
  if (!args) return;
  const rows: Row[] = [];
  // A workbook that stops the run is a failure, and the other workbooks still run.
  const stopped: string[] = [];
  const dir = mkdtempSync(join(tmpdir(), 'xlsx-bench-'));
  let server: Awaited<ReturnType<typeof createServer>> | undefined;
  try {
    server = await createServer({
      configFile: join(DEMO, 'vite.config.ts'),
      root: DEMO,
      server: { host: '127.0.0.1', port: 5299, strictPort: false },
      logLevel: 'error',
    });
    await server.listen();
    const base = server.resolvedUrls?.local[0] ?? 'http://127.0.0.1:5299/';
    for (const name of args.workbooks) {
      // Progress goes to stderr, as in `edit-bench.ts`. Stdout holds the report, and the heap
      // growth by constructor of the memory mode.
      console.error(`run    ${args.mode} ${name}`);
      try {
        rows.push(await runWorkbook(name, args, base, dir));
      } catch (error) {
        console.error(error);
        stopped.push(`${args.mode} ${name} stopped: ${(error as Error).message.split('\n')[0]}`);
      }
    }
  } finally {
    await server?.close().catch(() => undefined);
    rmSync(dir, { recursive: true, force: true });
  }
  const leaks = leakChecks(rows);
  const sizes = [
    args.workbooks.includes('grid') ? `${args.rows} rows` : '',
    args.workbooks.includes('sheets') ? `${BROWSER_SIZES.sheets} sheets` : '',
    args.workbooks.includes('strings') ? `${args.strings} strings` : '',
  ].filter(Boolean);
  const about =
    `Sizes: ${sizes.join(', ')}. ` + `Browser: Chromium ${browserVersion ?? '(none started)'}.`;
  console.log(report(args.mode, rows, leaks, stopped, about).join('\n'));
  if (args.out) {
    const values = rows.map((r) => ({ workbook: r.workbook, values: valuesOf(args.mode, r) }));
    const saved = { schema: 1, args, browserVersion, values, rows, leaks: leaks.failures, stopped };
    writeFileSync(args.out, `${JSON.stringify(saved, null, 2)}\n`);
  }
  process.exit(stopped.length + leaks.failures.length > 0 ? 1 : 0);
}

if (import.meta.main) await main().catch(stopOnUsageError);
