// End-to-end smoke check of the demo in Playwright Chromium.
//
//   bun run test:e2e                       Chromium
//   bun run test:e2e --browser firefox     another engine (repeat the option for more)
//
// The flows of the demo and of the editor run as happy-dom tests in `bun run test`: the start,
// the focus regions, view only, the theme, the saves, the number formats, typed entry, the
// headings switch, the questions before unsaved edits are lost, the work that overlaps a load,
// the download logic, and each fixture's open, edit, save, and reopen
// (`examples/vite/src/__tests__/demo-*.test.tsx`), and the editor tests of
// `packages/core/src/editor/__tests__/` and the adapters. This check keeps one short smoke run
// and the steps that need a real browser engine: paint and loaded fonts, the system clipboard
// and its events, focus that the engine moves, layout and computed styles, and a real download.
//
// It starts the demo's dev server once, and runs these steps in each engine, on one page of a
// fresh browser, one result line for each:
//
// - The sample workbook opens at start, and every sheet shows. Each sheet paints dark pixels on
//   the grid canvas, and no font face fails to load (`verify-demo-engine.ts`).
// - A value typed into an empty cell saves from the header's Save, the saved copy keeps the
//   value and every sheet, and the copy opens again from a picked file with the value.
// - The formula bar shows the whole line of the caret after line breaks, its own undo removes
//   one line, its line count does not cover the text, and a line of tabs keeps the caret in view
//   (`verify-demo-editing.ts`).
// - A press on the line count puts the caret at the end of the shown line, and the bar scrolls
//   to show it: on a long line, a line shorter than the next, a line of tabs, and in a scaled bar.
// - A copy of a cell with Control+C or Command+C, pasted with Control+V or Command+V, reaches
//   the grid as clipboard events with the private clipboard type, and undo takes it back
//   (`verify-demo-engine.ts`).
// - A cut of a cell with Control+X or Command+X, pasted into another cell, moves the value, and
//   undo takes it back. A headless Firefox run skips this step (`HEADLESS_SKIPS`).
// - The dashed outline of a pending copy is 2 pixels, with no animation, follows the scroll,
//   keeps its line with forced colors, leaves a print, and Escape ends it
//   (`verify-demo-clipboard-mark.ts`).
// - A copy right after a press on a cell, or after Shift+F6 into the grid, with no page
//   selection, also reaches the grid.
// - The grid keeps the focus for typing, Enter, copy, and paste inside editable host content,
//   and after a press with host text selected, in both layouts. Editing commands do not change
//   the grid or move the focus, and the formula bar and input method text still reach the cells
//   (`verify-demo-host.ts`).
// - On a page of its own, a host CSS reset changes no chrome style, with each popup and the
//   number format dialog open, and the chrome stylesheet changes no host style. In a host that
//   moves, scales, or clips the editor, the toolbar shows the groups of its layout width, a
//   toolbar list opens under its trigger inside the visible box, a press selects the cell under
//   the pointer, and the sheet tab menu opens at the pointer (`verify-demo-embedding.ts`).
// - On a page of its own, in hosts that size the editor by its content, the editor width stays
//   within the window, and the status strip keeps its width and its text in a host flex row
//   (`verify-demo-content-sized.ts`).
// - On a page of its own with `dir="rtl"` on the root element, a horizontal wheel scrolls the
//   grid and the headings follow, and a jump to AZ1 from the name box shows AZ. With a
//   right-to-left catalog (`?lang=ar`), the number format list and the tab menu open at the
//   right edge of their trigger (`verify-demo-rtl.ts`).
// - On a new page with no file save picker, the header's Save downloads one file whose bytes
//   hold the edit and every sheet. A phone-wide bar shows Save and names each control, the page
//   does not scroll sideways, and the text of a notice and of an error starts at the padding of
//   the message bar (`verify-demo-download.ts`).
//
// Any console error or page error fails the step, and so does an alert (the demo's or the
// editor's) that is still on the page when the step ends. A step that shows an alert on
// purpose closes it.
//
// To debug a failure, fix the first FAILED step first: the steps run in order on one page, so
// a failed step can make later steps fail too. Each failed step writes a screenshot of the
// page to `screenshots/demo-<engine>-<step>.png` (CI uploads the folder). Set `DEMO_HEADED=1`
// to watch the run in a browser window.
//
// The steps are the same in each engine. A step that an engine cannot pass with no window is
// listed in `HEADLESS_SKIPS` (`verify-demo-browser.ts`) with its cause, and its result line names
// that cause: no step is skipped silently. Native behavior that differs by engine, such as the
// undo steps of a text field, is checked with the rule of each engine, in the step. Firefox and
// WebKit have no file save picker, so the demo downloads the saved copy there. The step
// `save with no picker` proves that download in every engine. The other steps save through the
// stand-in picker in every engine, so they run the same way and read the same bytes.
//
// PPTX Editor runs `test:e2e` as Playwright Test specs in Chromium only. Here the checks are one
// script whose steps run in order on one page, so a flag picks the engines, and Chromium runs
// when there is none. DOCX Editor has no `test:e2e` script of this kind: its Playwright
// configurations run Chromium.
//
// Same shape as pptx-editor/scripts/verify-demo.ts at 79a5d22d (EigenPal). Differences: the
// demo saves through the file save picker, so the page of most steps gets a stand-in picker that
// keeps the bytes; it fails on the text of any alert on the page; and there is no private corpus
// mode and no parity screenshot mode.

import type { Browser, Page } from '@playwright/test';
import { mkdirSync, readdirSync, readFileSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import { createServer } from 'vite';
import { cellValue, sheetStale } from '../packages/core/src/cascade/index.ts';
import {
  cellKey,
  parseCellRef,
  readXlsx,
  sheetByName,
  sheetIndex,
  workbookIndex,
  xmlPartOf,
  type XlsxPackage,
} from '../packages/core/src/store/index.ts';
import { SAMPLE_BOLD_EDIT } from './sample-workbook/edits.ts';
import { SHEET_GUIDE } from './sample-workbook/showcase-sheets.ts';
import { checkPaint, verifyClipboard, verifyFocusCopy } from './verify-demo-engine.ts';
import { verifyClipboardMark } from './verify-demo-clipboard-mark.ts';
import { verifyContentSizedHosts } from './verify-demo-content-sized.ts';
import { verifyEmbedding } from './verify-demo-embedding.ts';
import { verifyRightToLeftHost } from './verify-demo-rtl.ts';
import { verifyEditableHost } from './verify-demo-host.ts';
import {
  alertNote,
  engineSummary,
  pageAlerts,
  runEngines,
  STEP_LIMIT_MS,
  stopReason,
  TimeLimitError,
  withTimeLimit,
  HEADED,
  runPasses,
  unknownSkips,
  launchProblem,
  launchRemedy,
  ENGINES,
  type EngineResult,
  goTo,
  skipIssue,
  skipCounter,
  stepResult,
  optionsOrExit,
  problemOf,
  screenshotOf,
  showsSheet,
  stubSavePicker,
  WAIT_MS,
  type Engine,
} from './verify-demo-browser.ts';
import { verifyDownloadSave } from './verify-demo-download.ts';
import { verifyFormulaBarLayout, verifyLineCountPress } from './verify-demo-editing.ts';

const ROOT = join(import.meta.dir, '..');
const DEMO = join(ROOT, 'examples/vite');
/** The committed sample workbook, which a step opens again as a picked file. */
const SAMPLE_BYTES = readFileSync(join(DEMO, 'public/sample.xlsx'));

const [SAMPLE_SHEET] = SHEET_GUIDE[0]!;
/** The sheet that the steps after `sample open` start on: the last sheet that it visits. */
const LAST_SHEET = SHEET_GUIDE.at(-1)![0];
/** The empty cell of the last sheet that the smoke edit fills. No other step reads it. */
const SMOKE_REF = 'AB400';

/** The number of copies that the stand-in picker kept. */
const savedCount = (page: Page): Promise<number> =>
  page.evaluate(() => (window as unknown as { __saved: unknown[] }).__saved.length);

/** Saves through the demo's own Save button in its header, and returns the kept bytes. */
async function save(page: Page): Promise<Uint8Array> {
  const before = await savedCount(page);
  await page.locator('.demo-topbar').getByRole('button', { name: 'Save', exact: true }).click();
  await page
    .waitForFunction(
      (n) => (window as unknown as { __saved: unknown[] }).__saved.length > n,
      before,
      { timeout: WAIT_MS }
    )
    .catch(async () => {
      throw new Error(`the Save wrote no copy${await alertNote(page)}`);
    });
  const saved = await page.evaluate(() => {
    const list = (window as unknown as { __saved: number[][] }).__saved;
    return list[list.length - 1]!;
  });
  return new Uint8Array(saved);
}

/** The stored text of the cell `ref` of `sheet` in `pkg`, or `undefined` for no cell. */
function textAt(pkg: XlsxPackage, sheet: string, ref: string): string | undefined {
  const index = workbookIndex(pkg)!;
  const part = xmlPartOf(pkg, sheetByName(index, sheet)?.partName);
  const at = parseCellRef(ref)!;
  const cell = part ? sheetIndex(part).cells.get(cellKey(at.row, at.col)) : undefined;
  if (!part || !cell) return undefined;
  const { value } = cellValue(pkg, index, cell.el, sheetStale(index, part.root));
  return value.kind === 'string' ? value.text : JSON.stringify(value);
}

/** The confirm dialogs that the page showed. Each one gets a yes. */
const dialogs = { seen: [] as string[] };

/** The sample opens at start, and every sheet shows and paints. */
async function verifySampleOpen(page: Page, url: string): Promise<string> {
  const t0 = performance.now();
  await page.goto(url);
  if (!(await showsSheet(page, SAMPLE_SHEET))) throw new Error(`no sheet ${SAMPLE_SHEET}`);
  const opened = Math.round(performance.now() - t0);
  for (const [name] of SHEET_GUIDE) {
    await page.getByRole('tab', { name, exact: true }).click();
    if (!(await showsSheet(page, name))) throw new Error(`tab ${name} shows no sheet`);
    await checkPaint(page).catch((cause: Error) => {
      throw new Error(`on ${name}, ${cause.message}`);
    });
  }
  return `opened in ${opened} ms, ${SHEET_GUIDE.length} sheets visited and painted`;
}

/**
 * A value typed into an empty cell saves, and the saved copy opens again from a picked file with
 * the value. The step ends on the last sheet, where the steps after it start.
 */
async function verifyTypeSaveReopen(page: Page): Promise<string> {
  const was = await goTo(page, SMOKE_REF);
  if (was !== '') throw new Error(`${SMOKE_REF} of the sample holds "${was}", not an empty cell`);
  const engine = page.context().browser()?.browserType().name() ?? 'browser';
  const text = `Typed in ${engine}`;
  await page.locator('.xlsx-scroller').first().focus();
  await page.keyboard.type(text);
  await page.keyboard.press('Enter');
  if ((await goTo(page, SMOKE_REF)) !== text) throw new Error(`${SMOKE_REF} did not take the text`);
  const bytes = await save(page);
  const read = readXlsx(bytes);
  if (!read.ok) throw new Error(`the saved copy does not read (${read.error.reason})`);
  const names = workbookIndex(read.pkg)?.sheets.map((s) => s.name);
  if (JSON.stringify(names) !== JSON.stringify(SHEET_GUIDE.map(([n]) => n)))
    throw new Error(`the saved copy has the sheets ${JSON.stringify(names)}`);
  if (textAt(read.pkg, LAST_SHEET, SMOKE_REF) !== text)
    throw new Error(`the saved copy has no "${text}" at ${SMOKE_REF}`);
  await page.getByRole('button', { name: 'Close the message' }).click();
  // The saved copy holds the edit, so the open asks nothing. The sample opens first, from a
  // picked file, so the grid shows the last sheet again only after the copy loads.
  const asked = dialogs.seen.length;
  const fileInput = page.locator('.demo-topbar input[type=file]');
  const mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  await fileInput.setInputFiles({ name: 'sample.xlsx', mimeType, buffer: SAMPLE_BYTES });
  if (!(await showsSheet(page, SAMPLE_SHEET))) throw new Error('the sample did not open again');
  await fileInput.setInputFiles({
    name: 'sample-edited.xlsx',
    mimeType,
    buffer: Buffer.from(bytes),
  });
  // The top bar names the file, with its extension apart, once the editor shows it.
  await page.waitForFunction(
    () =>
      document.querySelector('.demo-topbar .demo-file__name')?.textContent === 'sample-edited' &&
      document.querySelector('.demo-topbar .demo-file__ext')?.textContent === '.xlsx',
    null,
    { timeout: WAIT_MS }
  );
  await page.getByRole('tab', { name: LAST_SHEET, exact: true }).click();
  if (!(await showsSheet(page, LAST_SHEET))) throw new Error('the saved copy did not open again');
  const reopened = await goTo(page, SMOKE_REF);
  if (reopened !== text) throw new Error(`the saved copy opened with "${reopened}" in the bar`);
  if (dialogs.seen.length !== asked) throw new Error('the demo asked with no unsaved edit');
  return `typed "${text}" into ${SMOKE_REF}, saved it, and opened the saved copy again with it`;
}

/** How long a browser can take to close after a lost page. */
const BROWSER_CLOSE_MS = 10_000;

/** Runs the steps in one engine, on one page. */
async function runEngine(engine: Engine, url: string): Promise<EngineResult> {
  console.log(`== ${engine}`);
  dialogs.seen.length = 0;
  let browser: Browser;
  try {
    browser = await ENGINES[engine].launch({ headless: !HEADED });
  } catch (cause) {
    console.log(`launch: FAILED, ${launchProblem(cause)}`);
    console.log(`  ${launchRemedy(engine, launchProblem(cause))}`);
    return { passed: false, skipped: 0, stopped: 'did not run: the browser did not launch' };
  }
  const page = await browser.newPage({ viewport: { width: 1600, height: 900 } });
  await page.addInitScript(stubSavePicker);
  const errors: string[] = [];
  /** Counts each console error and page error of `watched` against the running step. */
  const watch = (watched: Page) => {
    watched.on('console', (m) => {
      if (m.type() === 'error') errors.push(m.text());
    });
    watched.on('pageerror', (e) => errors.push(e.message));
  };
  watch(page);
  page.on('dialog', (d) => {
    dialogs.seen.push(d.message());
    void d.accept();
  });
  const skips = skipCounter(engine);
  let anyFailed = false;
  let firstFailed: string | undefined;
  // The step at which the page closed or crashed. No later step can run on that page.
  let lostAt: string | null = null;
  let lostWhy = 'the page was lost';
  /**
   * Prints one result. An alert on the page or a new console error also fails it. True when
   * the result passes.
   */
  const report = async (name: string, problem: string | null, detail: string, since: number) => {
    const { shown, lost } = await pageAlerts(page);
    if (lost) lostAt = name;
    const newErrors = errors.slice(since);
    const { line, passed } = stepResult(name, problem ?? lost, shown, newErrors.length, detail);
    console.log(line);
    for (const e of newErrors.slice(0, 3)) console.log(`  ${e}`);
    return passed;
  };
  const names = new Set<string>();
  // A step that opens its own page hands it over here. A failed step then takes the screenshot
  // of that page, and the page closes after the result.
  let ownPage: { readonly page: Page; readonly close: () => Promise<void> } | null = null;
  const adopt = (own: Page, close: () => Promise<void>) => {
    ownPage = { page: own, close };
  };
  /** Runs one step and prints its result. True when the step passes or is a known skip. */
  const step = async (name: string, run: () => Promise<string>): Promise<boolean> => {
    names.add(name);
    if (lostAt !== null) return false;
    const issue = skipIssue(engine, name);
    if (issue) {
      skips.skip(name, issue);
      return true;
    }
    const since = errors.length;
    let passed: boolean;
    try {
      passed = await report(
        name,
        null,
        await withTimeLimit(run(), STEP_LIMIT_MS, 'the step'),
        since
      );
    } catch (cause) {
      // The hung work can still hold the page, so no later step runs on it.
      if (cause instanceof TimeLimitError) {
        lostAt = name;
        lostWhy = 'the step passed its time limit';
      }
      passed = await report(name, problemOf(cause), '', since);
    }
    anyFailed ||= !passed;
    if (!passed) firstFailed ??= name;
    const shown = ownPage as { page: Page; close: () => Promise<void> } | null;
    if (!passed) await screenshotOf(shown?.page ?? page, `${engine}-${name}`);
    await shown?.close().catch(() => {});
    ownPage = null;
    return passed;
  };
  try {
    // Every later step starts from the sample, so none runs when it does not open.
    const opened = await step('sample open', () => verifySampleOpen(page, url));
    if (!opened) {
      lostAt = 'sample open';
      lostWhy = 'the sample did not open';
      console.log(`later steps: skipped, because ${lostWhy}`);
    } else {
      await step('type, save, and reopen', () => verifyTypeSaveReopen(page));
      await step('formula bar layout', () => verifyFormulaBarLayout(page));
      await step('line count press', () => verifyLineCountPress(page));
      await step('copy and paste', () => verifyClipboard(page, 'copy'));
      await step('cut and paste', () => verifyClipboard(page, 'cut'));
      await step('clipboard outline', () => verifyClipboardMark(page));
      await step('copy after the focus', () => verifyFocusCopy(page));
      await step('editable host content', () => verifyEditableHost(page));
      await step('embedding', () => verifyEmbedding(browser, url, watch, adopt, SAMPLE_SHEET));
      await step('content-sized host', () =>
        verifyContentSizedHosts(browser, url, watch, adopt, SAMPLE_SHEET)
      );
      await step('right-to-left host', () =>
        verifyRightToLeftHost(browser, url, watch, adopt, SAMPLE_SHEET)
      );
      await step('save with no picker', () =>
        verifyDownloadSave(browser, url, watch, adopt, {
          sheet: SAMPLE_BOLD_EDIT.sheet,
          ref: SAMPLE_BOLD_EDIT.ref,
          firstSheet: SAMPLE_SHEET,
        })
      );
      if (lostAt !== null) console.log(`later steps: skipped, because ${lostWhy} at "${lostAt}"`);
      // A skip that names no step of a run would skip nothing, and hide its issue.
      const unknown = unknownSkips(engine, names);
      if (unknown.length) {
        console.log(`skip list: FAILED, HEADLESS_SKIPS names no step "${unknown.join('", "')}"`);
        anyFailed = true;
      }
    }
  } finally {
    await withTimeLimit(browser.close(), BROWSER_CLOSE_MS, 'the browser close').catch(() => {});
  }
  return {
    passed: !anyFailed,
    skipped: skips.count,
    ...(firstFailed ? { firstFailed } : {}),
    ...stopReason(lostAt, lostWhy),
  };
}

async function main(): Promise<void> {
  const { engines } = optionsOrExit(process.argv.slice(2));
  const server = await createServer({
    configFile: join(DEMO, 'vite.config.ts'),
    root: DEMO,
    // One address for the port check and the browser: on `localhost`, another server on the
    // other address family can take the page.
    server: { host: '127.0.0.1', port: 5299, strictPort: false },
    logLevel: 'error',
  });
  await server.listen();
  const url = server.resolvedUrls?.local[0] ?? 'http://127.0.0.1:5299/';
  // The screenshots of an earlier run go, so each one in the folder is from this run.
  const shots = join(ROOT, 'screenshots');
  mkdirSync(shots, { recursive: true });
  for (const f of readdirSync(shots, { withFileTypes: true }).filter((e) =>
    e.name.startsWith('demo-')
  ))
    rmSync(join(shots, f.name));
  // One engine at a time, so a run never holds two browsers.
  let results: [Engine, EngineResult][];
  try {
    results = await runEngines(engines, (engine) => runEngine(engine, url));
  } finally {
    await server.close();
  }
  // The summary names what each engine did not prove, so a skip never shows only in a step line.
  for (const [engine, r] of results) console.log(engineSummary(engine, r));
  process.exit(runPasses(results) ? 0 : 1);
}

await main();
