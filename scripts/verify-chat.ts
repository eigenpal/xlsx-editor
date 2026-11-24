// The happy-path acceptance check of the chat example (`examples/chat`) in Playwright Chromium.
//
//   bun run test:e2e:chat
//
// The chat example embeds the editor in a panel beside a chat, at about half the window, as a
// host app does. This check proves the path that a user of such a host takes, in a real browser
// engine: layout at the panel's widths, focus that moves between the chat and the editor,
// keyboard ownership between the host and the editor, theming, and a real download. The
// assistant's command map has its own unit test (`examples/chat/src/assistant.test.ts`).
//
// It starts the example's dev server once, and runs these steps in order on one page, one result
// line for each. Any console error or page error fails the step. A failed step writes a
// screenshot to `screenshots/chat-<step>.png`.
//
// - The page opens with the workbook card. The card lists the sheet tabs, and the host's first
//   sheet (Sales) is the shown sheet. The card paints a preview of the sheet.
// - Open slides the panel in beside the chat at half the window. The chat stays usable, and the
//   panel's chrome wears the chat's accent through the `--doc-*` tokens.
// - A chat command writes a cell through the editor API, the card shows "Edited", and its
//   preview paints the edit.
// - A press on the card of the open workbook puts the keyboard in the grid: typing edits a cell.
// - Escape in an open cell edit cancels the edit and keeps the panel. Escape in the chat closes
//   the panel and keeps the focus in the chat. Close gives the focus back to the card.
// - In view only, a chat edit is refused, and the chat shows the editor's own reason.
// - Download and Control+S (Command+S on macOS) in the grid each download one workbook whose
//   bytes hold the chat's edit.
// - Expand covers the chat, which turns inert, and Show chat brings it back.
// - The divider moves with the arrow keys.
// - Escape after a pointer opens a toolbar list, with the keyboard still in the chat, closes only
//   the list.
// - At a 1024 px window, the half panel keeps Bold in the toolbar, not only in More.
// - On a phone-wide window, the panel is a full-window sheet with Back, and the page does not
//   scroll sideways.
//
// Set `CHAT_HEADED=1` to watch the run in a browser window.

import { chromium, type Page } from '@playwright/test';
import { mkdirSync, readFileSync } from 'node:fs';
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
} from '../packages/core/src/store/index.ts';

const ROOT = join(import.meta.dir, '..');
const EXAMPLE = join(ROOT, 'examples/chat');
const SHOTS = join(ROOT, 'screenshots');
const WAIT_MS = 10_000;
const MOD = process.platform === 'darwin' ? 'Meta' : 'Control';

interface EditorProbe {
  readonly status: string;
  readonly sheet: string | null;
  readonly ref: string | null;
  readonly content: string | null;
  readonly revision: number;
}

/** The editor state that the example exposes on `window.__xlsxEditor` (its `onReady`). */
const probe = (page: Page): Promise<EditorProbe> =>
  page.evaluate(() => {
    const editor = (window as unknown as { __xlsxEditor?: { snapshot(): any } }).__xlsxEditor;
    const s = editor?.snapshot();
    return {
      status: s?.status ?? 'none',
      sheet: s?.sheets[s.activeSheet]?.name ?? null,
      ref: s?.activeCell?.ref ?? null,
      content: s?.activeCell?.content ?? null,
      revision: s?.revision ?? -1,
    };
  });

const panelOpen = async (page: Page): Promise<boolean> =>
  (await page.getAttribute('.shell', 'data-open')) !== null;

/** The pixels of the card's sheet preview, as a data URL. */
const previewPixels = (page: Page): Promise<string> =>
  page.evaluate(() => document.querySelector<HTMLCanvasElement>('.card__grid canvas')!.toDataURL());

/** Opens the panel from the card when an earlier step left it closed. */
async function ensureOpen(page: Page): Promise<void> {
  if (await panelOpen(page)) return;
  await page.click('.card');
  await page.waitForTimeout(500);
}

/** The element that holds the focus, as `tag.class`. */
const focused = (page: Page): Promise<string> =>
  page.evaluate(() => {
    const a = document.activeElement;
    return a ? `${a.tagName.toLowerCase()}.${(a as HTMLElement).className}` : 'none';
  });

function check(ok: boolean, what: string): void {
  if (!ok) throw new Error(what);
}

/** Sends `text` from the composer and waits until the reply has streamed in. */
async function say(page: Page, text: string): Promise<string> {
  const replies = await page.locator('.msg--assistant').count();
  await page.locator('.composer__input').fill(text);
  await page.keyboard.press('Enter');
  await page.locator('.msg--assistant').nth(replies).waitFor({ timeout: WAIT_MS });
  // The send button turns ready again when the reply has streamed in and the composer has text.
  await page.waitForFunction(
    () => !document.querySelector('.msg__thinking') && !!document.querySelector('.composer__input'),
    null,
    { timeout: WAIT_MS }
  );
  await page.waitForTimeout(700);
  return (await page.locator('.msg--assistant .msg__text').last().textContent()) ?? '';
}

/** The stored text of `ref` on `sheet` in workbook `bytes`. */
function storedText(bytes: Uint8Array, sheet: string, ref: string): string | undefined {
  const read = readXlsx(bytes);
  if (!read.ok) throw new Error(`the download does not read (${read.error.reason})`);
  const index = workbookIndex(read.pkg)!;
  const part = xmlPartOf(read.pkg, sheetByName(index, sheet)?.partName);
  const at = parseCellRef(ref)!;
  const cell = part ? sheetIndex(part).cells.get(cellKey(at.row, at.col)) : undefined;
  if (!part || !cell) return undefined;
  const { value } = cellValue(read.pkg, index, cell.el, sheetStale(index, part.root));
  return value.kind === 'number' ? String(value.value) : JSON.stringify(value);
}

/** Runs `act`, and returns the bytes of the one download that it starts. */
async function downloadOf(page: Page, act: () => Promise<void>): Promise<Uint8Array> {
  const download = page.waitForEvent('download', { timeout: WAIT_MS });
  await act();
  const file = await download;
  check(
    file.suggestedFilename() === 'Showcase.xlsx',
    `the file name is ${file.suggestedFilename()}`
  );
  return new Uint8Array(readFileSync(await file.path()));
}

type Step = readonly [name: string, run: (page: Page) => Promise<string | void>];

const STEPS: readonly Step[] = [
  [
    'opens with the workbook card',
    async (page) => {
      await page.waitForFunction(
        () => (window as any).__xlsxEditor?.snapshot().status === 'ready',
        null,
        { timeout: 30_000 }
      );
      await page.locator('.card').waitFor({ timeout: WAIT_MS });
      const tabs = await page.locator('.card__tab').allTextContents();
      check(tabs.slice(0, 2).join() === 'Welcome,Sales', `the card tabs are ${tabs.join()}`);
      const active = await page.locator('.card__tab[data-active]').textContent();
      check(active === 'Sales', `the card marks ${active}`);
      const preview = await page.locator('.card__grid canvas').boundingBox();
      check((preview?.width ?? 0) > 100, 'the card paints no sheet preview');
      check(!(await panelOpen(page)), 'the panel is open at start');
    },
  ],
  [
    'the card opens the panel beside the chat',
    async (page) => {
      await page.click('.card');
      await page.waitForTimeout(500);
      check(await panelOpen(page), 'the panel did not open');
      const [chat, panel] = await Promise.all([
        page.locator('.shell__chat').boundingBox(),
        page.locator('.shell__panel').boundingBox(),
      ]);
      const width = page.viewportSize()!.width;
      check(!!chat && !!panel, 'the chat or the panel has no box');
      check(Math.abs(panel!.width - width / 2) <= 2, `the panel is ${panel!.width} px wide`);
      check(chat!.x + chat!.width <= panel!.x + 1, 'the chat and the panel overlap');
      check(
        (await page.getAttribute('.shell__chat', 'inert')) === null,
        'the chat is inert beside the panel'
      );
      check(
        (await page.locator('.book .xlsx-scroller').getAttribute('aria-label')) ===
          'Worksheet Sales',
        'the grid does not show Sales'
      );
      // The editor chrome wears the chat accent: the pressed state of a toolbar toggle.
      const accent = await page.evaluate(() =>
        getComputedStyle(document.querySelector('.book')!).getPropertyValue('--doc-primary').trim()
      );
      const shell = await page.evaluate(() =>
        getComputedStyle(document.documentElement).getPropertyValue('--c-accent').trim()
      );
      check(accent === shell && accent !== '', `--doc-primary is ${accent}, not ${shell}`);
      return `panel ${Math.round(panel!.width)} of ${width} px`;
    },
  ],
  [
    'a chat command writes a cell',
    async (page) => {
      const before = await previewPixels(page);
      const text = await say(page, 'set B5 to 1200');
      check(text.startsWith('I wrote 1200 in B5'), `the reply is "${text}"`);
      const s = await probe(page);
      check(s.ref === 'B5' && s.content === '1200', `B5 holds ${s.content}`);
      check(
        ((await page.locator('.card__meta').textContent()) ?? '').includes('Edited'),
        'the card does not show Edited'
      );
      check((await previewPixels(page)) !== before, 'the card preview does not paint the edit');
    },
  ],
  [
    'the card puts the keyboard in the grid',
    async (page) => {
      const before = (await probe(page)).revision;
      await page.click('.card');
      check(
        (await focused(page)) === 'div.xlsx-scroller',
        `the focus is on ${await focused(page)}`
      );
      await page.keyboard.type('7');
      await page.keyboard.press('Enter');
      const s = await probe(page);
      check(s.revision === before + 1, `the revision is ${s.revision}, not ${before + 1}`);
      await page.keyboard.press(`${MOD}+z`);
    },
  ],
  [
    'Escape belongs to the grid, then to the chat',
    async (page) => {
      await page.keyboard.type('9');
      check((await probe(page)).content !== null, 'no cell edit opened');
      await page.keyboard.press('Escape');
      check(await panelOpen(page), 'Escape in a cell edit closed the panel');
      await page.click('.composer__input');
      await page.keyboard.press('Escape');
      await page.waitForTimeout(300);
      check(!(await panelOpen(page)), 'Escape in the chat did not close the panel');
      // The focus was in the chat, so it stays there.
      check((await focused(page)).startsWith('textarea'), `the focus is on ${await focused(page)}`);
      await page.click('.card');
      await page.waitForTimeout(400);
      await page.click('.book__head [aria-label="Close"]');
      await page.waitForTimeout(300);
      check(
        (await focused(page)) === 'button.card',
        `Close left the focus on ${await focused(page)}`
      );
      await page.click('.card');
      await page.waitForTimeout(400);
    },
  ],
  [
    'view only refuses a chat edit with the editor reason',
    async (page) => {
      await say(page, 'make it view only');
      check(
        (await page.getAttribute('.book__head [aria-label="View only"]', 'aria-pressed')) ===
          'true',
        'View only is not pressed'
      );
      const text = await say(page, 'set B5 to 9');
      check(text.startsWith('I could not do that.'), `the reply is "${text}"`);
      check((await probe(page)).content === '1200', 'the refused edit changed B5');
      await say(page, 'allow edits');
    },
  ],
  [
    'Download and the save keys download the edits',
    async (page) => {
      const fromButton = await downloadOf(page, () =>
        page.click('.book__head [aria-label="Download .xlsx"]')
      );
      check(storedText(fromButton, 'Sales', 'B5') === '1200', 'the download lacks B5');
      await page.click('.card');
      const fromKeys = await downloadOf(page, () => page.keyboard.press(`${MOD}+s`));
      check(storedText(fromKeys, 'Sales', 'B5') === '1200', 'the save key download lacks B5');
    },
  ],
  [
    'Expand covers the chat, and Show chat brings it back',
    async (page) => {
      await page.click('.book__head [aria-label="Expand"]');
      await page.waitForTimeout(500);
      check((await page.getAttribute('.shell__chat', 'inert')) !== null, 'the chat is not inert');
      const panel = await page.locator('.shell__panel').boundingBox();
      check(panel!.width === page.viewportSize()!.width, `the panel is ${panel!.width} px wide`);
      await page.click('.book__head [aria-label="Show chat"]');
      await page.waitForTimeout(500);
      check((await page.getAttribute('.shell__chat', 'inert')) === null, 'the chat stays inert');
    },
  ],
  [
    'the divider moves with the keyboard',
    async (page) => {
      const before = (await page.locator('.shell__panel').boundingBox())!.width;
      await page.focus('.shell__handle');
      for (let i = 0; i < 3; i++) await page.keyboard.press('ArrowLeft');
      await page.waitForTimeout(400);
      const after = (await page.locator('.shell__panel').boundingBox())!.width;
      check(after > before + 40, `the panel went from ${before} to ${after} px`);
      await page.dblclick('.shell__handle');
      await page.waitForTimeout(400);
    },
  ],
  [
    'Escape in a pointer-opened list closes only the list',
    async (page) => {
      await page.click('.composer__input');
      await page.click('.book__toolbar [aria-label^="Font size"]');
      await page.locator('[role="listbox"]').first().waitFor({ timeout: WAIT_MS });
      await page.keyboard.press('Escape');
      await page.waitForTimeout(300);
      check((await page.locator('[role="listbox"]').count()) === 0, 'the list stays open');
      check(await panelOpen(page), 'Escape for the list also closed the panel');
    },
  ],
  [
    'a half panel at 1024 px keeps Bold in the toolbar',
    async (page) => {
      await ensureOpen(page);
      await page.setViewportSize({ width: 1024, height: 768 });
      await page.waitForTimeout(600);
      const bold = page.locator('.book__toolbar [aria-label="Bold"]');
      check((await bold.count()) === 1 && (await bold.isVisible()), 'Bold is not in the bar');
    },
  ],
  [
    'a phone shows the panel as a sheet with Back',
    async (page) => {
      await ensureOpen(page);
      await page.setViewportSize({ width: 390, height: 844 });
      await page.waitForTimeout(600);
      const panel = await page.locator('.shell__panel').boundingBox();
      check(panel!.width === 390 && panel!.x === 0, 'the panel does not cover the window');
      const sideways = await page.evaluate(
        () => document.documentElement.scrollWidth > document.documentElement.clientWidth
      );
      check(!sideways, 'the page scrolls sideways');
      await page.click('.book__head [aria-label="Back to chat"]');
      await page.waitForTimeout(400);
      check(!(await panelOpen(page)), 'Back did not close the sheet');
    },
  ],
];

async function main(): Promise<void> {
  const server = await createServer({
    configFile: join(EXAMPLE, 'vite.config.ts'),
    root: EXAMPLE,
    server: { host: '127.0.0.1', port: 5298, strictPort: false },
    logLevel: 'error',
  });
  await server.listen();
  const url = server.resolvedUrls?.local[0] ?? 'http://127.0.0.1:5298/';
  mkdirSync(SHOTS, { recursive: true });
  const browser = await chromium.launch({ headless: process.env.CHAT_HEADED !== '1' });
  let failed = 0;
  try {
    const context = await browser.newContext({
      acceptDownloads: true,
      viewport: { width: 1440, height: 900 },
      reducedMotion: 'reduce',
    });
    const page = await context.newPage();
    const errors: string[] = [];
    page.on('console', (m) => {
      if (m.type() === 'error') errors.push(m.text());
    });
    page.on('pageerror', (e) => errors.push(e.message));
    await page.goto(url);
    for (const [name, run] of STEPS) {
      errors.length = 0;
      try {
        const note = await run(page);
        if (errors.length) throw new Error(`console: ${errors.join(' | ')}`);
        console.log(`${name}: ok${note ? ` (${note})` : ''}`);
      } catch (cause) {
        failed++;
        console.log(`${name}: FAILED, ${cause instanceof Error ? cause.message : String(cause)}`);
        const file = name.replace(/[^a-z0-9]+/gi, '-').toLowerCase();
        await page.screenshot({ path: join(SHOTS, `chat-${file}.png`) }).catch(() => {});
      }
    }
  } finally {
    await browser.close();
    await server.close();
  }
  console.log(
    failed ? `${failed} of ${STEPS.length} steps FAILED` : `${STEPS.length} steps passed`
  );
  process.exit(failed ? 1 : 0);
}

await main();
