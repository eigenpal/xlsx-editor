// The save of the demo in a browser with no file save picker, on its own page with no stand-in
// picker. It removes the native picker of Chromium, so each engine takes the path that Firefox
// and WebKit take. It reads the bytes of a real download, and it checks the layout of the top
// bar in a phone-wide window and of the message bar. The start state of Save, each way to save
// and to repeat a save, a retry after a failed save, and the notice and the question before the
// page closes are happy-dom tests (`examples/vite/src/__tests__/demo-download.test.tsx`).

import type { Browser, Page } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { cellFormat } from '../packages/core/src/cascade/index.ts';
import {
  attr,
  cellKey,
  parseCellRef,
  readXlsx,
  sheetByName,
  sheetIndex,
  workbookIndex,
  xmlPartOf,
} from '../packages/core/src/store/index.ts';
import { alertNote, alerts, showsSheet, WAIT_MS } from './verify-demo-browser.ts';
import { downloadNotice } from '../examples/vite/src/save.ts';
import { SHEET_GUIDE } from './sample-workbook/showcase-sheets.ts';

/** The name of the copy that a save of the sample downloads. */
const DOWNLOAD_NAME = 'sample-edited.xlsx';

/** What the step needs from the runner. */
export interface DownloadSteps {
  /** The sheet and the cell of the Bold edit. */
  readonly sheet: string;
  readonly ref: string;
  /** The sheet that the sample opens on. */
  readonly firstSheet: string;
}

/** A phone-wide bar shows Save and names each control, and the page does not scroll sideways. */
async function verifyPhoneBar(page: Page): Promise<void> {
  await page.setViewportSize({ width: 360, height: 740 });
  try {
    const header = page.locator('.demo-topbar').getByRole('button', { name: 'Save', exact: true });
    if (!(await header.isVisible())) throw new Error('a phone-wide bar hides Save');
    const wide = await page.evaluate(
      () => document.documentElement.scrollWidth > document.documentElement.clientWidth
    );
    if (wide) throw new Error('a phone-wide page scrolls sideways');
    // The bar hides the label text of these controls, and each control keeps its name.
    for (const [role, name] of [
      ['checkbox', 'Show all headings'],
      ['checkbox', 'View only'],
      ['button', 'Open a file'],
    ] as const) {
      if ((await page.getByRole(role, { name, exact: true }).count()) !== 1)
        throw new Error(`a phone-wide bar leaves the control "${name}" with no name`);
    }
  } finally {
    await page.setViewportSize({ width: 1600, height: 900 });
  }
}

/** The text of the message bar starts at the padding of the bar. */
async function verifyMessageAligned(page: Page, what: string): Promise<void> {
  const offset = await page.locator('.demo-message').evaluate((bar) => {
    const text = bar.querySelector('.demo-message__text');
    const left = parseFloat(getComputedStyle(bar).paddingLeft);
    return (text?.getBoundingClientRect().left ?? NaN) - bar.getBoundingClientRect().left - left;
  });
  if (Math.abs(offset) > 0.5)
    throw new Error(`the text of ${what} starts ${offset} px off the message bar`);
}

/** Whether the cell `ref` of `sheet` in the downloaded `bytes` is bold, after a read check. */
function boldIn(bytes: Uint8Array, sheet: string, ref: string): boolean {
  const read = readXlsx(bytes);
  if (!read.ok) throw new Error(`the downloaded copy does not read (${read.error.reason})`);
  const index = workbookIndex(read.pkg)!;
  const names = index.sheets.map((s) => s.name);
  if (JSON.stringify(names) !== JSON.stringify(SHEET_GUIDE.map(([n]) => n)))
    throw new Error(`the downloaded copy has the sheets ${JSON.stringify(names)}`);
  const part = xmlPartOf(read.pkg, sheetByName(index, sheet)?.partName);
  const at = parseCellRef(ref)!;
  const cell = part ? sheetIndex(part).cells.get(cellKey(at.row, at.col)) : undefined;
  return cellFormat(read.pkg, index, attr(cell?.el, 's')).format.font.bold;
}

/**
 * A real download from the header's Save after the Bold edit, whose bytes hold the edit and every
 * sheet, and the layout of the bars. `adopt` gives the page of the step to the runner, which takes
 * its screenshot and then closes it.
 */
export async function verifyDownloadSave(
  browser: Browser,
  url: string,
  watch: (page: Page) => void,
  adopt: (page: Page, close: () => Promise<void>) => void,
  demo: DownloadSteps
): Promise<string> {
  const context = await browser.newContext({
    acceptDownloads: true,
    viewport: { width: 1600, height: 900 },
  });
  const page = await context.newPage();
  adopt(page, () => context.close());
  watch(page);
  await page.addInitScript(() => {
    delete (window as unknown as { showSaveFilePicker?: unknown }).showSaveFilePicker;
  });
  await page.goto(url);
  if (!(await showsSheet(page, demo.firstSheet))) throw new Error('the sample did not open');
  if (await page.evaluate(() => 'showSaveFilePicker' in window))
    throw new Error('the page still has a file save picker');
  await verifyPhoneBar(page);
  await page.getByRole('tab', { name: demo.sheet, exact: true }).click();
  if (!(await showsSheet(page, demo.sheet))) throw new Error(`no sheet ${demo.sheet}`);
  await page.getByRole('button', { name: 'Bold', exact: true }).click();
  const downloads: unknown[] = [];
  page.on('download', (d) => downloads.push(d));
  const started = page.waitForEvent('download', { timeout: WAIT_MS }).catch(async () => {
    throw new Error(`the Save downloaded no file${await alertNote(page)}`);
  });
  const [download] = await Promise.all([
    started,
    page.locator('.demo-topbar').getByRole('button', { name: 'Save', exact: true }).click(),
  ]);
  // One save downloads one file: a second handler would start two.
  await Bun.sleep(1100);
  if (downloads.length !== 1) throw new Error(`the Save downloaded ${downloads.length} files`);
  if (download.suggestedFilename() !== DOWNLOAD_NAME)
    throw new Error(`the Save downloaded the name ${download.suggestedFilename()}`);
  const failure = await download.failure();
  if (failure) throw new Error(`the download failed: ${failure}`);
  const bytes = new Uint8Array(readFileSync(await download.path()));
  if (!boldIn(bytes, demo.sheet, demo.ref))
    throw new Error('the downloaded file does not hold the edit');
  const notice = downloadNotice(DOWNLOAD_NAME, true);
  await page
    .getByRole('status')
    .getByText(notice, { exact: true })
    .waitFor({ timeout: WAIT_MS })
    .catch(() => {
      throw new Error(`the Save did not say "${notice}"`);
    });
  await verifyMessageAligned(page, 'a notice');
  // An error shows in the same bar: a picked file that does not read.
  await page.locator('.demo-topbar input[type=file]').setInputFiles({
    name: 'broken.xlsx',
    mimeType: 'application/octet-stream',
    buffer: Buffer.from('not a workbook'),
  });
  await page.getByRole('alert').filter({ hasText: 'Could not open broken.xlsx' }).waitFor();
  await verifyMessageAligned(page, 'an error');
  await page.getByRole('button', { name: 'Close the message' }).click();
  const shown = await alerts(page);
  if (shown.length) throw new Error(`alert "${shown.join('; ')}"`);
  return 'a phone-wide bar kept Save and the names of its controls with no sideways scroll, the header Save downloaded one file that holds the edit and every sheet, and a notice and an error started at the padding of the message bar';
}
