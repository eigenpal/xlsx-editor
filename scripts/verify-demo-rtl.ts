// The editor on a right-to-left host page, in each engine, on a page of its own.
//
// - With `dir="rtl"` on the root element of the page, a horizontal wheel over the grid scrolls
//   it: the scroll offset moves from 0, and a column heading that did not show before shows.
//   The grid keeps a left-to-right layout (`direction: ltr` on `.xlsx-surface` in
//   `packages/core/src/styles/editor.css`), so its physical coordinates hold.
// - A jump to AZ1 from the name box shows the AZ column heading inside the visible part of the
//   grid, with the cell under it.

import type { Browser, Page } from '@playwright/test';
import { goTo, showsSheet } from './verify-demo-browser.ts';

/** The column headings inside the visible box of the scroller, from left to right. */
const shownHeadings = (page: Page) =>
  page.evaluate(() => {
    const scroller = document.querySelector('.xlsx-scroller') as HTMLElement;
    const box = scroller.getBoundingClientRect();
    return [...document.querySelectorAll<HTMLElement>('.xlsx-column-heads .xlsx-head')]
      .map((el) => ({ text: el.textContent ?? '', rect: el.getBoundingClientRect() }))
      .filter(({ rect }) => rect.width > 0 && rect.left >= box.left && rect.right <= box.right)
      .sort((a, b) => a.rect.left - b.rect.left)
      .map(({ text }) => text);
  });

const scrollLeft = (page: Page) =>
  page
    .locator('.xlsx-scroller')
    .first()
    .evaluate((el) => el.scrollLeft);

/**
 * The right-to-left host checks, on a page of their own: `adopt` gives the page to the runner,
 * which takes its screenshot on a failure and then closes it.
 */
export async function verifyRightToLeftHost(
  browser: Browser,
  url: string,
  watch: (page: Page) => void,
  adopt: (page: Page, close: () => Promise<void>) => void,
  sheet: string
): Promise<string> {
  const context = await browser.newContext({ viewport: { width: 1440, height: 900 } });
  const page = await context.newPage();
  adopt(page, () => context.close());
  watch(page);
  await page.addInitScript(() => {
    document.addEventListener('DOMContentLoaded', () =>
      document.documentElement.setAttribute('dir', 'rtl')
    );
  });
  await page.goto(url);
  if (!(await showsSheet(page, sheet))) throw new Error('the sample did not open');
  const direction = await page.evaluate(() => getComputedStyle(document.documentElement).direction);
  if (direction !== 'rtl') throw new Error(`the host page is ${direction}, not rtl`);
  const before = await shownHeadings(page);
  if (before.length === 0) throw new Error('the grid shows no column heading');

  const grid = page.locator('.xlsx-scroller').first();
  const box = (await grid.boundingBox())!;
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  await page.mouse.wheel(600, 0);
  await page
    .waitForFunction(
      () => (document.querySelector('.xlsx-scroller') as HTMLElement).scrollLeft > 0,
      undefined,
      { timeout: 5000 }
    )
    .catch(() => undefined);
  const scrolled = await scrollLeft(page);
  if (scrolled <= 0) throw new Error(`a horizontal wheel left the scroll offset at ${scrolled}`);
  // The headings paint after the scroll, in the next frame: a heading that did not show before
  // the scroll shows now.
  await page.waitForTimeout(200);
  const after = await shownHeadings(page);
  const added = after.filter((text) => !before.includes(text));
  if (added.length === 0)
    throw new Error(
      `after a scroll of ${scrolled} pixels, the headings still show ${after.join(' ')}`
    );

  await goTo(page, 'AZ1');
  await page.waitForFunction(
    () =>
      [...document.querySelectorAll('.xlsx-column-heads .xlsx-head')].some(
        (el) => el.textContent === 'AZ'
      ),
    undefined,
    { timeout: 5000 }
  );
  const jumped = await shownHeadings(page);
  if (!jumped.includes('AZ'))
    throw new Error(`after a jump to AZ1, the shown headings are ${jumped.join(' ')}`);
  return (
    `on a right-to-left page, a wheel scrolled the grid ${Math.round(scrolled)} pixels and the ` +
    `headings show ${added[0]}, and a jump to AZ1 shows AZ`
  );
}
