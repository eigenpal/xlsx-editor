// The status strip step of `verify-demo.ts`, kept apart so the runner stays under the line cap.
// It checks what only a browser engine lays out: with no message, the status strip takes no
// room between the formula bar and the grid, and its notices stay mounted as live status nodes;
// a message pushes the grid down by the height of the strip; and the grid keeps its scroll
// position and its selection while the message shows and after it goes. The texts and the live
// roles of the strip are happy-dom tests (`packages/react/src/__tests__/adapter.test.tsx` and
// `packages/vue/test/adapter.test.ts`).

import type { Browser, Page } from '@playwright/test';
import { showsSheet, WAIT_MS } from './verify-demo-browser.ts';

/** How long a set scroll must hold before the step starts. */
const HOLD_MS = 300;

/** What the page lays out around the status strip. */
interface StripLayout {
  /** The space from the bottom of the formula bar to the top of the grid scroller. */
  readonly gap: number;
  /** The top borders of the grid scroller and its parts under the strip's parent. */
  readonly borders: number;
  /** The height of the status strip. */
  readonly strip: number;
  /** The number of mounted status nodes in the strip that the page does not hide. */
  readonly live: number;
  readonly scrollTop: number;
  readonly scrollLeft: number;
}

function stripLayout(page: Page): Promise<StripLayout> {
  return page.evaluate(() => {
    const strip = document.querySelector<HTMLElement>('.xlsx-status-strip');
    const bar = document.querySelector<HTMLElement>('.xlsx-formula-bar');
    const scroller = document.querySelector<HTMLElement>('.xlsx-scroller');
    if (!strip || !bar || !scroller) throw new Error('no status strip, formula bar, or grid');
    let borders = 0;
    for (let el: HTMLElement | null = scroller; el && el !== strip.parentElement;) {
      borders += parseFloat(getComputedStyle(el).borderTopWidth) || 0;
      el = el.parentElement;
    }
    const live = [...strip.querySelectorAll<HTMLElement>('[role="status"]')].filter((n) => {
      const style = getComputedStyle(n);
      return style.display !== 'none' && style.visibility !== 'hidden' && !n.closest('[hidden]');
    }).length;
    return {
      gap: scroller.getBoundingClientRect().top - bar.getBoundingClientRect().bottom,
      borders,
      strip: strip.getBoundingClientRect().height,
      live,
      scrollTop: scroller.scrollTop,
      scrollLeft: scroller.scrollLeft,
    };
  });
}

/**
 * On a page of its own, with the sample as it opens, so no edit shows a notice: with no message,
 * the grid starts under the formula bar (at most 1 pixel plus the grid's top borders), and the
 * strip keeps its live status nodes. A refused formula shows in the strip, which then has a
 * height and pushes the grid down. Escape clears the message, and the strip takes no room
 * again. The grid of `sheets.scrolled` keeps its scroll position and its selection throughout.
 * `adopt` gives the page to the runner, which takes its screenshot on a failure and closes it.
 */
export async function verifyStatusStrip(
  browser: Browser,
  url: string,
  watch: (page: Page) => void,
  adopt: (page: Page, close: () => Promise<void>) => void,
  sheets: { readonly first: string; readonly scrolled: string }
): Promise<string> {
  const context = await browser.newContext({ viewport: { width: 1280, height: 800 } });
  const page = await context.newPage();
  adopt(page, () => context.close());
  watch(page);
  await page.goto(url);
  if (!(await showsSheet(page, sheets.first))) throw new Error('the sample did not open');
  await page.getByRole('tab', { name: sheets.scrolled, exact: true }).click();
  if (!(await showsSheet(page, sheets.scrolled))) throw new Error(`no sheet ${sheets.scrolled}`);
  const scroller = page.locator('.xlsx-scroller').first();
  const nameBox = page.getByRole('textbox', { name: 'Cell reference' });
  const content = page.getByRole('textbox', { name: 'Cell content' });
  // A scrolled grid, and a press on a cell in its middle, so the active cell is in view and no
  // later action scrolls to it. The grid sizes its scroll range and shows the sheet's own view
  // after the sheet shows, so the scroll is set again until it holds for a while.
  const holds = async () => {
    await scroller.evaluate((el) => {
      el.scrollTop = 600;
      el.scrollLeft = 200;
    });
    await page.waitForTimeout(HOLD_MS);
    return scroller.evaluate((el) => el.scrollTop > 0);
  };
  for (let tries = 0; !(await holds()); tries += 1)
    if (tries >= 20) throw new Error(`the grid of ${sheets.scrolled} did not keep a scroll`);
  const box = (await scroller.boundingBox())!;
  await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
  const ref = await nameBox.inputValue();
  const text = await content.inputValue();
  const empty = await stripLayout(page);
  if (empty.scrollTop === 0) throw new Error(`a press on ${ref} scrolled the grid to the top`);
  if (empty.strip !== 0) throw new Error(`with no message, the strip is ${empty.strip} px high`);
  if (empty.gap > 1 + empty.borders)
    throw new Error(
      `with no message, the grid starts ${empty.gap} px under the formula bar, ` +
        `not at most ${1 + empty.borders} px`
    );
  if (empty.live === 0) throw new Error('with no message, the strip has no live status node');
  const kept = (state: StripLayout, when: string) => {
    if (state.scrollTop !== empty.scrollTop || state.scrollLeft !== empty.scrollLeft)
      throw new Error(
        `${when}, the grid scrolled to ${state.scrollTop}, ${state.scrollLeft}, ` +
          `not ${empty.scrollTop}, ${empty.scrollLeft}`
      );
  };
  await scroller.focus();
  await page.keyboard.type('=SUM(');
  await page.keyboard.press('Enter');
  await page.locator('.xlsx-refusal').waitFor({ timeout: WAIT_MS });
  const shown = await stripLayout(page);
  if (!(shown.strip > 0)) throw new Error('with a message, the strip is 0 px high');
  if (shown.gap < shown.strip)
    throw new Error(`the ${shown.strip} px strip covers the grid (gap ${shown.gap} px)`);
  kept(shown, 'with a message');
  if ((await nameBox.inputValue()) !== ref)
    throw new Error(`with a message, the selection is ${await nameBox.inputValue()}, not ${ref}`);
  await page.keyboard.press('Escape');
  await page.locator('.xlsx-refusal').waitFor({ state: 'detached', timeout: WAIT_MS });
  const gone = await stripLayout(page);
  if (gone.strip !== 0) throw new Error(`after the message, the strip is ${gone.strip} px high`);
  kept(gone, 'after the message');
  if ((await nameBox.inputValue()) !== ref || (await content.inputValue()) !== text)
    throw new Error(`after the message, the selection or the text of ${ref} changed`);
  return (
    `the grid started ${empty.gap} px under the formula bar with no message, a message made ` +
    `the strip ${shown.strip} px high, and the grid kept its scroll ` +
    `(${empty.scrollTop}, ${empty.scrollLeft}) and ${ref} selected`
  );
}
