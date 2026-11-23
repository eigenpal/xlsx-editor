// The editor in hosts that size it by its content, in each engine, on a page of its own.
//
// - In an inline block and in a centered flex column, the editor width stays within the window
//   and does not grow for 2 seconds. With no width from the host, the editor is 0 pixels wide;
//   with a width that the host gives the viewport, it keeps that width, so the editor renders. The viewport class `xlsx-editor-viewport` keeps the width
//   of the editor from following its content (`packages/core/src/styles/editor.css`); without
//   it, the widths of the grid would feed back into the editor and grow it without limit.
// - The status strip, moved into a host flex row beside a button and into a column that aligns
//   its items to the start, keeps a width above zero, and the text of a refusal in it shows.
//
// The strip move changes the DOM that the adapter renders, so it is the last check of the page,
// and the runner closes the page after it.

import type { Browser, Page } from '@playwright/test';
import { showsSheet } from './verify-demo-browser.ts';

const INLINE_BLOCK = '#root { display: inline-block; height: 700px; }';
const CENTERED_COLUMN =
  '#root { display: flex; flex-direction: column; align-items: center; height: 700px; }';
/** The width that a host gives the viewport. */
const GIVEN = 900;
const GIVEN_CSS = `.demo-root { width: ${GIVEN}px; }`;

/**
 * Hosts that size the demo viewport (`.demo-root`, a child of `#root`) by its content, and the
 * editor width that each one must keep: none when the host gives no width.
 */
const HOSTS: readonly (readonly [string, string, number | null])[] = [
  ['an inline block', INLINE_BLOCK, null],
  ['a centered flex column', CENTERED_COLUMN, null],
  [`an inline block with a ${GIVEN}px viewport`, `${INLINE_BLOCK} ${GIVEN_CSS}`, GIVEN],
  [`a centered flex column with a ${GIVEN}px viewport`, `${CENTERED_COLUMN} ${GIVEN_CSS}`, GIVEN],
];

/** How long the width must stay the same. */
const SETTLE_MS = 2000;

const viewportWidth = (page: Page) =>
  page.locator('.xlsx-editor-viewport').evaluate((el) => el.getBoundingClientRect().width);

async function openSample(page: Page, url: string, sheet: string): Promise<void> {
  await page.goto(url);
  if (!(await showsSheet(page, sheet))) throw new Error('the sample did not open');
}

/** The editor width stays within the window, and stays the same, in each content-sized host. */
async function verifyBoundedWidth(page: Page, url: string, sheet: string): Promise<string[]> {
  const window = page.viewportSize()!.width;
  const seen: string[] = [];
  for (const [host, css, given] of HOSTS) {
    await openSample(page, url, sheet);
    await page.addStyleTag({ content: css });
    const first = await viewportWidth(page);
    await page.waitForTimeout(SETTLE_MS);
    const last = await viewportWidth(page);
    if (last > window)
      throw new Error(`in ${host}, the editor is ${last} pixels wide, wider than the window`);
    if (Math.abs(last - first) > 1)
      throw new Error(`in ${host}, the editor width went from ${first} to ${last} pixels`);
    if (given !== null && Math.abs(last - given) > 1)
      throw new Error(`in ${host}, the editor is ${last} pixels wide, not ${given}`);
    seen.push(`${host} ${Math.round(last)} px`);
  }
  return seen;
}

/** The status strip keeps its width and its text in a host flex row and a start-aligned column. */
async function verifyStripInHostFlex(page: Page, url: string, sheet: string): Promise<void> {
  await openSample(page, url, sheet);
  // A formula that cannot be read is refused, and the strip shows the reason.
  await page.locator('.xlsx-scroller').first().focus();
  await page.keyboard.type('=SUM(');
  await page.keyboard.press('Enter');
  await page.locator('.xlsx-refusal').waitFor();
  for (const [host, style] of [
    ['a flex row beside a button', 'display: flex'],
    [
      'a column that aligns to the start',
      'display: flex; flex-direction: column; align-items: flex-start',
    ],
  ] as const) {
    const sizes = await page.evaluate((css) => {
      const strip = document.querySelector('.xlsx-status-strip') as HTMLElement;
      // The first host wraps the strip with a button; the next one restyles that wrapper.
      let box = strip.parentElement!;
      if (!box.hasAttribute('data-host-box')) {
        box = document.createElement('div');
        box.setAttribute('data-host-box', '');
        const button = document.createElement('button');
        button.type = 'button';
        button.textContent = 'Host action';
        strip.before(box);
        box.append(button, strip);
      }
      box.setAttribute('style', css);
      const text = strip.querySelector('.xlsx-refusal') as HTMLElement;
      return {
        strip: strip.getBoundingClientRect().width,
        text: text.getBoundingClientRect().width,
        words: text.textContent ?? '',
        shown: getComputedStyle(text).visibility === 'visible',
      };
    }, style);
    if (sizes.strip <= 0) throw new Error(`in ${host}, the status strip is 0 pixels wide`);
    if (sizes.text <= 0 || !sizes.shown || sizes.words === '')
      throw new Error(`in ${host}, the refusal text of the status strip does not show`);
  }
}

/**
 * The content-sized host checks, on a page of their own: `adopt` gives the page to the runner,
 * which takes its screenshot on a failure and then closes it.
 */
export async function verifyContentSizedHosts(
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
  const widths = await verifyBoundedWidth(page, url, sheet);
  await verifyStripInHostFlex(page, url, sheet);
  return (
    `the editor width stayed within the window for ${SETTLE_MS} ms (${widths.join(', ')}), ` +
    'and the status strip kept its width and refusal text in a host flex row and column'
  );
}
