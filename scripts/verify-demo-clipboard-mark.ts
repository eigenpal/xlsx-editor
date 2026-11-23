// The dashed outline of a pending copy in the grid of the demo, as the browser engine draws it:
// its line, its place after a scroll, forced colors, and a print. A real copy starts it, and
// Escape ends it. The states that keep or end the outline, its edges on a frozen seam, and the
// Escape that the demo's message takes are happy-dom tests
// (`packages/core/src/editor/__tests__/clipboard-mark.test.ts` and
// `examples/vite/src/__tests__/demo-open.test.tsx`). The step leaves the workbook as it found it.

import type { Page } from '@playwright/test';
import { goTo } from './verify-demo-browser.ts';

/** The cell of the step: empty in the sample, and below the cells of the other steps. */
const SOURCE = 'Z510';

const WHOLE = 'bottom left right top';

interface Edge {
  readonly side: string;
  readonly x: number;
  readonly y: number;
}

/** The edges that the grid shows after its next paint. */
const readEdges = (page: Page): Promise<Edge[]> =>
  page.evaluate(async () => {
    for (let i = 0; i < 2; i += 1) await new Promise((done) => requestAnimationFrame(done));
    return [...document.querySelectorAll<HTMLElement>('.xlsx-overlay .xlsx-pending-edge')].map(
      (el) => {
        const box = el.getBoundingClientRect();
        return { side: el.dataset.side ?? '', x: box.x, y: box.y };
      }
    );
  });

/** Throws unless the outline shows exactly `sides`, in sorted order, or nothing. */
async function expectMark(page: Page, sides: string, when: string): Promise<void> {
  const shown = (await readEdges(page))
    .map((e) => e.side)
    .sort()
    .join(' ');
  if (shown !== sides) throw new Error(`${when}, the outline shows "${shown}", not "${sides}"`);
}

/** The line of the outline: dashed, with no animation, and a color that shows on its strip. */
async function verifyLine(page: Page, when: string): Promise<void> {
  // A scroll or a media change draws the outline again in the next frame, so read the edge
  // that is in the document then: an edge that the redraw removed has no computed style.
  const handle = await page.waitForFunction(() => {
    const el = document.querySelector('.xlsx-pending-edge[data-side=top]');
    if (!el?.isConnected) return undefined;
    const s = getComputedStyle(el);
    if (!s.borderTopStyle) return undefined;
    return {
      style: s.borderTopStyle,
      width: s.borderTopWidth,
      color: s.borderTopColor,
      strip: s.backgroundColor,
      animation: s.animationName,
    };
  });
  const line = (await handle.jsonValue())!;
  if (line.style !== 'dashed' || line.width !== '2px')
    throw new Error(`${when}, the line is ${line.width} ${line.style}, not 2px dashed`);
  if (line.animation !== 'none') throw new Error(`${when}, the line has an animation`);
  if (line.color === line.strip) throw new Error(`${when}, the dashes match their strip`);
}

/** The outline follows a scroll of the grid by the distance of the scroll. */
async function verifyScroll(page: Page): Promise<void> {
  const top = async () => (await readEdges(page)).find((e) => e.side === 'top')?.y;
  const scrollBy = (by: number) =>
    page.evaluate((px) => {
      const scroller = document.querySelector<HTMLElement>('.xlsx-scroller')!;
      const before = scroller.scrollTop;
      scroller.scrollTop = before + px;
      return scroller.scrollTop - before;
    }, by);
  const before = await top();
  const moved = await scrollBy(40);
  const after = await top();
  if (before === undefined || after === undefined || moved === 0)
    throw new Error('the scroll of the outline could not be read');
  if (Math.abs(before - moved - after) > 3)
    throw new Error(`a scroll of ${moved}px moved the outline from ${before} to ${after}`);
  await scrollBy(-moved);
}

/** Forced colors keep the line, and a print of the page leaves it out. */
async function verifyMedia(page: Page): Promise<string> {
  await page.emulateMedia({ forcedColors: 'active' });
  const forced = await page.evaluate(() => matchMedia('(forced-colors: active)').matches);
  if (forced) await verifyLine(page, 'with forced colors');
  await page.emulateMedia({ forcedColors: 'none' });
  await page.emulateMedia({ media: 'print' });
  // As in `verifyLine`: read the edge that is in the document after the redraw.
  const printed = await (
    await page.waitForFunction(() => {
      const el = document.querySelector('.xlsx-pending-edge');
      return (el?.isConnected && getComputedStyle(el).display) || undefined;
    })
  ).jsonValue();
  await page.emulateMedia({ media: 'screen' });
  if (printed !== 'none') throw new Error(`a print shows the outline (display ${printed})`);
  return forced ? 'forced colors keep the line' : 'this engine does not emulate forced colors';
}

/** The outline of a pending copy as the engine draws it, and Escape that ends it. */
export async function verifyClipboardMark(page: Page): Promise<string> {
  if ((await goTo(page, SOURCE)) !== '') throw new Error(`${SOURCE} of the sample is not empty`);
  await page.keyboard.press('ControlOrMeta+c');
  await expectMark(page, WHOLE, 'after a copy');
  await verifyLine(page, 'after a copy');
  await verifyScroll(page);
  const media = await verifyMedia(page);
  await page.keyboard.press('Escape');
  await expectMark(page, '', 'after Escape on a copy');
  return `the outline of a copy showed a 2px dashed line with no animation, followed the scroll, left the print, and Escape ended it; ${media}`;
}
