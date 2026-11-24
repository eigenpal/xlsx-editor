// The editor embedded in a host page, in each engine, on a page of its own. A host CSS reset
// changes no chrome style, in the chrome and with each popup and the number format dialog open.
// The chrome stylesheet changes nothing on the host page, and host content beside the chrome
// keeps the host styles. In a host container that moves, scales, or clips the editor, the
// toolbar shows the groups of its layout width, a toolbar list opens under its trigger and stays
// in the visible box (capped in a short window), a press on the grid selects the cell under the
// pointer, and the sheet tab menu opens at the pointer.
//
// After the embedding tests of pptx-editor/e2e/browser.spec.ts at 3bf89a88 (EigenPal):
// `HOST_RESET` and `chromeStyles` are copied with the class prefix changed. Intended
// differences: the reset also sets inherited text props on the body (`text-indent`,
// `word-spacing`, `tab-size`, `font-feature-settings`, and `font-variation-settings`), and the
// comparison reads them; a
// step of `scripts/verify-demo.ts` in every engine, not a Playwright Test spec in Chromium; the
// states are this editor's popups and dialog; there is no tooltip, because the chrome shows
// native titles; and the scaled host also checks a press on the grid, whose geometry is in
// layout pixels.

import type { Browser, Locator, Page } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { showsSheet } from './verify-demo-browser.ts';

const STYLESHEET = join(import.meta.dir, '../packages/core/src/styles/editor.css');

// A host reset in the shape of Tailwind's preflight (its element rules): no margins or padding,
// headings and lists without their defaults, controls that inherit the font, and pictures as
// blocks. The body also sets other inherited text props that a host page can set. The editor
// chrome must look the same under it.
const HOST_RESET = `
  *, ::after, ::before { box-sizing: border-box; border: 0 solid; margin: 0; padding: 0; }
  html { line-height: 1.5; font-family: ui-sans-serif, system-ui, sans-serif; }
  body {
    line-height: inherit; text-indent: 2em; word-spacing: 3px; tab-size: 2;
    font-feature-settings: "smcp"; font-variation-settings: "wght" 700;
  }
  hr { height: 0; color: inherit; border-top-width: 1px; }
  h1, h2, h3, h4, h5, h6 { font-size: inherit; font-weight: inherit; }
  a { color: inherit; text-decoration: inherit; }
  b, strong { font-weight: bolder; }
  small { font-size: 80%; }
  table { text-indent: 0; border-color: inherit; border-collapse: collapse; }
  ol, ul, menu { list-style: none; }
  img, svg, video, canvas, audio, iframe, embed, object { display: block; vertical-align: middle; }
  img, video { max-width: 100%; height: auto; }
  button, input, select, optgroup, textarea, ::file-selector-button {
    font: inherit; letter-spacing: inherit; color: inherit; border-radius: 0;
    background-color: transparent; opacity: 1;
  }
  textarea { resize: vertical; }
`;

/** The box and text styles of every chrome element, for a before-and-after comparison. */
async function chromeStyles(page: Page): Promise<Record<string, string>> {
  return page.evaluate(() => {
    const props = [
      'margin-top',
      'margin-right',
      'margin-bottom',
      'margin-left',
      'padding-top',
      'padding-right',
      'padding-bottom',
      'padding-left',
      'border-top-width',
      'border-right-width',
      'border-bottom-width',
      'border-left-width',
      'display',
      'width',
      'height',
      'max-width',
      'font-size',
      'font-weight',
      'font-family',
      'line-height',
      'letter-spacing',
      'word-spacing',
      'text-indent',
      'tab-size',
      'font-feature-settings',
      'font-variation-settings',
      'list-style-type',
      'color',
      'background-color',
      'border-radius',
      'text-decoration-line',
      'opacity',
      'vertical-align',
    ];
    const out: Record<string, string> = {};
    // The editor's parts and what they hold. Demo shell elements are host content, but a part
    // inside a shell wrapper is chrome.
    const PART =
      "[class^='xlsx-']:not([class^='xlsx-editor']), [class*=' xlsx-']:not([class^='xlsx-editor'])";
    const chrome = [...document.querySelectorAll('.xlsx-editor *')].filter((el) => {
      const part = el.closest(PART);
      const shell = el.closest('[class*="demo-"]');
      return !!part && (!shell || shell.contains(part));
    });
    chrome.forEach((el, i) => {
      const s = getComputedStyle(el);
      const parent = el.parentElement ? getComputedStyle(el.parentElement).display : '';
      const tag = `${i} ${el.tagName.toLowerCase()}.${el.getAttribute('class') ?? ''}`;
      for (const p of props) {
        // A flex or grid item ignores vertical-align.
        if (p === 'vertical-align' && /flex|grid/.test(parent)) continue;
        out[`${tag} ${p}`] = s.getPropertyValue(p);
      }
    });
    return out;
  });
}

/** Waits until no finite animation runs, so each box is final. */
const settled = (page: Page) =>
  page.waitForFunction(() =>
    document
      .getAnimations()
      .every((a) => a.playState !== 'running' || a.effect?.getComputedTiming().endTime === Infinity)
  );

/** Opens the sample on `page` and waits for its sheet and fonts. */
async function openSample(page: Page, url: string, sheet: string): Promise<void> {
  await page.goto(url);
  if (!(await showsSheet(page, sheet))) throw new Error('the sample did not open');
  await page.evaluate(() => document.fonts.ready.then(() => undefined));
}

const numberList = (page: Page) => page.getByRole('button', { name: /^Number format:/ });
const firstTab = (page: Page) => page.locator('.xlsx-sheet-tab').first();

/** The states of the chrome that the reset check compares. */
const STATES: readonly (readonly [string, (page: Page) => Promise<void>])[] = [
  ['the chrome', async () => {}],
  [
    'the font list',
    (p) => opens(p, p.getByRole('button', { name: /^Select font family/ }), '[data-popup]'),
  ],
  [
    'the fill color picker',
    (p) => opens(p, p.getByRole('button', { name: 'Fill color options' }), '[data-popup]'),
  ],
  ['the number format list', (p) => opens(p, numberList(p), '[data-popup]')],
  [
    'the insert menu',
    (p) => opens(p, p.getByRole('button', { name: 'Insert and delete' }), '[data-popup]'),
  ],
  ['the sheet tab menu', (p) => opens(p, firstTab(p), '.xlsx-sheet-menu', 'right')],
  ['the sheet rename field', (p) => opens(p, firstTab(p), '.xlsx-sheet-rename', 'double')],
  [
    'the number format dialog',
    async (p) => {
      await opens(p, numberList(p), '[data-popup]');
      await opens(p, p.getByRole('button', { name: 'More number formats…' }), 'dialog[open]');
    },
  ],
];

/** Presses `trigger` and waits until the part `selector` shows. */
async function opens(
  page: Page,
  trigger: Locator,
  selector: string,
  how: 'left' | 'right' | 'double' = 'left'
): Promise<void> {
  if (how === 'double') await trigger.dblclick();
  else await trigger.click({ button: how });
  await page.locator(selector).first().waitFor();
}

/** A host CSS reset changes no chrome style, in each state. */
async function verifyHostReset(page: Page, url: string, sheet: string): Promise<void> {
  for (const [name, open] of STATES) {
    await openSample(page, url, sheet);
    // The demo shell sets its own font and line height on the editor scope. Take them off, so
    // what the host page passes down reaches the editor as it does on a plain page. Keep its box,
    // so the grid keeps its height, and wait for the grid to lay out again.
    await page.evaluate(async () => {
      const shell = document.querySelector<HTMLElement>('.demo-root');
      shell?.classList.remove('demo-root');
      shell?.setAttribute('style', 'display: flex; flex-direction: column; height: 100dvh');
      for (let i = 0; i < 3; i += 1) await new Promise((done) => requestAnimationFrame(done));
    });
    await open(page).catch((cause: unknown) => {
      throw new Error(`${name} did not open: ${String(cause).split('\n')[0]}`);
    });
    // No title under a resting pointer, and every open animation done.
    await page.mouse.move(0, 0);
    await settled(page);
    const before = await chromeStyles(page);
    await page.addStyleTag({ content: HOST_RESET });
    const after = await chromeStyles(page);
    const changed = Object.keys(before).filter((k) => before[k] !== after[k]);
    if (Object.keys(before).length === 0) throw new Error(`${name}: no chrome element was found`);
    if (changed.length)
      throw new Error(
        `a host reset changed ${changed.length} chrome styles in ${name}: ` +
          changed
            .slice(0, 5)
            .map((k) => `${k}: ${before[k]} -> ${after[k]}`)
            .join('; ')
      );
  }
}

/** The chrome stylesheet changes nothing on a host page with no editor. */
async function verifyHostUntouched(page: Page, css: string): Promise<void> {
  const host =
    '<h1 class="dark">Title</h1><p>Text <a href="#x">link</a></p><button type="button">Go</button>' +
    '<input aria-label="Name"><select aria-label="Pick"><option>One</option></select>' +
    '<ul><li>Item</li></ul><table><tr><td>Cell</td></tr></table>' +
    '<svg width="20" height="20" aria-hidden="true"><rect width="20" height="20"/></svg>' +
    '<canvas width="20" height="20"></canvas><dialog open>Note</dialog>';
  const styles = (): Promise<Record<string, string>[]> =>
    page.evaluate(() =>
      [document.documentElement, document.body, ...document.body.querySelectorAll('*')].map(
        (el) => {
          const s = getComputedStyle(el);
          const out: Record<string, string> = {};
          for (const name of s) out[name] = s.getPropertyValue(name);
          for (const name of ['--doc-primary', '--doc-bg', '--doc-text'])
            out[name] = s.getPropertyValue(name);
          return out;
        }
      )
    );
  await page.setContent(`<!doctype html><html><head></head><body>${host}</body></html>`);
  const before = await styles();
  await page.setContent(
    `<!doctype html><html><head><style>${css}</style></head><body>${host}</body></html>`
  );
  const after = await styles();
  for (const [i, was] of before.entries())
    for (const [name, value] of Object.entries(was))
      if (after[i]?.[name] !== value)
        throw new Error(
          `the chrome stylesheet changed ${name} of host element ${i}: ${value} -> ${after[i]?.[name]}`
        );
}

/** Host classes that load before the editor stylesheet keep their styles in the editor scope. */
async function verifyHostClasses(page: Page, css: string): Promise<void> {
  const host = `
    .btn { padding: 8px 12px; font-size: 20px; }
    .list { list-style: none; margin: 0; padding: 0; }
    .note { margin: 0; }
    .title { font-size: 30px; margin: 0; }
    .icon { display: none; }`;
  await page.setContent(
    `<!doctype html><html><head><style>${host}</style><style>${css}</style></head><body>` +
      '<div class="xlsx-editor"><button class="btn">Go</button><ul class="list"><li>A</li></ul>' +
      '<p class="note">Note</p><h2 class="title">Title</h2>' +
      '<svg class="icon" width="10" height="10"></svg></div></body></html>'
  );
  const seen = await page.evaluate(() => {
    const get = (sel: string, prop: string): string =>
      getComputedStyle(document.querySelector(sel)!).getPropertyValue(prop);
    return JSON.stringify([
      [get('.btn', 'padding-top'), get('.btn', 'padding-left'), get('.btn', 'font-size')],
      [get('.list', 'list-style-type'), get('.list', 'margin-top'), get('.list', 'padding-left')],
      get('.note', 'margin-top'),
      [get('.title', 'font-size'), get('.title', 'margin-top')],
      get('.icon', 'display'),
    ]);
  });
  const want = JSON.stringify([
    ['8px', '12px', '20px'],
    ['none', '0px', '0px'],
    '0px',
    ['30px', '0px'],
    'none',
  ]);
  if (seen !== want) throw new Error(`host classes in the editor scope became ${seen}`);
}

/** The cell that a press at (`x`, `y`) of the stage, in layout pixels, selects. */
async function pressAt(page: Page, x: number, y: number, scale: number): Promise<string> {
  const stage = (await page.locator('.xlsx-stage').first().boundingBox())!;
  await page.mouse.click(stage.x + x * scale, stage.y + y * scale);
  return page.getByRole('textbox', { name: 'Cell reference' }).inputValue();
}

/**
 * Opens the number format list, or with `capped` the font list, and throws unless the list is
 * under its trigger and ends above `bottom`. With `capped`, it must also be capped to its room.
 */
async function expectListUnder(
  page: Page,
  host: string,
  bottom: () => Promise<number>,
  capped = false
): Promise<void> {
  // A frame that narrows the bar keeps the font list, the first list of the bar, in view.
  const trigger = capped
    ? page.getByRole('button', { name: /^Select font family/ })
    : numberList(page);
  await trigger.click();
  const list = page.locator('.xlsx-toolbar [data-popup]').first();
  await list.waitFor();
  await settled(page);
  const [t, l, limit] = [
    (await trigger.boundingBox())!,
    (await list.boundingBox())!,
    await bottom(),
  ];
  if (l.y < t.y + t.height - 0.5 || l.x > t.x + t.width || l.x + l.width < t.x)
    throw new Error(`in ${host}, the list is not under its trigger`);
  if (l.y + l.height > limit + 0.5)
    throw new Error(`in ${host}, the list ends at ${l.y + l.height}, past ${limit}`);
  if (capped && !(await list.evaluate((el) => el.hasAttribute('data-capped'))))
    throw new Error(`in ${host}, the list is not capped to its room`);
  await page.keyboard.press('Escape');
  await list.waitFor({ state: 'hidden' });
}

/**
 * The number of toolbar groups that the bar shows, not in its More panel, once it holds for ten
 * frames: the bar measures its groups again after a resize, in a later frame.
 */
const shownGroups = (page: Page) =>
  page
    .locator('.xlsx-toolbar')
    .first()
    .evaluate(async (bar) => {
      const count = () =>
        [...bar.querySelectorAll('[data-group]')].filter((g) => !g.closest('[role="dialog"]'))
          .length;
      let last = count();
      let same = 0;
      for (let frames = 0; same < 10 && frames < 600; frames += 1) {
        await new Promise((done) => requestAnimationFrame(done));
        const now = count();
        same = now === last ? same + 1 : 0;
        last = now;
      }
      if (same < 10) throw new Error('the toolbar group count did not hold for ten frames');
      return last;
    });

/** Opens the sheet tab menu at a pointer, and throws unless the menu starts at the pointer. */
async function expectMenuAtPointer(page: Page, host: string): Promise<void> {
  const tab = (await firstTab(page).boundingBox())!;
  const at = { x: tab.x + tab.width * 0.75, y: tab.y + tab.height / 2 };
  await page.mouse.click(at.x, at.y, { button: 'right' });
  const menu = page.locator('.xlsx-sheet-menu');
  await menu.waitFor();
  await settled(page);
  const box = (await menu.boundingBox())!;
  // Within a pixel and a half: the bar sits at a fractional offset.
  if (Math.abs(box.x - at.x) > 1.5)
    throw new Error(
      `in ${host}, the sheet tab menu starts at ${box.x}, not at the pointer ${at.x}`
    );
  await page.keyboard.press('Escape');
  await menu.waitFor({ state: 'hidden' });
}

/** A host frame around the editor: its CSS, its scale, and the window it is in. */
interface Frame {
  readonly host: string;
  readonly css: string;
  readonly scale: number;
  /** A short window, where only the toolbar shows: the list check alone runs. */
  readonly short?: boolean;
}

const FRAMES: readonly Frame[] = [
  {
    host: 'a host that moves the editor with a transform',
    css: '#root { transform: translate(37px, 23px); filter: saturate(1); height: 700px; }',
    scale: 1,
  },
  {
    host: 'a host that scales the editor by 1.25',
    css: '#root { transform: translate(37px, 23px) scale(1.25); transform-origin: 0 0; height: 600px; }',
    scale: 1.25,
  },
  {
    host: 'a host that clips the editor to 400 pixels',
    css: '#root { transform: translate(37px, 23px); overflow: hidden; height: 400px; }',
    scale: 1,
  },
  {
    host: 'a host that scales the editor by 2 in a short window',
    css: '#root { transform: scale(2); transform-origin: 0 0; width: 720px; height: 210px; }',
    scale: 2,
    short: true,
  },
];

/**
 * Lists, the toolbar groups, the grid, and the sheet tab menu in hosts that move, scale, and
 * clip the editor.
 */
async function verifyHostFrames(page: Page, url: string, sheet: string): Promise<void> {
  // Points of the stage that also show in the frame that clips it, each well inside its cell:
  // an engine can round a pointer on a cell edge to either side.
  const points = [
    [260, 100],
    [520, 140],
    [90, 52],
  ] as const;
  const size = page.viewportSize()!;
  await openSample(page, url, sheet);
  const plain: string[] = [];
  for (const [x, y] of points) plain.push(await pressAt(page, x, y, 1));
  const groups = await shownGroups(page);
  for (const { host, css, scale, short } of FRAMES) {
    const height = short ? 420 : size.height;
    // The short frame lays the editor out 720 pixels wide: the bar shows the groups that a
    // window of that width shows.
    let want = groups;
    if (short) {
      await page.setViewportSize({ width: 720, height });
      await openSample(page, url, sheet);
      want = await shownGroups(page);
    }
    await page.setViewportSize({ width: size.width, height });
    await openSample(page, url, sheet);
    // The demo shell fills the window; in a host frame it fills the frame, which stays in the
    // window, so the sheet tabs show.
    await page.addStyleTag({ content: `${css} .demo-root { height: 100% !important; }` });
    await settled(page);
    const clip = css.includes('overflow');
    const bottom = async () => {
      if (!clip) return height;
      const root = (await page.locator('#root').boundingBox())!;
      return Math.min(height, root.y + root.height);
    };
    // The bar measures its groups in its own pixels, so a scale does not change their number.
    const shown = await shownGroups(page);
    if (shown !== want)
      throw new Error(`in ${host}, the toolbar shows ${shown} groups, not ${want}`);
    await expectListUnder(page, host, bottom, short);
    if (short) continue;
    const seen: string[] = [];
    for (const [x, y] of points) seen.push(await pressAt(page, x, y, scale));
    if (seen.join() !== plain.join())
      throw new Error(`in ${host}, presses selected ${seen.join(' ')}, not ${plain.join(' ')}`);
    await expectMenuAtPointer(page, host);
  }
  await page.setViewportSize(size);
}

/**
 * The embedding checks, on a page of their own: `adopt` gives the page to the runner, which
 * takes its screenshot on a failure and then closes it.
 */
export async function verifyEmbedding(
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
  await verifyHostReset(page, url, sheet);
  await verifyHostFrames(page, url, sheet);
  const css = readFileSync(STYLESHEET, 'utf8');
  await verifyHostUntouched(page, css);
  await verifyHostClasses(page, css);
  return (
    `a host reset changed no chrome style in ${STATES.length} states, the stylesheet changed ` +
    'no host style, and toolbar groups, lists, presses, and the sheet tab menu kept their places ' +
    'in hosts that move, scale, and clip the editor'
  );
}
