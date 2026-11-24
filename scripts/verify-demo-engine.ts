// The checks of `verify-demo.ts` that differ most from one browser engine to another: the paint
// of the grid canvas, the loaded font faces, and the clipboard events of the grid. Kept apart
// so the runner stays under its line cap.

import type { Page } from '@playwright/test';
import { CLIPBOARD_TYPE } from '../packages/core/src/editor/clipboard.ts';
import { goTo, PLAIN_PASTE_CAUSE, skipIssue, type Engine } from './verify-demo-browser.ts';

/**
 * The fewest dark pixels (luminance under 50%) that a painted sheet has on its canvas. Gridlines
 * and the default fills are light, so a blank canvas has none. A dark fill or border also counts,
 * so the check shows paint, and it does not show text.
 */
export const MIN_DARK_PIXELS = 50;

/**
 * Checks that the grid canvas of the demo shows dark pixels and that no font face of the page
 * failed to load. A blank canvas, a canvas that the engine cannot read, or a face with the
 * status `error` throws. Returns the number of dark pixels.
 */
export async function checkPaint(page: Page): Promise<number> {
  const paint = await page.evaluate(async () => {
    await document.fonts.ready;
    const failed = [...document.fonts].filter((f) => f.status === 'error').map((f) => f.family);
    // The paint of a frame comes after the layout, so the check reads the next frame.
    await new Promise<void>((done) => requestAnimationFrame(() => done()));
    const canvas = document.querySelector<HTMLCanvasElement>('.demo-root .xlsx-grid canvas');
    if (!canvas) return { failed, width: 0, height: 0, dark: 0 };
    const { width, height } = canvas;
    const ctx = width && height ? canvas.getContext('2d') : null;
    const data = ctx ? ctx.getImageData(0, 0, width, height).data : new Uint8ClampedArray();
    let dark = 0;
    for (let i = 0; i < data.length; i += 4) {
      const luminance = 0.2126 * data[i]! + 0.7152 * data[i + 1]! + 0.0722 * data[i + 2]!;
      if (data[i + 3]! > 128 && luminance < 128) dark++;
    }
    return { failed, width, height, dark };
  });
  if (paint.failed.length)
    throw new Error(`the font faces ${paint.failed.join(', ')} failed to load`);
  if (!paint.width || !paint.height) throw new Error('the grid has no canvas with a size');
  if (paint.dark < MIN_DARK_PIXELS)
    throw new Error(
      `the grid canvas shows ${paint.dark} dark pixels, fewer than ${MIN_DARK_PIXELS}`
    );
  return paint.dark;
}

/** The cells of the clipboard steps: empty in the sample, and below the other steps' cells. */
const SOURCE = 'Z500';
const TARGET = 'Z501';

/**
 * Records the clipboard events of the page, with their target and types, for a failure. The
 * reader returns them as text, and whether the last paste had the private clipboard type.
 */
async function recordClipboard(
  page: Page
): Promise<() => Promise<{ events: string; internal: boolean }>> {
  await page.evaluate((privateType) => {
    const w = window as unknown as {
      __clip: string[];
      __clipEnd: AbortController;
      __internal: boolean;
    };
    w.__clip = [];
    w.__internal = false;
    w.__clipEnd = new AbortController();
    for (const type of ['copy', 'cut', 'paste'] as const)
      document.addEventListener(
        type,
        (e) => {
          const el = e.target as Element | null;
          const where = el?.closest?.('[role=grid]') ? ' in grid' : '';
          const target = `${el?.localName ?? 'none'}.${el?.className || '-'}${where}`;
          const text = e.clipboardData?.getData('text/plain') ?? '';
          const types = [...(e.clipboardData?.types ?? [])];
          if (type === 'paste') w.__internal = types.includes(privateType);
          w.__clip.push(`${type} on ${target} "${text}" [${types.join(' ')}]`);
        },
        { capture: true, signal: w.__clipEnd.signal }
      );
  }, CLIPBOARD_TYPE);
  // The step removes its listeners when it reads the events.
  return async () => {
    const seen = await page.evaluate(() => {
      const w = window as unknown as {
        __clip: string[];
        __clipEnd: AbortController;
        __internal: boolean;
      };
      w.__clipEnd.abort();
      return { list: w.__clip, internal: w.__internal };
    });
    return { events: `events: ${seen.list.join('; ') || 'none'}`, internal: seen.internal };
  };
}

/**
 * Copy and paste, or cut and paste, in the grid. A value typed into a cell is copied with
 * Control+C or Command+C, or cut with Control+X or Command+X, and pasted into the next cell with
 * Control+V or Command+V. The keys must reach the grid as `copy`, `cut`, and `paste` events: the
 * pasted cell holds the value, and a cut leaves the first cell empty. Undo then takes back each
 * edit, so the step leaves the workbook as it found it.
 */
export async function verifyClipboard(page: Page, mode: 'copy' | 'cut'): Promise<string> {
  for (const ref of [SOURCE, TARGET]) {
    const was = await goTo(page, ref);
    if (was !== '') throw new Error(`${ref} of the sample holds "${was}", not an empty cell`);
  }
  // Text of this engine and run, so a clipboard left by an earlier run cannot pass the step.
  const engine = page.context().browser()?.browserType().name() ?? 'browser';
  const text = `${mode === 'cut' ? 'Cut' : 'Copied'} in ${engine} ${Date.now()}`;
  await goTo(page, SOURCE);
  await page.keyboard.type(text);
  await page.keyboard.press('Enter');
  const events = await recordClipboard(page);
  await goTo(page, SOURCE);
  await page.keyboard.press(mode === 'cut' ? 'ControlOrMeta+x' : 'ControlOrMeta+c');
  await goTo(page, TARGET);
  await page.keyboard.press('ControlOrMeta+v');
  const pasted = await goTo(page, TARGET);
  const left = await goTo(page, SOURCE);
  const { events: seen, internal } = await events();
  if (pasted !== text)
    throw new Error(`a paste into ${TARGET} gave "${pasted}", not "${text}" (${seen})`);
  if (left !== (mode === 'cut' ? '' : text))
    throw new Error(`after the ${mode} and the paste, ${SOURCE} holds "${left}" (${seen})`);
  // Undo takes back the paste and the typed value. A cut and its paste can be one or two
  // undo steps, so the step stops when both cells are empty again.
  const undo = page.getByRole('button', { name: 'Undo', exact: true });
  let undos = 0;
  const empty = async () => (await goTo(page, SOURCE)) === '' && (await goTo(page, TARGET)) === '';
  while (!(await empty())) {
    if (++undos > 3) throw new Error(`three undos did not take back the ${mode} and the paste`);
    await undo.click();
  }
  const did = mode === 'cut' ? 'moved' : 'copied';
  // A paste with no private type takes the plain text lane: the value, with no cell format.
  // Only the clipboard of a run whose cut step is skipped may lose the type.
  if (!internal && !skipIssue(engine as Engine, 'cut and paste'))
    throw new Error(`the paste in ${engine} has no private clipboard type (${seen})`);
  const lane = internal ? '' : `, as plain text (no private clipboard type: ${PLAIN_PASTE_CAUSE})`;
  return `Mod+${mode === 'cut' ? 'X' : 'C'} and Mod+V ${did} ${SOURCE} into ${TARGET}, and ${undos} undos took back the edits${lane}`;
}

/** The menu bar of the demo top bar, and a choice of one row of one of its menus. */
function menuBarOf(page: Page) {
  const bar = page.getByRole('menubar', { name: 'Menu bar' });
  const choose = async (menu: string, row: string, role: 'menuitem' | 'menuitemcheckbox') => {
    await bar.getByRole('menuitem', { name: menu, exact: true }).click();
    await bar.getByRole('menu', { name: menu }).getByRole(role, { name: row, exact: true }).click();
  };
  return { bar, choose };
}

/** Undo until SOURCE and TARGET are empty again, at most `limit` times. Returns the count. */
async function undoToEmpty(page: Page, limit: number, what: string): Promise<number> {
  const undo = page.getByRole('button', { name: 'Undo', exact: true });
  let undos = 0;
  const empty = async () => (await goTo(page, SOURCE)) === '' && (await goTo(page, TARGET)) === '';
  while (!(await empty())) {
    if (++undos > limit) throw new Error(`${limit} undos did not take back ${what}`);
    await undo.click();
  }
  return undos;
}

/** The content of the cell editor once it is `want`, or after five seconds. */
async function contentSoon(page: Page, ref: string, want: string): Promise<string> {
  await page
    .waitForFunction(
      (text) =>
        (document.querySelector('[aria-label="Cell content"]') as HTMLTextAreaElement | null)
          ?.value === text,
      want,
      { timeout: 5000 }
    )
    .catch(() => undefined);
  return goTo(page, ref);
}

/**
 * The menu bar's clipboard rows in the demo top bar: Edit > Copy and Edit > Paste copy a typed
 * value into the next cell, and Edit > Cut and Edit > Paste move it. The test context grants the
 * clipboard read that Edit > Paste needs, and takes the grant back after the step. Undo then
 * takes back each edit, so the step leaves the workbook as it found it.
 */
export async function verifyMenuClipboard(page: Page): Promise<string> {
  const context = page.context();
  const origin = new URL(page.url()).origin;
  await context.grantPermissions(['clipboard-read', 'clipboard-write'], { origin });
  try {
    for (const ref of [SOURCE, TARGET]) {
      const was = await goTo(page, ref);
      if (was !== '') throw new Error(`${ref} of the sample holds "${was}", not an empty cell`);
    }
    const { choose } = menuBarOf(page);
    const engine = context.browser()?.browserType().name() ?? 'browser';
    const text = `Menu copy in ${engine} ${Date.now()}`;
    await goTo(page, SOURCE);
    await page.keyboard.type(text);
    await page.keyboard.press('Enter');
    await goTo(page, SOURCE);
    await choose('Edit', 'Copy', 'menuitem');
    await goTo(page, TARGET);
    await choose('Edit', 'Paste', 'menuitem');
    const pasted = await contentSoon(page, TARGET, text);
    if (pasted !== text)
      throw new Error(`Edit > Paste into ${TARGET} gave "${pasted}", not "${text}"`);
    const copyUndos = await undoToEmpty(page, 3, 'the menu copy and paste');
    // Edit > Cut and Edit > Paste move the value: the source cell is empty after the paste.
    const moved = `Menu cut in ${engine} ${Date.now()}`;
    await goTo(page, SOURCE);
    await page.keyboard.type(moved);
    await page.keyboard.press('Enter');
    await goTo(page, SOURCE);
    await choose('Edit', 'Cut', 'menuitem');
    await goTo(page, TARGET);
    await choose('Edit', 'Paste', 'menuitem');
    const arrived = await contentSoon(page, TARGET, moved);
    const left = await goTo(page, SOURCE);
    if (arrived !== moved)
      throw new Error(
        `Edit > Paste after Edit > Cut gave "${arrived}" in ${TARGET}, not "${moved}"`
      );
    if (left !== '') throw new Error(`Edit > Cut and Edit > Paste left "${left}" in ${SOURCE}`);
    const cutUndos = await undoToEmpty(page, 3, 'the menu cut and paste');
    return `Edit > Copy and Edit > Paste copied ${SOURCE} into ${TARGET}, Edit > Cut and Edit > Paste moved it, and ${copyUndos + cutUndos} undos took back the edits`;
  } finally {
    await context.clearPermissions();
  }
}

/**
 * Format > Bold in the demo top bar: a press makes the cell bold, which the row then shows
 * checked, and a second press takes it back. Undo then takes back the edits.
 */
export async function verifyMenuBold(page: Page): Promise<string> {
  const { bar, choose } = menuBarOf(page);
  const was = await goTo(page, TARGET);
  if (was !== '') throw new Error(`${TARGET} of the sample holds "${was}", not an empty cell`);
  await page.keyboard.type('Bold');
  await page.keyboard.press('Enter');
  // Format > Bold: the row is a checkbox item, checked on a bold cell.
  const boldRow = async () => {
    const format = bar.getByRole('menuitem', { name: 'Format', exact: true });
    await format.click();
    const row = bar.getByRole('menuitemcheckbox', { name: 'Bold', exact: true });
    const checked = await row.getAttribute('aria-checked');
    // A second press on the trigger closes the menu.
    if ((await format.getAttribute('aria-expanded')) === 'true') await format.click();
    return checked;
  };
  await goTo(page, TARGET);
  if ((await boldRow()) !== 'false') throw new Error('Format > Bold is checked on a plain cell');
  await choose('Format', 'Bold', 'menuitemcheckbox');
  if ((await boldRow()) !== 'true') throw new Error('Format > Bold is not checked after a press');
  await goTo(page, TARGET);
  await choose('Format', 'Bold', 'menuitemcheckbox');
  if ((await boldRow()) !== 'false') throw new Error('a second Format > Bold left the cell bold');
  const undos = await undoToEmpty(page, 4, 'the menu edits');
  return `Format > Bold checked and unchecked its row, and ${undos} undos took back the edits`;
}

/**
 * Copy right after the grid takes the focus, with no page selection: by a press on a cell, and
 * by Shift+F6 from the footer. An engine can send a clipboard event to the node of the page
 * selection, not to the focused grid, and the step before this one types into a cell first, so
 * it cannot show this case. Each copy is pasted into an empty cell, and undo takes back the edits.
 */
export async function verifyFocusCopy(page: Page): Promise<string> {
  const nameBox = page.getByRole('textbox', { name: 'Cell reference' });
  const content = page.getByRole('textbox', { name: 'Cell content' });
  const engine = page.context().browser()?.browserType().name() ?? 'browser';
  const scroller = page.locator('.xlsx-scroller').first();
  // A cell near the top left, under the headings, that the press reaches with no scroll.
  const pressCell = async () => {
    await scroller.evaluate(async (el) => {
      el.scrollTo(0, 0);
      for (let i = 0; i < 2; i += 1) await new Promise((done) => requestAnimationFrame(done));
    });
    const stage = (await page.locator('.xlsx-stage').first().boundingBox())!;
    await page.mouse.click(stage.x + 260, stage.y + 100);
    return nameBox.inputValue();
  };
  // No page selection, and with `blur` no focus: the state of a page that the user has not
  // touched.
  const clearSelection = (blur: boolean) =>
    page.evaluate((andBlur) => {
      document.getSelection()?.removeAllRanges();
      if (andBlur) (document.activeElement as HTMLElement | null)?.blur?.();
    }, blur);
  const done: string[] = [];
  for (const how of ['a press', 'Shift+F6'] as const) {
    if ((await goTo(page, TARGET)) !== '') throw new Error(`${TARGET} of the sample is not empty`);
    const ref = await pressCell();
    const was = await content.inputValue();
    const text = `Copied after ${how} in ${engine} ${Date.now()}`;
    await page.keyboard.type(text);
    await page.keyboard.press('Enter');
    await clearSelection(true);
    if (how === 'a press') {
      if ((await pressCell()) !== ref) throw new Error(`a second press did not select ${ref}`);
    } else {
      await goTo(page, ref);
      // F6 moves the focus to the footer, and Shift+F6 brings it back to the grid.
      await page.keyboard.press('F6');
      await clearSelection(false);
      await page.keyboard.press('Shift+F6');
    }
    const events = await recordClipboard(page);
    await page.keyboard.press('ControlOrMeta+c');
    await goTo(page, TARGET);
    await page.keyboard.press('ControlOrMeta+v');
    const pasted = await goTo(page, TARGET);
    const { events: seen } = await events();
    if (pasted !== text)
      throw new Error(
        `a copy after ${how} pasted "${pasted}" into ${TARGET}, not "${text}" (${seen})`
      );
    const undo = page.getByRole('button', { name: 'Undo', exact: true });
    let undos = 0;
    while ((await goTo(page, TARGET)) !== '' || (await goTo(page, ref)) !== was) {
      if (++undos > 3) throw new Error(`three undos did not take back the copy after ${how}`);
      await undo.click();
    }
    done.push(`after ${how} on ${ref}`);
  }
  return `Mod+C copied with no page selection ${done.join(' and ')}, and undo took back the edits`;
}
