// The grid on a host page with editable content or a text selection, in each engine. A press on
// the grid moves the page selection into the grid (`clipboard-events.ts`), and an engine moves
// the focus to an editing host that holds the selection. So this step presses real keys, and it
// reads the cells and the focus after each one. Editing commands run with keys and with
// `execCommand`, because the keys are editing commands on macOS only. Only Chromium composes
// through an input method protocol. The other engines open the cell editor with the same key
// event and insert the text.

import type { Page } from '@playwright/test';
import { goTo } from './verify-demo-browser.ts';

/** The ways that a host page can surround the grid. */
const LAYOUTS = [
  { name: 'inside editable host content', editable: true, selected: false },
  { name: 'after a press with host text selected', editable: false, selected: true },
  { name: 'inside editable host content with its text selected', editable: true, selected: true },
] as const;

/** Presses a cell near the top left of the grid, and returns the reference of the name box. */
async function pressCell(page: Page): Promise<string> {
  await page
    .locator('.xlsx-scroller')
    .first()
    .evaluate(async (el) => {
      el.scrollTo(0, 0);
      for (let i = 0; i < 2; i += 1) await new Promise((done) => requestAnimationFrame(done));
    });
  const stage = (await page.locator('.xlsx-stage').first().boundingBox())!;
  await page.mouse.click(stage.x + 260, stage.y + 100);
  return page.getByRole('textbox', { name: 'Cell reference' }).inputValue();
}

/** Throws when the grid does not have the focus after `what`. */
async function expectGridFocus(page: Page, what: string): Promise<void> {
  const at = await page.evaluate(() => {
    const el = document.activeElement;
    return el?.classList.contains('xlsx-scroller') ? '' : `${el?.localName}.${el?.className}`;
  });
  if (at) throw new Error(`after ${what}, the focus is on ${at}, not on the grid`);
}

/** The text and the element count of the grid, which editing commands must not change. */
const gridContent = (page: Page) =>
  page
    .locator('.xlsx-scroller')
    .first()
    .evaluate(async (grid) => {
      await new Promise((done) => requestAnimationFrame(done));
      return `${grid.querySelectorAll('*').length}:${grid.textContent}`;
    });

/** Editing commands with the page selection in the grid: nothing changes, and the focus stays. */
async function expectNoEditingCommands(page: Page, name: string): Promise<void> {
  const before = await gridContent(page);
  await page.keyboard.press('Control+k');
  await page.keyboard.press('Control+d');
  await expectGridFocus(page, 'Control+K and Control+D');
  await page
    .locator('.xlsx-scroller')
    .first()
    .evaluate((grid) => {
      const doc = grid.ownerDocument;
      doc.getSelection()!.collapse(grid, 0);
      for (const command of ['delete', 'forwardDelete', 'insertText'])
        doc.execCommand(command, false, 'x');
    });
  await expectGridFocus(page, 'the editing commands');
  if ((await gridContent(page)) !== before)
    throw new Error(`${name}: an editing command changed the grid`);
}

/** Replaces the content of the active cell with `text` through the formula bar. */
async function typeInFormulaBar(page: Page, text: string): Promise<void> {
  await page.getByRole('textbox', { name: 'Cell content' }).click();
  await page.keyboard.press('ControlOrMeta+a');
  await page.keyboard.type(text);
  await page.keyboard.press('Enter');
}

/** Opens the cell editor with the key event of an input method, and commits `text` in it. */
async function typeWithInputMethod(page: Page, engine: string, text: string): Promise<void> {
  await page
    .locator('.xlsx-scroller')
    .first()
    .evaluate((grid) => {
      (grid as HTMLElement).focus();
      const init = { key: 'Process', bubbles: true, cancelable: true };
      grid.dispatchEvent(new KeyboardEvent('keydown', init));
    });
  const inEditor = await page.evaluate(
    () => document.activeElement?.classList.contains('xlsx-cell-editor') === true
  );
  if (!inEditor) throw new Error('the key event of an input method opened no cell editor');
  if (engine === 'chromium') {
    const session = await page.context().newCDPSession(page);
    await session.send('Input.imeSetComposition', {
      text,
      selectionStart: text.length,
      selectionEnd: text.length,
    });
    await session.send('Input.insertText', { text });
    await session.detach();
  } else await page.keyboard.insertText(text);
  await page.keyboard.press('Enter');
}

/** Adds the host text next to the surface, makes the host editable, and selects the text. */
const setUpHost = (page: Page, editable: boolean, selected: boolean) =>
  page
    .locator('.xlsx-scroller')
    .first()
    .evaluate(
      (grid, opts) => {
        const host = grid.closest('.xlsx-surface')!.parentElement!;
        const text = host.ownerDocument.createElement('p');
        text.id = 'verify-host-text';
        text.textContent = 'Host text';
        text.style.cssText = 'position:absolute;right:24px;bottom:24px;margin:0';
        host.append(text);
        if (opts.editable) host.setAttribute('contenteditable', 'true');
        if (opts.selected) host.ownerDocument.getSelection()!.selectAllChildren(text);
      },
      { editable, selected }
    );

/** Removes the host text and the editable state, and returns the text that the host kept. */
const tearDownHost = (page: Page) =>
  page
    .locator('.xlsx-scroller')
    .first()
    .evaluate((grid) => {
      const host = grid.closest('.xlsx-surface')!.parentElement!;
      host.removeAttribute('contenteditable');
      const text = host.ownerDocument.getElementById('verify-host-text');
      const kept = text?.textContent ?? '';
      text?.remove();
      host.ownerDocument.getSelection()?.removeAllRanges();
      return kept;
    });

/** A press on the grid that has the focus moves a page selection on host text into the grid. */
async function verifyPressOnFocusedGrid(page: Page): Promise<void> {
  await pressCell(page);
  await setUpHost(page, false, true);
  try {
    await expectGridFocus(page, 'the selection of host text');
    await pressCell(page);
    const inGrid = await page
      .locator('.xlsx-scroller')
      .first()
      .evaluate((grid) => grid.contains(grid.ownerDocument.getSelection()?.anchorNode ?? null));
    if (!inGrid) throw new Error('a press on the focused grid left the selection on host text');
  } finally {
    await tearDownHost(page);
  }
}

/**
 * Typing, Enter, copy, paste, editing commands, the formula bar, and input method text, for each
 * layout of `LAYOUTS`. The grid keeps the focus after each key, the cells hold the entered
 * values, the host text does not change, and undo takes back the edits.
 */
export async function verifyEditableHost(page: Page): Promise<string> {
  const engine = page.context().browser()?.browserType().name() ?? 'browser';
  const undo = page.getByRole('button', { name: 'Undo', exact: true });
  const done: string[] = [];
  for (const { name, editable, selected } of LAYOUTS) {
    const ref = await pressCell(page);
    const match = /^([A-Z]+)(\d+)$/.exec(ref);
    if (!match) throw new Error(`a press selected "${ref}", not one cell`);
    const cells = [0, 1, 2, 3, 4].map((i) => `${match[1]}${Number(match[2]) + i}`);
    const was: string[] = [];
    for (const cell of cells) was.push(await goTo(page, cell));
    const first = `First ${engine} ${Date.now()}`;
    const second = `Second ${engine} ${Date.now()}`;
    const barText = `Bar ${engine} ${Date.now()}`;
    const composed = `日本語 ${Date.now()}`;
    let kept: string;
    await setUpHost(page, editable, selected);
    try {
      if ((await pressCell(page)) !== ref) throw new Error(`a press did not select ${ref}`);
      await expectGridFocus(page, 'the press');
      await page.keyboard.type(first);
      await page.keyboard.press('Enter');
      await expectGridFocus(page, 'the first Enter');
      await page.keyboard.type(second);
      await page.keyboard.press('Enter');
      await expectGridFocus(page, 'the second Enter');
      if ((await pressCell(page)) !== ref) throw new Error(`a press did not select ${ref} again`);
      await page.keyboard.press('ControlOrMeta+c');
      await page.keyboard.press('ArrowDown');
      await page.keyboard.press('ArrowDown');
      await page.keyboard.press('ControlOrMeta+v');
      await expectGridFocus(page, 'the paste');
      await expectNoEditingCommands(page, name);
      await goTo(page, cells[3]!);
      await typeInFormulaBar(page, barText);
      await goTo(page, cells[4]!);
      await typeWithInputMethod(page, engine, composed);
      await expectGridFocus(page, 'the input method text');
    } finally {
      kept = await tearDownHost(page);
    }
    if (kept !== 'Host text') throw new Error(`${name}: the host text became "${kept}"`);
    const want = [first, second, first, barText, composed];
    for (const [i, cell] of cells.entries()) {
      const got = await goTo(page, cell);
      if (got !== want[i]) throw new Error(`${name}: ${cell} holds "${got}", not "${want[i]}"`);
    }
    const restored = async () => {
      for (const [i, cell] of cells.entries())
        if ((await goTo(page, cell)) !== was[i]) return false;
      return true;
    };
    let undos = 0;
    while (!(await restored())) {
      if (++undos > 8) throw new Error(`${name}: eight undos did not take back the edits`);
      await undo.click();
    }
    done.push(name);
  }
  await verifyPressOnFocusedGrid(page);
  return `typing, Enter, copy, paste, editing commands, the formula bar, and input method text worked ${done.join(', ')}, and a press on the focused grid took the page selection`;
}
