// The formula bar steps of `verify-demo.ts`, kept apart so the runner stays under the line cap.
// They check what only a browser engine lays out: the scroll of the bar, the place of the caret
// after a press on the line count, and the line count beside the text. The edits of the bar, the
// name box, the selection, and the heading fits are happy-dom tests
// (`packages/react/src/__tests__/formula-bar.test.tsx`,
// `packages/core/src/editor/__tests__/formula-bar-edit.test.ts`, `xlsx-editor.test.ts`, and
// `fit.test.ts`).

import type { Page } from '@playwright/test';

/** The two fields of the formula bar, and a go-to through the name box. */
function barFields(page: Page) {
  const nameBox = page.getByRole('textbox', { name: 'Cell reference' });
  const content = page.getByRole('textbox', { name: 'Cell content' });
  const goTo = async (ref: string) => {
    await nameBox.click();
    await nameBox.fill(ref);
    await page.keyboard.press('Enter');
  };
  return { nameBox, content, goTo };
}

/**
 * The layout of the formula bar during an edit of Z200: after line breaks, the one-row field
 * shows the whole line of the caret; a press on the line count of line 2 puts the caret at its
 * end; the field's own undo removes one line, by the rule of each engine; the line count does
 * not cover the text; and a line of tabs keeps the caret in view. An undo restores Z200.
 */
export async function verifyFormulaBarLayout(page: Page): Promise<string> {
  const { content, goTo } = barFields(page);
  await goTo('Z200');
  const before = await content.inputValue();
  if (before !== '') throw new Error(`Z200 of the sample holds "${before}", not an empty cell`);
  await content.click();
  await content.fill('=1+2');
  await page.keyboard.press('Enter');
  await goTo('Z200');
  if ((await content.inputValue()) !== '=1+2') throw new Error('the bar did not commit =1+2');
  // After Alt+Enter, the one-row field shows the whole line of the caret, not part of two.
  await content.click();
  await page.keyboard.press('End');
  // A line wider than the field scrolls it; the next line shows from its start.
  await page.keyboard.type(' long'.repeat(60));
  for (const step of ['Alt+Enter', 'x', 'Alt+Enter', 'y']) await page.keyboard.press(step);
  const lines = await content.evaluate((el) => {
    const line = parseFloat(getComputedStyle(el).lineHeight);
    return { top: el.scrollTop, line, left: el.scrollLeft };
  });
  if (lines.left !== 0)
    throw new Error(`the bar hid the start of line 3 (scrollLeft ${lines.left})`);
  if (lines.top !== 2 * lines.line)
    throw new Error(`the bar showed scrollTop ${lines.top}, not line 3 (${2 * lines.line})`);
  // A press on the line count puts the caret at the end of the line that the field shows.
  // Home moves the caret away from the end of line 2, so only the press can bring it back.
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('Home');
  // The line count takes no pointer events, so the press goes to the box under it, as a user's does.
  const lineCount = (await page.locator('.xlsx-cell-content-lines').first().boundingBox())!;
  await page.mouse.click(lineCount.x + lineCount.width / 2, lineCount.y + lineCount.height / 2);
  const counted = await content.evaluate((node) => {
    const el = node as HTMLTextAreaElement;
    return {
      caret: el.selectionStart,
      end: el.value.indexOf('\n', el.value.indexOf('\n') + 1),
      focused: el === document.activeElement,
    };
  });
  if (!counted.focused || counted.caret !== counted.end)
    throw new Error(
      `a press on the line count put the caret at ${counted.caret}, not ${counted.end}`
    );
  // The caret goes back to the end of the text, where the undo check below starts.
  await page.keyboard.press('ControlOrMeta+End');
  // The field's own undo removes typed text and a line break, as in the cell editor.
  const typed = await content.inputValue();
  for (let i = 0; i < 2; i += 1) await page.keyboard.press('ControlOrMeta+z');
  const undone = await content.inputValue();
  // Chromium and Firefox record each inserted break as its own undo step. WebKit can group a
  // run of typing with the breaks into one step, as in any text field there, so its undo may
  // remove more: at least the break and character that one Chromium step removes, and never
  // more than the typing of this step.
  const engine = page.context().browser()?.browserType().name();
  const oneLineLess = undone === typed.slice(0, -2);
  const grouped =
    engine === 'webkit' &&
    undone.length <= typed.length - 2 &&
    typed.startsWith(undone) &&
    undone.startsWith('=1+2');
  if (!oneLineLess && !grouped)
    throw new Error(
      `an undo in the bar gave "${undone}" for "${typed}", not ${
        engine === 'webkit' ? 'at least one line less within the typing' : 'one line less'
      }`
    );
  const undid = oneLineLess ? 'a line break' : 'the typed run, as one step';
  // The line count sits beside the text, so a caret at the end of a long line is not under it.
  await page.keyboard.press('Alt+Enter');
  await page.keyboard.type(' word'.repeat(90));
  const over = await content.evaluate((el) => {
    const count = el.parentElement?.querySelector('.xlsx-cell-content-lines');
    return count ? el.getBoundingClientRect().right - count.getBoundingClientRect().left : 0;
  });
  if (over > 0) throw new Error(`the line count covers ${over} px of the bar's text`);
  // A line of tabs is wider than a canvas measures it: the caret stays in view.
  await page.keyboard.press('Alt+Enter');
  await page.keyboard.insertText('\t'.repeat(120));
  await page.keyboard.type('tab');
  const tabs = await page
    .waitForFunction(() => (document.activeElement as HTMLElement).scrollLeft > 0, null, {
      timeout: 2000,
    })
    .then(() => true)
    .catch(() => false);
  if (!tabs) {
    const state = await content.evaluate((el) => [el.scrollLeft, el.scrollWidth, el.clientWidth]);
    throw new Error(`the bar scrolled a line of tabs back past its caret (${state.join(', ')})`);
  }
  await page.keyboard.press('Escape');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  if ((await content.inputValue()) !== before) throw new Error('an undo did not restore Z200');
  return `the bar showed the line of the caret, a press on the line count of line 2 put the caret at its end, its undo removed ${undid}, the line count left the text clear, a line of tabs kept the caret in view, and an undo restored Z200`;
}

/**
 * A press on the line count of a field without the keyboard puts the caret at the end of the
 * first line, which the field shows, and the field scrolls to show it: on a long line and on a
 * line of tabs. The step leaves Z300 as it found it.
 */
export async function verifyLineCountPress(page: Page): Promise<string> {
  const { content, goTo } = barFields(page);
  // A press on the line count of a field without the keyboard puts the caret at the end of the
  // first line, which the field shows. The field stays on that line, and it scrolls right to
  // show the caret at the end of a line wider than the field.
  await goTo('Z300');
  const z300 = await content.inputValue();
  if (z300 !== '') throw new Error(`Z300 of the sample holds "${z300}", not an empty cell`);
  const long = ' long'.repeat(80);
  await content.click();
  await page.keyboard.insertText(long);
  await page.keyboard.press('Alt+Enter');
  await page.keyboard.insertText('gamma');
  await page.keyboard.press('Enter');
  await goTo('Z300');
  // The caret goes to the start, so the field shows the start of the line before the press.
  // The check proves the result of the press. It does not prove which code scrolls the field.
  const start = await content.evaluate((node) => {
    const el = node as HTMLTextAreaElement;
    el.setSelectionRange(0, 0);
    return el.scrollLeft;
  });
  if (start !== 0) throw new Error(`the bar showed Z300 at scrollLeft ${start}, not at the start`);
  const countBox = (await page.locator('.xlsx-cell-content-lines').first().boundingBox())!;
  await page.mouse.click(countBox.x + countBox.width / 2, countBox.y + countBox.height / 2);
  const pressed = await content.evaluate((node) => {
    const el = node as HTMLTextAreaElement;
    return [el.selectionStart, el.scrollTop, el.scrollLeft];
  });
  if (pressed[0] !== long.length || pressed[1] !== 0 || !(pressed[2] > 0))
    throw new Error(
      `a press on the line count gave caret ${pressed.join(', ')} (offset, top, left)`
    );
  await page.keyboard.press('Escape');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  if ((await content.inputValue()) !== z300) throw new Error('an undo did not restore Z300');
  // A first line shorter than the second: the field keeps the scroll of the old caret at the
  // end of line 2, and the press still puts the caret of line 1 at the right edge, with the
  // text before it in view.
  await content.click();
  const shorter = ' short'.repeat(60);
  await page.keyboard.insertText(shorter);
  await page.keyboard.press('Alt+Enter');
  await page.keyboard.insertText(' longer'.repeat(120));
  await page.keyboard.press('Enter');
  await goTo('Z300');
  const shortBox = (await page.locator('.xlsx-cell-content-lines').first().boundingBox())!;
  await page.mouse.click(shortBox.x + shortBox.width / 2, shortBox.y + shortBox.height / 2);
  const edge = await content.evaluate((node) => {
    const el = node as HTMLTextAreaElement;
    const style = getComputedStyle(el);
    const context = document.createElement('canvas').getContext('2d')!;
    context.font = style.font;
    const line = el.value.slice(0, el.selectionStart);
    const x = parseFloat(style.paddingLeft) + context.measureText(line).width - el.scrollLeft;
    return { caret: el.selectionStart, x: Math.round(x), width: el.clientWidth };
  });
  if (edge.caret !== shorter.length || edge.x < edge.width / 2 || edge.x > edge.width)
    throw new Error(
      `a press on the line count of a short line gave caret ${edge.caret} at x ${edge.x} of ${edge.width}`
    );
  await page.keyboard.press('Escape');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  if ((await content.inputValue()) !== z300) throw new Error('an undo did not restore Z300');
  // On a first line of tabs, the widest line, the press scrolls the field to its end, so the
  // caret at the end of that line is in view. The tab stops come from layout, not a guess.
  await content.click();
  const tabLine = `${'\t'.repeat(120)}end`;
  await page.keyboard.insertText(tabLine);
  await page.keyboard.press('Alt+Enter');
  await page.keyboard.insertText('b');
  await page.keyboard.press('Enter');
  await goTo('Z300');
  await content.evaluate((node) => (node as HTMLTextAreaElement).setSelectionRange(0, 0));
  const tabBox = (await page.locator('.xlsx-cell-content-lines').first().boundingBox())!;
  await page.mouse.click(tabBox.x + tabBox.width / 2, tabBox.y + tabBox.height / 2);
  const tabbed = await content.evaluate((node) => {
    const el = node as HTMLTextAreaElement;
    return [el.selectionStart, el.scrollLeft, el.scrollWidth - el.clientWidth];
  });
  if (tabbed[0] !== tabLine.length || tabbed[1] < tabbed[2] - 2)
    throw new Error(
      `a press on the line count of tabs gave ${tabbed.join(', ')} (offset, left, end)`
    );
  // A scaled formula bar gives the same scroll: the tab line is measured in layout pixels.
  await page.keyboard.press('Escape');
  const bar = page.locator('.xlsx-formula-bar').first();
  await bar.evaluate((el) => ((el as HTMLElement).style.transform = 'scale(0.5)'));
  await content.evaluate((node) => (node as HTMLTextAreaElement).setSelectionRange(0, 0));
  const scaledBox = (await page.locator('.xlsx-cell-content-lines').first().boundingBox())!;
  await page.mouse.click(scaledBox.x + scaledBox.width / 2, scaledBox.y + scaledBox.height / 2);
  const scaled = await content.evaluate((node) => {
    const el = node as HTMLTextAreaElement;
    return [el.selectionStart, el.scrollLeft, el.scrollWidth - el.clientWidth];
  });
  await bar.evaluate((el) => ((el as HTMLElement).style.transform = ''));
  if (scaled[0] !== tabLine.length || scaled[1] < scaled[2] - 2)
    throw new Error(
      `a press on the line count of a scaled bar gave ${scaled.join(', ')} (offset, left, end)`
    );
  await page.keyboard.press('Escape');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  if ((await content.inputValue()) !== z300) throw new Error('an undo did not restore Z300');
  return 'a press on the line count put the caret at the end of a long line, a line shorter than the next, and a line of tabs, in view, also in a scaled bar, and an undo restored Z300';
}
