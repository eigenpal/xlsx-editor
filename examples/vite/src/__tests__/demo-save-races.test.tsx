/** @jsxImportSource react */
// What the demo keeps when a save overlaps other work, in happy-dom (`demo-harness.tsx`): a later
// edit keeps the message of a failed or cancelled save; during the download of the sample at start
// and during a headings switch, each Save says why it waits and no save starts; while a workbook
// is open and a picked file runs its steps before its load, the toolbar Save says that a workbook
// is loading; and after a save that the browser refused for a late dialog, the next press writes
// the kept bytes. These were steps of the browser check (`scripts/verify-demo.ts`).

import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, waitFor } from '@testing-library/react';
import { SAVE_FAILED } from '../save.ts';
import { mountDemo, settle, WAIT_MS } from './demo-harness.tsx';
import {
  closeAlert,
  countingPicker,
  EDIT,
  largeWorkbook,
  openFile,
  pressHeadings,
  pressSave,
  pressToolbarSave,
  savesEnd,
  setPicker,
  smallWorkbook,
  watchSaving,
} from './demo-race-helpers.ts';

afterEach(cleanup);

describe('a save that overlaps other work', () => {
  test('a later edit does not replace the error of a failed write or the silence of a cancel', async () => {
    const demo = await mountDemo();
    await demo.headingsTo(false);
    // The demo reports a failed write on the console, and this test causes that failure.
    const consoleError = console.error;
    console.error = (...args: unknown[]) => {
      if (!String(args[0]).startsWith('The save of the demo failed')) consoleError(...args);
    };
    try {
      for (const mode of ['failed', 'cancelled'] as const) {
        await openFile(demo, mode, smallWorkbook(mode));
        demo.typeInGrid('first-edit');
        let picked = false;
        let release = () => {};
        const held = new Promise<void>((resolve) => (release = resolve));
        setPicker(async () => {
          picked = true;
          await held;
          if (mode === 'cancelled') throw new DOMException('The user cancelled.', 'AbortError');
          return {
            createWritable: async () => {
              throw new DOMException('The disk is full.', 'NotReadableError');
            },
          };
        });
        pressSave(demo);
        await waitFor(() => expect(picked).toBe(true), { timeout: WAIT_MS });
        demo.typeInGrid('later-edit');
        pressSave(demo);
        release();
        await savesEnd(demo);
        if (mode === 'failed') await demo.alertSays(SAVE_FAILED);
        await settle(200);
        expect(document.body.textContent).not.toContain('The save holds the workbook');
        setPicker(demo.services.keepingPicker);
        await demo.save();
      }
    } finally {
      console.error = consoleError;
      setPicker(demo.services.keepingPicker);
    }
  });

  test('during the download of the sample, the headings switch and each Save say why they wait, and no save starts', async () => {
    let release = () => {};
    const demo = await mountDemo({
      beforeStart: (services) => (release = services.hold('sample.xlsx')),
      waitForSample: false,
    });
    const counts = countingPicker(false);
    const savingSeen = watchSaving(demo);
    try {
      pressHeadings(demo);
      await demo.alertSays('Could not switch the headings: a workbook is opening');
      expect(demo.headingsBox.checked).toBe(true);
      closeAlert(demo);
      // The download is part of the open: a save key on the body stays in the page and says why.
      act(() => (document.activeElement as HTMLElement | null)?.blur());
      const onBody = new KeyboardEvent('keydown', {
        key: 's',
        ctrlKey: true,
        bubbles: true,
        cancelable: true,
      });
      act(() => void document.body.dispatchEvent(onBody));
      expect(onBody.defaultPrevented).toBe(true);
      await demo.alertSays('Could not save: a workbook is opening');
      closeAlert(demo);
      act(() => demo.grid.focus());
      demo.saveKey();
      await demo.alertSays('Could not save: a workbook is opening');
      await waitFor(() => expect(demo.toolbarSave.getAttribute('aria-disabled')).toBe('true'), {
        timeout: WAIT_MS,
      });
      // No workbook is open yet, so the toolbar gives that reason, not the text of `opening`.
      expect(demo.toolbarSave.getAttribute('title')).toContain('The editor is not ready');
      closeAlert(demo);
      // The page takes a press of the disabled toolbar Save, and says why it waits.
      pressToolbarSave(demo);
      await demo.alertSays('Could not save: a workbook is opening');
      // A save that started shows its status after 300 ms, and opens the file dialog at its end.
      await settle(1000);
      expect(counts.calls).toBe(0);
      expect(savingSeen()).toBe(false);
      closeAlert(demo);
      release();
      await demo.showsSheet('Welcome');
    } finally {
      release();
      savingSeen();
      setPicker(demo.services.keepingPicker);
    }
  });

  test('while a workbook is open and a picked file runs its steps before its load, the toolbar Save says that a workbook is loading', async () => {
    const demo = await mountDemo();
    // The option is on, so the picked file is rewritten in steps before its load starts.
    await demo.headingsTo(true);
    const counts = countingPicker(false);
    try {
      demo.pickFile('slow.xlsx', largeWorkbook('Slow'));
      // The title as it first shows, and whether the sample was still on screen then.
      const seen = { title: null as string | null, sampleShown: false };
      await waitFor(
        () => {
          if (demo.toolbarSave.getAttribute('aria-disabled') !== 'true') throw new Error('enabled');
          seen.title ??= demo.toolbarSave.getAttribute('title');
          seen.sampleShown ||= !demo.loading() && demo.shownSheet() === 'Welcome';
        },
        { timeout: WAIT_MS, interval: 1 }
      );
      expect(seen.sampleShown).toBe(true);
      expect(seen.title ?? '').toContain('A workbook is loading');
      pressToolbarSave(demo);
      await demo.alertSays('Could not save: a workbook is opening');
      await demo.showsSheet('Slow');
      expect(counts.calls).toBe(0);
    } finally {
      setPicker(demo.services.keepingPicker);
    }
  });

  test('during a headings switch, Control+S, the toolbar Save, and the top bar Save give one text, and no save starts', async () => {
    const demo = await mountDemo();
    await demo.headingsTo(false);
    await openFile(demo, 'Switch', largeWorkbook('Switch'));
    const counts = countingPicker(false);
    const savingSeen = watchSaving(demo);
    const text = 'Could not save: the headings are switching';
    try {
      pressHeadings(demo);
      await waitFor(() => expect(demo.busy()).toContain('Switching the headings'), {
        timeout: WAIT_MS,
        interval: 1,
      });
      demo.saveKey();
      await demo.alertSays(text);
      closeAlert(demo);
      pressToolbarSave(demo);
      await demo.alertSays(text);
      closeAlert(demo);
      pressSave(demo);
      await demo.alertSays(text);
      // The switch still runs, so the presses landed inside it.
      expect(demo.busy()).toContain('Switching the headings');
      // The status strip of the editor shows no refusal of its own.
      expect(document.querySelector('.xlsx-refusal')?.textContent ?? '').toBe('');
      await waitFor(() => expect(demo.busy()).not.toContain('Switching the headings'), {
        timeout: WAIT_MS,
      });
      expect(demo.headingsBox.checked).toBe(true);
      expect(counts.calls).toBe(0);
      expect(savingSeen()).toBe(false);
    } finally {
      savingSeen();
      setPicker(demo.services.keepingPicker);
    }
  });

  for (const press of ['key', 'toolbar Save'] as const)
    test(`after a save that the browser refused for a late dialog, the ${press} writes the kept bytes`, async () => {
      const demo = await mountDemo();
      await demo.headingsTo(true);
      // Each press opens its own sheet, so the test does not pass on another workbook.
      const sheet = press === 'key' ? 'Ready' : 'Toolbar';
      await openFile(demo, sheet, largeWorkbook(sheet));
      demo.typeInGrid(EDIT);
      const counts = countingPicker(true);
      // With no recent click, the demo keeps the bytes and asks for a press of Save.
      demo.services.activation = false;
      let savingSeen = () => false;
      try {
        pressSave(demo);
        await demo.statusSays('The workbook is ready to save');
        demo.services.activation = true;
        act(() => demo.grid.focus());
        await savesEnd(demo);
        savingSeen = watchSaving(demo);
        if (press === 'key') demo.saveKey();
        else pressToolbarSave(demo);
        await demo.statusSays('to kept.xlsx');
        // A new save that started would still show its status after 300 ms.
        await settle(1000);
        expect(savingSeen()).toBe(false);
        expect(counts.calls).toBe(2);
      } finally {
        demo.services.activation = true;
        savingSeen();
        setPicker(demo.services.keepingPicker);
      }
    });
});
