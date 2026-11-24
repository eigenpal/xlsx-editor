/** @jsxImportSource react */
// What the demo keeps when an open overlaps the load of a large workbook, in happy-dom
// (`demo-harness.tsx`). Each test acts while the editor still loads, and fails when the load ended
// before that action, because the test then proves nothing. The browser check
// (`scripts/verify-demo.ts`) slowed the CPU of Chromium for these steps, so they ran in Chromium
// only; here the load yields to the test between its steps, in every run.

import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, waitFor } from '@testing-library/react';
import { mountDemo, sheetsOf, WAIT_MS } from './demo-harness.tsx';
import {
  holdsText,
  largeWorkbook,
  openFile,
  savesEnd,
  setPicker,
  smallWorkbook,
  stillLoading,
  timedOpen,
} from './demo-race-helpers.ts';

afterEach(cleanup);

describe('an open during a load', () => {
  test('a picked file that does not open during a load leaves the loaded workbook, its name, and its heading flag', async () => {
    const demo = await mountDemo();
    // The workbook before shows its headings, and the loaded workbook hides them.
    await openFile(demo, 'Before', smallWorkbook('Before', { headings: true }));
    const large = largeWorkbook('Loaded', { headings: false });
    demo.pickFile('loaded.xlsx', large);
    await demo.loadStarts();
    demo.pickFile('broken.xlsx', 'not a workbook');
    // Whether the load still ran when the refusal first showed.
    let loadingAtRefusal: boolean | undefined;
    await waitFor(
      () => {
        demo.alertNow('Could not open broken.xlsx');
        loadingAtRefusal ??= demo.loading();
      },
      { timeout: WAIT_MS, interval: 1 }
    );
    if (!loadingAtRefusal)
      throw new Error(
        'the load ended before the picked file was refused, so the test proves nothing'
      );
    await demo.showsSheet('Loaded');
    expect(demo.workbookName()).toBe('loaded.xlsx');
    const saved = await demo.save();
    expect(demo.services.suggested.at(-1)).toBe('loaded-edited.xlsx');
    expect(holdsText(saved, 'showRowColHeaders="0"')).toBe(true);
  });

  test('Save presses during a load do not delay the open, and never write the workbook that the load replaces', async () => {
    const demo = await mountDemo();
    // The first open warms the page, so the two measured opens are alike.
    await timedOpen(demo, 'Warm');
    const quietMs = await timedOpen(demo, 'Quiet');
    await openFile(demo, 'First', smallWorkbook('First'));
    const counts = { duringLoad: 0, pickedDuringLoad: 0, refusedDuringLoad: 0 };
    // The browser refuses a file dialog that opens long after the press.
    demo.services.activation = false;
    setPicker(async () => {
      if (demo.loading()) counts.pickedDuringLoad += 1;
      throw new DOMException('The picker needs a recent user action.', 'SecurityError');
    });
    const presses = setInterval(() => {
      const during = demo.loading();
      demo.headerSave.click();
      demo.toolbarSave.click();
      if (!during) return;
      counts.duringLoad += 1;
      if (demo.alerts().some((a) => a.includes('a workbook is opening')))
        counts.refusedDuringLoad += 1;
    }, 20);
    let pressedMs = 0;
    try {
      pressedMs = await timedOpen(demo, 'Second');
    } finally {
      clearInterval(presses);
      demo.services.activation = true;
      setPicker(demo.services.keepingPicker);
    }
    // A save that was still running ends here, so the next press starts from a quiet page.
    await savesEnd(demo);
    expect(counts.duringLoad).toBeGreaterThan(0);
    expect(counts.pickedDuringLoad).toBe(0);
    expect(counts.refusedDuringLoad).toBeGreaterThan(0);
    expect(pressedMs).toBeLessThanOrEqual(quietMs * 1.5);
    if (demo.view.queryByRole('button', { name: 'Close the message' })) demo.closeMessage();
    expect(sheetsOf(await demo.save())).toEqual(['Second']);
  });

  test('keys on the body during a load stay in the page, and each says why', async () => {
    const demo = await mountDemo();
    // After its dialog, the file input can keep the focus. The open gives the focus to the grid.
    act(() => document.querySelector<HTMLElement>('.demo-topbar input[type=file]')!.focus());
    await openFile(demo, 'Keys', smallWorkbook('Keys'));
    await waitFor(() => expect(document.activeElement?.closest('.xlsx-scroller')).not.toBeNull(), {
      timeout: WAIT_MS,
    });
    const keys: { key: string; prevented: boolean; onBody: boolean }[] = [];
    const onKey = (e: KeyboardEvent) =>
      keys.push({
        key: e.key.toLowerCase(),
        prevented: e.defaultPrevented,
        onBody: e.target === document.body,
      });
    window.addEventListener('keydown', onKey, true);
    const blur = () => act(() => (document.activeElement as HTMLElement | null)?.blur());
    try {
      demo.pickFile('keys-large.xlsx', largeWorkbook('KeysLarge'));
      await demo.loadStarts();
      blur();
      demo.saveKey(document.body);
      demo.alertNow('a workbook is opening');
      demo.key('x', {}, document.body);
      demo.alertNow('The key did nothing');
      stillLoading(demo, 'the keys were pressed');
      await demo.showsSheet('KeysLarge');
    } finally {
      window.removeEventListener('keydown', onKey, true);
    }
    for (const [key, count] of [
      ['s', 1],
      ['x', 1],
    ] as const) {
      const presses = keys.filter((k) => k.key === key);
      expect([key, presses.length]).toEqual([key, count]);
      expect(presses.every((k) => k.onBody && k.prevented)).toBe(true);
    }
    await demo.save();
  });
});
