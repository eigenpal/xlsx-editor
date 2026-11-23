/** @jsxImportSource react */
// What the demo keeps when an open overlaps the load of a large workbook, in happy-dom
// (`demo-harness.tsx`). Each test acts while the editor still loads, and fails when the load ended
// before that action, because the test then proves nothing. The browser check
// (`scripts/verify-demo.ts`) slowed the CPU of Chromium for these steps, so they ran in Chromium
// only; here the load yields to the test between its steps, in every run.

import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, waitFor } from '@testing-library/react';
import { flagsOf, mountDemo, settle, sheetsOf, WAIT_MS } from './demo-harness.tsx';
import {
  EDIT,
  holdsText,
  largeWorkbook,
  openFile,
  pressHeadings,
  savesEnd,
  setPicker,
  smallWorkbook,
  stillLoading,
  timedOpen,
} from './demo-race-helpers.ts';

afterEach(cleanup);

describe('an open during a load', () => {
  test('edits stay protected during a headings reload: a workbook that opens during it asks first', async () => {
    const demo = await mountDemo();
    await demo.headingsTo(false);
    await openFile(demo, 'Reload', largeWorkbook('Reload'));
    demo.typeInGrid(EDIT);
    const asked = demo.dialogs.seen.length;
    demo.dialogs.accept = false;
    try {
      pressHeadings(demo);
      await demo.loadStarts();
      demo.pickFile('second.xlsx', smallWorkbook('Second'));
      await demo.statusSays('Your edits are kept');
    } finally {
      demo.dialogs.accept = true;
    }
    expect(demo.dialogs.seen.length).toBe(asked + 1);
    // The picked file asked after its checks, while the reload still ran: else the test proves
    // nothing.
    expect(demo.dialogs.whileLoading.at(-1)).toBe(true);
    expect(demo.shownSheet()).toBe('Reload');
    expect(demo.leaveAsks()).toBe(true);
    expect(holdsText(await demo.save(), EDIT)).toBe(true);
  });

  test('a picked file that does not open during a load leaves the name and the flags of the loaded workbook', async () => {
    const demo = await mountDemo();
    await demo.headingsTo(true);
    await openFile(demo, 'Before', smallWorkbook('Before'));
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
    expect(flagsOf(saved)).toBe(flagsOf(large));
  });

  test('Save presses during a load do not delay the open, and never write the workbook that the load replaces', async () => {
    const demo = await mountDemo();
    await demo.headingsTo(true);
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
    await demo.headingsTo(false);
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

  test('a no to a newer file also stops the earlier accepted file before its load', async () => {
    const demo = await mountDemo();
    // With the option on, the accepted file is rewritten in steps before its load starts.
    await demo.headingsTo(true);
    await openFile(demo, 'Kept', smallWorkbook('Kept'));
    demo.typeInGrid(EDIT);
    const asked = demo.dialogs.seen.length;
    try {
      demo.pickFile('earlier.xlsx', largeWorkbook('Earlier'));
      await waitFor(() => expect(demo.dialogs.seen.length).toBe(asked + 1), {
        timeout: WAIT_MS,
        interval: 1,
      });
      demo.dialogs.accept = false;
      // The newer file asks only while the earlier file has not loaded: a load that ran would
      // have dropped the edit, and a load that runs answers for it.
      demo.pickFile('newer.xlsx', smallWorkbook('Newer'));
      await waitFor(() => expect(demo.dialogs.seen.length).toBe(asked + 2), {
        timeout: WAIT_MS,
        interval: 1,
      });
    } finally {
      demo.dialogs.accept = true;
    }
    // The earlier file ends its steps here. An open that still counts would replace the sheet.
    await settle(500);
    await waitFor(() => expect(demo.busy()).toBe(''), { timeout: WAIT_MS });
    expect(demo.shownSheet()).toBe('Kept');
    expect(demo.workbookName()).toBe('kept.xlsx');
    expect(demo.leaveAsks()).toBe(true);
    expect(holdsText(await demo.save(), EDIT)).toBe(true);
  });
});
