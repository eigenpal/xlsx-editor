/** @jsxImportSource react */
// The save of the demo in a browser with no file save picker, in happy-dom (`demo-harness.tsx`):
// Save starts enabled; each way to save, and each way to repeat a save, starts one download that
// holds the workbook at its save; a retry after a failed save starts a download; and the notice
// names the question before the page closes only when unsaved edits make the page ask. The link
// of each download is recorded here. The browser check (`scripts/verify-demo.ts`) keeps one real
// download, whose bytes it reads, and the layout of the page in a phone-wide window.

import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, waitFor } from '@testing-library/react';
import { SAMPLE_BOLD_EDIT } from '../../../../scripts/sample-workbook/edits.ts';
import { downloadNotice, SAVE_REPEAT_MS } from '../save.ts';
import { formatAt, mountDemo, readSaved, settle, WAIT_MS, type Demo } from './demo-harness.tsx';

const NAME = 'sample-edited.xlsx';
const { sheet, ref } = SAMPLE_BOLD_EDIT;

/** The downloads that the page started: the name and the bytes of each linked file. */
function recordDownloads() {
  const blobs = new Map<string, Blob>();
  const all: { name: string; blob: Blob | undefined }[] = [];
  const held = { create: URL.createObjectURL, revoke: URL.revokeObjectURL };
  const click = HTMLAnchorElement.prototype.click;
  let refuseNext = false;
  URL.createObjectURL = (blob: Blob) => {
    if (refuseNext) {
      refuseNext = false;
      throw new Error('the test refuses the download');
    }
    const url = `blob:demo/${blobs.size}`;
    blobs.set(url, blob);
    return url;
  };
  URL.revokeObjectURL = () => {};
  HTMLAnchorElement.prototype.click = function (this: HTMLAnchorElement) {
    all.push({ name: this.download, blob: blobs.get(this.getAttribute('href') ?? '') });
  };
  return {
    all,
    refuseNext: () => (refuseNext = true),
    restore: () => {
      URL.createObjectURL = held.create;
      URL.revokeObjectURL = held.revoke;
      HTMLAnchorElement.prototype.click = click;
    },
  };
}

let downloads: ReturnType<typeof recordDownloads> | null = null;
afterEach(() => {
  cleanup();
  downloads?.restore();
  downloads = null;
});

/** Runs `act`, and returns the bytes of the one file that it downloads. */
async function downloadOf(
  demo: Demo,
  how: string,
  press: () => void,
  unsaved = true
): Promise<Uint8Array> {
  const all = downloads!.all;
  const before = all.length;
  press();
  await waitFor(() => expect([how, all.length]).toEqual([how, before + 1]), { timeout: WAIT_MS });
  // One save downloads one file: a repeat or a second handler would start two. The wait also
  // ends the time in which the next save of this test counts as a repeat.
  await settle(SAVE_REPEAT_MS + 100);
  expect([how, all.length - before]).toEqual([how, 1]);
  const { name, blob } = all.at(-1)!;
  expect([how, name]).toEqual([how, NAME]);
  // The whole text: the notice with no unsaved edit is the start of the other notice.
  await waitFor(() => expect([how, demo.notice()]).toEqual([how, downloadNotice(NAME, unsaved)]), {
    timeout: WAIT_MS,
  });
  return new Uint8Array(await blob!.arrayBuffer());
}

const click = (el: Element) => act(() => void fireEvent.click(el));

/** A double click: two clicks, then the double click event. */
const doubleClick = (el: Element) => () => {
  click(el);
  click(el);
  act(() => void fireEvent.doubleClick(el));
};

/** A held Enter on a button: each key event that the page does not cancel clicks the button. */
const heldEnter = (el: HTMLElement) => () => {
  act(() => el.focus());
  for (let press = 0; press < 3; press += 1) {
    const e = new KeyboardEvent('keydown', {
      key: 'Enter',
      repeat: press > 0,
      bubbles: true,
      cancelable: true,
    });
    act(() => void el.dispatchEvent(e));
    if (!e.defaultPrevented) click(el);
  }
};

/** A download from the top bar, in view only mode or not. The notice must match the question. */
async function noticeMatches(
  demo: Demo,
  how: string,
  viewOnly: boolean,
  unsaved: boolean
): Promise<void> {
  if (demo.viewOnlyBox.checked !== viewOnly) click(demo.viewOnlyBox);
  // An old notice with the same text must not pass for the notice of this save.
  if (demo.view.queryByRole('button', { name: 'Close the message' })) demo.closeMessage();
  await downloadOf(demo, how, () => click(demo.headerSave), unsaved);
  expect([how, demo.leaveAsks()]).toEqual([how, unsaved]);
  if (demo.viewOnlyBox.checked) click(demo.viewOnlyBox);
}

describe('a save with no file save picker', () => {
  test('starts enabled, and each save and each repeat of a save starts one download', async () => {
    downloads = recordDownloads();
    const demo = await mountDemo({ picker: false });
    expect('showSaveFilePicker' in window).toBe(false);
    const save = demo.headerSave;
    expect([
      save.getAttribute('aria-disabled'),
      save.getAttribute('aria-describedby'),
      save.getAttribute('title'),
    ]).toEqual([null, null, 'Save']);
    expect(document.body.textContent).not.toContain('Nothing was saved');
    // A screen reader announces a change in a live region that is already on the page.
    const region = document.querySelector<HTMLElement>('.demo-message__status')!;
    expect([region.getAttribute('role'), region.textContent]).toEqual(['status', '']);
    await demo.tab(sheet);

    await noticeMatches(demo, 'save with no edit', false, false);
    // The first notice filled the region that was on the page before the save.
    expect(document.querySelector('.demo-message__status')).toBe(region);
    expect(region.textContent).not.toBe('');

    const repeats: readonly [string, () => void][] = [
      ['double click on the top bar Save', doubleClick(demo.headerSave)],
      ['double click on the toolbar Save', doubleClick(demo.toolbarSave)],
      ['held Enter on the top bar Save', heldEnter(demo.headerSave)],
      ['held Enter on the toolbar Save', heldEnter(demo.toolbarSave)],
      [
        'held Control+S',
        () => {
          act(() => demo.grid.focus());
          for (let press = 0; press < 3; press += 1) demo.saveKey(demo.grid, { repeat: press > 0 });
        },
      ],
    ];
    for (const [how, press] of repeats) {
      if (demo.view.queryByRole('button', { name: 'Close the message' })) demo.closeMessage();
      await downloadOf(demo, how, press, false);
    }

    // A retry right after a failed save starts a download.
    const consoleError = console.error;
    console.error = () => {};
    try {
      downloads.refuseNext();
      click(demo.headerSave);
      await demo.alertSays('Nothing was saved');
    } finally {
      console.error = consoleError;
    }
    await downloadOf(demo, 'retry right after a failed save', () => click(demo.headerSave), false);
    await noticeMatches(demo, 'view only save with no edit', true, false);
  }, 120_000);

  test('the top bar, the toolbar, and the keys each download the edit, and leaving still asks', async () => {
    downloads = recordDownloads();
    const demo = await mountDemo({ picker: false });
    await demo.tab(sheet);
    const saves: readonly [string, () => void][] = [
      ["header's Save", () => click(demo.headerSave)],
      ["toolbar's Save", () => click(demo.toolbarSave)],
      [
        'Control+S',
        () => {
          act(() => demo.grid.focus());
          demo.saveKey();
        },
      ],
    ];
    let expected = false;
    for (const [how, press] of saves) {
      demo.bold();
      expected = !expected;
      expect([how, demo.leaveAsks()]).toEqual([how, true]);
      const bytes = await downloadOf(demo, how, press);
      const pkg = readSaved(bytes);
      expect([how, formatAt(pkg, sheet, ref).bold]).toEqual([how, expected]);
      // A download cannot confirm the file, so leaving still asks.
      expect([how, demo.leaveAsks()]).toEqual([how, true]);
    }
    await noticeMatches(demo, 'view only save over unsaved edits', true, true);
    expect(demo.alerts()).toEqual([]);
  }, 120_000);
});
