/** @jsxImportSource react */
// The opens of the demo over the open workbook, in happy-dom (`demo-harness.tsx`): a file that
// does not open changes nothing; opening another workbook over unsaved edits asks first, from
// the list, from a picked file, over a headings switch, for a value typed with no Enter, and
// after a cancelled save; and the newest pick is the workbook that opens. These were steps of
// the browser check (`scripts/verify-demo.ts`).

import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, waitFor } from '@testing-library/react';
import { buildXlsx } from '../../../../scripts/fixtures/builder.ts';
import { SAMPLE_BOLD_EDIT } from '../../../../scripts/sample-workbook/edits.ts';
import { SAMPLE } from '../sample.ts';
import {
  fixtureBytes,
  flagsOf,
  formatAt,
  mountDemo,
  readSaved,
  sampleBytes,
  sampleWithNoSheet,
  settle,
  WAIT_MS,
  type Demo,
} from './demo-harness.tsx';

afterEach(cleanup);

const { sheet, ref } = SAMPLE_BOLD_EDIT;
const boldIn = (bytes: Uint8Array) => formatAt(readSaved(bytes), sheet, ref).bold;

/** Mounts the demo with the headings option off, on the sheet of the Bold edit. */
async function onEditSheet(): Promise<Demo> {
  const demo = await mountDemo();
  await demo.setHeadings(false);
  demo.closeMessage();
  await demo.tab(sheet);
  return demo;
}

/** The last value that `typeWithNoEnter` typed. */
let typed = 0;

/**
 * Types a new value into the active cell with no Enter, so the cell editor stays open. Each value
 * differs from the one before, because a commit of the same value is no edit.
 */
const typeWithNoEnter = (demo: Demo, text = String((typed += 1))) => demo.typeInGrid(text, false);

describe('a file that does not open', () => {
  test('changes nothing, keeps its message over a change of selection, and a close from the keyboard keeps the focus', async () => {
    const demo = await onEditSheet();
    demo.bold();
    const flags = flagsOf(sampleBytes());
    // The headings option is on, so the save reads the flags that the demo keeps.
    await demo.setHeadings(true);
    for (const [name, bytes] of [
      ['broken.xlsx', 'not a workbook'],
      ['no-sheet.xlsx', sampleWithNoSheet()],
    ] as const) {
      demo.pickFile(name, bytes);
      await demo.alertSays(`Could not open ${name}`);
      expect(demo.shownSheet()).toBe(sheet);
      // A change of selection in the grid is no edit, so the message stays.
      act(() => demo.grid.focus());
      demo.key('ArrowDown', {}, demo.grid);
      demo.key('ArrowUp', {}, demo.grid);
      expect(demo.alerts().length).toBeGreaterThan(0);
      if (name === 'broken.xlsx') {
        // Enter on the close button presses it.
        act(
          () => void fireEvent.click(demo.view.getByRole('button', { name: 'Close the message' }))
        );
      } else {
        // Escape during a cell edit cancels the edit and keeps the message; the next one closes it.
        demo.typeInGrid('7', false);
        demo.key('Escape');
        expect(demo.alerts().length).toBeGreaterThan(0);
        demo.key('Escape');
      }
      expect(demo.alerts()).toEqual([]);
      const at = document.activeElement;
      expect(!!at?.closest('.xlsx-editor') && !at.closest('.demo-topbar, [role="toolbar"]')).toBe(
        true
      );
      if (name !== 'broken.xlsx') {
        // The demo takes the Escape that closes the message, so it does not arm the grid exit.
        expect(document.querySelector('.xlsx-live')?.textContent).not.toContain(
          'to leave the grid'
        );
        demo.key('Tab');
        expect(document.activeElement?.closest('.xlsx-scroller')).not.toBeNull();
        demo.key('Tab', { shiftKey: true });
      }
      const after = await demo.save();
      expect(boldIn(after)).toBe(true);
      expect(flagsOf(after)).toBe(flags);
      expect(demo.workbookList.value).toBe(SAMPLE);
    }
  });
});

describe('the question before unsaved edits are lost', () => {
  test('from the list: a no keeps the workbook and fetches nothing', async () => {
    const demo = await onEditSheet();
    demo.bold();
    const fetched = demo.services.fetched.length;
    expect(await demo.askedOnOpen(() => demo.selectWorkbook('basic-values'))).toBe(1);
    expect(demo.dialogs.seen.at(-1)).toContain('saved them to a file');
    expect(demo.services.fetched.slice(fetched).some((u) => u.includes('basic-values'))).toBe(
      false
    );
    expect(demo.shownSheet()).toBe(sheet);
    expect(demo.workbookList.value).toBe(SAMPLE);
  });

  test('from a file: only after the file passes the checks, and a no keeps the edit and the flags', async () => {
    const demo = await onEditSheet();
    demo.bold();
    await demo.setHeadings(true);
    const broken = await demo.askedOnOpen(async () => {
      demo.pickFile('broken.xlsx', 'not a workbook');
      await demo.alertSays('Could not open broken.xlsx');
    });
    expect(broken).toBe(0);
    const wasBold = demo.boldShown();
    expect(
      await demo.askedOnOpen(() => demo.pickFile('basic-values.xlsx', fixtureBytes('basic-values')))
    ).toBe(1);
    expect(demo.shownSheet()).toBe(sheet);
    const saved = await demo.save();
    expect(boldIn(saved)).toBe(wasBold);
    expect(flagsOf(saved)).toBe(flagsOf(sampleBytes()));
  });

  test('over a headings switch: an edit, a typed value, and an edit after a save each ask', async () => {
    const demo = await onEditSheet();
    const toList = () => demo.selectWorkbook('basic-values');
    demo.bold();
    await demo.setHeadings(true);
    expect(await demo.askedOnOpen(toList)).toBe(1);
    await demo.save();
    typeWithNoEnter(demo);
    await demo.setHeadings(false);
    expect(await demo.askedOnOpen(toList)).toBe(1);
    await demo.save();
    demo.bold();
    await demo.save();
    await demo.setHeadings(true);
    demo.bold();
    expect(await demo.askedOnOpen(toList)).toBe(1);
  });

  test('a typed value asks, a formula that does not commit stops the open, and an edit after a cancelled save asks', async () => {
    const demo = await onEditSheet();
    const opens = [
      ['the list', () => demo.selectWorkbook('basic-values')],
      ['a file', () => demo.pickFile('basic-values.xlsx', fixtureBytes('basic-values'))],
    ] as const;
    for (const [from, open] of opens) {
      await demo.save();
      typeWithNoEnter(demo);
      expect([from, await demo.askedOnOpen(open)]).toEqual([from, 1]);
    }
    for (const [from, open] of opens) {
      await demo.save();
      typeWithNoEnter(demo, '=SUM(');
      const fetched = demo.services.fetched.length;
      const asked = await demo.askedOnOpen(async () => {
        open();
        await demo.alertSays('finish or cancel the cell edit');
      });
      expect([from, asked]).toEqual([from, 0]);
      if (from === 'the list')
        expect(demo.services.fetched.slice(fetched).some((u) => u.includes('basic-values'))).toBe(
          false
        );
      expect(demo.shownSheet()).toBe(sheet);
      // The cell editor keeps the formula.
      expect(document.querySelector<HTMLTextAreaElement>('.xlsx-cell-editor')?.value).toBe('=SUM(');
      demo.key('Escape');
    }
    await demo.save();
    demo.bold();
    demo.services.cancelNext = true;
    act(() => void fireEvent.click(demo.headerSave));
    await waitFor(() => expect(demo.services.cancelNext).toBe(false), { timeout: WAIT_MS });
    expect(await demo.askedOnOpen(() => demo.selectWorkbook('basic-values'))).toBe(1);
  });
});

describe('the Escape that closes a message', () => {
  /** A clipboard event on the grid, with a clipboard that keeps what the editor sets. */
  const clip = (demo: Demo, type: 'copy' | 'cut') => {
    const data = new Map<string, string>();
    const clipboardData = {
      setData: (t: string, v: string) => void data.set(t, v),
      getData: (t: string) => data.get(t) ?? '',
      get types() {
        return [...data.keys()];
      },
    };
    const e = new Event(type, { bubbles: true, cancelable: true });
    Object.defineProperty(e, 'clipboardData', { value: clipboardData });
    act(() => void demo.grid.dispatchEvent(e));
  };
  const sides = () =>
    [...document.querySelectorAll<HTMLElement>('.xlsx-overlay .xlsx-pending-edge')]
      .map((el) => el.dataset.side)
      .sort()
      .join(' ');

  test('keeps the outline of a pending copy or cut, and the next Escape ends it', async () => {
    const demo = await mountDemo();
    for (const type of ['copy', 'cut'] as const) {
      demo.goTo('Z510');
      clip(demo, type);
      await settle(50);
      expect([type, sides()]).toEqual([type, 'bottom left right top']);
      await demo.save();
      await demo.statusSays('Saved ');
      act(() => demo.grid.focus());
      demo.key('Escape', {}, demo.grid);
      await settle(50);
      expect(demo.notice()).toBe('');
      expect([type, sides()]).toEqual([type, 'bottom left right top']);
      demo.key('Escape', {}, demo.grid);
      await settle(50);
      expect([type, sides()]).toEqual([type, '']);
    }
  });
});

describe('the newest pick', () => {
  test('a file picked during the slow fetch of a fixture stays open when the fetch ends', async () => {
    const demo = await mountDemo();
    const release = demo.services.hold('basic-values');
    demo.selectWorkbook('basic-values');
    demo.pickFile('picked.xlsx', buildXlsx({ sheets: [{ name: 'Picked', rows: '' }] }));
    await demo.showsSheet('Picked');
    release();
    // The fetch ends here. An older fetch that still counts would replace the sheet.
    await settle(500);
    expect(demo.shownSheet()).toBe('Picked');
    expect(demo.view.getByRole('tab', { name: 'Picked' })).toBeTruthy();
  });
});
