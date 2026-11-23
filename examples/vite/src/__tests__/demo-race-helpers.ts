// The helpers of the race tests of the demo (`demo-races.test.tsx`, `demo-save-races.test.tsx`):
// a workbook whose load takes many steps, and stand-in file dialogs that count or hold.

import { act, fireEvent, waitFor } from '@testing-library/react';
import { expect } from 'bun:test';
import { readZip } from '../../../../packages/ooxml-core/src/zip.ts';
import { buildXlsx, c, row, type FixtureSheet } from '../../../../scripts/fixtures/builder.ts';
import { WAIT_MS, type Demo } from './demo-harness.tsx';

export const FIXTURE = 'basic-values';
/** The list pick that comes after {@link FIXTURE}. */
export const NEWER_FIXTURE = 'structure';
export const EDIT = 'edit-kept-over-the-reload';

/** The rows of {@link largeWorkbook}, built once. */
let largeRows: string | undefined;

/** A workbook of one sheet `name` whose load takes many steps, each of which yields. */
export function largeWorkbook(name: string, extra: Partial<FixtureSheet> = {}): Uint8Array {
  if (largeRows === undefined) {
    const rows: string[] = [];
    for (let r = 1; r <= 12_000; r++) {
      const cells: string[] = [];
      for (let x = 0; x < 20; x++) cells.push(c(`${String.fromCharCode(65 + x)}${r}`, r * x));
      rows.push(row(r, cells));
    }
    largeRows = rows.join('');
  }
  return buildXlsx({ sheets: [{ name, rows: largeRows, ...extra }] });
}

export const smallWorkbook = (name: string): Uint8Array =>
  buildXlsx({ sheets: [{ name, rows: row(1, [c('A1', 'small')]) }] });

/** Whether any part of a saved copy holds `text`. */
export function holdsText(bytes: Uint8Array, text: string): boolean {
  const zip = readZip(bytes);
  if (!zip.ok) throw new Error(`the saved copy does not unzip (${zip.reason})`);
  const decoder = new TextDecoder();
  return [...zip.entries.values()].some((part) => decoder.decode(part).includes(text));
}

/** Fails when the load already ended, because the test then proves nothing. */
export function stillLoading(demo: Demo, what: string): void {
  if (!demo.loading()) throw new Error(`the load ended before ${what}, so the test proves nothing`);
}

/** Opens `bytes` from a picked file, and waits for the sheet `name`. */
export async function openFile(demo: Demo, name: string, bytes: Uint8Array): Promise<void> {
  demo.pickFile(`${name.toLowerCase()}.xlsx`, bytes);
  await demo.showsSheet(name);
}

/** The time in ms that the demo takes to open the large workbook of the sheet `sheet`. */
export async function timedOpen(demo: Demo, sheet: string): Promise<number> {
  const file = largeWorkbook(sheet);
  const start = performance.now();
  await openFile(demo, sheet, file);
  return performance.now() - start;
}

/** Clicks the headings box with no wait for the end of the switch. */
export const pressHeadings = (demo: Demo) => act(() => void fireEvent.click(demo.headingsBox));

/** Presses the Save button of the top bar. */
export const pressSave = (demo: Demo) => act(() => void fireEvent.click(demo.headerSave));

/** Presses the Save button of the toolbar, also while it is disabled. */
export const pressToolbarSave = (demo: Demo) => act(() => void fireEvent.click(demo.toolbarSave));

/** Replaces the file save picker for the rest of a test. */
export function setPicker(picker: unknown): void {
  (window as unknown as { showSaveFilePicker: unknown }).showSaveFilePicker = picker;
}

/**
 * A stand-in file dialog that counts its calls, and with `refuseFirst` refuses the first call as
 * a browser refuses a dialog long after its click.
 */
export function countingPicker(refuseFirst: boolean): { calls: number } {
  const counts = { calls: 0 };
  setPicker(async () => {
    counts.calls += 1;
    if (refuseFirst && counts.calls === 1)
      throw new DOMException('The picker needs a recent user action.', 'SecurityError');
    return {
      name: 'kept.xlsx',
      createWritable: async () => ({ write: async () => {}, close: async () => {} }),
    };
  });
  return counts;
}

/** Records whether the status of a save shows, until the returned call. */
export function watchSaving(demo: Demo): () => boolean {
  let seen = false;
  const timer = setInterval(() => {
    if (demo.busy().includes('Saving')) seen = true;
  }, 20);
  return () => {
    clearInterval(timer);
    return seen;
  };
}

/** Waits until no save runs. */
export const savesEnd = (demo: Demo) =>
  waitFor(() => expect(demo.busy()).not.toContain('Saving'), { timeout: WAIT_MS });

/** Closes the message of the demo, so the next message is a new one. */
export function closeAlert(demo: Demo): void {
  demo.closeMessage();
  expect(demo.alerts()).toEqual([]);
}
