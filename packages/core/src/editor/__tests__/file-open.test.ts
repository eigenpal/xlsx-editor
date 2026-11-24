// The packaged File > Open of the menu bar (`openPickedWorkbook`): a picked file that does not
// open keeps the open workbook, its state, and its open cell edit, and shows `open-failed`; one
// read of the file; a later pick or load wins; and Open is disabled while a load runs.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { openPickedWorkbook } from '../file-open.ts';
import { loadReadCount } from '../load-steps.ts';
import { toolbarCommandState } from '../toolbar-commands.ts';
import { createXlsxEditor, type XlsxEditorInstance } from '../index.ts';
import type { FontSource } from '../../layout/index.ts';
import { flushEditor } from './flush-editor.ts';

const fontsDir = join(
  dirname(Bun.resolveSync('@docx-editor.dev/fonts/package.json', import.meta.dir)),
  'assets'
);
const FONTS: FontSource[] = [
  {
    family: 'Carlito',
    bold: false,
    italic: false,
    bytes: new Uint8Array(readFileSync(join(fontsDir, 'Carlito-Regular.ttf'))),
    origin: 'substitute',
  },
];

let viewers: XlsxEditorInstance[] = [];
afterEach(() => {
  for (const v of viewers) v.destroy();
  viewers = [];
});

async function open(bytes?: Uint8Array) {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.appendChild(host);
  const viewer = createXlsxEditor({ fonts: FONTS });
  viewers.push(viewer);
  viewer.attach(host);
  await viewer.load(
    bytes ?? buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 2)]) }] })
  );
  flushEditor(viewer);
  return viewer;
}

describe('packaged File > Open', () => {
  test('a failed pick on a dirty workbook fires no change, and keeps its revision', async () => {
    const viewer = await open();
    expect(viewer.beginEdit('dirty').ok).toBe(true);
    expect(viewer.commitEdit().ok).toBe(true);
    const revision = viewer.snapshot().revision;
    expect(revision).toBeGreaterThan(0);
    const changes: unknown[] = [];
    const failures: unknown[] = [];
    viewer.on('change', (change) => void changes.push(change));
    viewer.on('openFailure', (failure) => void failures.push(failure));
    await openPickedWorkbook(viewer, new TextEncoder().encode('not a workbook'));
    expect(changes).toEqual([]);
    expect(failures).toEqual([]);
    expect(viewer.snapshot().revision).toBe(revision);
    expect(viewer.snapshot().canUndo).toBe(true);
  });

  test('a failed pick on a failed workbook does not fire openFailure again', async () => {
    const hidden = buildXlsx({
      sheets: [{ name: 'H', state: 'hidden', rows: row(1, [c('A1', 1)]) }],
    });
    const viewer = await open(hidden);
    const failures: unknown[] = [];
    const changes: unknown[] = [];
    viewer.on('openFailure', (failure) => void failures.push(failure));
    viewer.on('change', (change) => void changes.push(change));
    await openPickedWorkbook(viewer, new TextEncoder().encode('not a workbook'));
    expect(viewer.snapshot().status).toBe('failed');
    expect(failures).toEqual([]);
    expect(changes).toEqual([]);
  });

  test('a pick whose file read fails after a host load leaves the host load', async () => {
    const viewer = await open();
    let fail: (e: Error) => void = () => {};
    const slow = {
      arrayBuffer: () => new Promise<ArrayBuffer>((_, reject) => (fail = reject)),
    } as unknown as Blob;
    const pick = openPickedWorkbook(viewer, slow);
    await viewer.load(buildXlsx({ sheets: [{ name: 'Host', rows: row(1, [c('A1', 1)]) }] }));
    fail(new Error('gone'));
    await pick;
    expect(viewer.snapshot().status).toBe('ready');
    expect(viewer.snapshot().sheets.map((s) => s.name)).toEqual(['Host']);
    expect(viewer.snapshot().refusal).toBeNull();
  });

  test('a picked file that the editor cannot open keeps the open workbook and its edits', async () => {
    const viewer = await open();
    expect(viewer.beginEdit('kept').ok).toBe(true);
    expect(viewer.commitEdit().ok).toBe(true);
    viewer.select({ row: 0, col: 0 });
    const rejections: string[] = [];
    viewer.on('rejection', (refusal) => void rejections.push(refusal.code));
    const statuses: string[] = [];
    viewer.subscribe(() => void statuses.push(viewer.snapshot().status));
    const reads = loadReadCount();
    const after = await openPickedWorkbook(viewer, new TextEncoder().encode('not a workbook'));
    expect(loadReadCount() - reads).toBe(1);
    // The read shows the loading status, then the open workbook is ready again.
    expect(statuses).toContain('loading');
    expect(after.status).toBe('ready');
    expect(viewer.snapshot().refusal?.code).toBe('open-failed');
    expect(rejections).toEqual(['open-failed']);
    expect(viewer.snapshot().activeCell?.content).toBe('kept');
    expect(viewer.snapshot().canUndo).toBe(true);
  });

  test('a failed pick keeps the open cell edit with its typed text', async () => {
    const viewer = await open();
    expect(viewer.beginEdit('typed').ok).toBe(true);
    await openPickedWorkbook(viewer, new TextEncoder().encode('not a workbook'));
    expect(viewer.snapshot().editing).toBe(true);
    expect(viewer.snapshot().editText).toBe('typed');
    expect(viewer.snapshot().refusal?.code).toBe('open-failed');
    // A pick that opens drops the edit, as a load does.
    await openPickedWorkbook(
      viewer,
      buildXlsx({ sheets: [{ name: 'N', rows: row(1, [c('A1', 1)]) }] })
    );
    expect(viewer.snapshot().editing).toBe(false);
  });

  test('a failed pick keeps a failed workbook failed, with its failure', async () => {
    const hidden = buildXlsx({
      sheets: [{ name: 'H', state: 'hidden', rows: row(1, [c('A1', 1)]) }],
    });
    const viewer = await open(hidden);
    expect(viewer.snapshot().failure?.code).toBe('no-worksheet');
    await openPickedWorkbook(viewer, new TextEncoder().encode('not a workbook'));
    expect(viewer.snapshot().status).toBe('failed');
    expect(viewer.snapshot().failure?.code).toBe('no-worksheet');
    expect(viewer.snapshot().refusal?.code).toBe('open-failed');
  });

  test('in an empty editor, a pick gives the snapshot that load gives', async () => {
    const hidden = buildXlsx({
      sheets: [{ name: 'H', state: 'hidden', rows: row(1, [c('A1', 1)]) }],
    });
    const empty = () => {
      const viewer = createXlsxEditor({ fonts: FONTS });
      viewers.push(viewer);
      return viewer;
    };
    for (const bytes of [hidden, new TextEncoder().encode('not a workbook')]) {
      const [a, b] = [empty(), empty()];
      const loaded = await a.load(bytes);
      const picked = await openPickedWorkbook(b, bytes);
      expect(picked).toEqual(loaded);
    }
  });

  test('a file that the browser cannot read shows open-failed and keeps the workbook', async () => {
    const viewer = await open();
    const unreadable = { arrayBuffer: () => Promise.reject(new Error('gone')) } as unknown as Blob;
    await openPickedWorkbook(viewer, unreadable);
    expect(viewer.snapshot().status).toBe('ready');
    expect(viewer.snapshot().refusal?.code).toBe('open-failed');
  });

  test('a host load during the file read of a pick wins', async () => {
    const viewer = await open();
    let give: (b: ArrayBuffer) => void = () => {};
    const slow = {
      arrayBuffer: () => new Promise<ArrayBuffer>((resolve) => (give = resolve)),
    } as unknown as Blob;
    const pick = openPickedWorkbook(viewer, slow);
    expect(viewer.snapshot().status).toBe('loading');
    await viewer.load(buildXlsx({ sheets: [{ name: 'Host', rows: row(1, [c('A1', 1)]) }] }));
    const late = buildXlsx({ sheets: [{ name: 'Late', rows: row(1, [c('A1', 1)]) }] });
    give(late.slice().buffer);
    await pick;
    expect(viewer.snapshot().sheets.map((s) => s.name)).toEqual(['Host']);
  });

  test('a bad pick during the read of another pick keeps the workbook and edit from before both', async () => {
    const viewer = await open();
    expect(viewer.beginEdit('typed').ok).toBe(true);
    const bad = () => openPickedWorkbook(viewer, new TextEncoder().encode('not a workbook'));
    const first = bad();
    // While the first pick loads, Open is disabled, so the menu cannot start a second pick.
    expect(
      toolbarCommandState(viewer.snapshot(), 'file.open', { open: () => {} }).disabledReason?.key
    ).toBe('disabledReason.loading');
    await bad();
    await first;
    expect(viewer.snapshot().status).toBe('ready');
    expect(viewer.snapshot().sheets.map((s) => s.name)).toEqual(['S']);
    expect(viewer.snapshot().editText).toBe('typed');
    expect(viewer.snapshot().refusal?.code).toBe('open-failed');
    // A pick that opens, then a bad pick: the opened workbook stands.
    await openPickedWorkbook(
      viewer,
      buildXlsx({ sheets: [{ name: 'N', rows: row(1, [c('A1', 1)]) }] })
    );
    await bad();
    expect(viewer.snapshot().sheets.map((s) => s.name)).toEqual(['N']);
    expect(viewer.snapshot().status).toBe('ready');
  });

  test('a detach during the read of a pick reports the dropped edit', async () => {
    const viewer = await open();
    expect(viewer.beginEdit('typed').ok).toBe(true);
    let give: (b: ArrayBuffer) => void = () => {};
    const slow = {
      arrayBuffer: () => new Promise<ArrayBuffer>((resolve) => (give = resolve)),
    } as unknown as Blob;
    const pick = openPickedWorkbook(viewer, slow);
    viewer.detach();
    give(new TextEncoder().encode('not a workbook').slice().buffer);
    await pick;
    expect(viewer.snapshot().editDropped).toBe(true);
    expect(viewer.snapshot().refusal?.code).toBe('open-failed');
  });

  test('in an empty editor, a file that the browser cannot read fails with the code read', async () => {
    const viewer = createXlsxEditor({ fonts: FONTS });
    viewers.push(viewer);
    const unreadable = { arrayBuffer: () => Promise.reject(new Error('gone')) } as unknown as Blob;
    const after = await openPickedWorkbook(viewer, unreadable);
    expect(after.status).toBe('failed');
    expect(after.failure?.code).toBe('read');
  });

  test('with an editor that the engine did not make, an unreadable file does not reject', async () => {
    const viewer = await open();
    const foreign = {
      load: (b: Uint8Array) => viewer.load(b),
      snapshot: () => viewer.snapshot(),
    } as unknown as XlsxEditorInstance;
    const unreadable = { arrayBuffer: () => Promise.reject(new Error('gone')) } as unknown as Blob;
    expect((await openPickedWorkbook(foreign, unreadable)).status).toBe('ready');
  });

  test('a file that the editor can open replaces the workbook with one read', async () => {
    const viewer = await open();
    const reads = loadReadCount();
    const next = buildXlsx({ sheets: [{ name: 'Next', rows: row(1, [c('A1', 9)]) }] });
    expect((await openPickedWorkbook(viewer, next)).status).toBe('ready');
    expect(loadReadCount() - reads).toBe(1);
    expect(viewer.snapshot().sheets.map((s) => s.name)).toEqual(['Next']);
  });

  test('of two quick picks, the last wins, and a later load wins over a pick', async () => {
    const viewer = await open();
    const pick = (name: string) =>
      openPickedWorkbook(viewer, buildXlsx({ sheets: [{ name, rows: row(1, [c('A1', 1)]) }] }));
    const first = pick('First');
    await pick('Second');
    await first;
    expect(viewer.snapshot().sheets.map((s) => s.name)).toEqual(['Second']);
    // A broken pick that a host load replaces shows no refusal.
    const broken = openPickedWorkbook(viewer, new TextEncoder().encode('not a workbook'));
    await viewer.load(buildXlsx({ sheets: [{ name: 'Host', rows: row(1, [c('A1', 1)]) }] }));
    await broken;
    expect(viewer.snapshot().sheets.map((s) => s.name)).toEqual(['Host']);
    expect(viewer.snapshot().refusal).toBeNull();
  });
});
