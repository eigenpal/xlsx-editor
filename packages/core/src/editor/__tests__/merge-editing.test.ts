// The merge commands of the editor: `mergeCells` merges the selection and selects the merged cell
// with its top-left cell active; a merge over other values is refused with `merge-drops-values`
// until the caller passes `dropValues`; a merge and center also centers; `unmergeCells` unmerges;
// one undo puts everything back; the saved workbook keeps the result; view mode and a protected
// sheet refuse.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx, c, row, type FixtureSheet } from '../../../../../scripts/fixtures/builder.ts';
import {
  createXlsxEditor,
  mergeCommandState,
  refusalText,
  type MergeCommand,
  type XlsxEditorInstance,
} from '../index.ts';
import { flushEditor } from './flush-editor.ts';
import {
  attr,
  child,
  children,
  readXlsx,
  sheetByName,
  workbookIndex,
  xmlPartOf,
} from '../../store/index.ts';
import type { FontSource } from '../../layout/index.ts';
import { englishText } from '../english.ts';

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

/** A1 holds `top`, C1 holds 5, and A2 is empty. */
async function mount(sheet: Partial<FixtureSheet> = {}) {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.appendChild(host);
  const viewer = createXlsxEditor({ fonts: FONTS });
  viewers.push(viewer);
  viewer.attach(host);
  await viewer.load(
    buildXlsx({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 'top'), c('C1', 5)]), ...sheet }],
    })
  );
  const scroller = host.querySelector('.xlsx-scroller') as HTMLElement;
  Object.defineProperty(scroller, 'clientWidth', { get: () => 800 });
  Object.defineProperty(scroller, 'clientHeight', { get: () => 400 });
  flushEditor(viewer);
  return { viewer, host, scroller };
}

/** The saved worksheet: its merged ranges, and the stored cells as `ref=value`. */
function saved(viewer: XlsxEditorInstance) {
  const r = readXlsx(viewer.save()!);
  if (!r.ok) throw new Error('reopen');
  const root = xmlPartOf(r.pkg, sheetByName(workbookIndex(r.pkg)!, 'S')!.partName)!.root;
  const merges = children(child(root, 'x:mergeCells'), 'x:mergeCell').map((m) => attr(m, 'ref'));
  const cells = children(child(root, 'x:sheetData'), 'x:row').flatMap((r) =>
    children(r, 'x:c').flatMap((cell) => {
      const v = child(cell, 'x:v') ?? child(child(cell, 'x:is'), 'x:t');
      const text = v?.children[0]?.type === 'text' ? v.children[0].value : undefined;
      return text === undefined ? [] : [`${attr(cell, 'r')}=${text}`];
    })
  );
  return { merges, cells };
}

const code = (r: ReturnType<XlsxEditorInstance['mergeCells']>) => (r.ok ? null : r.refusal.code);

describe('mergeCells', () => {
  test('merges the selection and selects the merged cell with its top-left cell active', async () => {
    const { viewer } = await mount();
    viewer.select({ row: 1, col: 1 }, { row: 0, col: 0 });
    expect(viewer.mergeCells().ok).toBe(true);
    flushEditor(viewer);
    expect(saved(viewer).merges).toEqual(['A1:B2']);
    const s = viewer.snapshot();
    expect(s.selection?.active).toEqual({ row: 0, col: 0 });
    expect(s.selectionRef).toBe('A1:B2');
  });

  test('the view stays on the top-left cell, not on the far corner of the range', async () => {
    const { viewer, scroller } = await mount();
    viewer.select({ row: 59, col: 1 }, { row: 0, col: 0 });
    flushEditor(viewer);
    expect(scroller.scrollTop).toBe(0);
    expect(viewer.mergeCells().ok).toBe(true);
    flushEditor(viewer);
    expect(scroller.scrollTop).toBe(0);
  });

  test('a merge over other values is refused until the caller drops them', async () => {
    const { viewer } = await mount();
    viewer.select({ row: 0, col: 0 }, { row: 0, col: 2 });
    const r = viewer.mergeCells();
    expect(code(r)).toBe('merge-drops-values');
    expect(saved(viewer).merges).toEqual([]);
    expect(r.ok ? '' : refusalText(r.refusal, englishText)).toBe(
      'The cells were not merged: cells other than the top-left cell hold values, such as C1. A merge keeps only the value of the top-left cell, so clear those cells first.'
    );
    expect(viewer.mergeCells({ dropValues: true }).ok).toBe(true);
    expect(saved(viewer)).toEqual({ merges: ['A1:C1'], cells: ['A1=top'] });
  });

  test('one undo puts back the merge and the dropped values', async () => {
    const { viewer } = await mount();
    viewer.select({ row: 0, col: 2 }, { row: 0, col: 0 });
    const before = saved(viewer);
    expect(viewer.mergeCells({ dropValues: true, center: true }).ok).toBe(true);
    expect(viewer.undo().ok).toBe(true);
    expect(saved(viewer)).toEqual(before);
    expect(viewer.redo().ok).toBe(true);
    expect(saved(viewer)).toEqual({ merges: ['A1:C1'], cells: ['A1=top'] });
    expect(viewer.snapshot().selection?.active).toEqual({ row: 0, col: 0 });
  });

  test('dropValues clears every other cell of a range of many rows', async () => {
    const { viewer } = await mount({
      rows: row(1, [c('A1', 'top'), c('C1', 5)]) + row(3, [c('A3', 7), c('C3', 8)]),
    });
    viewer.select({ row: 0, col: 0 }, { row: 2, col: 2 });
    expect(viewer.mergeCells({ dropValues: true }).ok).toBe(true);
    expect(saved(viewer)).toEqual({ merges: ['A1:C3'], cells: ['A1=top'] });
  });

  test('an open cell edit commits first, and the merge keeps its text', async () => {
    const { viewer } = await mount({ rows: row(1, [c('A1', 'top')]) });
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 1 });
    expect(viewer.beginEdit('new').ok).toBe(true);
    expect(viewer.mergeCells().ok).toBe(true);
    // The committed text goes to the shared string table: A1 stores its index, 0.
    expect(saved(viewer)).toEqual({ merges: ['A1:B2'], cells: ['A1=0'] });
  });

  test('merge and center centers the merged cell', async () => {
    const { viewer } = await mount();
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 1 });
    expect(viewer.mergeCells({ center: true }).ok).toBe(true);
    expect(viewer.snapshot().activeFormat?.horizontal).toBe('center');
  });

  test('refuses one cell, an overlap, view mode, and a protected sheet', async () => {
    const one = await mount();
    one.viewer.select({ row: 1, col: 1 });
    expect(code(one.viewer.mergeCells())).toBe('merge-one-cell');
    const overlap = await mount({ merges: ['B2:C3'] });
    overlap.viewer.select({ row: 0, col: 0 }, { row: 1, col: 1 });
    expect(code(overlap.viewer.mergeCells())).toBe('merge-overlap');
    const view = await mount();
    view.viewer.setMode('view');
    view.viewer.select({ row: 0, col: 0 }, { row: 1, col: 1 });
    expect(code(view.viewer.mergeCells())).toBe('view-mode');
    const locked = await mount({ beforeMerges: '<sheetProtection sheet="1"/>' });
    locked.viewer.select({ row: 0, col: 0 }, { row: 1, col: 1 });
    const r = locked.viewer.mergeCells();
    expect(r.ok ? null : [r.refusal.code, r.refusal.lock]).toEqual(['locked', 'merge']);
    expect(r.ok ? '' : refusalText(r.refusal, englishText)).toContain('protected');
  });
});

describe('unmergeCells', () => {
  test('unmerges the merged range of the active cell, and one undo merges it again', async () => {
    const { viewer } = await mount({ merges: ['A1:B2'] });
    viewer.select({ row: 0, col: 0 });
    expect(viewer.unmergeCells().ok).toBe(true);
    expect(saved(viewer).merges).toEqual([]);
    expect(viewer.undo().ok).toBe(true);
    expect(saved(viewer).merges).toEqual(['A1:B2']);
  });

  test('a selection with no merged range changes nothing and makes no undo step', async () => {
    const { viewer } = await mount();
    viewer.select({ row: 3, col: 3 });
    expect(viewer.unmergeCells().ok).toBe(true);
    expect(viewer.snapshot().canUndo).toBe(false);
  });
});

describe('mergeCommandState', () => {
  test('names the first cell that a merge would empty, so the host can ask first', async () => {
    const { viewer } = await mount({
      rows: row(1, [c('A1', 'top'), c('C1', 5)]) + row(2, [c('B2', 1)]),
    });
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 2 });
    expect(mergeCommandState(viewer.snapshot(), 'merge').dropsValuesAt).toBe('C1');
    expect(mergeCommandState(viewer.snapshot(), 'unmerge').dropsValuesAt).toBe(null);
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    expect(mergeCommandState(viewer.snapshot(), 'merge').dropsValuesAt).toBe(null);
    viewer.select({ row: 0, col: 0 }, { row: 1_048_575, col: 16_383 });
    expect(mergeCommandState(viewer.snapshot(), 'merge').dropsValuesAt).toBe('C1');
  });

  test('an open cell edit counts as it would be after its commit, as the command reads it', async () => {
    const typed = await mount({ rows: row(1, [c('A1', 'top')]) });
    typed.viewer.select({ row: 0, col: 1 }, { row: 0, col: 0 });
    expect(typed.viewer.beginEdit('typed').ok).toBe(true);
    expect(mergeCommandState(typed.viewer.snapshot(), 'merge').dropsValuesAt).toBe('B1');
    expect(code(typed.viewer.mergeCells())).toBe('merge-drops-values');
    const emptied = await mount();
    emptied.viewer.select({ row: 0, col: 2 }, { row: 0, col: 0 });
    expect(emptied.viewer.beginEdit('').ok).toBe(true);
    expect(mergeCommandState(emptied.viewer.snapshot(), 'merge').dropsValuesAt).toBe(null);
    expect(emptied.viewer.mergeCells().ok).toBe(true);
    const far = await mount({ rows: row(1, [c('A1', 'top')]) });
    far.viewer.select({ row: 1_048_575, col: 16_383 }, { row: 0, col: 0 });
    expect(far.viewer.beginEdit('x').ok).toBe(true);
    expect(mergeCommandState(far.viewer.snapshot(), 'merge').dropsValuesAt).toBe('XFD1048576');
  });

  test('dropsValuesAt follows each key of an open edit, in the grid and in the formula bar', async () => {
    const { viewer, host } = await mount({ rows: row(1, [c('A1', 'top'), c('C1', 5)]) });
    const drops = () => mergeCommandState(viewer.snapshot(), 'merge').dropsValuesAt;
    viewer.select({ row: 0, col: 1 }, { row: 0, col: 0 });
    expect(viewer.beginEdit('').ok).toBe(true);
    expect(drops()).toBe(null);
    const input = host.querySelector('textarea') as HTMLTextAreaElement;
    input.value = 'x';
    input.dispatchEvent(
      new (host.ownerDocument.defaultView as unknown as typeof globalThis).Event('input')
    );
    expect(drops()).toBe('B1');
    expect(viewer.setEditText("'").ok).toBe(true);
    expect(drops()).toBe(null);
    expect(viewer.mergeCells().ok).toBe(true);
    // A second edit starts with no draft: its first key publishes the state again.
    viewer.select({ row: 0, col: 3 }, { row: 0, col: 2 });
    expect(viewer.setEditText('').ok).toBe(true);
    expect(drops()).toBe(null);
    expect(viewer.setEditText('7').ok).toBe(true);
    expect(drops()).toBe('D1');
  });

  test('an edit with the text that the cell holds writes nothing, and counts as the cell', async () => {
    const { viewer } = await mount({ rows: row(1, [c('A1', 'top'), c('B1', 'old')]) });
    viewer.select({ row: 0, col: 1 }, { row: 0, col: 0 });
    expect(viewer.beginEdit().ok).toBe(true);
    expect(mergeCommandState(viewer.snapshot(), 'merge').dropsValuesAt).toBe('B1');
    expect(code(viewer.mergeCells())).toBe('merge-drops-values');
  });

  test('merge and center with dropValues agrees with its state near the format limit', async () => {
    const rows =
      row(1, [c('A1', 'top'), c('B1', 1)]) +
      Array.from({ length: 16_383 }, (_, i) => row(i + 2, [c(`B${i + 2}`, i)])).join('');
    const { viewer } = await mount({ rows });
    viewer.select({ row: 0, col: 0 }, { row: 16_383, col: 1 });
    const s = mergeCommandState(viewer.snapshot(), 'mergeCenter');
    expect([s.enabled, s.dropsValuesAt]).toEqual([true, 'B1']);
    expect(viewer.mergeCells({ center: true, dropValues: true }).ok).toBe(true);
  });

  test('a protected sheet refuses a merge with dropValues by the merge rule', async () => {
    const { viewer } = await mount({ beforeMerges: '<sheetProtection sheet="1"/>' });
    viewer.select({ row: 0, col: 0 }, { row: 0, col: 2 });
    const r = viewer.mergeCells({ dropValues: true });
    expect(r.ok ? null : [r.refusal.code, r.refusal.lock]).toEqual(['locked', 'merge']);
  });

  const state = (viewer: XlsxEditorInstance, command: MergeCommand) => {
    const s = mergeCommandState(viewer.snapshot(), command);
    return s.enabled ? 'enabled' : s.disabledReason?.key;
  };

  test('follows the selection, the merged ranges, view mode, and protection', async () => {
    const { viewer } = await mount({ merges: ['D4:E5'] });
    viewer.select({ row: 0, col: 0 });
    expect(state(viewer, 'merge')).toBe('disabledReason.mergeOneCell');
    expect(state(viewer, 'unmerge')).toBe('disabledReason.noMergedRange');
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 1 });
    expect(state(viewer, 'merge')).toBe('enabled');
    expect(state(viewer, 'mergeCenter')).toBe('enabled');
    viewer.select({ row: 3, col: 3 });
    expect(state(viewer, 'merge')).toBe('disabledReason.mergeOverlap');
    expect(state(viewer, 'unmerge')).toBe('enabled');
    viewer.setMode('view');
    expect(state(viewer, 'unmerge')).toBe('disabledReason.viewMode');
    const locked = await mount({ beforeMerges: '<sheetProtection sheet="1"/>' });
    locked.viewer.select({ row: 0, col: 0 }, { row: 1, col: 1 });
    expect(state(locked.viewer, 'merge')).toBe('disabledReason.mergeLocked');
  });

  test('a one-cell or overlapping merged range of the file counts, as the command reads it', async () => {
    const one = await mount({ merges: ['B2'] });
    one.viewer.select({ row: 0, col: 0 }, { row: 2, col: 2 });
    expect(state(one.viewer, 'merge')).toBe('disabledReason.mergeOverlap');
    expect(code(one.viewer.mergeCells())).toBe('merge-overlap');
    expect(state(one.viewer, 'unmerge')).toBe('enabled');
    expect(one.viewer.unmergeCells().ok).toBe(true);
    expect(state(one.viewer, 'merge')).toBe('enabled');
    const cell = await mount({ merges: ['C3'] });
    cell.viewer.select({ row: 2, col: 2 });
    expect(state(cell.viewer, 'merge')).toBe('disabledReason.mergeOneCell');
    expect(code(cell.viewer.mergeCells())).toBe('merge-one-cell');
    const both = await mount({ merges: ['D4:E5', 'E5:F6'] });
    both.viewer.select({ row: 5, col: 5 }, { row: 6, col: 6 });
    expect(state(both.viewer, 'merge')).toBe('disabledReason.mergeOverlap');
    expect(code(both.viewer.mergeCells())).toBe('merge-overlap');
  });

  test('merge and center follows the format limit of the selection', async () => {
    const { viewer } = await mount();
    viewer.select({ row: 0, col: 0 }, { row: 200_000, col: 50 });
    expect(viewer.snapshot().formatRefusal?.code).toBe('range-too-large');
    expect(state(viewer, 'mergeCenter')).toBe('disabledReason.rangeTooLarge');
    expect(state(viewer, 'merge')).toBe('enabled');
  });

  test('a merge and an unmerge change the state', async () => {
    const { viewer } = await mount();
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 1 });
    expect(viewer.mergeCells().ok).toBe(true);
    expect(state(viewer, 'unmerge')).toBe('enabled');
    expect(state(viewer, 'merge')).toBe('disabledReason.mergeOverlap');
    expect(viewer.unmergeCells().ok).toBe(true);
    expect(state(viewer, 'merge')).toBe('enabled');
  });
});
