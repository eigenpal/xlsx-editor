// The decimal steps of a selection: the limit of the scan, the cells of a wide table,
// and an empty active position among more cells, which the step never writes.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import {
  createXlsxEditor,
  runToolbarCommand,
  toolbarCommandState,
  type NumberLocale,
  type XlsxEditorInstance,
} from '../index.ts';
import { flushEditor } from './flush-editor.ts';
import { decimalStepPreview } from './decimal-step-preview.ts';
import { decimalStepOps, isSingleSelection, selectionDecimalSteps } from '../decimal-selection.ts';
import { xmlPartOf } from '../../store/index.ts';
import { formatCellRef } from '../../store/cell-ref.ts';
import { selectBand } from '../selection.ts';
import { refusalText, type EditRefusal } from '../cell-input.ts';
import { englishText } from '../english.ts';
import {
  DECIMAL_SELECTION_PROBLEM_KEYS,
  DECIMAL_STEP_PROBLEM_KEYS,
  decimalReasonOf,
  type DecimalSelectionProblem,
} from '../decimal-reasons.ts';
import { decimalStepsDetail } from '../decimal-selection.ts';
import type { FontSource } from '../../layout/index.ts';
import { readXlsx } from '../../store/package/package.ts';
import { workbookIndex } from '../../store/workbook/workbook-index.ts';
import { WorkbookDocumentStore } from '../../store/store/document-store.ts';

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
const EN: NumberLocale = { decimal: '.', group: ',' };

let editors: XlsxEditorInstance[] = [];
afterEach(() => {
  for (const e of editors) e.destroy();
  editors = [];
});

/**
 * A sheet with `rows`, the cell formats `cellXfs` after the default, column A in `colStyle`,
 * and the custom number formats `numFmts`.
 */
async function sheetWith(
  rows: string,
  cellXfs: readonly string[],
  colStyle?: number,
  numFmts?: readonly string[]
) {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.appendChild(host);
  const editor = createXlsxEditor({ fonts: FONTS });
  editors.push(editor);
  editor.attach(host);
  const cols =
    colStyle === undefined
      ? {}
      : {
          cols: `<cols><col min="1" max="1" width="10" style="${colStyle}" customWidth="1"/></cols>`,
        };
  const styles = { cellXfs, ...(numFmts ? { numFmts } : {}) };
  await editor.load(buildXlsx({ sheets: [{ name: 'S', rows, ...cols }], styles }));
  flushEditor(editor);
  return editor;
}

const xfOf = (id: number) =>
  `<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`;

describe('General numbers of a range', () => {
  /** The number format codes of E3 to E6 (or `refs`). */
  const codesOf = (editor: XlsxEditorInstance, refs: readonly [number, number][]) =>
    refs.map(([r, col]) => {
      editor.select({ row: r, col });
      return editor.snapshot().activeFormat!.numberFormat.code;
    });
  const e3e6: [number, number][] = [
    [2, 4],
    [3, 4],
    [4, 4],
    [5, 4],
  ];

  test('every General number of a range takes one code, so the range lines up', async () => {
    const editor = await sheetWith(
      row(3, [c('E3', 5)]) +
        row(4, [c('E4', 6.5)]) +
        row(5, [c('E5', 8)]) +
        row(6, [c('E6', 9.25)]),
      []
    );
    const step = () => {
      editor.select({ row: 2, col: 4 }, { row: 5, col: 4 });
      const preview = decimalStepPreview(editor.snapshot(), 1);
      expect(runToolbarCommand(editor, 'number.increaseDecimals').ok).toBe(true);
      return preview;
    };
    // The preview of the active E3 gives the code that the step writes on every cell.
    expect(step()).toEqual({ ok: true, code: '0.0' });
    expect(codesOf(editor, e3e6)).toEqual(['0.0', '0.0', '0.0', '0.0']);
    expect(step()).toEqual({ ok: true, code: '0.00' });
    expect(codesOf(editor, e3e6)).toEqual(['0.00', '0.00', '0.00', '0.00']);
  });

  test('an active date gives way to the first General number, and keeps its format', async () => {
    const editor = await sheetWith(
      row(3, [c('E3', 45000, { s: 1 })]) + row(4, [c('E4', 6.5)]) + row(5, [c('E5', 8)]),
      [xfOf(14)]
    );
    editor.select({ row: 2, col: 4 }, { row: 4, col: 4 });
    expect(runToolbarCommand(editor, 'number.increaseDecimals').ok).toBe(true);
    // E4 shows 6.5, one place: every General number takes two.
    expect(codesOf(editor, e3e6.slice(0, 3))).toEqual(['mm-dd-yy', '0.00', '0.00']);
  });

  test('a decrease from an active General number with no places lets each step alone', async () => {
    // The shared step from the active 5 fails, so each other General number steps from its own.
    const editor = await sheetWith(row(3, [c('E3', 5)]) + row(4, [c('E4', 6.5)]), []);
    editor.select({ row: 2, col: 4 }, { row: 3, col: 4 });
    const state = toolbarCommandState(editor.snapshot(), 'number.decreaseDecimals');
    expect(state.enabled).toBe(true);
    expect(decimalStepPreview(editor.snapshot(), -1)).toEqual({ ok: false, problem: 'noDecimals' });
    expect(runToolbarCommand(editor, 'number.decreaseDecimals').ok).toBe(true);
    expect(codesOf(editor, e3e6.slice(0, 2))).toEqual(['General', '0']);
  });

  test('after a failed shared step, each General number steps from its own places', async () => {
    const editor = await sheetWith(
      row(3, [c('E3', 5)]) + row(4, [c('E4', 6.5)]) + row(5, [c('E5', 7.25)]),
      []
    );
    editor.select({ row: 2, col: 4 }, { row: 4, col: 4 });
    expect(runToolbarCommand(editor, 'number.decreaseDecimals').ok).toBe(true);
    expect(codesOf(editor, e3e6.slice(0, 3))).toEqual(['General', '0', '0.0']);
  });

  test('General integers have no places to remove, so a decrease stays disabled', async () => {
    const editor = await sheetWith(row(3, [c('E3', 5)]) + row(4, [c('E4', 6)]), []);
    editor.select({ row: 2, col: 4 }, { row: 3, col: 4 });
    const state = toolbarCommandState(editor.snapshot(), 'number.decreaseDecimals');
    expect(state.disabledReason?.key).toBe('disabledReason.decimalsNoDecimalsSelection');
  });

  test('an active text cell, then 5, then 6.5: a decrease steps 6.5 alone', async () => {
    const editor = await sheetWith(
      row(2, [c('E2', 'txt')]) + row(3, [c('E3', 5)]) + row(4, [c('E4', 6.5)]),
      []
    );
    editor.select({ row: 1, col: 4 }, { row: 3, col: 4 });
    expect(toolbarCommandState(editor.snapshot(), 'number.decreaseDecimals').enabled).toBe(true);
    expect(runToolbarCommand(editor, 'number.decreaseDecimals').ok).toBe(true);
    expect(codesOf(editor, e3e6.slice(0, 2))).toEqual(['General', '0']);
  });
  test('each column shares the step of its own first General number', async () => {
    // B3 (1234.5678) is active; E3 to E6 hold 5, 6.5, 8, and 9.25.
    const editor = await sheetWith(
      row(3, [c('B3', 1234.5678), c('E3', 5)]) +
        row(4, [c('E4', 6.5)]) +
        row(5, [c('E5', 8)]) +
        row(6, [c('E6', 9.25)]),
      []
    );
    editor.select({ row: 2, col: 1 }, { row: 9, col: 4 });
    // B3 shows 1234.568 in its column: three places.
    expect(decimalStepPreview(editor.snapshot(), 1)).toEqual({ ok: true, code: '0.0000' });
    expect(runToolbarCommand(editor, 'number.increaseDecimals').ok).toBe(true);
    // B3 steps from its own places; column E takes the step of E3.
    expect(codesOf(editor, [[2, 1], ...e3e6])).toEqual(['0.0000', '0.0', '0.0', '0.0', '0.0']);
  });

  test('a decrease can show an integer with a place, from its column', async () => {
    const editor = await sheetWith(row(3, [c('E3', 6.25)]) + row(4, [c('E4', 7)]), []);
    editor.select({ row: 2, col: 4 }, { row: 3, col: 4 });
    expect(runToolbarCommand(editor, 'number.decreaseDecimals').ok).toBe(true);
    // E4 (7) has no place to remove, but follows E3: it shows 7.0.
    expect(codesOf(editor, e3e6.slice(0, 2))).toEqual(['0.0', '0.0']);
  });

  test('a General number in exponent form steps alone', async () => {
    // D5 shows 1.23E+08 in its column; D3 and D4 show fixed numbers.
    const rows =
      row(3, [c('D3', 0.333333333333)]) + row(4, [c('D4', 1.5)]) + row(5, [c('D5', 123456789.123)]);
    const d3d5: [number, number][] = [
      [2, 3],
      [3, 3],
      [4, 3],
    ];
    const alone = await sheetWith(rows, []);
    alone.select({ row: 4, col: 3 });
    const own = decimalStepPreview(alone.snapshot(), 1);
    const code = own?.ok ? own.code : '';
    expect(code.endsWith('E+00')).toBe(true);
    for (const active of [2, 4]) {
      const editor = await sheetWith(rows, []);
      // The active cell first: D3 (fixed) or D5 (exponent form).
      editor.select({ row: active, col: 3 }, { row: active === 2 ? 4 : 2, col: 3 });
      expect(runToolbarCommand(editor, 'number.increaseDecimals').ok).toBe(true);
      const [d3, d4, d5] = codesOf(editor, d3d5);
      expect(d4).toBe(d3!);
      expect(d3).not.toContain('E');
      expect(d5).toBe(code);
    }
  });
});

describe('the reason of a selection with no cell that can step', () => {
  test('the text of a refusal gives the reason of the disabled control', async () => {
    // A1 holds 1.5 in `0` (id 1); A2 and A3 hold text.
    const editor = await sheetWith(
      row(1, [c('A1', 1.5, { s: 1 })]) + row(2, [c('A2', 'a')]) + row(3, [c('A3', 'b')]),
      [xfOf(1)]
    );
    const text = () => {
      const r = runToolbarCommand(editor, 'number.decreaseDecimals');
      return r.ok ? null : refusalText(r.refusal, englishText);
    };
    editor.select({ row: 0, col: 0 });
    expect(text()).toBe(
      'The decimal places were not changed. This number format has no decimal places to remove.'
    );
    editor.select({ row: 1, col: 0 }, { row: 2, col: 0 });
    expect(text()).toBe(
      'The decimal places were not changed. ' +
        'No selected cell holds a number, so there are no decimal places to change.'
    );
  });

  test('a decrease on fractions and a code with no places names the fraction', async () => {
    // B10 and B11 show fractions (ids 12 and 13); B12 has `#,##0,"K"`, with no places.
    const editor = await sheetWith(
      row(10, [c('B10', 1.5, { s: 1 })]) +
        row(11, [c('B11', 2.25, { s: 2 })]) +
        row(12, [c('B12', 12000, { s: 3 })]),
      [xfOf(12), xfOf(13), xfOf(164)],
      undefined,
      ['<numFmt numFmtId="164" formatCode="#,##0,&quot;K&quot;"/>']
    );
    editor.select({ row: 9, col: 1 }, { row: 11, col: 1 });
    const reason = toolbarCommandState(editor.snapshot(), 'number.decreaseDecimals').disabledReason;
    expect(reason?.key).toBe('disabledReason.decimalsFractionSelection');
    const r = runToolbarCommand(editor, 'number.decreaseDecimals');
    expect(r.ok ? null : r.refusal.detail).toBe('fraction');
    // One cell with no place to remove gives the reason of that cell.
    editor.select({ row: 11, col: 1 });
    expect(
      toolbarCommandState(editor.snapshot(), 'number.decreaseDecimals').disabledReason?.key
    ).toBe('disabledReason.decimalsNoDecimals');
  });

  test('an increase and a decrease name the same cell', async () => {
    // B2 shows a date with places in its seconds; B3 shows a fraction.
    const editor = await sheetWith(
      row(2, [c('B2', 1.5, { s: 1 })]) + row(3, [c('B3', 2.25, { s: 2 })]),
      [xfOf(164), xfOf(12)],
      undefined,
      ['<numFmt numFmtId="164" formatCode="[h]:mm:ss.00"/>']
    );
    editor.select({ row: 1, col: 1 }, { row: 2, col: 1 });
    for (const id of ['number.increaseDecimals', 'number.decreaseDecimals'] as const)
      expect(toolbarCommandState(editor.snapshot(), id).disabledReason?.key).toBe(
        'disabledReason.decimalsDateSelection'
      );
  });

  for (const [name, numFmts, key, detail] of [
    ['a custom id with no code', [], 'decimalsMissingCodeSelection', 'missingCode'],
    [
      'a code that the editor cannot read',
      ['<numFmt numFmtId="170" formatCode="0.00[[["/>'],
      'decimalsUnreadableSelection',
      'unreadable',
    ],
  ] as const)
    test(`a decrease on \`0\` and ${name} names the active cell`, async () => {
      // B1 has \`0\` (id 1), with no place to remove; B2 has id 170.
      const editor = await sheetWith(
        row(1, [c('B1', 5, { s: 1 })]) + row(2, [c('B2', 6.5, { s: 2 })]),
        [xfOf(1), xfOf(170)],
        undefined,
        numFmts
      );
      const reason = (active: number) => {
        editor.select({ row: active, col: 1 }, { row: 1 - active, col: 1 });
        const r = runToolbarCommand(editor, 'number.decreaseDecimals');
        return [
          toolbarCommandState(editor.snapshot(), 'number.decreaseDecimals').disabledReason?.key,
          r.ok ? null : r.refusal.detail,
        ];
      };
      // B1 active: a selected cell shows no places, which is not true of B2.
      expect(reason(0)).toEqual([
        'disabledReason.decimalsNoDecimalsCellSelection',
        'cellNoDecimals',
      ]);
      expect(reason(1)).toEqual([`disabledReason.${key}`, detail]);
    });

  test('a text cell with an explicit code keeps its format, alone or among more cells', async () => {
    // B1 holds the text 'note' in `0.00` (id 2); B2 shows a date (id 14).
    const editor = await sheetWith(
      row(1, [c('B1', 'note', { s: 1 })]) + row(2, [c('B2', 45000, { s: 2 })]),
      [xfOf(2), xfOf(14)]
    );
    editor.select({ row: 0, col: 1 }, { row: 1, col: 1 });
    const snapshot = editor.snapshot();
    expect(decimalStepsDetail(snapshot.decimalSteps).activeWritten).toBe(false);
    expect(decimalStepPreview(snapshot, 1)).toEqual({ ok: false, problem: 'noNumber' });
    const reason = toolbarCommandState(snapshot, 'number.increaseDecimals').disabledReason;
    expect(reason?.key).toBe('disabledReason.decimalsDateSelection');
    const r = runToolbarCommand(editor, 'number.increaseDecimals');
    expect(r.ok ? null : r.refusal.detail).toBe('date');
    // B1 alone acts as a range of text cells: it holds no number, so its code does not step.
    editor.select({ row: 0, col: 1 });
    const alone = editor.snapshot();
    expect(decimalStepsDetail(alone.decimalSteps).activeWritten).toBe(false);
    expect(decimalStepPreview(alone, 1)).toEqual({ ok: false, problem: 'noNumber' });
    const one = runToolbarCommand(editor, 'number.increaseDecimals');
    expect(one.ok ? null : one.refusal.detail).toBe('noNumbers');
    expect(editor.snapshot().activeFormat!.numberFormat.code).toBe('0.00');
  });

  test('a General section is not a format with no places to remove', async () => {
    // B1 has `0.00;General`, with places in its first section; B2 shows a date (id 14).
    const generalSection = (b2: number) =>
      sheetWith(
        row(1, [c('B1', 1.5, { s: 1 })]) + row(2, [c('B2', 45000, { s: 2 })]),
        [xfOf(164), xfOf(b2)],
        undefined,
        ['<numFmt numFmtId="164" formatCode="0.00;General"/>']
      );
    const reason = async (b2: number, active: number) => {
      const editor = await generalSection(b2);
      editor.select({ row: active, col: 1 }, { row: 1 - active, col: 1 });
      const r = runToolbarCommand(editor, 'number.decreaseDecimals');
      return [
        toolbarCommandState(editor.snapshot(), 'number.decreaseDecimals').disabledReason?.key,
        r.ok ? null : r.refusal.detail,
      ];
    };
    expect(await reason(14, 0)).toEqual([
      'disabledReason.decimalsGeneralSectionSelection',
      'generalSection',
    ]);
    expect(await reason(14, 1)).toEqual(['disabledReason.decimalsDateSelection', 'date']);
    // B2 in \`0\` (id 1), with no place to remove, names itself.
    expect(await reason(1, 1)).toEqual([
      'disabledReason.decimalsNoDecimalsCellSelection',
      'cellNoDecimals',
    ]);
  });

  test('an increase with different problems names the active cell', async () => {
    // B1 has the most places that the editor writes; B2 shows a date (id 14).
    const max = `0.${'0'.repeat(30)}`;
    const editor = await sheetWith(
      row(1, [c('B1', 1.5, { s: 1 })]) + row(2, [c('B2', 45000, { s: 2 })]),
      [xfOf(164), xfOf(14)],
      undefined,
      [`<numFmt numFmtId="164" formatCode="${max}"/>`]
    );
    editor.select({ row: 0, col: 1 }, { row: 1, col: 1 });
    // Not "no selected number format can take another place": that is not why B2 cannot step.
    const reason = toolbarCommandState(editor.snapshot(), 'number.increaseDecimals').disabledReason;
    expect(reason).toEqual({
      key: 'disabledReason.decimalsMaxDecimalsCellSelection',
      params: { limit: '30' },
    });
    const r = runToolbarCommand(editor, 'number.increaseDecimals');
    expect(r.ok ? null : refusalText(r.refusal, englishText)).toBe(
      'The decimal places were not changed. ' +
        'A selected cell already has the most decimal places that the editor writes: 30.'
    );
    // The reason comes from the typed field: `detail` is for logs and never becomes user text.
    const generic =
      'The decimal places were not changed: ' +
      'no selected cell has a number format whose decimal places can change.';
    expect(refusalText({ code: 'decimals-unchanged', detail: 'date' }, englishText)).toBe(generic);
    // A reason with a key that no decimal step gives shows the generic text too.
    const other = { key: 'disabledReason.viewMode' } as const;
    expect(refusalText({ code: 'decimals-unchanged', reason: other }, englishText)).toBe(generic);
    // B2 active names the date.
    editor.select({ row: 1, col: 1 }, { row: 0, col: 1 });
    expect(
      toolbarCommandState(editor.snapshot(), 'number.increaseDecimals').disabledReason?.key
    ).toBe('disabledReason.decimalsDateSelection');
    // Cells that agree keep the reason of the first.
    editor.select({ row: 0, col: 1 });
    expect(
      toolbarCommandState(editor.snapshot(), 'number.increaseDecimals').disabledReason?.key
    ).toBe('disabledReason.decimalsMaxDecimals');
  });

  test('a decrease on one General integer speaks of the cell', async () => {
    const editor = await sheetWith(row(1, [c('A1', 7)]), []);
    editor.select({ row: 0, col: 0 });
    const reason = toolbarCommandState(editor.snapshot(), 'number.decreaseDecimals').disabledReason;
    expect(reason?.key).toBe('disabledReason.decimalsNoShownPlaces');
    expect(decimalStepPreview(editor.snapshot(), -1)).toEqual({ ok: false, problem: 'noPlaces' });
    const r = runToolbarCommand(editor, 'number.decreaseDecimals');
    expect(r.ok ? null : r.refusal.detail).toBe('noPlaces');
  });

  test('an empty active position names no reason: its column format is never read', async () => {
    // Column A shows a date (built-in 14); A1 has no `x:c`, and A2 holds General text.
    const editor = await sheetWith(row(2, [c('A2', 'text')]), [xfOf(14)], 1);
    editor.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    const reason = toolbarCommandState(editor.snapshot(), 'number.increaseDecimals').disabledReason;
    expect(reason?.key).toBe('disabledReason.decimalsNoNumberSelection');
    const r = runToolbarCommand(editor, 'number.increaseDecimals');
    expect(r.ok ? null : [r.refusal.code, r.refusal.detail]).toEqual([
      'decimals-unchanged',
      'noNumbers',
    ]);
  });

  test('cells with no number name no format reason: text, a boolean, and an error', async () => {
    // B1 is text in the Text format `@` (id 49), B2 is TRUE, and B3 is #DIV/0!.
    const editor = await sheetWith(
      row(1, [c('B1', '00123', { s: 1 })]) +
        row(2, ['<c r="B2" t="b"><v>1</v></c>']) +
        row(3, ['<c r="B3" t="e"><v>#DIV/0!</v></c>']),
      [xfOf(49)]
    );
    editor.select({ row: 0, col: 1 }, { row: 2, col: 1 });
    for (const id of ['number.increaseDecimals', 'number.decreaseDecimals'] as const)
      expect(toolbarCommandState(editor.snapshot(), id).disabledReason?.key).toBe(
        'disabledReason.decimalsNoNumberSelection'
      );
    // One such cell alone disables both steps too, whatever its format, in words about that cell.
    for (const at of [0, 1, 2]) {
      editor.select({ row: at, col: 1 });
      for (const id of ['number.increaseDecimals', 'number.decreaseDecimals'] as const)
        expect(toolbarCommandState(editor.snapshot(), id).disabledReason?.key).toBe(
          'disabledReason.decimalsNoNumber'
        );
    }
    const r = runToolbarCommand(editor, 'number.increaseDecimals');
    expect(r.ok ? null : refusalText(r.refusal, englishText)).toBe(
      'The decimal places were not changed. ' +
        'This cell does not hold a number, so it has no decimal places to change.'
    );
  });

  test('a date cell (t="d") steps from the serial that the grid shows', async () => {
    const date = '<c r="B1" t="d" s="1"><v>2023-03-15T00:00:00</v></c>';
    const number = await sheetWith(row(1, [date]), [xfOf(2)]);
    number.select({ row: 0, col: 1 });
    expect(runToolbarCommand(number, 'number.increaseDecimals').ok).toBe(true);
    expect(number.snapshot().activeFormat!.numberFormat.code).toBe('0.000');
    const shortDate = await sheetWith(row(1, [date]), [xfOf(14)]);
    shortDate.select({ row: 0, col: 1 });
    expect(
      toolbarCommandState(shortDate.snapshot(), 'number.increaseDecimals').disabledReason?.key
    ).toBe('disabledReason.decimalsDate');
  });

  test('one General text cell does not step, and one empty General cell still takes 0.0', async () => {
    // B1 holds the text 'note' in General; B2 has no `x:c` element.
    const editor = await sheetWith(row(1, [c('B1', 'note')]), []);
    editor.select({ row: 0, col: 1 });
    const r = runToolbarCommand(editor, 'number.increaseDecimals');
    expect(r.ok ? null : r.refusal.detail).toBe('noNumbers');
    expect(editor.snapshot().activeFormat!.numberFormat.code).toBe('General');
    editor.select({ row: 1, col: 1 });
    expect(runToolbarCommand(editor, 'number.increaseDecimals').ok).toBe(true);
    expect(editor.snapshot().activeFormat!.numberFormat.code).toBe('0.0');
  });
});

describe('an empty active position', () => {
  test('an empty active position among more cells is not written, and the preview says so', async () => {
    // Column A has the cell format `[Red]0.00` (index 2); A1 has no `x:c` element.
    const steps = async (a2: string, a1 = '') => {
      const win = new Window();
      const doc = win.document as unknown as Document;
      const host = doc.createElement('div');
      doc.body.appendChild(host);
      const editor = createXlsxEditor({ fonts: FONTS });
      editors.push(editor);
      editor.attach(host);
      const xf = (id: number) =>
        `<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`;
      await editor.load(
        buildXlsx({
          sheets: [
            {
              name: 'S',
              rows: row(1, a1 ? [a1] : []) + row(2, [a2]),
              cols: '<cols><col min="1" max="1" width="10" style="1" customWidth="1"/></cols>',
            },
          ],
          styles: {
            numFmts: ['<numFmt numFmtId="164" formatCode="[Red]0.00"/>'],
            cellXfs: [xf(164)],
          },
        })
      );
      flushEditor(editor);
      editor.select({ row: 0, col: 0 }, { row: 1, col: 0 });
      return editor;
    };
    const codeAt = (editor: XlsxEditorInstance, r: number) => {
      editor.select({ row: r, col: 0 });
      return editor.snapshot().activeFormat!.numberFormat.code;
    };
    const empty = await steps(c('A2', 5));
    expect(decimalStepsDetail(empty.snapshot().decimalSteps).activeWritten).toBe(false);
    expect(decimalStepPreview(empty.snapshot(), 1)).toEqual({ ok: false, problem: 'noNumber' });
    expect(runToolbarCommand(empty, 'number.increaseDecimals').ok).toBe(true);
    // A1 keeps the column format, and A2 (General, 5) steps to one place.
    expect([codeAt(empty, 0), codeAt(empty, 1)]).toEqual(['[Red]0.00', '0.0']);
    // A2 holds text with General: no selected cell holds a number.
    const text = await steps(c('A2', 'text'));
    const reason = toolbarCommandState(text.snapshot(), 'number.increaseDecimals').disabledReason;
    expect(reason?.key).toBe('disabledReason.decimalsNoNumberSelection');
    expect(decimalStepPreview(text.snapshot(), 1)?.ok).toBe(false);
    // A1 exists with the cell format but holds no number: among more cells, it is not written.
    const bare = await steps(c('A2', 5), '<c r="A1" s="1"/>');
    expect(decimalStepsDetail(bare.snapshot().decimalSteps).activeWritten).toBe(false);
    expect(runToolbarCommand(bare, 'number.increaseDecimals').ok).toBe(true);
    expect([codeAt(bare, 0), codeAt(bare, 1)]).toEqual(['[Red]0.00', '0.0']);
    // A1 holds a number in the cell format: the preview and the command write it.
    const written = await steps(c('A2', 5), '<c r="A1" s="1"><v>1.5</v></c>');
    expect(decimalStepsDetail(written.snapshot().decimalSteps).activeWritten).toBe(true);
    expect(decimalStepPreview(written.snapshot(), 1)).toEqual({ ok: true, code: '[Red]0.000' });
    expect(runToolbarCommand(written, 'number.increaseDecimals').ok).toBe(true);
    expect([codeAt(written, 0), codeAt(written, 1)]).toEqual(['[Red]0.000', '0.0']);
  });
});

describe('decimal steps of a selection', () => {
  /** A package with `rows` rows of `cols` numbers and the merged ranges `merges`, and its sheet. */
  function table(rows: number, cols: number, merges?: readonly string[]) {
    const xml = Array.from({ length: rows }, (_, r) =>
      row(
        r + 1,
        Array.from(
          { length: cols },
          (_, k) => `<c r="${formatCellRef({ row: r, col: k })}"><v>1</v></c>`
        )
      )
    ).join('');
    const sheet = { name: 'S', rows: xml, ...(merges ? { merges } : {}) };
    const read = readXlsx(buildXlsx({ sheets: [sheet] }));
    if (!read.ok) throw new Error('the workbook does not open');
    const index = workbookIndex(read.pkg)!;
    return { pkg: read.pkg, index, entry: index.sheets[0]! };
  }
  const box = (top: number, left: number, bottom: number, right: number) => ({
    sheetIndex: 0,
    active: { row: top, col: left },
    anchor: { row: top, col: left },
    range: { start: { row: top, col: left }, end: { row: bottom, col: right } },
  });

  test('a selection that holds more cells than the limit refuses the step', () => {
    const { pkg, index, entry } = table(3, 6_000);
    // 3 rows of 5,500 selected cells: 16,500 cells, past the limit.
    const selection = box(0, 0, 2, 5_499);
    const steps = selectionDecimalSteps(pkg, index, entry.partName, selection, EN);
    const refusal: EditRefusal = {
      code: 'decimals-too-large',
      detail: 'tooLarge',
      reason: { key: 'disabledReason.decimalsTooLarge', params: { limit: '16,384' } },
    };
    expect(steps).toEqual({ increase: refusal, decrease: refusal });
    expect(isSingleSelection(xmlPartOf(pkg, entry.partName)?.root, selection)).toBe(false);
    const r = decimalStepOps(pkg, index, entry, selection, EN, 1);
    // The command gives the refusal of the state.
    expect(r.ok ? null : r.refusal).toEqual(refusal);
    // The limit counts the selected cells, not the rows: 3 rows of 5,000 cells step.
    expect(
      selectionDecimalSteps(pkg, index, entry.partName, box(0, 0, 2, 4_999), EN).increase
    ).toBeNull();
  });

  test('one merged range over more cells than the limit steps its anchor', () => {
    // 16,500 cells under one merged range: the step reads only the anchor A1.
    const last = formatCellRef({ row: 2, col: 5_499 });
    const { pkg, index, entry } = table(3, 5_500, [`A1:${last}`]);
    const selection = box(0, 0, 2, 5_499);
    const steps = selectionDecimalSteps(pkg, index, entry.partName, selection, EN);
    expect([
      isSingleSelection(xmlPartOf(pkg, entry.partName)?.root, selection),
      steps.increase,
    ]).toEqual([true, null]);
    const r = decimalStepOps(pkg, index, entry, selection, EN, 1);
    expect(r.ok && r.ops.length).toBe(1);
  });

  test('an active cell under a merged range of a column band is not read or written', () => {
    // A1:B2 is merged; B1 (1.5) and B3 (2.5) have `0.00`. A band of column B makes B1 active.
    const xml =
      row(1, ['<c r="B1" s="1"><v>1.5</v></c>']) + row(3, ['<c r="B3" s="1"><v>2.5</v></c>']);
    const read = readXlsx(
      buildXlsx({
        sheets: [{ name: 'S', rows: xml, merges: ['A1:B2'] }],
        styles: { cellXfs: [xfOf(2)] },
      })
    );
    if (!read.ok) throw new Error('the workbook does not open');
    const index = workbookIndex(read.pkg)!;
    const entry = index.sheets[0]!;
    const selection = selectBand(0, 'columns', 1, 1);
    expect(selection.active).toEqual({ row: 0, col: 1 });
    const steps = selectionDecimalSteps(read.pkg, index, entry.partName, selection, EN);
    expect([steps.increase, decimalStepsDetail(steps).activeWritten]).toEqual([null, false]);
    const r = decimalStepOps(read.pkg, index, entry, selection, EN, 1);
    // One op, on B3 only.
    expect(r.ok && r.ops.length).toBe(1);
    expect(JSON.stringify(r.ok ? r.ops : [])).toContain('B3');
    expect(JSON.stringify(r.ok ? r.ops : [])).not.toContain('B1');
  });

  test('a code with a control character keeps its format, and the other cells step', () => {
    // A1 (1.5) has `0.00&#9;`, a tab that a saved code cannot hold; B1 (2.5) has `0.00`.
    const xml = row(1, ['<c r="A1" s="1"><v>1.5</v></c>', '<c r="B1" s="2"><v>2.5</v></c>']);
    const read = readXlsx(
      buildXlsx({
        sheets: [{ name: 'S', rows: xml }],
        styles: {
          numFmts: ['<numFmt numFmtId="164" formatCode="0.00&#9;"/>'],
          cellXfs: [xfOf(164), xfOf(2)],
        },
      })
    );
    if (!read.ok) throw new Error('the workbook does not open');
    const index = workbookIndex(read.pkg)!;
    const entry = index.sheets[0]!;
    const selection = box(0, 0, 0, 1);
    expect(selectionDecimalSteps(read.pkg, index, entry.partName, selection, EN).increase).toBe(
      null
    );
    const r = decimalStepOps(read.pkg, index, entry, selection, EN, 1);
    if (!r.ok) throw new Error('the step is refused');
    expect(JSON.stringify(r.ops)).not.toContain('A1');
    expect(new WorkbookDocumentStore(read.pkg).transact(r.ops).ok).toBe(true);
    // The cell alone is refused, with the reason of a code that the editor cannot read.
    expect(
      selectionDecimalSteps(read.pkg, index, entry.partName, box(0, 0, 0, 0), EN).increase?.detail
    ).toBe('unreadable');
  });

  test('a column of a wide table steps, and cells beside it do not count', () => {
    const { pkg, index, entry } = table(1_000, 50);
    const selection = box(0, 2, 399, 2);
    expect(selectionDecimalSteps(pkg, index, entry.partName, selection, EN).increase).toBeNull();
    const r = decimalStepOps(pkg, index, entry, selection, EN, 1);
    expect(r.ok && r.ops.length).toBe(1);
    // A whole column too.
    expect(
      selectionDecimalSteps(pkg, index, entry.partName, box(0, 2, 1_048_575, 2), EN).increase
    ).toBeNull();
  });
});

describe('the parameters of the decimal reasons', () => {
  test('every reason whose text has a limit gets one from decimalReasonOf', () => {
    const en = JSON.parse(
      readFileSync(join(import.meta.dir, '../../../../i18n/en.json'), 'utf8')
    ) as { disabledReason: Record<string, string> };
    for (const [single, keys] of [
      [true, DECIMAL_STEP_PROBLEM_KEYS],
      [false, DECIMAL_SELECTION_PROBLEM_KEYS],
    ] as const)
      for (const [problem, key] of Object.entries(keys)) {
        const text = en.disabledReason[key.slice('disabledReason.'.length)]!;
        const r = decimalReasonOf(problem as DecimalSelectionProblem, single);
        expect([key, r.params?.limit !== undefined]).toEqual([key, text.includes('{limit}')]);
      }
    // The limit is grouped in the number locale.
    expect(decimalReasonOf('tooLong', true)).toEqual({
      key: 'disabledReason.decimalsTooLong',
      params: { limit: '255' },
    });
    const de: NumberLocale = { decimal: ',', group: '.' };
    expect(decimalReasonOf('tooLarge', false, de).params).toEqual({
      limit: '16.384',
    });
  });
});
