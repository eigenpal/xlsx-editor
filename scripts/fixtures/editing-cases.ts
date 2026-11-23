// The fixtures of the change `workbook-structure-editing` (design D5): sheet ops, a clipboard
// paste, and a resize. Each one's edited copy goes through the save check. Every cached value
// is the true result of its formula before the edit.
//
// A clipboard edit lowers through the editor's own clipboard code (`WorkbookClipboard`), which
// is DOM-free, so the edited copy holds exactly the ops that a paste in the editor makes.

import {
  parseCellRef,
  workbookIndex,
  sheetByName,
  WorkbookDocumentStore,
  type WorkbookOp,
  type XlsxPackage,
} from '../../packages/core/src/store/index.ts';
import { EN_US } from '../../packages/core/src/cascade/index.ts';
import { WorkbookClipboard, type PasteData } from '../../packages/core/src/editor/clipboard.ts';
import type { CellEditing } from '../../packages/core/src/editor/cell-editing.ts';
import { ConditionalCut } from '../../packages/core/src/editor/conditional-cut.ts';
import type { EditRefusal } from '../../packages/core/src/editor/cell-input.ts';
import type { GridPoint } from '../../packages/core/src/contracts/index.ts';
import { decimalStepOps } from '../../packages/core/src/editor/decimal-selection.ts';
import { c, row, type FixtureWorkbook } from './builder.ts';

/** The sheet-ops fixture: two hidden sheets that a visible summary refers to. */
export function sheetOpsWorkbook(): FixtureWorkbook {
  return {
    appProps: true,
    sheets: [
      {
        name: 'Summary',
        cols: '<cols><col min="1" max="1" width="18" customWidth="1"/></cols>',
        rows: [
          row(1, [c('A1', 'Summary'), c('B1', 'Value')]),
          row(2, [c('A2', 'North'), c('B2', 10, { f: 'Data!B2' })]),
          row(3, [c('A3', 'South'), c('B3', 20, { f: 'Data!B3' })]),
          row(4, [c('A4', 'Total (name)'), c('B4', 30, { f: 'Total' })]),
          row(5, [c('A5', 'Scratch x2'), c('B5', 10, { f: 'Scratch!A1*2' })]),
          row(6, [c('A6', 'Half'), c('B6', 15, { f: 'SUM(Data!B2:B3)/2' })]),
        ].join(''),
      },
      {
        name: 'Data',
        state: 'hidden',
        rows: [
          row(1, [c('A1', 'Region'), c('B1', 'Amount')]),
          row(2, [c('A2', 'North'), c('B2', 10)]),
          row(3, [c('A3', 'South'), c('B3', 20)]),
          row(4, [c('A4', 'Rate x10'), c('B4', 100, { f: 'Rate*10' })]),
          row(5, [c('A5', 'First amount')]),
          row(6, [c('A6', 'Scratch value')]),
        ].join(''),
        afterMerges:
          '<hyperlinks><hyperlink ref="A5" location="Data!B2"/>' +
          '<hyperlink ref="A6" location="Scratch!A1"/></hyperlinks>',
      },
      { name: 'Scratch', state: 'hidden', rows: row(1, [c('A1', 5)]) },
    ],
    definedNames:
      '<definedName name="Rate" localSheetId="1">Data!$B$2</definedName>' +
      '<definedName name="Temp" localSheetId="2">Scratch!$A$1</definedName>' +
      '<definedName name="Total">SUM(Data!$B$2:$B$3)</definedName>',
  };
}

/**
 * The sheet-add fixture: a workbook whose `docProps/app.xml` lists two sheets, so the edited
 * copy, with a third sheet, saves sheet titles that no longer match the sheets.
 */
export function sheetAddWorkbook(): FixtureWorkbook {
  return {
    appProps: true,
    sheets: [
      {
        name: 'Report',
        rows: [
          row(1, [c('A1', 'Report'), c('B1', 'Value')]),
          row(2, [c('A2', 'Total'), c('B2', 12, { f: 'Data!A1*2' })]),
        ].join(''),
      },
      // Hidden, so the reference PDF prints one page.
      { name: 'Data', state: 'hidden', rows: row(1, [c('A1', 6)]) },
    ],
  };
}

/** The clipboard fixture: a small table with a header format, a number format, and formulas. */
export function clipboardWorkbook(): FixtureWorkbook {
  const head = (ref: string, text: string) => c(ref, text, { s: 1 });
  const line = (r: number, item: string, qty: number, price: number) =>
    row(r, [
      c(`A${r}`, item),
      c(`B${r}`, qty),
      c(`C${r}`, price, { s: 2 }),
      c(`D${r}`, qty * price, { f: `B${r}*C${r}`, s: 2 }),
    ]);
  return {
    sheets: [
      {
        name: 'Clip',
        rows: [
          row(1, [head('A1', 'Item'), head('B1', 'Qty'), head('C1', 'Price'), head('D1', 'Total')]),
          line(2, 'Pens', 4, 1.5),
          line(3, 'Paper', 2, 4.25),
          line(4, 'Ink', 3, 12),
        ].join(''),
      },
    ],
    styles: {
      fonts: [
        '<font><b/><sz val="11"/><color theme="1"/><name val="Carlito"/><family val="2"/><scheme val="minor"/></font>',
      ],
      fills: [
        '<fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/><bgColor indexed="64"/></patternFill></fill>',
      ],
      cellXfs: [
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>',
        '<xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
      ],
    },
  };
}

/** The resize fixture: a column entry that the edit splits, and rows with and without `x:row`. */
export function resizeWorkbook(): FixtureWorkbook {
  return {
    sheets: [
      {
        name: 'Sizes',
        cols: '<cols><col min="1" max="4" width="12" customWidth="1"/></cols>',
        rows: [
          row(1, [
            c('A1', 'Width 12'),
            c('B1', 'Width 20'),
            c('C1', 'Width 20'),
            c('D1', 'Width 12'),
          ]),
          row(2, [c('A2', 'Height 30')]),
          row(3, [c('A3', 'Height 30')], 'ht="24" customHeight="1"'),
          row(4, [c('A4', 'Default'), c('F4', 'F')]),
        ].join(''),
      },
    ],
  };
}

/** A copy of `from` (or a cut) pasted at the cell `to`, as the editor's clipboard lowers it. */
export interface ClipboardFixtureEdit {
  readonly sheet: string;
  readonly copy: string;
  readonly paste: string;
  readonly cut?: boolean;
}

/** Plain text pasted at the cell `at`, as text from another program. */
export interface PlainPasteFixtureEdit {
  readonly sheet: string;
  readonly pasteText: string;
  readonly at: string;
}

/**
 * The ops of one clipboard edit, read from `pkg`: the ops that `WorkbookClipboard` runs for a
 * copy (or a cut) of a range and a paste at a cell, or a plain-text paste.
 */
export function clipboardOps(
  pkg: XlsxPackage,
  edit: ClipboardFixtureEdit | PlainPasteFixtureEdit
): WorkbookOp[] {
  const index = workbookIndex(pkg);
  const entry = index && sheetByName(index, edit.sheet);
  if (!index || !entry) throw new Error(`clipboard edit: no sheet ${edit.sheet}`);
  const point = (ref: string): GridPoint => {
    const p = parseCellRef(ref);
    if (!p) throw new Error(`clipboard edit: bad reference ${ref}`);
    return p;
  };
  const store = new WorkbookDocumentStore(pkg);
  let range = { start: point('A1'), end: point('A1') };
  let ran: readonly WorkbookOp[] | null = null;
  let refusal: EditRefusal | null = null;
  const edits = {
    commitOpenEdit: () => ({ ok: true }),
    succeeded: () => undefined,
    refuse: (r: EditRefusal) => {
      refusal = r;
      return { ok: false, refusal: r };
    },
    run: (ops: readonly WorkbookOp[]) => {
      ran = ops;
      return { ok: true };
    },
  } as unknown as CellEditing;
  const clipboard = new WorkbookClipboard({
    edits,
    repaint: () => {},
    locale: () => EN_US,
    merges: () => [],
    select: () => undefined,
    announce: () => undefined,
    t: (key) => key,
    conditionalCut: new ConditionalCut(() => undefined),
    target: () => ({
      store,
      pkg,
      index,
      sheet: index.sheets.indexOf(entry),
      entry,
      selection: { sheetIndex: 0, active: range.start, anchor: range.start, range },
      focus: range.start,
      revision: 0,
    }),
  });
  let data: PasteData;
  if ('copy' in edit) {
    const [a, b] = edit.copy.split(':') as [string, string?];
    range = { start: point(a), end: point(b ?? a) };
    const copied = edit.cut ? clipboard.cut() : clipboard.copy();
    if (!copied.ok) throw new Error(`clipboard edit: the copy is refused (${copied.refusal.code})`);
    data = { text: copied.content.text, token: copied.content.token };
    range = { start: point(edit.paste), end: point(edit.paste) };
  } else {
    data = { text: edit.pasteText };
    range = { start: point(edit.at), end: point(edit.at) };
  }
  clipboard.paste(data);
  if (refusal)
    throw new Error(`clipboard edit: the paste is refused (${(refusal as EditRefusal).code})`);
  if (!ran) throw new Error('clipboard edit: the paste made no ops');
  return [...(ran as readonly WorkbookOp[])];
}

/**
 * A decimal step of a selection, as the toolbar runs it: `range` is the selection and `active`
 * its active cell. Each cell steps from its own number format (`decimalStepOps`).
 */
export interface DecimalStepFixtureEdit {
  readonly sheet: string;
  readonly decimalStep: 1 | -1;
  readonly range: string;
  readonly active: string;
}

/** The store ops of a decimal step, read from `pkg` after the earlier edits. */
export function decimalStepFixtureOps(
  pkg: XlsxPackage,
  edit: DecimalStepFixtureEdit
): WorkbookOp[] {
  const index = workbookIndex(pkg);
  const entry = index && sheetByName(index, edit.sheet);
  if (!index || !entry) throw new Error(`decimal step: no sheet ${edit.sheet}`);
  const [a, b = a] = edit.range.split(':');
  const start = parseCellRef(a!);
  const end = parseCellRef(b!);
  const active = parseCellRef(edit.active);
  if (!start || !end || !active) throw new Error(`decimal step: bad range ${edit.range}`);
  const selection = { sheetIndex: entry.index, active, anchor: start, range: { start, end } };
  const r = decimalStepOps(pkg, index, entry, selection, EN_US, edit.decimalStep);
  if (!r.ok) throw new Error(`decimal step: refused with ${r.refusal.code}`);
  return [...r.ops];
}
