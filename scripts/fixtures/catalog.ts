// The synthetic fixtures: one workbook per feature, committed under `fixtures/`, each with the
// edits that the save check applies to its edited copy (design D5a). CI's gates run on these,
// never on a private corpus. Each task that adds a rendering feature adds its fixtures here.
//
// A formula's cached value must be its true result: the reference PDFs print recalculated
// formula results (measured on basic-values: A5 prints 84), and the engine shows the cached
// value (design D3).

import {
  WorkbookDocumentStore,
  type FormatPatch,
  type StructuralOpName,
  type WorkbookOp,
  type XlsxPackage,
} from '../../packages/core/src/store/index.ts';
import { formatOps } from '../../packages/core/src/automation/index.ts';
import {
  NUMBER_FORMAT_CHOICES,
  builtinOrCode,
  type NumberFormatChoiceValue,
} from '../../packages/core/src/editor/number-format-commands.ts';
import { valuePatch } from '../../packages/core/src/editor/toolbar-commands.ts';
import { c, row, type FixtureWorkbook } from './builder.ts';
import {
  formulaOverflowWorkbook,
  gridStructureWorkbook,
  textLayoutWorkbook,
  wrappedFormulaWorkbook,
  wrappedOnlyFormulaWorkbook,
} from './layout-cases.ts';
import { structureWorkbook } from './structure-case.ts';
import { conditionalRulesWorkbook } from './conditional-cases.ts';
import {
  clipboardOps,
  clipboardWorkbook,
  resizeWorkbook,
  sheetAddWorkbook,
  sheetOpsWorkbook,
  decimalStepFixtureOps,
  type ClipboardFixtureEdit,
  type DecimalStepFixtureEdit,
  type PlainPasteFixtureEdit,
} from './editing-cases.ts';
import {
  STYLE_RESOLUTION_CASES,
  cellStylesWorkbook,
  fillsPatternsWorkbook,
} from './fill-style-cases.ts';
import {
  FORMAT_CASES,
  FORMAT_CASES_1904,
  FORMAT_CASES_MORE,
  formatSheet,
  styleSheet,
} from './cases.ts';

/**
 * One edit of a fixture's edited copy: set one cell to a value or a formula, apply a format
 * patch to a range, as the format commands lower it (`formatOps`), insert or delete rows or
 * columns (`start` is the 1-based first row or column), add, rename, move, or delete a sheet
 * (`sheet` names the new sheet of an add), paste a copy or plain text as the editor's clipboard
 * does, or resize columns or rows (`first` and `last` are 1-based).
 */
export type FixtureEdit =
  | { readonly sheet: string; readonly ref: string; readonly value: number | string | boolean }
  /** Formula text without the leading `=`. */
  | { readonly sheet: string; readonly ref: string; readonly formula: string }
  | { readonly sheet: string; readonly ref: string; readonly format: FormatPatch }
  | {
      readonly sheet: string;
      readonly structure: StructuralOpName;
      readonly start: number;
      readonly count: number;
    }
  | { readonly sheet: string; readonly sheetOp: 'add'; readonly index: number }
  | { readonly sheet: string; readonly sheetOp: 'rename'; readonly name: string }
  | { readonly sheet: string; readonly sheetOp: 'move'; readonly index: number }
  | { readonly sheet: string; readonly sheetOp: 'delete' }
  | ClipboardFixtureEdit
  | PlainPasteFixtureEdit
  | DecimalStepFixtureEdit
  | {
      readonly sheet: string;
      readonly resize: 'columns' | 'rows';
      readonly first: number;
      readonly last: number;
      /** Characters for columns, points for rows. */
      readonly size: number;
    };

export interface Fixture {
  readonly name: string;
  /** What the fixture shows, in one sentence. */
  readonly purpose: string;
  readonly workbook: () => FixtureWorkbook;
  /** The edits of the edited copy, applied as one transaction through the store ops. */
  readonly edits: readonly FixtureEdit[];
  /**
   * The file extension of the fixture and of its copies. Absent: `xlsx`. Set `xlsm` for a
   * workbook whose main part has the macro-enabled content type (`FixtureWorkbook.macroEnabled`).
   * A desktop application expects the extension `.xlsm` for this content type.
   * `scripts/fixtures/build.ts` refuses an extension that does not agree with the workbook.
   */
  readonly extension?: 'xlsm';
}

/** The file extensions of the fixtures. */
export type FixtureExtension = 'xlsx' | 'xlsm';

/**
 * The file extension of the fixture `name`. A name that the catalog does not have gets `xlsx`, so
 * a message about a missing fixture can name a file.
 */
export const fixtureExtension = (name: string): FixtureExtension =>
  FIXTURES.find((f) => f.name === name)?.extension ?? 'xlsx';

/**
 * The file name of the fixture `name`, or of one of its copies: `<name><copy>.<extension>`. For
 * example: `basic-values.xlsx`, `macro-enabled.xlsm`, and `macro-enabled.saved.xlsm` (the copy
 * `.saved`). Each script and test that names a fixture file or a copy uses this name.
 */
export const fixtureFile = (name: string, copy = ''): string =>
  `${name}${copy}.${fixtureExtension(name)}`;

/**
 * The store ops of a fixture's edits. A clipboard edit reads its source from `pkg` after the
 * earlier edits, as the editor reads the workbook at the paste: the earlier ops are applied to
 * a store first, and a refusal of them throws.
 */
export function fixtureEditOps(fixture: Fixture, pkg: XlsxPackage): WorkbookOp[] {
  const out: WorkbookOp[] = [];
  for (const e of fixture.edits) {
    if ('copy' in e || 'pasteText' in e || 'decimalStep' in e) {
      let current = pkg;
      if (out.length > 0) {
        const store = new WorkbookDocumentStore(pkg);
        const r = store.transact(out);
        if (!r.ok) throw new Error(`${fixture.name}: an edit before a clipboard edit is refused`);
        current = store.package;
      }
      out.push(
        ...('decimalStep' in e ? decimalStepFixtureOps(current, e) : clipboardOps(current, e))
      );
    } else out.push(...editOps(fixture, e));
  }
  return out;
}

/** The store ops of one edit that is not a clipboard edit. */
function editOps(
  fixture: Fixture,
  e: Exclude<FixtureEdit, ClipboardFixtureEdit | PlainPasteFixtureEdit | DecimalStepFixtureEdit>
): WorkbookOp[] {
  if ('sheetOp' in e) {
    switch (e.sheetOp) {
      case 'add':
        return [{ op: 'addSheet', name: e.sheet, index: e.index }];
      case 'rename':
        return [{ op: 'renameSheet', sheet: e.sheet, name: e.name }];
      case 'move':
        return [{ op: 'moveSheet', sheet: e.sheet, index: e.index }];
      case 'delete':
        return [{ op: 'deleteSheet', sheet: e.sheet }];
    }
  }
  if ('resize' in e)
    return [
      e.resize === 'columns'
        ? { op: 'setColumnWidth', sheet: e.sheet, first: e.first, last: e.last, width: e.size }
        : { op: 'setRowHeight', sheet: e.sheet, first: e.first, last: e.last, height: e.size },
    ];
  if ('format' in e) {
    const lowered = formatOps(e.sheet, e.ref, e.format);
    if (!lowered.ok)
      throw new Error(`${fixture.name}: the format edit is refused (${lowered.rejection.detail})`);
    return [...lowered.ops];
  }
  if ('structure' in e)
    return [{ op: e.structure, sheet: e.sheet, start: e.start, count: e.count }];
  return [valueOp(e)];
}

function valueOp(
  e: Extract<FixtureEdit, { readonly value: unknown } | { readonly formula: string }>
): WorkbookOp {
  return {
    op: 'setCellValue',
    sheet: e.sheet,
    ref: e.ref,
    value:
      'formula' in e
        ? { type: 'formula', text: e.formula }
        : typeof e.value === 'number'
          ? { type: 'number', value: e.value }
          : typeof e.value === 'boolean'
            ? { type: 'boolean', value: e.value }
            : { type: 'string', value: e.value },
  };
}

/**
 * The number format that the toolbar list writes for one entry in the default number locale, by
 * the toolbar path: `valuePatch` gives the entry's code, and `builtinOrCode` writes the
 * built-in id in its place. `formatSheet` defines its custom codes from id 164 up, so the
 * fixture workbook redefines no built-in id, and the map of defined ids is empty.
 */
function choiceFormat(value: NumberFormatChoiceValue): NonNullable<FormatPatch['numberFormat']> {
  const patch = valuePatch('number.format', value)?.numberFormat;
  if (!patch || patch.code === undefined) throw new Error(`The list entry ${value} has no code.`);
  const id = NUMBER_FORMAT_CHOICES.find((c) => c.value === value)?.id ?? null;
  return id === null ? { code: patch.code } : builtinOrCode(id, patch.code, new Map());
}

const enc = new TextEncoder();

export const FIXTURES: readonly Fixture[] = [
  {
    name: 'basic-values',
    purpose:
      'Each stored cell type: shared and inline strings, numbers, booleans, errors, and formulas with cached values.',
    workbook: () => ({
      strings: ['Shared text', 'Label'],
      sheets: [
        {
          name: 'Values',
          cols: '<cols><col min="1" max="1" width="16" customWidth="1"/></cols>',
          rows: [
            row(1, [c('A1', 0, { t: 's' }), c('B1', 'Inline text')]),
            row(2, [c('A2', 42), c('B2', -3.5), c('C2', 1234567.891)]),
            row(3, [c('A3', true), c('B3', false)]),
            row(4, [c('A4', '#DIV/0!', { t: 'e' }), c('B4', '#N/A', { t: 'e' })]),
            row(5, [
              c('A5', 84, { f: 'A2*2' }),
              c('B5', 'Shared text', { t: 'str', f: 'INDEX(A1:B1,1,1)' }),
            ]),
            row(6, [c('A6', 1, { t: 's' })]),
          ].join(''),
        },
      ],
    }),
    // A value edit and a formula edit: the save sets the stale flag of design D3.
    edits: [
      { sheet: 'Values', ref: 'A2', value: 7 },
      { sheet: 'Values', ref: 'A5', formula: 'A2*3' },
    ],
  },
  {
    name: 'formulas',
    purpose:
      'Formulas with cached values, a shared formula, and a calculation chain: the edit overwrites the shared formula master, which unshares its children (design D4) and drops the chain (design D5).',
    workbook: () => ({
      sheets: [
        {
          name: 'Formulas',
          rows: [
            row(1, [c('A1', 2), c('B1', 4, { f: 'A1*2' })]),
            row(2, [c('A2', 3), '<c r="B2"><f t="shared" ref="B2:B5" si="0">A2*2</f><v>6</v></c>']),
            row(3, [c('A3', 5), '<c r="B3"><f t="shared" si="0"/><v>10</v></c>']),
            row(4, [c('A4', 7), '<c r="B4"><f t="shared" si="0"/><v>14</v></c>']),
            row(5, [c('A5', 11), '<c r="B5"><f t="shared" si="0"/><v>22</v></c>']),
            row(6, [c('B6', 56, { f: 'SUM(B1:B5)' })]),
            row(7, [c('B7', 'n=2', { t: 'str', f: '"n="&A1' })]),
          ].join(''),
        },
      ],
      extraParts: [
        [
          'xl/calcChain.xml',
          'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml',
          enc.encode(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
              '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
              ['B1', 'B2', 'B3', 'B4', 'B5', 'B6', 'B7']
                .map((r) => `<c r="${r}" i="1"/>`)
                .join('') +
              '</calcChain>'
          ),
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain',
        ],
      ],
    }),
    // Overwriting the shared formula master writes each child as its own formula first.
    edits: [{ sheet: 'Formulas', ref: 'B2', value: 100 }],
  },
  {
    name: 'conditional-rules',
    purpose:
      'Conditional formatting `cellIs` rules with literal operands: each operator kind, text and boolean values, two true rules, `stopIfTrue`, an error value, and the fill, font, and number format of a differential format.',
    workbook: conditionalRulesWorkbook,
    edits: [{ sheet: 'Rules', ref: 'C1', value: 'edited' }],
  },
  {
    name: 'preserved-parts',
    purpose:
      'Parts and elements the engine does not model survive a save: conditional formatting, extension data, a binary part of an unknown relationship type, and hyperlinks with an external and a script target.',
    workbook: () => ({
      sheets: [
        {
          name: 'Kept',
          rows: row(1, [c('A1', 'Kept content'), c('B1', 5)]),
          afterMerges:
            '<conditionalFormatting sqref="B1"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>3</formula></cfRule></conditionalFormatting>' +
            '<hyperlinks><hyperlink ref="A1" r:id="rIdLink1"/><hyperlink ref="B1" r:id="rIdLink2"/></hyperlinks>',
          rels: [
            [
              'rIdLink1',
              'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
              'https://example.com/report',
              true,
            ],
            [
              'rIdLink2',
              'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
              'javascript:alert(1)',
              true,
            ],
          ],
          ext: '<ext uri="{6F3E9C2A-0B1D-4E5F-8A7B-1C2D3E4F5A6B}" xmlns:fx="urn:example:xlsx-editor-fixture"><fx:kept value="1"/></ext>',
        },
      ],
      styles: { dxfs: ['<dxf><font><b/></font></dxf>'] },
      extraParts: [
        [
          'xl/fixtureData.bin',
          'application/octet-stream',
          enc.encode('inert binary bytes'),
          'http://example.com/relationships/fixture-data',
        ],
      ],
    }),
    edits: [{ sheet: 'Kept', ref: 'B1', value: 6 }],
  },
  {
    name: 'number-formats',
    purpose:
      'Built-in and custom number formats, sections, conditions, colors, text sections, and dates and times in the 1900 date system, one case per cell (scripts/fixtures/cases.ts).',
    workbook: () => {
      const sheet = formatSheet(FORMAT_CASES);
      return {
        sheets: [{ name: 'Formats', rows: sheet.rows, cols: sheet.cols }],
        styles: { numFmts: sheet.numFmts, cellXfs: sheet.cellXfs },
      };
    },
    edits: [{ sheet: 'Formats', ref: 'A1', value: 'edited' }],
  },
  {
    name: 'number-formats-more',
    purpose:
      'More number format cases: hours below 10, positive values of the locale-dependent built-ins, serials that are no date, and a numbered color.',
    workbook: () => {
      const sheet = formatSheet(FORMAT_CASES_MORE);
      return {
        sheets: [{ name: 'Formats', rows: sheet.rows, cols: sheet.cols }],
        styles: { numFmts: sheet.numFmts, cellXfs: sheet.cellXfs },
      };
    },
    // The number format commands of the toolbar: the Scientific entry (built-in id 11), the
    // Accounting, Currency, and Long date entries (the codes that the list writes), and one decimal
    // increase of B6:B11 with the date at B8 active: B8 to B10 show a date or a time and keep their
    // formats, and B6, B7, and B11 each step from their own codes.
    edits: [
      { sheet: 'Formats', ref: 'A1', value: 'edited' },
      { sheet: 'Formats', ref: 'B4', format: { numberFormat: choiceFormat('scientific') } },
      { sheet: 'Formats', ref: 'B5', format: { numberFormat: choiceFormat('accounting') } },
      { sheet: 'Formats', ref: 'B7', format: { numberFormat: choiceFormat('currency') } },
      { sheet: 'Formats', ref: 'B8', format: { numberFormat: choiceFormat('longDate') } },
      { sheet: 'Formats', decimalStep: 1, range: 'B6:B11', active: 'B8' },
    ],
  },
  {
    name: 'dates-1904',
    purpose: 'Date serials and an ISO 8601 date cell in the 1904 date system.',
    workbook: () => {
      const sheet = formatSheet(FORMAT_CASES_1904);
      return {
        date1904: true,
        sheets: [{ name: 'Dates', rows: sheet.rows, cols: sheet.cols }],
        styles: { numFmts: sheet.numFmts, cellXfs: sheet.cellXfs },
      };
    },
    edits: [{ sheet: 'Dates', ref: 'A1', value: 'edited' }],
  },
  {
    name: 'colors-fonts',
    purpose:
      'Theme, tinted, indexed, rgb, and automatic colors on fills and fonts, theme fonts, a font without applyFont, and a thin border, one case per cell.',
    workbook: () => {
      const sheet = styleSheet();
      return {
        sheets: [
          {
            name: 'Colors',
            rows: sheet.rows,
            cols: '<cols><col min="1" max="1" width="34" customWidth="1"/><col min="2" max="2" width="14" customWidth="1"/></cols>',
          },
        ],
        theme: { major: 'Liberation Sans' },
        styles: {
          fonts: sheet.fonts,
          fills: sheet.fills,
          borders: sheet.borders,
          cellXfs: sheet.cellXfs,
        },
      };
    },
    // A value edit and a format edit over styled and empty cells: the save writes new font,
    // fill, border, number format, and cell format records, and a new `x:numFmts`.
    edits: [
      { sheet: 'Colors', ref: 'A1', value: 'edited' },
      {
        sheet: 'Colors',
        ref: 'B2:C3',
        format: {
          bold: true,
          italic: true,
          underline: true,
          fontName: 'Courier New',
          fontSize: 12,
          fontColor: '1F4E79',
          fill: { type: 'solid', color: 'FFF2CC' },
          borders: { edges: 'outline', style: 'thin', color: '2F5597' },
          horizontal: 'center',
          vertical: 'top',
          wrapText: true,
          numberFormat: { code: '0.00;[Red]-0.00' },
        },
      },
    ],
  },
  {
    name: 'text-layout',
    purpose:
      'Cell text layout: alignment, overflow into empty neighbors, wrapping, indent, rotation, stacked text, shrinking, fitted numbers, rich text, and rows that take the height of their fonts (scripts/fixtures/layout-cases.ts).',
    workbook: textLayoutWorkbook,
    edits: [{ sheet: 'Text', ref: 'C1', value: 8 }],
  },
  {
    name: 'formula-overflow',
    purpose:
      'Text overflow that stops at a formula cell: one with a cached number (B1 caches 2), and one with a cached empty string that shows nothing (B3, whose empty `<v></v>` is an empty string value). Both stops come from cached values, so the rule for a formula with no cached value rests on the unit tests alone (scripts/fixtures/layout-cases.ts).',
    workbook: formulaOverflowWorkbook,
    edits: [{ sheet: 'Overflow', ref: 'D1', value: 9 }],
  },
  {
    name: 'wrapped-formula',
    purpose:
      'A formula cell whose formula and cached value are only in an `mc:AlternateContent` child: the grid shows the value of the `mc:Fallback` (B1 caches 4), because its `mc:Choice` requires a namespace that no consumer understands. A direct formula (B2) is the control (scripts/fixtures/layout-cases.ts).',
    workbook: wrappedFormulaWorkbook,
    edits: [{ sheet: 'Wrapped', ref: 'A3', value: 'edited' }],
  },
  {
    name: 'wrapped-only-formula',
    purpose:
      'A workbook whose only formula is in an `mc:AlternateContent` child (B1 caches 4 in its `mc:Fallback`). No cell has a direct formula, so the saved copy of an edit has `fullCalcOnLoad` only when the formula scan finds the wrapped formula (scripts/fixtures/layout-cases.ts).',
    workbook: wrappedOnlyFormulaWorkbook,
    edits: [{ sheet: 'Wrapped', ref: 'A2', value: 'edited' }],
  },
  {
    name: 'grid-structure',
    purpose:
      'Row heights, a hidden row and column, merged ranges with fills and borders, and a picture and preset shapes on two-cell, one-cell, and absolute anchors (scripts/fixtures/layout-cases.ts).',
    workbook: gridStructureWorkbook,
    edits: [{ sheet: 'Grid', ref: 'A1', value: 'edited' }],
  },
  {
    name: 'structure',
    purpose:
      'Rows and columns that move every reference (design D4): the edit deletes row 4 and inserts column C of Data. It covers cells and row properties that move; a reference to the deleted row that becomes #REF! on its own sheet, on another sheet (Refs, hidden so that the reference PDF is one page), and in a defined name; ranges that shrink (SUM(B2:B5)) and grow (SUM(Data!B2:C5) across the inserted column); absolute references; a shared formula that stays shared with a moved range; a same-sheet shared formula (Data column H) whose child that refers to the deleted row is unshared, while a later child stays shared; plain cross-sheet formulas on Refs that move or become #REF! (no shared formula holds a sheet-qualified reference, see scripts/fixtures/structure-case.ts); a merged range that shrinks; a column entry that the insertion splits; a hyperlink that moves; a defined name that moves; and drawings on two-cell, one-cell, and absolute anchors.',
    workbook: structureWorkbook,
    edits: [
      { sheet: 'Data', structure: 'deleteRows', start: 4, count: 1 },
      { sheet: 'Data', structure: 'insertColumns', start: 3, count: 1 },
    ],
  },
  {
    name: 'sheet-ops',
    purpose:
      'Sheet ops (design D1 of the change workbook-structure-editing): the edit renames the hidden sheet Data to Sales Data, adds Notes, moves Summary to the end, and deletes the hidden sheet Scratch. Formulas across sheets, a workbook-scoped and two sheet-scoped defined names, and an internal hyperlink follow the sheets, and docProps/app.xml keeps its sheet titles (scripts/fixtures/editing-cases.ts).',
    workbook: sheetOpsWorkbook,
    edits: [
      { sheet: 'Summary', ref: 'A8', value: 'edited' },
      { sheet: 'Data', sheetOp: 'rename', name: 'Sales Data' },
      { sheet: 'Notes', sheetOp: 'add', index: 1 },
      { sheet: 'Summary', sheetOp: 'move', index: 3 },
      { sheet: 'Scratch', sheetOp: 'delete' },
    ],
  },
  {
    name: 'sheet-add',
    purpose:
      'A sheet add that changes the sheet count (design D1 of the change workbook-structure-editing): the edit adds the sheet Added at the end of two sheets, and docProps/app.xml keeps its two sheet titles, so the save check opens a workbook whose sheet count and titles differ (scripts/fixtures/editing-cases.ts).',
    workbook: sheetAddWorkbook,
    edits: [
      { sheet: 'Report', ref: 'A4', value: 'edited' },
      { sheet: 'Added', sheetOp: 'add', index: 2 },
    ],
  },
  {
    name: 'clipboard-edit',
    purpose:
      'Clipboard paste (design D2 of the change workbook-structure-editing), lowered through the editor clipboard: a copy of a header and two rows with formulas and cell formats, a cut of a row, and a plain-text paste of typed input (scripts/fixtures/editing-cases.ts).',
    workbook: clipboardWorkbook,
    edits: [
      { sheet: 'Clip', ref: 'A6', value: 'edited' },
      { sheet: 'Clip', copy: 'A1:D3', paste: 'A7' },
      { sheet: 'Clip', copy: 'A4:D4', paste: 'F2', cut: true },
      {
        sheet: 'Clip',
        pasteText: 'Plain\t12\t=B11*2\r\n-5\t007\t2024-03-01\r\n12%\t$5.50',
        at: 'A11',
      },
    ],
  },
  {
    name: 'resize-edit',
    purpose:
      'Column widths and row heights set by a resize (design D3 of the change workbook-structure-editing): the edit splits a column entry, adds one for a column with none, and sets the height of a row with and a row without an x:row. The labels in row 1 and column A name the sizes of the edited copy, so the original prints Width 20 and Height 30 at their old sizes (scripts/fixtures/editing-cases.ts).',
    workbook: resizeWorkbook,
    edits: [
      { sheet: 'Sizes', ref: 'A1', value: 'edited' },
      { sheet: 'Sizes', resize: 'columns', first: 2, last: 3, size: 20 },
      { sheet: 'Sizes', resize: 'columns', first: 6, last: 6, size: 4 },
      { sheet: 'Sizes', resize: 'rows', first: 2, last: 3, size: 30 },
      { sheet: 'Sizes', resize: 'rows', first: 8, last: 8, size: 40 },
    ],
  },
  {
    name: 'fills-patterns',
    purpose:
      'Each pattern type with set and absent colors, the same pattern in adjacent cells and merged ranges, and linear and path gradients (scripts/fixtures/fill-style-cases.ts).',
    workbook: fillsPatternsWorkbook,
    edits: [{ sheet: 'Fills', ref: 'A1', value: 'edited' }],
  },
  {
    name: 'cell-styles',
    purpose:
      'Cell formats that omit a record or set their own, with each apply attribute state, under cell style formats (scripts/fixtures/fill-style-cases.ts). The edit sets italic and wrap on cells that omit their font and alignment.',
    workbook: cellStylesWorkbook,
    edits: [
      { sheet: 'Styles', ref: styleRef('font omit apply absent'), format: { italic: true } },
      { sheet: 'Styles', ref: styleRef('alignment omit apply absent'), format: { wrapText: true } },
      { sheet: 'Styles', ref: 'C1', value: 'edited' },
    ],
  },
  {
    name: 'macro-enabled',
    purpose:
      'A macro-enabled workbook with no macro part. Its main part has the macro-enabled content type (scripts/fixtures/builder.ts, `macroEnabled`). The save keeps that content type. A script cannot make a valid macro part, so the unit tests of packages/core cover a workbook with one.',
    extension: 'xlsm',
    workbook: () => ({
      macroEnabled: true,
      sheets: [
        {
          name: 'Budget',
          cols: '<cols><col min="1" max="1" width="14" customWidth="1"/></cols>',
          rows: [
            row(1, [c('A1', 'Region'), c('B1', 'Amount')]),
            row(2, [c('A2', 'North'), c('B2', 120)]),
            row(3, [c('A3', 'South'), c('B3', 80)]),
            row(4, [c('A4', 'Total'), c('B4', 200, { f: 'SUM(B2:B3)' })]),
          ].join(''),
        },
      ],
    }),
    edits: [{ sheet: 'Budget', ref: 'C1', value: 'edited' }],
  },
];

function styleRef(label: string): string {
  const hit = STYLE_RESOLUTION_CASES.find((k) => k.label === label);
  if (!hit) throw new Error(`no cell style case ${label}`);
  return hit.ref;
}
