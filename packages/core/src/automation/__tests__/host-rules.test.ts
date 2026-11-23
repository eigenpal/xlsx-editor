// The stated rules of the automation host, each at its edge: every limit at its value and one
// past it, the grid corners, address forms and sheet-name quoting, number format reads, the
// opening worksheet, the index of a store refusal, and the refusals that the planner makes
// before the store (`host.ts`, "REFUSED BEFORE THE STORE"), counted in sheet index builds.

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildXlsx, c, row, type FixtureSheet } from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx, WorkbookDocumentStore, xmlPartOf } from '../../store/index.ts';
import { cellChildTestRecorder, withPart } from '../../store/store/cells.ts';
import { MAX_CELL_TEXT } from '../../store/store/ops.ts';
import { sheetIndexTestRecorder } from '../../store/workbook/sheet-index.ts';
import {
  MAX_ARRAY_RANGES,
  MAX_BATCH_CELLS,
  MAX_BATCH_LISTED_WORKSHEETS,
  MAX_BATCH_MERGE_CHECKS,
  MAX_BATCH_OPERATIONS,
  MAX_BATCH_WRITE_CELLS,
  MAX_CELL_CHILDREN,
  type AutomationBatchResponse,
  type AutomationHost,
  type AutomationOperation,
} from '../index.ts';
import { createStoreAutomationHost } from '../host.ts';
import { MAX_MERGES } from '../../layout/panes.ts';
import { MAX_RANGE_CELLS } from '../range-address.ts';

function open(bytes: Uint8Array): AutomationHost {
  const read = readXlsx(bytes);
  if (!read.ok) throw new Error('the test workbook does not open');
  return createStoreAutomationHost(new WorkbookDocumentStore(read.pkg));
}

const book = (sheets: FixtureSheet[], extra: Parameters<typeof buildXlsx>[0] = { sheets: [] }) =>
  open(buildXlsx({ ...extra, sheets }));

const data = (rows = '', beforeMerges?: string) =>
  book([{ name: 'Data', rows, ...(beforeMerges ? { beforeMerges } : {}) }]);

const read = (address: string, extra: Partial<AutomationOperation> = {}): AutomationOperation =>
  ({ op: 'readRange', sheet: { name: 'Data' }, address, values: true, ...extra }) as never;

const writeGrid = (address: string, values: unknown[][]): AutomationOperation =>
  ({ op: 'writeValues', sheet: { name: 'Data' }, address, values }) as never;

const num = (value: number) => ({ type: 'number', value });

function refusal(response: AutomationBatchResponse) {
  expect(response.ok).toBe(false);
  const at = response.results.findIndex((r) => r.status === 'error');
  const result = response.results[at];
  return { at, error: result?.status === 'error' ? result.error : undefined };
}

function okValue(response: AutomationBatchResponse, at = 0): unknown {
  expect(response.ok).toBe(true);
  const result = response.results[at];
  return result?.status === 'ok' ? result.value : undefined;
}

const valuesOf = (response: AutomationBatchResponse, at = 0): unknown =>
  (okValue(response, at) as { values?: unknown } | undefined)?.values;

describe('limits at their value and one past it', () => {
  test('the limits are the numbers that the docs state', () => {
    expect({
      MAX_BATCH_OPERATIONS,
      MAX_BATCH_CELLS,
      MAX_BATCH_WRITE_CELLS,
      MAX_BATCH_LISTED_WORKSHEETS,
      MAX_ARRAY_RANGES,
      MAX_RANGE_CELLS,
      MAX_CELL_TEXT,
      MAX_CELL_CHILDREN,
      MAX_MERGES,
      MAX_BATCH_MERGE_CHECKS,
    }).toEqual({
      MAX_BATCH_OPERATIONS: 10_000,
      MAX_BATCH_CELLS: 1_000_000,
      MAX_BATCH_WRITE_CELLS: 10_000,
      MAX_BATCH_LISTED_WORKSHEETS: 100_000,
      MAX_ARRAY_RANGES: 10_000,
      MAX_RANGE_CELLS: 100_000,
      MAX_CELL_TEXT: 32_767,
      MAX_CELL_CHILDREN: 1_000,
      MAX_MERGES: 100_000,
      MAX_BATCH_MERGE_CHECKS: 50_000_000,
    });
  });

  test('operations', () => {
    const host = data();
    const at = (n: number) => Array.from({ length: n }, () => read('A1', { values: false }));
    expect(host.execute({ operations: at(MAX_BATCH_OPERATIONS) }).ok).toBe(true);
    expect(refusal(host.execute({ operations: at(MAX_BATCH_OPERATIONS + 1) })).error).toMatchObject(
      { code: 'resource-limit', detail: 'operations' }
    );
  });

  test('cells of one range, of the reads and writes, and of the writes', () => {
    const host = data();
    const tall = `A1:A${MAX_RANGE_CELLS}`;
    expect(host.execute({ operations: [read(tall)] }).ok).toBe(true);
    expect(
      refusal(host.execute({ operations: [read(`A1:A${MAX_RANGE_CELLS + 1}`)] })).error?.detail
    ).toBe('range-cells');
    const reads = (n: number) => Array.from({ length: n }, () => read(tall));
    expect(host.execute({ operations: reads(MAX_BATCH_CELLS / MAX_RANGE_CELLS) }).ok).toBe(true);
    expect(
      refusal(
        host.execute({ operations: [...reads(MAX_BATCH_CELLS / MAX_RANGE_CELLS), read('B1')] })
      ).error?.detail
    ).toBe('batch-cells');
    const column = (n: number) => Array.from({ length: n }, () => [null]);
    const full = writeGrid(`C1:C${MAX_BATCH_WRITE_CELLS}`, column(MAX_BATCH_WRITE_CELLS));
    expect(host.execute({ operations: [full] }).ok).toBe(true);
    expect(
      refusal(host.execute({ operations: [full, writeGrid('D1', [[null]])] })).error?.detail
    ).toBe('write-cells');
  });

  test('worksheets that the lists give', () => {
    const sheets = (n: number) =>
      Array.from({ length: n }, (_, i): FixtureSheet => ({ name: `S${i}`, rows: '' }));
    const per = 10;
    const lists = MAX_BATCH_LISTED_WORKSHEETS / per;
    expect(lists).toBeLessThanOrEqual(MAX_BATCH_OPERATIONS);
    const ops = Array.from({ length: lists }, () => ({ op: 'listWorksheets' }) as const);
    expect(book(sheets(per)).execute({ operations: ops }).ok).toBe(true);
    expect(refusal(book(sheets(per + 1)).execute({ operations: ops })).error).toMatchObject({
      code: 'resource-limit',
      detail: 'batch-worksheets',
    });
  });

  test('array ranges of a worksheet', () => {
    const ranges = (n: number) =>
      Array.from({ length: n }, (_, i) =>
        row(i + 10, [`<c r="D${i + 10}"><f t="array" ref="D${i + 10}">1</f><v>1</v></c>`])
      ).join('');
    expect(data(ranges(MAX_ARRAY_RANGES)).execute({ operations: [read('A1')] }).ok).toBe(true);
    expect(
      refusal(data(ranges(MAX_ARRAY_RANGES + 1)).execute({ operations: [read('A1')] })).error
        ?.detail
    ).toBe('formula-ranges');
  });

  test('the corners of the grid', () => {
    const host = data();
    const address = (text: string) => okValue(host.execute({ operations: [read(text)] }));
    expect(address('XFD1048576')).toMatchObject({ address: 'Data!XFD1048576' });
    expect(address('A1048576')).toMatchObject({ address: 'Data!A1048576' });
    expect(address('XFD1')).toMatchObject({ address: 'Data!XFD1' });
    for (const outside of ['A1048577', 'XFE1', 'A0', 'XFD1048577'])
      expect(refusal(host.execute({ operations: [read(outside)] })).error?.code).toBe(
        'invalid-argument'
      );
  });
});

describe('addresses', () => {
  test('a rectangle reads in any corner order, letter case, and with $ marks', () => {
    const host = data(row(1, [c('A1', 1), c('B1', 2)]) + row(2, [c('A2', 3), c('B2', 4)]));
    for (const text of ['B2:A1', 'a1:$b$2', '$A$1:B2', 'A2:B1'])
      expect(okValue(host.execute({ operations: [read(text)] }))).toEqual({
        address: 'Data!A1:B2',
        values: [
          [1, 2],
          [3, 4],
        ],
      });
  });

  test('an address of three parts, a sheet name, a whole row, or a name is refused', () => {
    const host = data();
    for (const text of ['A1:B2:C3', 'Data!A1', '1:1', 'Total'])
      expect(refusal(host.execute({ operations: [read(text)] })).error?.code).toBe('unsupported');
  });

  test('a sheet name is quoted when formula text needs it', () => {
    const names: [string, string][] = [
      ['Plain', 'Plain!A1'],
      ['A1', "'A1'!A1"],
      ['TRUE', "'TRUE'!A1"],
      ['false', "'false'!A1"],
      ['R1C1', "'R1C1'!A1"],
      ['C3', "'C3'!A1"],
      ['C', "'C'!A1"],
      ['r', "'r'!A1"],
      ["a''b", "'a''''b'!A1"],
      ['My Sheet', "'My Sheet'!A1"],
      ["It's", "'It''s'!A1"],
      ['1st', "'1st'!A1"],
    ];
    const host = book(names.map(([name]) => ({ name, rows: '' })));
    for (const [name, address] of names)
      expect(
        okValue(host.execute({ operations: [{ ...read('A1'), sheet: { name } } as never] }))
      ).toMatchObject({ address });
  });
});

describe('stored values that cannot be read', () => {
  test('a cell the sheet index cannot place, an unknown error code, or a malformed number', () => {
    const cases = [
      row(1, [c('A1', 1), c('A1', 2)]),
      row(1, ['<c r="A1" t="e"><v>#FOO!</v></c>']),
      row(1, ['<c r="A1"><v>abc</v></c>']),
    ];
    for (const rows of cases)
      expect(refusal(data(rows).execute({ operations: [read('A1')] })).error?.code).toBe(
        'unsupported'
      );
  });

  test('an address longer than any cell address is refused', () => {
    const long = `A${'0'.repeat(40)}1`;
    expect(refusal(data().execute({ operations: [read(long)] })).error?.code).toBe('unsupported');
  });
});

describe('number formats', () => {
  const formats = (cellXfs: string[], numFmts: string[] = []) =>
    open(
      buildXlsx({
        sheets: [
          {
            name: 'Data',
            rows: row(
              1,
              cellXfs.map((_, i) => `<c r="${'ABCDEFGH'[i]}1" s="${i + 1}"/>`)
            ),
          },
        ],
        styles: { cellXfs, ...(numFmts.length > 0 ? { numFmts } : {}) },
      })
    );
  const xf = (id: number) => `<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0" xfId="0"/>`;
  const readFormat = (host: AutomationHost, address: string) =>
    host.execute({ operations: [read(address, { values: false, numberFormat: true })] });

  test('a read gives the stored code, or the ECMA-376 code of a built-in id', () => {
    const host = formats([xf(14), xf(10), xf(164)], ['<numFmt numFmtId="164" formatCode="0.0"/>']);
    expect(okValue(readFormat(host, 'A1:D1'))).toMatchObject({
      numberFormat: [['mm-dd-yy', '0.00%', '0.0', 'General']],
    });
  });

  test('a formula cell gives its format, though its value read is refused', () => {
    const host = formats([xf(10)]);
    const formula = data(row(1, ['<c r="A1" s="0"><f>1+1</f><v>2</v></c>']));
    expect(okValue(readFormat(host, 'A1'))).toMatchObject({ numberFormat: [['0.00%']] });
    expect(okValue(readFormat(formula, 'A1'))).toMatchObject({ numberFormat: [['General']] });
    expect(refusal(formula.execute({ operations: [read('A1')] })).error?.detail).toBe(
      'formula-cell'
    );
  });

  test('a locale built-in id, a malformed or undefined id, or a cell format out of range is refused', () => {
    const cases: [ReturnType<typeof formats>, string][] = [
      [formats([xf(30)]), 'A1'],
      [formats([xf(55)]), 'A1'],
      [formats([xf(200)]), 'A1'],
      [data(row(1, ['<c r="A1" s="40"/>'])), 'A1'],
      [formats(['<xf numFmtId="abc" fontId="0" fillId="0" borderId="0" xfId="0"/>']), 'A1'],
    ];
    for (const [host, address] of cases)
      expect(refusal(readFormat(host, address)).error).toMatchObject({
        code: 'unsupported',
        detail: 'unreadable-format',
      });
  });
});

describe('the opening worksheet', () => {
  const sheets = (hidden: number[] = []): FixtureSheet[] =>
    ['One', 'Two', 'Three'].map((name, i) => ({
      name,
      rows: '',
      ...(hidden.includes(i) ? { state: 'hidden' as const } : {}),
    }));
  const active = (host: AutomationHost) =>
    okValue(host.execute({ operations: [{ op: 'getWorksheet', sheet: { active: true } }] }));

  test('is the active tab when it is a visible worksheet', () => {
    expect(active(book(sheets(), { sheets: [], activeTab: 1 }))).toMatchObject({ name: 'Two' });
    expect(active(book(sheets(), { sheets: [], activeTab: 2 }))).toMatchObject({ name: 'Three' });
  });

  test('is the first visible worksheet when the active tab is hidden', () => {
    expect(active(book(sheets([0, 1]), { sheets: [], activeTab: 1 }))).toMatchObject({
      name: 'Three',
    });
  });

  test('is the first visible worksheet when the active tab is very hidden', () => {
    const veryHidden = sheets().map((sheet, i) =>
      i === 1 ? { ...sheet, state: 'veryHidden' as const } : sheet
    );
    expect(active(book(veryHidden, { sheets: [], activeTab: 1 }))).toMatchObject({ name: 'One' });
  });

  test('is missing when no worksheet is visible', () => {
    expect(
      refusal(
        book(sheets([0, 1, 2]), { sheets: [], activeTab: 0 }).execute({
          operations: [{ op: 'getWorksheet', sheet: { active: true } }],
        })
      ).error
    ).toMatchObject({ code: 'item-not-found', detail: 'no-active' });
  });
});

describe('a chartsheet', () => {
  // A chartsheet is a sheet of the workbook but not a worksheet (`x:chartsheet`, ECMA-376 Part 1, 18.3.1.12),
  // so the host never lists it, finds it by name, or opens on it.
  const CHARTSHEET =
    '<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<sheetViews><sheetView workbookViewId="0"/></sheetViews></chartsheet>';
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const withChartsheet = () =>
    book(
      [
        { name: 'Data', rows: '' },
        { name: 'Two', rows: '' },
      ],
      {
        sheets: [],
        // The chartsheet is the third sheet and the active tab.
        activeTab: 2,
        extraSheets: '<sheet name="Chart1" sheetId="9" r:id="rIdX1"/>',
        extraParts: [
          [
            'xl/chartsheets/sheet1.xml',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml',
            new TextEncoder().encode(CHARTSHEET),
            `${REL}/chartsheet`,
          ],
        ],
      }
    );

  test('is not in the worksheet list', () => {
    const list = okValue(withChartsheet().execute({ operations: [{ op: 'listWorksheets' }] }));
    expect((list as { name: string }[]).map((sheet) => sheet.name)).toEqual(['Data', 'Two']);
  });

  test('is not found by its name', () => {
    const response = withChartsheet().execute({
      operations: [{ op: 'getWorksheet', sheet: { name: 'Chart1' } }],
    });
    expect(refusal(response).error).toMatchObject({ code: 'item-not-found', detail: 'name' });
  });

  test('is not the opening worksheet when it is the active tab', () => {
    const response = withChartsheet().execute({
      operations: [{ op: 'getWorksheet', sheet: { active: true } }],
    });
    expect(okValue(response)).toMatchObject({ name: 'Data' });
  });
});

describe('store refusals', () => {
  test('a refusal names the first refused write', () => {
    const host = data('', '<sheetProtection sheet="1"/>');
    const one = writeGrid('A1', [[num(1)]]);
    const { at, error } = refusal(host.execute({ operations: [one, one, one] }));
    expect(at).toBe(0);
    expect(error).toMatchObject({ code: 'transaction-refused', detail: 'locked' });
  });

  test('a write reports a change, and a read does not', () => {
    const host = data();
    expect(host.execute({ operations: [read('A1')] }).changed).toBe(false);
    expect(host.execute({ operations: [writeGrid('A1', [[num(1)]])] }).changed).toBe(true);
    expect(host.execute({ operations: [writeGrid('A1', [[num(1)]])] }).changed).toBe(false);
  });
});

describe('refused before the store', () => {
  const recorder = sheetIndexTestRecorder();
  const column = (last: unknown) =>
    writeGrid(
      `G1:G${MAX_BATCH_WRITE_CELLS}`,
      Array.from({ length: MAX_BATCH_WRITE_CELLS }, (_, i) => [
        i === MAX_BATCH_WRITE_CELLS - 1 ? last : num(i),
      ])
    );

  test('a text longer than a cell holds, after 9,999 writes, in one planning pass', () => {
    const host = data();
    const long = { type: 'string', value: 'x'.repeat(MAX_CELL_TEXT + 1) };
    recorder.reset();
    const { at, error } = refusal(host.execute({ operations: [read('A1'), column(long)] }));
    expect({ at, error }).toMatchObject({
      at: 1,
      error: { code: 'transaction-refused', detail: 'text-too-long' },
    });
    expect(recorder.builds).toBeLessThan(5);
    const most = { type: 'string', value: 'x'.repeat(MAX_CELL_TEXT) };
    expect(host.execute({ operations: [column(most)] }).ok).toBe(true);
  });

  test('a cell whose children the edit cannot rewrite, after 9,999 writes', () => {
    const blocked = `<c r="G${MAX_BATCH_WRITE_CELLS}"><v>2</v><e:w xmlns:e="urn:example:e"><v>3</v></e:w></c>`;
    const host = data(row(MAX_BATCH_WRITE_CELLS, [blocked]));
    recorder.reset();
    expect(refusal(host.execute({ operations: [column(num(1))] })).error).toMatchObject({
      code: 'transaction-refused',
      detail: 'cell-alternate-content',
    });
    expect(recorder.builds).toBeLessThan(5);
  });
});

describe('a shared string table that cannot take text', () => {
  test('refuses the first write of text before the store, and nothing applies', () => {
    const read0 = readXlsx(buildXlsx({ strings: ['a'], sheets: [{ name: 'Data', rows: '' }] }));
    if (!read0.ok) throw new Error('the test workbook does not open');
    const part = xmlPartOf(read0.pkg, '/xl/sharedStrings.xml');
    if (!part) throw new Error('the test workbook has no shared string table');
    // A table whose root is not a shared string table, as a damaged part reads.
    const damaged = { ...part, root: { ...part.root, kind: 'generic' } } as never;
    const store = new WorkbookDocumentStore(withPart(read0.pkg, damaged));
    const before = store.revision;
    let transactions = 0;
    const transact = store.transact.bind(store);
    store.transact = (...args) => ((transactions += 1), transact(...args));
    const host = createStoreAutomationHost(store);
    const text = { type: 'string', value: 'x' };
    expect(
      refusal(
        host.execute({ operations: [writeGrid('A1', [[num(1)]]), writeGrid('B1', [[text]])] })
      )
    ).toMatchObject({
      at: 1,
      error: { code: 'transaction-refused', detail: 'unsupported-package' },
    });
    expect(transactions).toBe(0);
    expect(store.revision).toBe(before);
    expect(host.execute({ operations: [writeGrid('A1', [[num(1)]])] }).ok).toBe(true);
  });
});

describe('one store pass for each worksheet', () => {
  const recorder = sheetIndexTestRecorder();
  const MOVES = 1000;
  const filled = (name: string, beforeMerges?: string): FixtureSheet => ({
    name,
    rows: Array.from({ length: 1000 }, (_, i) => row(i + 1, [c(`B${i + 1}`, i)])).join(''),
    ...(beforeMerges ? { beforeMerges } : {}),
  });
  const moves = (last?: AutomationOperation) =>
    Array.from({ length: MOVES }, (_, i) =>
      i === MOVES - 1 && last
        ? last
        : ({
            op: 'writeValues',
            sheet: { name: i % 2 === 0 ? 'Data' : 'Other' },
            address: `A${i + 1}`,
            values: [[num(i)]],
          } as never)
    );

  test('writes that move between two worksheets cost a few passes, not one for each move', () => {
    const host = book([filled('Data'), filled('Other')]);
    recorder.reset();
    expect(host.execute({ operations: moves() }).ok).toBe(true);
    expect(recorder.builds).toBeLessThan(10);
    const other = (address: string) =>
      ({ op: 'readRange', sheet: { name: 'Other' }, address, values: true }) as never;
    expect(valuesOf(host.execute({ operations: [read('A1:A3'), other('A1:A2')] }))).toEqual([
      [0],
      [''],
      [2],
    ]);
    expect(valuesOf(host.execute({ operations: [other('A1:A2')] }))).toEqual([[''], [1]]);
  });

  test('a refused last write is found in a few passes for each worksheet', () => {
    const host = book([
      filled('Data'),
      filled('Other'),
      { name: 'Locked', rows: '', beforeMerges: '<sheetProtection sheet="1"/>' },
    ]);
    const locked = {
      op: 'writeValues',
      sheet: { name: 'Locked' },
      address: 'A1',
      values: [[num(1)]],
    } as never;
    recorder.reset();
    const { at, error } = refusal(host.execute({ operations: moves(locked) }));
    expect({ at, error }).toMatchObject({
      at: MOVES - 1,
      error: { code: 'transaction-refused', detail: 'locked' },
    });
    expect(recorder.builds).toBeLessThan(60);
  });

  test('two writes to one cell keep their order', () => {
    const host = book([filled('Data'), filled('Other')]);
    const other = {
      op: 'writeValues',
      sheet: { name: 'Other' },
      address: 'A1',
      values: [[num(9)]],
    };
    expect(
      host.execute({
        operations: [writeGrid('A1', [[num(1)]]), other as never, writeGrid('A1', [[num(2)]])],
      }).ok
    ).toBe(true);
    expect(valuesOf(host.execute({ operations: [read('A1')] }))).toEqual([[2]]);
  });
});

describe('a value that the read cannot give', () => {
  const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
  const wrapped = (r: string) =>
    `<c r="${r}"><mc:AlternateContent xmlns:mc="${MC}"><mc:Fallback><v>3</v></mc:Fallback>` +
    '</mc:AlternateContent></c>';

  test('a cell whose value is only in a wrapper refuses its read, so an empty read never erases it', () => {
    const host = data(row(1, [wrapped('A1'), c('B1', 4)]));
    expect(refusal(host.execute({ operations: [read('A1:B1')] })).error).toMatchObject({
      code: 'unsupported',
      detail: 'unreadable-value',
    });
    expect(valuesOf(host.execute({ operations: [read('B1')] }))).toEqual([[4]]);
    expect(
      okValue(host.execute({ operations: [read('A1', { values: false, numberFormat: true })] }))
    ).toBeDefined();
  });

  test('a value deep in an element of another namespace refuses a write after a read', () => {
    // The wrapper has an element grandchild, so the read check keeps its answer for the
    // wrapper, and the write check of the cell takes that answer.
    const deep = '<e:w xmlns:e="urn:example:e"><e:a><v>3</v></e:a></e:w>';
    const host = data(row(1, [`<c r="A1"><v>2</v>${deep}</c>`]));
    expect(refusal(host.execute({ operations: [read('A1')] })).error?.detail).toBe(
      'unreadable-value'
    );
    for (let pass = 0; pass < 2; pass += 1)
      expect(
        refusal(host.execute({ operations: [writeGrid('A1', [[num(7)]])] })).error
      ).toMatchObject({ code: 'transaction-refused', detail: 'cell-alternate-content' });
  });

  test('a value in an element of another namespace also refuses its read', () => {
    const host = data(row(1, ['<c r="A1"><e:w xmlns:e="urn:example:e"><v>3</v></e:w></c>']));
    expect(refusal(host.execute({ operations: [read('A1')] })).error?.detail).toBe(
      'unreadable-value'
    );
  });

  // A string item whose text is in a child that the read does not take reads as "" or as part of
  // its text, so the read refuses it, for an inline string and for a shared string alike.
  const alternate = (text: string) =>
    `<mc:AlternateContent xmlns:mc="${MC}"><mc:Fallback><t>${text}</t></mc:Fallback>` +
    '</mc:AlternateContent>';
  const hidden = [
    `<si>${alternate('all')}</si>`,
    `<si><t>a</t><r>${alternate('b')}</r></si>`,
    '<si><t>a</t><e:w xmlns:e="urn:example:e"><t>b</t></e:w></si>',
    `<si><t>a${alternate('b')}</t></si>`,
    // A plain run after the run or child that hides text does not make the item readable.
    `<si><r>${alternate('a')}</r><r><t>b</t></r></si>`,
    '<si><e:w xmlns:e="urn:example:e"><t>a</t></e:w><r><t>b</t></r></si>',
    // An element inside the x:t of a run.
    `<si><t>a</t><r><t>b${alternate('c')}</t></r></si>`,
  ];

  test('an inline string with text in a wrapper refuses its read', () => {
    for (const item of hidden) {
      const is = item.replace(/^<si>/, '<is>').replace(/<\/si>$/, '</is>');
      const host = data(row(1, [`<c r="A1" t="inlineStr">${is}</c>`, c('B1', 4)]));
      expect(refusal(host.execute({ operations: [read('A1:B1')] })).error).toMatchObject({
        code: 'unsupported',
        detail: 'unreadable-value',
      });
      expect(valuesOf(host.execute({ operations: [read('B1')] }))).toEqual([[4]]);
    }
  });

  test('a shared string cell whose item holds text in a wrapper refuses its read', () => {
    const rich = '<si><t>x</t><r><rPr><b/></rPr><t>y</t></r><rPh sb="0" eb="1"><t>z</t></rPh></si>';
    const columns = 'ABCDEFG';
    const cells = hidden.map((_, i) => c(`${columns[i]}1`, i + 1, { t: 's' }));
    const host = book([{ name: 'Data', rows: row(1, [...cells, c('H1', 0, { t: 's' })]) }], {
      sheets: [],
      rawStrings: [rich, ...hidden],
    });
    // Each item is read twice: the second read takes the answer that the first one kept.
    for (const address of [...columns].map((col) => `${col}1`))
      for (let pass = 0; pass < 2; pass += 1)
        expect(refusal(host.execute({ operations: [read(address)] })).error).toMatchObject({
          code: 'unsupported',
          detail: 'unreadable-value',
        });
    // A rich string with runs and a phonetic run reads as its text.
    expect(valuesOf(host.execute({ operations: [read('H1')] }))).toEqual([['xy']]);
  });
});

describe('a wide child of another namespace', () => {
  // One cell with a child in another namespace of 200,000 empty elements, read and written 10,000
  // times in one batch: the cell child checks walk that subtree once, not once per operation.
  const WIDE = 200_000;
  const wide = () =>
    data(row(1, [`<c r="A1"><v>1</v><q:w xmlns:q="urn:q">${'<q:e/>'.repeat(WIDE)}</q:w></c>`]));
  const repeated = (operation: AutomationOperation) =>
    Array.from({ length: MAX_BATCH_OPERATIONS }, () => operation);

  test('10,000 reads of the cell walk its subtree once', () => {
    const host = wide();
    const recorder = cellChildTestRecorder();
    recorder.reset();
    expect(host.execute({ operations: repeated(read('A1')) }).ok).toBe(true);
    expect(recorder.walks).toBe(1);
    expect(recorder.nodes).toBe(WIDE);
  });

  test('10,000 writes to the cell walk its subtree once', () => {
    const host = wide();
    const recorder = cellChildTestRecorder();
    recorder.reset();
    expect(host.execute({ operations: repeated(writeGrid('A1', [[num(2)]])) }).ok).toBe(true);
    expect(recorder.walks).toBe(1);
    expect(recorder.nodes).toBe(WIDE);
    expect(valuesOf(host.execute({ operations: [read('A1')] }))).toEqual([[2]]);
  });
  test('a read and a write of the cell in one batch walk its subtree once', () => {
    const host = wide();
    const recorder = cellChildTestRecorder();
    recorder.reset();
    const operations = [read('A1'), writeGrid('A1', [[num(2)]])];
    expect(host.execute({ operations }).ok).toBe(true);
    expect(recorder.walks).toBe(1);
  });
});

describe('a cell with many children', () => {
  // A cell with 200,000 direct children in another namespace: each operation on a cell reads its
  // children, so 10,000 operations on it would block the batch. The host refuses the cell.
  const many = (count: number) =>
    data(row(1, [`<c r="A1"><v>1</v>${'<q:e xmlns:q="urn:q"/>'.repeat(count - 1)}</c>`]));

  test('a read and a write of a cell past the limit refuse, and name the cell', () => {
    const host = many(200_000);
    for (const operation of [read('A1:B2'), writeGrid('A1', [[num(2)]])])
      expect(refusal(host.execute({ operations: [operation] })).error).toMatchObject({
        code: 'resource-limit',
        detail: 'cell-children',
        cell: 'A1',
      });
    // A number format read and a write of null do not read the children.
    expect(
      host.execute({ operations: [read('A1', { values: false, numberFormat: true } as never)] }).ok
    ).toBe(true);
    expect(host.execute({ operations: [writeGrid('A1', [[null]])] }).ok).toBe(true);
  });

  test('a cell at the limit is read and written', () => {
    const host = many(MAX_CELL_CHILDREN);
    const operations = [read('A1'), writeGrid('A1', [[num(2)]])];
    expect(host.execute({ operations }).ok).toBe(true);
    expect(valuesOf(host.execute({ operations: [read('A1')] }))).toEqual([[2]]);
  });
});

describe('a worksheet with many merged ranges', () => {
  // Each write compares its bounds with each merged range, so the host caps their count.
  const merged = (count: number) =>
    book([
      {
        name: 'Data',
        rows: '',
        merges: Array.from({ length: count }, (_, i) => `A${i + 2}:B${i + 2}`),
      },
    ]);

  test('a value write past the limit refuses, and a read passes', () => {
    const host = merged(MAX_MERGES + 1);
    expect(
      refusal(host.execute({ operations: [writeGrid('F1', [[num(1)]])] })).error
    ).toMatchObject({ code: 'resource-limit', detail: 'merged-ranges' });
    expect(host.execute({ operations: [read('F1')] }).ok).toBe(true);
  });

  test('a worksheet at the limit takes the write', () => {
    const host = merged(MAX_MERGES);
    expect(host.execute({ operations: [writeGrid('F1', [[num(1)]])] }).ok).toBe(true);
    expect(valuesOf(host.execute({ operations: [read('F1')] }))).toEqual([[1]]);
  });

  // A 9 KB workbook with 99,999 merged ranges that all cover A1:B2, and 9,999 writes to A1 in one
  // batch: each write checks each merged range, which blocked the batch for 11 s before the
  // batch counted the checks.
  test('the writes of a batch check at most MAX_BATCH_MERGE_CHECKS merged ranges', () => {
    const count = 99_999;
    const host = book([{ name: 'Data', rows: '', merges: Array(count).fill('A1:B2') }]);
    const writes = Array.from({ length: 9_999 }, (_, i) => writeGrid('A1', [[num(i)]]));
    const past = Math.floor(MAX_BATCH_MERGE_CHECKS / count);
    expect(refusal(host.execute({ operations: writes }))).toMatchObject({
      at: past,
      error: { code: 'resource-limit', detail: 'merge-checks' },
    });
    expect(host.execute({ operations: writes.slice(0, past) }).ok).toBe(true);
  });

  test('a batch at the merge check limit takes its writes, and one check more refuses', () => {
    // 9,999 writes to a worksheet with 5,000 merged ranges, then one write to a worksheet with
    // 5,000 or 5,001: the batch makes exactly the limit of checks, or one more.
    const per = MAX_BATCH_MERGE_CHECKS / MAX_BATCH_OPERATIONS;
    const mergesOf = (count: number) =>
      Array.from({ length: count }, (_, i) => `A${i + 2}:B${i + 2}`);
    const two = (last: number) =>
      book([
        { name: 'Data', rows: '', merges: mergesOf(per) },
        { name: 'More', rows: '', merges: mergesOf(last) },
      ]);
    const writes = [
      ...Array.from({ length: MAX_BATCH_OPERATIONS - 1 }, (_, i) =>
        writeGrid(`F${i + 1}`, [[null]])
      ),
      { op: 'writeValues', sheet: { name: 'More' }, address: 'F1', values: [[null]] } as never,
    ];
    expect(two(per).execute({ operations: writes }).ok).toBe(true);
    expect(refusal(two(per + 1).execute({ operations: writes }))).toMatchObject({
      at: MAX_BATCH_OPERATIONS - 1,
      error: { code: 'resource-limit', detail: 'merge-checks' },
    });
  });
});

describe('a write that a merged range meets', () => {
  const sheet = (merges: string[], rows = '') => book([{ name: 'Data', rows, merges }]);
  const hidden = (host: AutomationHost, address: string, values: unknown[][]) =>
    refusal(host.execute({ operations: [writeGrid(address, values)] })).error;
  const n = (value: number | null) => (value === null ? null : num(value));

  test('a value under the top-left cell of a merge refuses, and names its cell', () => {
    expect(hidden(sheet(['B3:C4']), 'B4', [[num(1)]])).toMatchObject({
      code: 'unsupported',
      detail: 'merged-cell',
      cell: 'B4',
    });
  });

  test('values in several rows of a merge refuse at the first hidden cell', () => {
    expect(
      hidden(sheet(['A1:B2']), 'A1:B2', [
        [n(1), n(2)],
        [null, null],
      ])?.cell
    ).toBe('B1');
    expect(
      hidden(sheet(['B5:C6']), 'A4:C6', [
        [n(5), null, null],
        [null, null, n(7)],
        [null, null, null],
      ])?.cell
    ).toBe('C5');
    const full = [0, 1, 2].map((r) => [0, 1, 2].map((c) => n(r * 3 + c)));
    expect(hidden(sheet(['A1:C3']), 'A1:C3', full)?.cell).toBe('B1');
  });

  test('a count past 255 before a merge still finds the values that the merge hides', () => {
    // 255 values before IV1:SR1, an empty top-left cell, and 256 hidden values: a count of 8 bits
    // would give the same value at both ends of the merge.
    const values = Array.from({ length: 512 }, (_, i) => n(i === 255 ? null : i));
    expect(hidden(sheet(['IV1:SR1']), 'A1:SR1', [values])).toMatchObject({
      code: 'unsupported',
      detail: 'merged-cell',
      cell: 'IW1',
    });
  });

  test('a hidden value in the second merge that a write meets refuses', () => {
    const grid = [
      [null, null],
      [null, null],
      [null, n(1)],
    ];
    expect(hidden(sheet(['A1:B1', 'A3:B3']), 'A1:B3', grid)?.cell).toBe('B3');
  });

  test('values before the merge and an empty write in it do not move the refused cell', () => {
    // A1 is above and to the left of B2:C3, and C2 and C3 are hidden: the count of the merge is 2.
    const grid = [
      [n(1), null, null],
      [null, null, n(2)],
      [null, null, n(3)],
    ];
    expect(hidden(sheet(['B2:C3']), 'A1:C3', grid)?.cell).toBe('C2');
    // An empty write stores no value, so the hidden value after it is the refused cell.
    const empty = { type: 'empty' };
    expect(hidden(sheet(['A1:C1']), 'A1:C1', [[null, empty, n(1)]])?.cell).toBe('C1');
  });

  test('each grid of one or two values in a 4 by 4 write refuses where a merge hides it', () => {
    // The merge B2:C3 hides B3, C2, and C3. A cell-by-cell oracle gives the first hidden value in
    // reading order, or none.
    const host = sheet(['B2:C3']);
    const hides = (at: number) => {
      const [r, c] = [Math.floor(at / 4), at % 4];
      return r >= 1 && r <= 2 && c >= 1 && c <= 2 && !(r === 1 && c === 1);
    };
    const name = (at: number) => `${'ABCD'[at % 4]}${Math.floor(at / 4) + 1}`;
    for (let first = 0; first < 16; first += 1)
      for (let second = first; second < 16; second += 1) {
        const grid = [0, 1, 2, 3].map((y) =>
          [0, 1, 2, 3].map((x) => (y * 4 + x === first || y * 4 + x === second ? n(1) : null))
        );
        const response = host.execute({ operations: [writeGrid('A1:D4', grid)] });
        const hidden = [first, second].find(hides);
        const key = `${name(first)} ${name(second)}`;
        expect([key, response.ok ? undefined : refusal(response).error?.cell]).toEqual([
          key,
          hidden === undefined ? undefined : name(hidden),
        ]);
      }
  });
});

describe('a read refusal names its cell', () => {
  test('a formula cell in the third column and second row is C2', () => {
    const host = data(row(2, ['<c r="C2"><f>1+1</f><v>2</v></c>']));
    expect(refusal(host.execute({ operations: [read('A1:D4')] })).error).toMatchObject({
      detail: 'formula-cell',
      cell: 'C2',
    });
  });

  test('an empty cell that the file holds reads as an empty string', () => {
    const host = data(row(1, ['<c r="A1" s="0"/>']));
    expect(valuesOf(host.execute({ operations: [read('A1:B1')] }))).toEqual([['', '']]);
  });
});

describe('a written grid of the wrong size', () => {
  test('a row count or a column count that differs from the range is refused', () => {
    const host = data(row(1, [c('A1', 1)]));
    for (const grid of [[[num(1), num(2)]], [[num(1)], [num(2)], [num(3)]], [[num(1)], [num(2)]]])
      expect(refusal(host.execute({ operations: [writeGrid('A1:B2', grid)] })).error).toMatchObject(
        { code: 'invalid-argument', detail: 'values' }
      );
    expect(valuesOf(host.execute({ operations: [read('A1')] }))).toEqual([[1]]);
  });
});

describe('a number format read where values cannot be read', () => {
  const formatRead = (address: string) => read(address, { values: false, numberFormat: true });

  test('a cell of an array formula range gives its format', () => {
    const host = data(row(5, ['<c r="A5"><f t="array" ref="A5:B6">1</f><v>1</v></c>']));
    expect(refusal(host.execute({ operations: [read('B6')] })).error?.detail).toBe('formula-cell');
    expect(okValue(host.execute({ operations: [formatRead('A5:B6')] }))).toMatchObject({
      numberFormat: [
        ['General', 'General'],
        ['General', 'General'],
      ],
    });
  });

  test('a worksheet whose array range cannot be read gives its formats', () => {
    const host = data(row(5, ['<c r="B5"><f t="array" ref="B5:B6:B9">1</f><v>1</v></c>']));
    expect(refusal(host.execute({ operations: [read('A1')] })).error?.detail).toBe(
      'formula-ranges'
    );
    expect(okValue(host.execute({ operations: [formatRead('A1:B5')] }))).toMatchObject({
      numberFormat: Array.from({ length: 5 }, () => ['General', 'General']),
    });
  });
});

describe('a null write over an array formula range', () => {
  test('is refused, as the hint says, and a range beside it is not', () => {
    const host = data(row(5, ['<c r="A5"><f t="array" ref="A5:B6">1</f><v>1</v></c>']));
    expect(
      refusal(host.execute({ operations: [writeGrid('B6:C6', [[null, num(1)]])] })).error
    ).toMatchObject({ code: 'unsupported', detail: 'formula-cell' });
    expect(host.execute({ operations: [writeGrid('C6', [[num(1)]])] }).ok).toBe(true);
  });
});

describe('the cell that a refusal names', () => {
  const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
  const grid = (value: unknown) => Array.from({ length: 3 }, () => [value, value, value]);
  const at = (cell: string) => data(row(1, [c('A1', 1)]) + row(2, [cell]));

  test('a read names the first cell that it cannot give', () => {
    const cases: [string, string][] = [
      ['<c r="B2"><f>1</f><v>2</v></c>', 'formula-cell'],
      [
        `<c r="B2"><mc:AlternateContent xmlns:mc="${MC}"><mc:Fallback><v>3</v></mc:Fallback></mc:AlternateContent></c>`,
        'unreadable-value',
      ],
      ['<c r="B2" t="e"><v>#FOO!</v></c>', 'unreadable-value'],
      ['<c r="B2" t="s"><v>9</v></c>', 'unreadable-value'],
      ['<c r="B2" t="d"><v>x</v></c>', 'unreadable-value'],
      [
        '<c r="B2" t="inlineStr"><is><t>a</t><e:w xmlns:e="urn:example:e"><t>b</t></e:w></is></c>',
        'unreadable-value',
      ],
    ];
    for (const [cell, detail] of cases)
      expect([cell, refusal(at(cell).execute({ operations: [read('A1:C3')] })).error]).toEqual([
        cell,
        expect.objectContaining({ code: 'unsupported', detail, cell: 'B2' }),
      ]);
    const format = read('A1:C3', { values: false, numberFormat: true });
    expect(refusal(at('<c r="B2" s="40"/>').execute({ operations: [format] })).error).toEqual(
      expect.objectContaining({ detail: 'unreadable-format', cell: 'B2' })
    );
  });

  test('an array range names its first cell inside the range of the operation', () => {
    const host = data(row(5, ['<c r="A5"><f t="array" ref="A5:B6">1</f><v>1</v></c>']));
    for (const [address, cell, rows, cols] of [
      ['B4:C9', 'B5', 6, 2],
      ['A6:A9', 'A6', 4, 1],
      ['A1:F9', 'A5', 9, 6],
    ] as const) {
      const nulls = Array.from({ length: rows }, () => Array.from({ length: cols }, () => null));
      for (const operation of [read(address), writeGrid(address, nulls)])
        expect(refusal(host.execute({ operations: [operation] })).error).toEqual(
          expect.objectContaining({ detail: 'formula-cell', cell })
        );
    }
  });

  test('a write names the first cell that it cannot write', () => {
    const cases: [string, unknown, string][] = [
      ['<c r="B2"><f>1</f><v>2</v></c>', num(1), 'formula-cell'],
      [
        '<c r="B2"><v>2</v><e:w xmlns:e="urn:example:e"><v>3</v></e:w></c>',
        num(1),
        'cell-alternate-content',
      ],
      ['<c r="C2"/>', { type: 'string', value: 'x'.repeat(MAX_CELL_TEXT + 1) }, 'text-too-long'],
    ];
    for (const [cell, value, detail] of cases)
      expect(
        refusal(at(cell).execute({ operations: [writeGrid('A1:C3', grid(value))] })).error
      ).toEqual(
        expect.objectContaining({ detail, cell: detail === 'text-too-long' ? 'A1' : 'B2' })
      );
  });

  test('a refusal that is not about one cell names no cell', () => {
    const host = data();
    const { error } = refusal(host.execute({ operations: [read('A1:B2:C3')] }));
    expect(error && 'cell' in error).toBe(false);
  });
});

describe('a worksheet whose x:sheetData is not in place', () => {
  const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

  /** The Data worksheet with its `x:sheetData` element replaced by `change(sheetData)`. */
  function moved(change: (sheetData: string) => string): AutomationHost {
    const files = unzipSync(buildXlsx({ sheets: [{ name: 'Data', rows: row(1, [c('A1', 5)]) }] }));
    const name = 'xl/worksheets/sheet1.xml';
    files[name] = strToU8(
      strFromU8(files[name]!)
        .replace('<worksheet ', `<worksheet xmlns:mc="${MC}" `)
        .replace(/<sheetData>[\s\S]*<\/sheetData>/, change)
    );
    return open(zipSync(files));
  }

  // A consumer that takes the wrapper reads 5 in A1: an empty read would erase it.
  const wrapped = (sheetData: string) =>
    `<mc:AlternateContent><mc:Choice Requires="mc">${sheetData}</mc:Choice></mc:AlternateContent>`;

  test('refuses each read and each write, as the store refuses each write', () => {
    for (const host of [moved(wrapped), moved(() => '')])
      for (const operation of [
        read('A1'),
        read('A1', { values: false, numberFormat: true }),
        writeGrid('A1', [[num(1)]]),
      ])
        expect(refusal(host.execute({ operations: [operation] })).error).toMatchObject({
          code: 'unsupported',
          detail: 'unreadable-sheet',
        });
    // The address alone needs no cell.
    const address = read('A1', { values: false });
    expect(okValue(moved(wrapped).execute({ operations: [address] }))).toEqual({
      address: 'Data!A1',
    });
  });
});
