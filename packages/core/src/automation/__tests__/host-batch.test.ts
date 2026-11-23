// The work and the order of one batch. The reads of a batch resolve each cell, row, and cell
// format once (`host-memo.ts`, `ReadMemo`). A hostile workbook of a few kilobytes can put 200,000
// text nodes in one `x:v`, or 400,000 attributes on one `x:c` or `x:row`. Without the memo,
// 10,000 reads of such a cell, or one read of a row of 16,384 empty cells, block the main thread
// for more than 5 s. The checks count the work, as `cellChildTestRecorder` does: a count is exact
// on every machine, a time is not. A long sheet name makes its address text once for each batch,
// and a selector name is looked up once for each batch, with no fold for an exact name. A locked
// cell refuses in the planner with no search over the batch, and the protection of a worksheet
// and the lock of each cell format are read once for each batch. A store refusal that the planner
// does not make is found by a bounded search, and the writes through two worksheet entries that
// name one part keep their order.

import { describe, expect, test } from 'bun:test';
import { buildXlsx, row } from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx, WorkbookDocumentStore } from '../../store/index.ts';
import {
  MAX_BATCH_OPERATIONS,
  type AutomationBatchResponse,
  type AutomationHost,
  type AutomationOperation,
} from '../index.ts';
import { createStoreAutomationHost } from '../host.ts';
import { hostWorkTestRecorder } from '../host-memo.ts';

const xf = (id: number) => `<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0" xfId="0"/>`;

function open(bytes: Uint8Array): AutomationHost {
  const read = readXlsx(bytes);
  if (!read.ok) throw new Error('the test workbook does not open');
  return createStoreAutomationHost(new WorkbookDocumentStore(read.pkg));
}

const data = (rows: string) =>
  open(buildXlsx({ sheets: [{ name: 'Data', rows }], styles: { cellXfs: [xf(10)] } }));

const num = (value: number) => ({ type: 'number', value });

const read = (address: string, values: boolean, numberFormat: boolean): AutomationOperation =>
  ({ op: 'readRange', sheet: { name: 'Data' }, address, values, numberFormat }) as never;

const repeated = (operation: AutomationOperation) =>
  Array.from({ length: MAX_BATCH_OPERATIONS }, () => operation);

/** The value of each result, or the refusal. */
function results(response: AutomationBatchResponse): unknown[] {
  expect(response.ok).toBe(true);
  return response.results.map((result) => (result.status === 'ok' ? result.value : result));
}

const okValue = (response: AutomationBatchResponse): unknown => results(response)[0];

const valuesOf = (response: AutomationBatchResponse): unknown =>
  (okValue(response) as { values?: unknown } | undefined)?.values;

/** Attributes of another namespace on one element: `count` of them, each a few bytes. */
const manyAttributes = (count: number) =>
  `xmlns:q="urn:q" ${Array.from({ length: count }, (_, i) => `q:a${i}=""`).join(' ')}`;

describe('the reads of one batch', () => {
  test('10,000 reads of a value split in 200,000 text nodes resolve it once', () => {
    // Text and CDATA sections in turn: the reader keeps each one as its own text node.
    const PAIRS = 100_000;
    const host = data(row(1, [`<c r="A1" t="str"><v>${'a<![CDATA[b]]>'.repeat(PAIRS)}</v></c>`]));
    const recorder = hostWorkTestRecorder();
    recorder.reset();
    const values = results(host.execute({ operations: repeated(read('A1', true, false)) }));
    expect(values).toHaveLength(MAX_BATCH_OPERATIONS);
    expect(values[0]).toEqual({ address: 'Data!A1', values: [['ab'.repeat(PAIRS)]] });
    expect(values.at(-1)).toEqual(values[0]);
    expect(recorder.values).toBe(1);
  });

  test('10,000 reads of a cell with 400,000 attributes read its value and format once', () => {
    const host = data(row(1, [`<c r="A1" s="1" ${manyAttributes(400_000)}><v>2</v></c>`]));
    const recorder = hostWorkTestRecorder();
    recorder.reset();
    const values = results(host.execute({ operations: repeated(read('A1', true, true)) }));
    expect(values.at(-1)).toEqual({
      address: 'Data!A1',
      values: [[2]],
      numberFormat: [['0.00%']],
    });
    expect([recorder.values, recorder.styles, recorder.codes]).toEqual([1, 1, 1]);
  });

  test('one read of a row of 16,384 empty cells reads the row format once', () => {
    const attrs = `s="1" customFormat="1" ${manyAttributes(400_000)}`;
    const host = data(row(1, ['<c r="A1"><v>1</v></c>'], attrs));
    const recorder = hostWorkTestRecorder();
    recorder.reset();
    const [range] = results(host.execute({ operations: [read('B1:XFD1', false, true)] })) as {
      numberFormat: string[][];
    }[];
    expect(range!.numberFormat).toEqual([Array.from({ length: 16_383 }, () => '0.00%')]);
    // Each empty cell takes the row format: one read of the row, one cell format resolved.
    expect([recorder.rows, recorder.styles, recorder.codes]).toEqual([1, 0, 1]);
  });

  test('10,000 format reads in a row with 400,000 attributes read the row format once', () => {
    const attrs = `s="1" customFormat="1" ${manyAttributes(400_000)}`;
    const host = data(row(1, ['<c r="A1"><v>1</v></c>'], attrs));
    const recorder = hostWorkTestRecorder();
    recorder.reset();
    const values = results(host.execute({ operations: repeated(read('B1', false, true)) }));
    expect(values.at(-1)).toEqual({ address: 'Data!B1', numberFormat: [['0.00%']] });
    expect([recorder.rows, recorder.codes]).toEqual([1, 1]);
  });

  test('two rows with their own row formats give each empty cell the format of its row', () => {
    const rows =
      row(1, ['<c r="A1"><v>1</v></c>'], 's="1" customFormat="1"') +
      row(2, ['<c r="A2"><v>2</v></c>'], 's="2" customFormat="1"');
    const host = open(
      buildXlsx({ sheets: [{ name: 'Data', rows }], styles: { cellXfs: [xf(10), xf(14)] } })
    );
    const recorder = hostWorkTestRecorder();
    recorder.reset();
    expect(results(host.execute({ operations: [read('B1:B2', false, true)] }))).toEqual([
      { address: 'Data!B1:B2', numberFormat: [['0.00%'], ['mm-dd-yy']] },
    ]);
    // The memo keeps the format of each row element: two rows, two reads.
    expect(recorder.rows).toBe(2);
  });

  test('each batch reads again, and a cell and its row keep their own formats', () => {
    const host = data(row(1, ['<c r="A1" s="0"><v>1</v></c>'], 's="1" customFormat="1"'));
    const recorder = hostWorkTestRecorder();
    recorder.reset();
    for (let batch = 0; batch < 2; batch += 1)
      expect(results(host.execute({ operations: [read('A1:B2', true, true)] }))).toEqual([
        {
          address: 'Data!A1:B2',
          values: [
            [1, ''],
            ['', ''],
          ],
          numberFormat: [
            ['General', '0.00%'],
            ['General', 'General'],
          ],
        },
      ]);
    // Two batches: the cell, its format index, the row, and the two cell formats, twice.
    expect([recorder.values, recorder.styles, recorder.rows, recorder.codes]).toEqual([2, 2, 2, 4]);
  });
});

describe('the sheet name of an address', () => {
  test('10,000 reads of a long quoted sheet name make its address text once', () => {
    const name = "a'b".repeat(5_000);
    const host = open(buildXlsx({ sheets: [{ name, rows: row(1, ['<c r="A1"><v>1</v></c>']) }] }));
    const recorder = hostWorkTestRecorder();
    recorder.reset();
    const address = { op: 'readRange', sheet: { active: true }, address: 'B2' } as never;
    const values = results(host.execute({ operations: repeated(address) }));
    expect(values.at(-1)).toEqual({ address: `'${"a''b".repeat(5_000)}'!B2` });
    expect(recorder.sheets).toBe(1);
  });
});

describe('the worksheet of a selector name', () => {
  // The open path does not limit the length of `x:sheet/@name`, and the fold of a long name
  // without ASCII letters takes about 2 ms. Without the memo, 10,000 reads by name take 20 s.
  const name = 'ä'.repeat(20_000);
  const host = () =>
    open(buildXlsx({ sheets: [{ name, rows: row(1, ['<c r="A1"><v>1</v></c>']) }] }));
  const byName = (selector: string): AutomationOperation =>
    ({ op: 'readRange', sheet: { name: selector }, address: 'A1', values: true }) as never;

  test('10,000 reads by one name look it up once, and each other spelling once', () => {
    const recorder = hostWorkTestRecorder();
    recorder.reset();
    const upper = 'Ä'.repeat(20_000);
    const operations = [...repeated(byName(name)).slice(2), byName(upper), byName(upper)];
    const values = results(host().execute({ operations }));
    expect(values.at(-1)).toEqual(values[0]);
    expect((values[0] as { values: unknown }).values).toEqual([[1]]);
    // The exact name needs no fold; only the other spelling is folded.
    expect([recorder.names, recorder.folds]).toEqual([2, 1]);
  });

  test('the list, a read by id, and a read by the exact name fold no name', () => {
    // A fold of every worksheet name in each batch blocks for seconds when a name has millions
    // of letters, so only a name with no exact match is folded.
    const workbook = host();
    const recorder = hostWorkTestRecorder();
    recorder.reset();
    const [list] = results(workbook.execute({ operations: [{ op: 'listWorksheets' } as never] }));
    const id = (list as { id: string }[])[0]!.id;
    const byId = { op: 'readRange', sheet: { id }, address: 'A1', values: true } as never;
    for (let batch = 0; batch < 3; batch += 1)
      expect(results(workbook.execute({ operations: [byId, byName(name)] }))).toHaveLength(2);
    expect(recorder.folds).toBe(0);
  });

  test('two worksheets that differ only by case: the exact name wins, else none', () => {
    const workbook = open(
      buildXlsx({
        sheets: [
          { name: 'Data', rows: row(1, ['<c r="A1"><v>1</v></c>']) },
          { name: 'DATA', rows: row(1, ['<c r="A1"><v>2</v></c>']) },
        ],
      })
    );
    const get = (selector: string) =>
      ({ op: 'getWorksheet', sheet: { name: selector }, orNull: true }) as never;
    const found = results(
      workbook.execute({ operations: [get('DATA'), get('Data'), get('data')] })
    );
    expect(found.map((sheet) => (sheet as { name: string } | null)?.name ?? null)).toEqual([
      'DATA',
      'Data',
      null,
    ]);
  });

  test('a name that the workbook does not hold refuses on each use, and folds once', () => {
    const recorder = hostWorkTestRecorder();
    recorder.reset();
    const missing = { op: 'getWorksheet', sheet: { name: 'Other' }, orNull: true } as never;
    expect(results(host().execute({ operations: [missing, missing, missing] }))).toEqual([
      null,
      null,
      null,
    ]);
    expect(recorder.names).toBe(1);
  });
});

describe('a locked cell of a protected worksheet', () => {
  // Column A is unlocked by its column format, every other cell is locked.
  const unlocked =
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">' +
    '<protection locked="0"/></xf>';
  const guarded = () =>
    open(
      buildXlsx({
        sheets: [
          {
            name: 'Data',
            rows: '',
            cols: '<cols><col min="1" max="1" width="9" style="1"/></cols>',
            beforeMerges: '<sheetProtection sheet="1"/>',
          },
        ],
        styles: { cellXfs: [unlocked] },
      })
    );
  const write = (address: string) =>
    ({ op: 'writeValues', sheet: { name: 'Data' }, address, values: [[num(1)]] }) as never;
  const column = Array.from({ length: MAX_BATCH_OPERATIONS - 1 }, (_, i) => write(`A${i + 1}`));

  test('refuses in the planner, with its cell, and the store makes no search', () => {
    const host = guarded();
    const recorder = hostWorkTestRecorder();
    recorder.reset();
    const response = host.execute({ operations: [...column, write('B1')] });
    expect(response.ok).toBe(false);
    const at = response.results.findIndex((result) => result.status === 'error');
    const result = response.results[at];
    expect([at, result?.status === 'error' ? result.error : undefined]).toMatchObject([
      MAX_BATCH_OPERATIONS - 1,
      { code: 'transaction-refused', detail: 'locked', cell: 'B1' },
    ]);
    expect(recorder.searches).toBe(0);
    // Column A takes the unlocked column format, and B1 format 0: one lock read for each.
    expect(recorder.locks).toBe(2);
    // Nothing applied: column A is still empty.
    expect(valuesOf(host.execute({ operations: [read('A1', true, false)] }))).toEqual([['']]);
  });

  test('takes the writes to the unlocked column, and reads the protection once', () => {
    const host = guarded();
    const recorder = hostWorkTestRecorder();
    recorder.reset();
    expect(host.execute({ operations: [...column, write('A10000')] }).ok).toBe(true);
    // One scan of the worksheet root for `x:sheetProtection` holds for each write of the batch,
    // and one read of the lock of the column format holds for each new cell.
    expect([recorder.guarded, recorder.locks]).toEqual([1, 1]);
    expect(valuesOf(host.execute({ operations: [read('A10000', true, false)] }))).toEqual([[1]]);
  });

  test('the protection of one worksheet does not answer for another', () => {
    const host = open(
      buildXlsx({
        sheets: [
          {
            name: 'Data',
            rows: '',
            cols: '<cols><col min="1" max="1" width="9" style="1"/></cols>',
            beforeMerges: '<sheetProtection sheet="1"/>',
          },
          { name: 'Open', rows: '' },
        ],
        styles: { cellXfs: [unlocked] },
      })
    );
    const recorder = hostWorkTestRecorder();
    recorder.reset();
    const toOpen = { op: 'writeValues', sheet: { name: 'Open' }, address: 'B1' };
    const operations = [write('A1'), { ...toOpen, values: [[num(1)]] } as never];
    expect(host.execute({ operations }).ok).toBe(true);
    expect(recorder.guarded).toBe(2);
  });

  test('9,999 writes to a cell with 50,000 attributes read its format and lock once', () => {
    const host = open(
      buildXlsx({
        sheets: [
          {
            name: 'Data',
            rows: row(1, [`<c r="A1" ${manyAttributes(50_000)} s="1"/>`]),
            beforeMerges: '<sheetProtection sheet="1"/>',
          },
        ],
        styles: { cellXfs: [unlocked] },
      })
    );
    const recorder = hostWorkTestRecorder();
    recorder.reset();
    // An unknown last operation ends the batch after the planner, before the store, whose own
    // check of each write is older work.
    const writes = Array.from({ length: MAX_BATCH_OPERATIONS - 1 }, () => write('A1'));
    const response = host.execute({ operations: [...writes, { op: 'unknown' } as never] });
    expect(response.results.at(-1)).toMatchObject({
      status: 'error',
      error: { code: 'unknown-operation' },
    });
    expect([recorder.styles, recorder.locks]).toEqual([1, 1]);
  });
});

describe('a store refusal that the planner does not make', () => {
  test('is found at its write by a search of at most 3 store passes', () => {
    // Two `x:sheet` entries named `Dup`: the planner finds the part by its id, and the store,
    // which finds a worksheet by name, refuses the write. No planner check makes this refusal.
    const host = open(
      buildXlsx({
        sheets: [
          { name: 'Data', rows: '' },
          { name: 'Dup', rows: '' },
        ],
        extraSheets: '<sheet name="Dup" sheetId="9" r:id="rId2"/>',
      })
    );
    const [list] = results(host.execute({ operations: [{ op: 'listWorksheets' } as never] }));
    const dup = (list as { id: string; name: string }[])[1]!;
    expect(dup.name).toBe('Dup');
    const data = (address: string) =>
      ({ op: 'writeValues', sheet: { name: 'Data' }, address, values: [[num(1)]] }) as never;
    const toDup = { op: 'writeValues', sheet: { id: dup.id }, address: 'A1', values: [[num(1)]] };
    const recorder = hostWorkTestRecorder();
    recorder.reset();
    const response = host.execute({
      operations: [data('A1'), data('A2'), data('A3'), toDup as never, data('A4')],
    });
    expect(response.ok).toBe(false);
    const at = response.results.findIndex((result) => result.status === 'error');
    expect([at, response.results[at]]).toMatchObject([
      3,
      { error: { code: 'transaction-refused', detail: 'unknown-sheet' } },
    ]);
    // Five writes: a binary search over the prefixes takes at most 3 passes.
    expect(recorder.searches).toBeLessThanOrEqual(3);
    expect(valuesOf(host.execute({ operations: [read('A1', true, false)] }))).toEqual([['']]);
  });

  test('is found at each place in the batch', () => {
    const host = open(
      buildXlsx({
        sheets: [
          { name: 'Data', rows: '' },
          { name: 'Dup', rows: '' },
        ],
        extraSheets: '<sheet name="Dup" sheetId="9" r:id="rId2"/>',
      })
    );
    const [list] = results(host.execute({ operations: [{ op: 'listWorksheets' } as never] }));
    const id = (list as { id: string }[])[1]!.id;
    const data = { op: 'writeValues', sheet: { name: 'Data' }, address: 'A1', values: [[num(1)]] };
    const toDup = { op: 'writeValues', sheet: { id }, address: 'A1', values: [[num(1)]] };
    for (let at = 0; at < 4; at += 1) {
      const operations = Array.from({ length: 4 }, (_, i) => (i === at ? toDup : data) as never);
      const response = host.execute({ operations });
      expect([at, response.results.findIndex((result) => result.status === 'error')]).toEqual([
        at,
        at,
      ]);
    }
  });
});

describe('two worksheet entries that name one part', () => {
  const twoNames = () =>
    open(
      buildXlsx({
        sheets: [{ name: 'A', rows: '' }],
        extraSheets: '<sheet name="B" sheetId="9" r:id="rId1"/>',
      })
    );

  test('the writes through both names keep their order', () => {
    const host = twoNames();
    const write = (name: string, value: number): AutomationOperation =>
      ({ op: 'writeValues', sheet: { name }, address: 'A1', values: [[num(value)]] }) as never;
    expect(host.execute({ operations: [write('B', 1), write('A', 2), write('B', 3)] }).ok).toBe(
      true
    );
    for (const name of ['A', 'B'])
      expect(
        valuesOf(
          host.execute({
            operations: [
              { op: 'readRange', sheet: { name }, address: 'A1', values: true } as never,
            ],
          })
        )
      ).toEqual([[3]]);
  });

  test('the list gives both names, and the id of the part names the first entry', () => {
    const host = twoNames();
    const list = okValue(host.execute({ operations: [{ op: 'listWorksheets' } as never] })) as {
      id: string;
      name: string;
    }[];
    expect(list.map((sheet) => sheet.name)).toEqual(['A', 'B']);
    expect(list[1]!.id).toBe(list[0]!.id);
    const byId = { op: 'getWorksheet', sheet: { id: list[0]!.id } } as never;
    expect(okValue(host.execute({ operations: [byId] }))).toEqual(list[0]);
  });
});
