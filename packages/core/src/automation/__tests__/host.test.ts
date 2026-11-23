// The automation host over a store, after PPTX Editor's `host.test.ts` and
// `host-untrusted.test.ts` in `packages/core/src/automation/__tests__/` at d4502552: the batch
// rule, the refusal of one operation, the revision check, the limits, and untrusted requests.

import { describe, expect, test } from 'bun:test';
import {
  buildXlsx,
  c,
  row,
  type FixtureWorkbook,
} from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx, WorkbookDocumentStore } from '../../store/index.ts';
import {
  MAX_ARRAY_RANGES,
  MAX_BATCH_CELLS,
  MAX_BATCH_OPERATIONS,
  MAX_BATCH_WRITE_CELLS,
  openAutomationHost,
  type AutomationBatchResponse,
  type AutomationHost,
  type AutomationOperation,
} from '../index.ts';
import { createStoreAutomationHost } from '../host.ts';

function workbook(dataRows = '', beforeMerges?: string): Uint8Array {
  return buildXlsx({
    sheets: [
      {
        name: 'Data',
        rows:
          row(1, [c('A1', 42), c('B1', 'text'), c('C1', 3, { f: 'A1/14' })]) +
          row(2, [c('A2', '#DIV/0!', { t: 'e' }), c('B2', true)]) +
          dataRows,
        ...(beforeMerges ? { beforeMerges } : {}),
      },
      { name: 'Other', rows: '' },
      { name: 'data', state: 'hidden', rows: '' },
    ],
  });
}

/** A host over a workbook of the fixture builder. */
function openBuilt(book: FixtureWorkbook) {
  return open(buildXlsx(book));
}

function open(bytes = workbook()): { host: AutomationHost; store: WorkbookDocumentStore } {
  const read = readXlsx(bytes);
  if (!read.ok) throw new Error('the test workbook does not open');
  const store = new WorkbookDocumentStore(read.pkg);
  return { host: createStoreAutomationHost(store), store };
}

const read = (address: string, sheet = 'Data'): AutomationOperation => ({
  op: 'readRange',
  sheet: { name: sheet },
  address,
  values: true,
});

const readFormats = (address: string): AutomationOperation => ({
  op: 'readRange',
  sheet: { name: 'Data' },
  address,
  numberFormat: true,
});

const write = (address: string, value: number | string, sheet = 'Data'): AutomationOperation => ({
  op: 'writeValues',
  sheet: { name: sheet },
  address,
  values: [[typeof value === 'number' ? { type: 'number', value } : { type: 'string', value }]],
});

function valueAt(response: AutomationBatchResponse, at: number): unknown {
  const result = response.results[at];
  return result?.status === 'ok' ? (result.value as { values?: unknown }).values : undefined;
}

function formatsAt(response: AutomationBatchResponse): unknown {
  const result = response.results[0];
  return result?.status === 'ok' ? (result.value as { numberFormat?: unknown }).numberFormat : 0;
}

/** The failing operation of a refused batch, and its error. */
function refusal(response: AutomationBatchResponse) {
  expect(response.ok).toBe(false);
  expect(response.changed).toBe(false);
  const at = response.results.findIndex((r) => r.status === 'error');
  for (const [i, other] of response.results.entries())
    if (i !== at) expect(other.status).toBe('skipped');
  const result = response.results[at];
  return { at, error: result?.status === 'error' ? result.error : undefined };
}

describe('reads', () => {
  test('a read gives the stored values, an error as its text, and "" for an empty cell', () => {
    const { host } = open();
    const response = host.execute({ operations: [read('A1:B3')] });
    expect(valueAt(response, 0)).toEqual([
      [42, 'text'],
      ['#DIV/0!', true],
      ['', ''],
    ]);
    const result = response.results[0];
    expect(result?.status === 'ok' && (result.value as { address: string }).address).toBe(
      'Data!A1:B3'
    );
  });

  test('a read of a formula cell, or of an array formula range, is refused', () => {
    const array = row(5, ['<c r="A5"><f t="array" ref="A5:B6">1</f><v>1</v></c>']);
    const { host } = open(workbook(array));
    expect(refusal(host.execute({ operations: [read('C1')] })).error?.detail).toBe('formula-cell');
    expect(refusal(host.execute({ operations: [read('B6')] })).error?.detail).toBe('formula-cell');
    expect(valueAt(host.execute({ operations: [read('C5')] }), 0)).toEqual([['']]);
  });

  test('an array range that cannot be read refuses each value read and write of its worksheet', () => {
    for (const ref of ['A1:ZZZZ9', 'ZZZZ1:B6', 'B5:B6:B9']) {
      const array = row(5, [`<c r="B5"><f t="array" ref="${ref}">1</f><v>1</v></c>`]);
      const { host, store } = open(workbook(array));
      for (const operation of [read('A1'), write('A1', 1)]) {
        const { error } = refusal(host.execute({ operations: [operation] }));
        expect(error?.detail).toBe('formula-ranges');
      }
      expect(store.revision).toBe(0);
    }
  });

  test('a stored value or a sheet that the host cannot read is refused', () => {
    const cases: [string, string][] = [
      ['<c r="A1" t="e"><v>#FOO!</v></c>', 'unreadable-value'],
      ['<c r="A1" t="s"><v>99</v></c>', 'unreadable-value'],
      ['<c r="A2"><v>5</v></c>', 'unreadable-sheet'],
    ];
    for (const [cell, detail] of cases) {
      const { host } = openBuilt({ sheets: [{ name: 'Data', rows: row(1, [cell]) }] });
      expect(refusal(host.execute({ operations: [read('A1:A2')] })).error?.detail).toBe(detail);
    }
    const date = row(1, ['<c r="A1" t="d"><v>1904-01-03T00:00:00</v></c>']);
    const { host } = openBuilt({ sheets: [{ name: 'Data', rows: date }], date1904: true });
    expect(valueAt(host.execute({ operations: [read('A1')] }), 0)).toEqual([[2]]);
  });

  test('a number format read gives the stored code, General, or a refusal', () => {
    const xf = (id: number) =>
      `<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`;
    const styles = {
      numFmts: ['<numFmt numFmtId="164" formatCode="0.000"/>'],
      cellXfs: [xf(164), xf(31), xf(200), xf(10)],
    };
    const formats = (cells: string[], address: string, cols?: string) => {
      const sheet = { name: 'Data', rows: row(1, cells), ...(cols ? { cols } : {}) };
      return openBuilt({ sheets: [sheet], styles }).host.execute({
        operations: [readFormats(address)],
      });
    };
    const stored = formats(['<c r="A1" s="1"><v>1</v></c>', '<c r="B1"><v>1</v></c>'], 'A1:B1');
    expect(formatsAt(stored)).toEqual([['0.000', 'General']]);
    const column = '<cols><col min="3" max="3" width="9" style="4" customWidth="1"/></cols>';
    expect(formatsAt(formats([], 'C1', column))).toEqual([['0.00%']]);
    // A built-in id whose code depends on the locale, an id with no code, and a style index out
    // of range.
    for (const s of [2, 3, 99]) {
      const response = formats([`<c r="A1" s="${s}"><v>1</v></c>`], 'A1');
      expect(refusal(response).error?.detail).toBe('unreadable-format');
    }
  });

  test('a worksheet with more array ranges than the limit refuses its value reads', () => {
    const cells: string[] = [];
    for (let i = 0; i <= MAX_ARRAY_RANGES; i += 1)
      cells.push(`<c r="D${i + 10}"><f t="array" ref="D${i + 10}">1</f><v>1</v></c>`);
    const { host } = open(workbook(cells.map((cell, i) => row(i + 10, [cell])).join('')));
    expect(refusal(host.execute({ operations: [read('A1')] })).error?.detail).toBe(
      'formula-ranges'
    );
  });

  test('a read after a write in the same batch is refused, and an address read is not', () => {
    const { host } = open();
    const response = host.execute({ operations: [write('A1', 1), read('A1')] });
    expect(refusal(response)).toMatchObject({ at: 1, error: { detail: 'read-after-write' } });
    const formats = host.execute({ operations: [write('A1', 1), readFormats('A1')] });
    expect(refusal(formats)).toMatchObject({ at: 1, error: { detail: 'read-after-write' } });
    const address = host.execute({
      operations: [write('A1', 1), { op: 'readRange', sheet: { name: 'Data' }, address: 'A1' }],
    });
    expect(address.ok).toBe(true);
  });
});

describe('worksheets', () => {
  test('a name matches without case, and names that differ only in case match exactly', () => {
    const { host } = open();
    const get = (name: string) =>
      host.execute({ operations: [{ op: 'getWorksheet', sheet: { name } }] }).results[0];
    expect(get('Other')).toMatchObject({ status: 'ok', value: { name: 'Other' } });
    expect(get('OTHER')).toMatchObject({ status: 'ok', value: { name: 'Other' } });
    expect(get('data')).toMatchObject({ status: 'ok', value: { name: 'data' } });
    expect(get('DATA')).toMatchObject({ status: 'error', error: { code: 'item-not-found' } });
    const orNull = host.execute({
      operations: [{ op: 'getWorksheet', sheet: { name: 'Nope' }, orNull: true }],
    });
    expect(orNull.results[0]).toEqual({ status: 'ok', value: null });
  });

  test('the active worksheet is the active tab, or the first visible one', () => {
    const active = (activeTab: number) => {
      const sheets = [
        { name: 'Data', rows: '' },
        { name: 'Other', rows: '' },
        { name: 'Hidden', state: 'hidden' as const, rows: '' },
      ];
      const { host } = openBuilt({ sheets, activeTab });
      const response = host.execute({
        operations: [{ op: 'getWorksheet', sheet: { active: true } }],
      });
      return response.results[0];
    };
    expect(active(1)).toMatchObject({ status: 'ok', value: { name: 'Other' } });
    expect(active(2)).toMatchObject({ status: 'ok', value: { name: 'Data' } });
  });

  test('an address quotes a sheet name that reads as a reference, a boolean, or has a quote', () => {
    const cases = [
      ['A1', "'A1'!B2"],
      ['R1C1', "'R1C1'!B2"],
      ['TRUE', "'TRUE'!B2"],
      ["it's", "'it''s'!B2"],
      ['Plain_1', 'Plain_1!B2'],
    ];
    for (const [name, address] of cases) {
      const { host } = openBuilt({ sheets: [{ name: name!, rows: '' }] });
      const response = host.execute({
        operations: [{ op: 'readRange', sheet: { name: name! }, address: 'B2' }],
      });
      expect(response.results[0]).toEqual({ status: 'ok', value: { address } });
    }
  });

  test('the list gives the worksheets in tab order with ids that a lookup finds', () => {
    const { host } = open();
    const list = host.execute({ operations: [{ op: 'listWorksheets' }] }).results[0];
    const reads = (list?.status === 'ok' ? list.value : []) as { id: string; name: string }[];
    expect(reads.map((r) => r.name)).toEqual(['Data', 'Other', 'data']);
    const found = host.execute({
      operations: [{ op: 'getWorksheet', sheet: { id: reads[1]!.id } }],
    });
    expect(found.results[0]).toMatchObject({ status: 'ok', value: { name: 'Other' } });
  });
});

describe('writes', () => {
  test('writes commit in one transaction, save, and open again with their values', () => {
    const { host, store } = open();
    const before = store.revision;
    const response = host.execute({
      operations: [
        write('A1', 7),
        write('D1', ' new text '),
        {
          op: 'writeValues',
          sheet: { name: 'Data' },
          address: 'A2:B2',
          values: [[{ type: 'empty' }, { type: 'boolean', value: false }]],
        },
      ],
    });
    expect(response).toMatchObject({ ok: true, changed: true, revision: before + 1 });
    expect(host.revision()).toBe(before + 1);
    const saved = host.save();
    if (!saved.ok) throw new Error('save failed');
    const again = open(saved.bytes).host;
    expect(valueAt(again.execute({ operations: [read('A1:D2')] }), 0)).toBeUndefined();
    expect(valueAt(again.execute({ operations: [read('A1:B2')] }), 0)).toEqual([
      [7, 'text'],
      ['', false],
    ]);
    expect(valueAt(again.execute({ operations: [read('D1')] }), 0)).toEqual([[' new text ']]);
  });

  test('a write to a formula cell or into an array range is refused, and nothing applies', () => {
    const array = row(5, ['<c r="A5"><f t="array" ref="A5:B6">1</f><v>1</v></c>']);
    const { host, store } = open(workbook(array));
    const before = store.revision;
    expect(refusal(host.execute({ operations: [write('A1', 1), write('C1', 1)] }))).toMatchObject({
      at: 1,
      error: { code: 'unsupported', detail: 'formula-cell' },
    });
    expect(refusal(host.execute({ operations: [write('B6', 1)] })).error?.detail).toBe(
      'formula-cell'
    );
    expect(store.revision).toBe(before);
  });

  test('a store refusal names the write that it refused', () => {
    const { host } = open(workbook('', '<sheetProtection sheet="1"/>'));
    const operations = [write('A1', 1, 'Other'), write('A1', 1), write('B1', 1, 'Other')];
    const { at, error } = refusal(host.execute({ operations }));
    expect(at).toBe(1);
    expect(error).toMatchObject({ code: 'transaction-refused', detail: 'locked', cell: 'A1' });
  });

  test('a write that the store takes one cell at a time is refused before the store runs', () => {
    // `mc:Choice/@Requires` names a prefix, which must be declared (ECMA-376 Part 3, 7.6).
    const alternate =
      '<c r="A5"><v>2</v><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/' +
      'markup-compatibility/2006" xmlns:x14="urn:example:x14"><mc:Choice Requires="x14">' +
      '<e:x xmlns:e="urn:example:extension">keep</e:x></mc:Choice></mc:AlternateContent></c>';
    const { host, store } = open(workbook(row(5, [alternate])));
    let transactions = 0;
    const transact = store.transact.bind(store);
    store.transact = (...args) => ((transactions += 1), transact(...args));
    const many: AutomationOperation = {
      op: 'writeValues',
      sheet: { name: 'Data' },
      address: `E1:E${MAX_BATCH_WRITE_CELLS - 1}`,
      values: Array.from({ length: MAX_BATCH_WRITE_CELLS - 1 }, (_, i) => [
        { type: 'number', value: i } as const,
      ]),
    };
    const long = write('B1', 'x'.repeat(32_768));
    expect(refusal(host.execute({ operations: [many, long, write('A1', 1)] }))).toMatchObject({
      at: 1,
      error: { code: 'transaction-refused', detail: 'text-too-long' },
    });
    expect(refusal(host.execute({ operations: [write('A1', 1), write('A5', 2)] }))).toMatchObject({
      at: 1,
      error: { code: 'transaction-refused', detail: 'cell-alternate-content' },
    });
    expect(transactions).toBe(0);
  });

  test('a batch planned against an older revision is refused', () => {
    const { host } = open();
    const first = host.execute({ operations: [write('A1', 1)] });
    const stale = host.execute({ operations: [write('A1', 2)], expectedRevision: 0 });
    expect(first.revision).toBe(1);
    expect(refusal(stale).error?.code).toBe('stale-revision');
  });
});

describe('limits and untrusted requests', () => {
  test('a batch at each limit runs', () => {
    const { host } = open();
    const list = Array.from({ length: MAX_BATCH_OPERATIONS }, () => ({ op: 'listWorksheets' }));
    expect(host.execute({ operations: list as AutomationOperation[] }).ok).toBe(true);
    const reads = Array.from({ length: MAX_BATCH_CELLS / 100_000 }, () => read('A3:A100002'));
    expect(host.execute({ operations: reads }).ok).toBe(true);
    const rows = MAX_BATCH_WRITE_CELLS;
    const values = Array.from({ length: rows }, () => [{ type: 'number', value: 1 }] as const);
    const full = { op: 'writeValues', sheet: { name: 'Data' }, address: `E1:E${rows}`, values };
    expect(host.execute({ operations: [full as AutomationOperation] }).ok).toBe(true);
  });

  test('each limit refuses the batch', () => {
    const { host } = open();
    const many = Array.from({ length: MAX_BATCH_OPERATIONS + 1 }, () => read('A1'));
    expect(refusal(host.execute({ operations: many })).error?.detail).toBe('operations');
    const big = Array.from({ length: MAX_BATCH_CELLS / 100_000 + 1 }, () => read('A1:A100000'));
    expect(refusal(host.execute({ operations: big })).error?.detail).toBe('batch-cells');
    expect(refusal(host.execute({ operations: [read('A1:B100000')] })).error?.detail).toBe(
      'range-cells'
    );
    const rows = MAX_BATCH_WRITE_CELLS + 1;
    const wide = {
      op: 'writeValues',
      sheet: { name: 'Data' },
      address: `E1:E${rows}`,
      values: Array.from({ length: rows }, () => [null]),
    } as const;
    expect(refusal(host.execute({ operations: [wide] })).error?.detail).toBe('write-cells');
    const sheets = Array.from({ length: 11 }, (_, i) => ({ name: `S${i}`, rows: '' }));
    const list = Array.from({ length: MAX_BATCH_OPERATIONS }, () => ({ op: 'listWorksheets' }));
    const listed = openBuilt({ sheets }).host.execute({
      operations: list as AutomationOperation[],
    });
    expect(refusal(listed).error?.detail).toBe('batch-worksheets');
  });

  test('bytes over a reader limit open as resource-limit, and other bad bytes as invalid', () => {
    const deep = '<a:x xmlns:a="urn:x">'.repeat(300) + '</a:x>'.repeat(300);
    const beforeMerges = `<extLst><ext uri="x">${deep}</ext></extLst>`;
    const bytes = buildXlsx({ sheets: [{ name: 'Data', rows: '', beforeMerges }] });
    expect(openAutomationHost(bytes)).toEqual({
      ok: false,
      code: 'resource-limit',
      reason: 'open.xml-depth-limit',
      limit: 'xml.maxDepth',
    });
    expect(openAutomationHost(new Uint8Array([1, 2, 3]))).toEqual({
      ok: false,
      code: 'invalid-argument',
      reason: 'open.zip-corrupt',
    });
  });

  test('a malformed operation, selector, address, or value refuses with its code', () => {
    const { host } = open();
    const cases: [unknown, string][] = [
      [{ op: 'deleteWorksheet' }, 'unknown-operation'],
      [{ op: 'getWorksheet', sheet: { name: 'Data', id: 'x' } }, 'invalid-argument'],
      [{ ...read('A1'), address: 'Data!A1' }, 'unsupported'],
      [{ ...read('A1'), address: 'A1:XFE1' }, 'invalid-argument'],
      [{ ...write('A1', 1), values: [[{ type: 'formula', text: '1+1' }]] }, 'invalid-argument'],
      [{ ...write('A1', 1), values: [[{ type: 'number', value: Infinity }]] }, 'invalid-argument'],
      [{ ...write('A1', 1), values: [[{ type: 'string', value: 5 }]] }, 'invalid-argument'],
      [{ ...write('A1', 1), values: [[{ type: 'boolean', value: 'false' }]] }, 'invalid-argument'],
      [{ ...write('A1', 1), values: [[{ value: 1 }]] }, 'invalid-argument'],
      [{ ...write('A1', 1), address: 'A1B' }, 'unsupported'],
      [{ ...write('A1:B1', 1), values: [[null]] }, 'invalid-argument'],
    ];
    for (const [operation, code] of cases)
      expect(
        refusal(host.execute({ operations: [operation as AutomationOperation] })).error?.code
      ).toBe(code as never);
    const notList = host.execute({ operations: 'x' } as never);
    expect(refusal(notList).error?.code).toBe('invalid-argument');
  });

  test('a disposed host refuses each batch and each save', () => {
    const { host } = open();
    host.dispose();
    host.dispose();
    expect(refusal(host.execute({ operations: [read('A1')] })).error?.code).toBe('disposed');
    expect(host.save()).toMatchObject({ ok: false, error: { code: 'disposed' } });
  });
});
