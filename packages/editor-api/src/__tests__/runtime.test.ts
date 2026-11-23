/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See LICENSE.md in this package.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The server runtime: reads as stored, value writes through the store, load and sync rules,
// refusals with stable codes, the run lifetime, and save and reopen.

import { describe, expect, spyOn, test } from 'bun:test';
import { zipSync } from 'fflate';
import { buildXlsx, c, row } from '../../../../scripts/fixtures/builder.ts';
import { XlsxEditor, type XlsxEditorServerRuntime } from '../index.ts';
import { resetDevWarnings } from '../runtime/dev-warnings.ts';
import { expectNoEcho, expectRefusal, expectThrows, sampleWorkbook } from './support.ts';

async function open(bytes = sampleWorkbook()): Promise<XlsxEditorServerRuntime> {
  return XlsxEditor.createServer(bytes);
}

/** The values of `address` on `sheet`, read in a run of their own. */
async function valuesOf(runtime: XlsxEditorServerRuntime, sheet: string, address: string) {
  return runtime.run(async (context) => {
    const range = context.workbook.worksheets.getItem(sheet).getRange(address).load('values');
    await context.sync();
    return range.values;
  });
}

/** Run `edit` on the range `address` of `Data`, then sync. */
async function edit(
  runtime: XlsxEditorServerRuntime,
  address: string,
  set: (range: ReturnType<typeof rangeOf>) => void
): Promise<void> {
  await runtime.run(async (context) => {
    set(context.workbook.worksheets.getItem('Data').getRange(address));
    await context.sync();
  });
}
type Context = Parameters<Parameters<XlsxEditorServerRuntime['run']>[0]>[0];
const rangeOf = (context: Context) => context.workbook.worksheets.getItem('Data').getRange('A1');

describe('reading', () => {
  test('lists every worksheet in tab order, hidden ones included', async () => {
    const runtime = await open();
    const names = await runtime.run(async (context) => {
      const sheets = context.workbook.worksheets.load('items/name');
      await context.sync();
      return sheets.items.map((sheet) => sheet.name);
    });
    expect(names).toEqual(['Data', 'My Sheet', 'Hidden']);
  });

  test('reads values as stored: numbers, strings, booleans, errors, and empty cells', async () => {
    const runtime = await open();
    expect(await valuesOf(runtime, 'Data', 'A1:F1')).toEqual([
      [42, 'shared text', 'inline text', true, '#DIV/0!', ''],
    ]);
    // A date cell gives its serial number.
    expect(await valuesOf(runtime, 'Data', 'A3:B3')).toEqual([[45000, '']]);
  });

  test('reads the number format code of each cell, and the address with its sheet', async () => {
    const runtime = await open();
    const read = await runtime.run(async (context) => {
      const range = context.workbook.worksheets.getItem('data').getRange('$a$2:a3');
      const quoted = context.workbook.worksheets.getItem('My Sheet').getRange('A1');
      range.load(['address', 'numberFormat']);
      context.load(quoted, { select: 'address' });
      await context.sync();
      return { address: range.address, formats: range.numberFormat, quoted: quoted.address };
    });
    expect(read).toEqual({
      address: 'Data!A2:A3',
      formats: [['0.00%'], ['mm-dd-yy']],
      quoted: "'My Sheet'!A1",
    });
  });

  test('the active worksheet is the one the workbook opens on', async () => {
    const runtime = await open();
    const name = await runtime.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet().load('name');
      await context.sync();
      return sheet.name;
    });
    expect(name).toBe('Data');
  });

  test('a lookup that finds nothing refuses, or answers a null object', async () => {
    const runtime = await open();
    const error = await expectRefusal(
      runtime.run(async (context) => {
        context.workbook.worksheets.getItem('Secret Name').load('name');
        await context.sync();
      }),
      'ItemNotFound'
    );
    expectNoEcho(error, ['Secret Name']);
    await runtime.run(async (context) => {
      const sheet = context.workbook.worksheets.getItemOrNullObject('Nope');
      expectThrows(() => sheet.isNullObject, 'PropertyNotLoaded');
      await context.sync();
      expect(sheet.isNullObject).toBe(true);
      expectThrows(() => sheet.name, 'PropertyNotLoaded');
    });
  });

  test('a read of a formula cell refuses, and names no cell text', async () => {
    const runtime = await open();
    const error = await expectRefusal(
      valuesOf(runtime, 'Data', 'A2:B2'),
      'NotSupported',
      'range.formula-cell'
    );
    expectNoEcho(error, ['A1/3.5']);
  });

  test('a property that no load asked for throws PropertyNotLoaded', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const range = rangeOf(context).load('address');
      await context.sync();
      expectThrows(() => range.values, 'PropertyNotLoaded');
    });
  });
});

describe('writing', () => {
  test('numbers, booleans, text, and clears save and open again', async () => {
    const runtime = await open();
    await edit(runtime, 'A1:D1', (range) => {
      range.values = [[7, false, 'new text', '']];
    });
    await edit(runtime, 'F5', (range) => {
      range.values = [['a note']];
    });
    const again = await open(await runtime.save());
    expect(await valuesOf(again, 'Data', 'A1:D1')).toEqual([[7, false, 'new text', '']]);
    expect(await valuesOf(again, 'Data', 'F5')).toEqual([['a note']]);
    // A null leaves its cell as it is.
    await edit(again, 'A1:B1', (range) => {
      range.values = [[null, 3]];
    });
    expect(await valuesOf(again, 'Data', 'A1:B1')).toEqual([[7, 3]]);
  });

  test('a sync is all or nothing', async () => {
    const runtime = await open();
    await expectRefusal(
      runtime.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem('Data');
        sheet.getRange('A1').values = [[1]];
        sheet.getRange('B2').values = [[2]];
        await context.sync();
      }),
      'NotSupported',
      'range.formula-cell'
    );
    expect(await valuesOf(runtime, 'Data', 'A1')).toEqual([[42]]);
  });

  test('formulas, text that reads as a value, and number formats are refused', async () => {
    const runtime = await open();
    const cases: [unknown, string][] = [
      ['=SUM(A1:A3)', 'write.formula-not-supported'],
      ['12', 'write.text-reads-as-value'],
      ['TRUE', 'write.text-reads-as-value'],
      ['-x', 'write.text-reads-as-value'],
      ['+x', 'write.text-reads-as-value'],
      ["'x", 'write.text-reads-as-value'],
      ['#N/A', 'write.text-reads-as-value'],
      ['12%', 'write.text-reads-as-value'],
      ['2026-01-31', 'write.text-reads-as-value'],
      [' 12 ', 'write.text-reads-as-value'],
      ['true ', 'write.text-reads-as-value'],
    ];
    for (const [value, diagnostic] of cases)
      await expectRefusal(
        edit(runtime, 'G1', (range) => {
          range.values = [[value]];
        }),
        'NotSupported',
        diagnostic
      );
    expect(await valuesOf(runtime, 'Data', 'G1')).toEqual([['']]);
  });

  test('a value array of the wrong shape or type refuses with InvalidArgument', async () => {
    const runtime = await open();
    const bad: [unknown, string][] = [
      [[[1]], 'write.array-shape'],
      [[[1, 2, 3]], 'write.array-shape'],
      [[[{}, 1]], 'argument.not-accepted'],
      [[[NaN, 1]], 'argument.not-accepted'],
      [[[Infinity, 1]], 'argument.not-accepted'],
    ];
    for (const [values, diagnostic] of bad)
      await expectRefusal(
        edit(runtime, 'A1:B1', (range) => {
          range.values = values as never;
        }),
        'InvalidArgument',
        diagnostic
      );
    await expectRefusal(
      edit(runtime, 'A1:B1', (range) => {
        range.values = 5 as never;
      }),
      'NotSupported',
      'write.single-value'
    );
  });

  test('the store refusals keep their public codes', async () => {
    const protectedSheet = { name: 'Data', rows: row(1, [c('A1', 1)]) };
    const locked = await open(
      buildXlsx({ sheets: [{ ...protectedSheet, beforeMerges: '<sheetProtection sheet="1"/>' }] })
    );
    await expectRefusal(
      edit(locked, 'A1', (range) => {
        range.values = [[2]];
      }),
      'NotSupported',
      'edit.locked'
    );
    await expectRefusal(
      edit(await open(), 'G1', (range) => {
        range.values = [['x'.repeat(32_768)]];
      }),
      'InvalidArgument',
      'write.text-too-long'
    );
  });

  test('a write after another run changed the workbook refuses with StaleDocument', async () => {
    const runtime = await open();
    let release!: () => void;
    const gate = new Promise<void>((resolve) => (release = resolve));
    const first = runtime.run(async (context) => {
      const range = rangeOf(context).load('values');
      await context.sync();
      await gate;
      range.values = [[1]];
      await context.sync();
    });
    await edit(runtime, 'A1', (range) => {
      range.values = [[2]];
    });
    release();
    await expectRefusal(first, 'StaleDocument');
    expect(await valuesOf(runtime, 'Data', 'A1')).toEqual([[2]]);
  });

  for (const order of ['a load and then a write', 'a write and then a load'])
    test(`a sync with ${order} after another run changed the workbook is stale`, async () => {
      // A sync with a load and a write still sends the revision that the run read.
      const runtime = await open();
      let release!: () => void;
      const gate = new Promise<void>((resolve) => (release = resolve));
      const first = runtime.run(async (context) => {
        rangeOf(context).load('values');
        await context.sync();
        await gate;
        const load = () =>
          context.workbook.worksheets.getItem('Data').getRange('B3').load('address');
        if (order === 'a load and then a write') load();
        rangeOf(context).values = [[1]];
        if (order === 'a write and then a load') load();
        await context.sync();
      });
      await edit(runtime, 'A1', (range) => {
        range.values = [[2]];
      });
      release();
      await expectRefusal(first, 'StaleDocument');
      expect(await valuesOf(runtime, 'Data', 'A1')).toEqual([[2]]);
    });

  test('a sync with nothing queued keeps the revision that a later write checks', async () => {
    const runtime = await open();
    let release!: () => void;
    const gate = new Promise<void>((resolve) => (release = resolve));
    const first = runtime.run(async (context) => {
      rangeOf(context).load('values');
      await context.sync();
      await gate;
      await context.sync();
      rangeOf(context).values = [[1]];
      await context.sync();
    });
    await edit(runtime, 'A1', (range) => {
      range.values = [[2]];
    });
    release();
    await expectRefusal(first, 'StaleDocument');
    expect(await valuesOf(runtime, 'Data', 'A1')).toEqual([[2]]);
  });

  test('a set makes the loaded value out of date', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const loaded = rangeOf(context).load('values');
      await context.sync();
      loaded.values = [[5]];
      expectThrows(() => loaded.values, 'PropertyNotLoaded');
      await context.sync();
      // A set after a load that no sync sent yet leaves the value not loaded after the sync.
      const queued = context.workbook.worksheets.getItem('Data').getRange('G1').load('values');
      queued.values = [[6]];
      await context.sync();
      expectThrows(() => queued.values, 'PropertyNotLoaded');
    });
  });

  test('a refused sync empties the queue, and a run drops what no sync sent', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      rangeOf(context).values = [['=1']];
      await expectRefusal(context.sync(), 'NotSupported');
      await context.sync();
    });
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      await runtime.run(async (context) => {
        context.workbook.worksheets.getItem('Data').getRange('G1').values = [[1]];
      });
      expect(warn).toHaveBeenCalledTimes(1);
    } finally {
      warn.mockRestore();
    }
    expect(await valuesOf(runtime, 'Data', 'A1:G1')).toEqual([
      [42, 'shared text', 'inline text', true, '#DIV/0!', '', ''],
    ]);
  });

  test('a load of values after a write in the same sync refuses', async () => {
    const runtime = await open();
    await expectRefusal(
      runtime.run(async (context) => {
        const range = rangeOf(context);
        range.values = [[1]];
        range.load('values');
        await context.sync();
      }),
      'NotSupported',
      'sync.read-after-write'
    );
  });
});

describe('load options and lifetime', () => {
  test('a load option outside the subset refuses', async () => {
    const runtime = await open();
    const options: [unknown, string, string?][] = [
      [{ values: true }, 'NotSupported', 'load.property-tree-not-supported'],
      [{ select: 'values', top: 1 }, 'InvalidArgument', 'argument.not-accepted'],
      [{ select: 'values', skip: 1 }, 'InvalidArgument', 'argument.not-accepted'],
      [{ select: 'values', expand: 'x' }, 'NotSupported', 'load.option-not-supported'],
      [{ select: 'values', expand: ['x'] }, 'NotSupported', 'load.option-not-supported'],
      [{ select: 'values', other: 1 }, 'InvalidArgument', 'argument.not-accepted'],
      ['values/text', 'InvalidArgument'],
      ['formulas', 'InvalidArgument'],
      [7, 'InvalidArgument'],
    ];
    for (const [option, code, diagnostic] of options) {
      const error = await expectRefusal(
        runtime.run(async (context) => {
          rangeOf(context).load(option as never);
          await context.sync();
        }),
        code as never,
        diagnostic
      );
      if (diagnostic) expect(error.debugInfo.innerError?.message).toMatch(/select|Use /);
      expectNoEcho(error, ['other']);
    }
  });

  test('an address with a sheet name, a whole column, or no address refuses', async () => {
    const runtime = await open();
    for (const address of ['Data!A1', 'A:A', undefined])
      await expectRefusal(
        runtime.run(async (context) => {
          context.workbook.worksheets.getItem('Data').getRange(address).load('values');
          await context.sync();
        }),
        'NotSupported'
      );
  });

  test('sync answers its pass-through value, with nothing queued and after a batch', async () => {
    const runtime = await open();
    const tag = { tag: 'kept' };
    await runtime.run(async (context) => {
      expect(await context.sync(tag)).toBe(tag);
      const range = rangeOf(context).load('values');
      expect(await context.sync('after a load')).toBe('after a load');
      expect(range.values).toEqual([[42]]);
      expect(await context.sync()).toBeUndefined();
    });
  });

  test('a proxy of an ended run, and a disposed runtime, refuse', async () => {
    const runtime = await open();
    const kept = await runtime.run(async (context) => rangeOf(context));
    expectThrows(() => kept.load('values'), 'InvalidObjectPath');
    expectThrows(() => kept.values, 'InvalidObjectPath');
    runtime.dispose();
    runtime.dispose();
    await expectRefusal(
      runtime.run(async () => 1),
      'RuntimeDisposed'
    );
    await expectRefusal(runtime.save(), 'RuntimeDisposed');
  });

  test('bytes that are not a workbook, or that exceed a reader limit, refuse', async () => {
    const corrupt = await expectRefusal(
      open(new Uint8Array([1, 2, 3])),
      'InvalidArgument',
      'open.zip-corrupt'
    );
    expect(corrupt.limit).toBeUndefined();
    expect(corrupt.target).toBe('createServer');
    const deep = '<a:x xmlns:a="urn:x">'.repeat(300) + '</a:x>'.repeat(300);
    const beforeMerges = `<extLst><ext uri="x">${deep}</ext></extLst>`;
    const bytes = buildXlsx({ sheets: [{ name: 'Data', rows: '', beforeMerges }] });
    const limited = await expectRefusal(
      open(bytes),
      'ResourceLimitExceeded',
      'open.xml-depth-limit'
    );
    expect(limited.limit).toBe('xml.maxDepth');
    expect(limited.debugInfo.innerError?.message).toContain('fixed');
  });

  test('each refused open names its diagnostic, and a hint with no file content', async () => {
    const text = (value: string) => new TextEncoder().encode(value);
    const types =
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="xml" ContentType="application/xml"/></Types>';
    const cases: [Uint8Array, string][] = [
      [zipSync({ 'secret.txt': text('x') }), 'open.content-types-invalid'],
      [zipSync({ '[Content_Types].xml': text(types) }), 'open.missing-workbook'],
      [zipSync({ '../secret.xml': text('x') }), 'open.zip-invalid-entry-name'],
      [zipSync({ '[Content_Types].xml': text('<Types') }), 'open.xml-malformed'],
      [
        zipSync({ '[Content_Types].xml': text('<!DOCTYPE Types []><Types/>') }),
        'open.xml-dtd-forbidden',
      ],
    ];
    for (const [bytes, diagnostic] of cases) {
      const error = await expectRefusal(open(bytes), 'InvalidArgument', diagnostic);
      expect(error.limit).toBeUndefined();
      expect(error.debugInfo.innerError?.message.length).toBeGreaterThan(0);
      expectNoEcho(error, ['secret']);
    }
  });
});
