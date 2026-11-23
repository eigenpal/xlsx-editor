/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See LICENSE.md in this package.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The stated rules of the runtime, one test each: the stale check of a sync, the public code of
// each host refusal, a set that makes a loaded value out of date, the queue after a refusal and
// after the run, `toJSON`, the warning for writes that no sync sent, null objects, the names an
// error never echoes, and the argument checks of each call.

import { describe, expect, spyOn, test } from 'bun:test';
import { zipSync } from 'fflate';
import { buildXlsx, c, row } from '../../../../scripts/fixtures/builder.ts';
import {
  XlsxEditor,
  type XlsxEditorError,
  type XlsxEditorErrorCode,
  type XlsxEditorServerRuntime,
} from '../index.ts';
import { resetDevWarnings } from '../runtime/dev-warnings.ts';
import { expectNoEcho, expectRefusal, expectThrows, sampleWorkbook } from './support.ts';

type Context = Parameters<Parameters<XlsxEditorServerRuntime['run']>[0]>[0];

const open = (bytes = sampleWorkbook()) => XlsxEditor.createServer(bytes);
const data = (context: Context) => context.workbook.worksheets.getItem('Data');

async function valuesOf(runtime: XlsxEditorServerRuntime, address: string): Promise<unknown> {
  return runtime.run(async (context) => {
    const range = data(context).getRange(address).load('values');
    await context.sync();
    return range.values;
  });
}

describe('the stale check of a sync', () => {
  test('a write after another run wrote refuses with StaleDocument and both revisions', async () => {
    const runtime = await open();
    const error = await expectRefusal(
      runtime.run(async (context) => {
        const range = data(context).getRange('A1').load('values');
        await context.sync();
        await runtime.run(async (other) => {
          data(other).getRange('A1').values = [[1]];
          await other.sync();
        });
        range.values = [[(range.values[0]![0] as number) + 1]];
        await context.sync();
      }),
      'StaleDocument'
    );
    expect(error).toMatchObject({ expectedRevision: 0, actualRevision: 1 });
    expect(await valuesOf(runtime, 'A1')).toEqual([[1]]);
  });

  test('a sync that only reads is never stale, and its revision is the one a write checks', async () => {
    const runtime = await open();
    const values = await runtime
      .run(async (context) => {
        await runtime.run(async (other) => {
          data(other).getRange('A1').values = [[5]];
          await other.sync();
        });
        const range = data(context).getRange('A1').load('values');
        await context.sync();
        range.values = [[6]];
        await context.sync();
        return range.values;
      })
      .catch((error: unknown) => error);
    // The set made the loaded value out of date, so the read after the sync refuses.
    expect(values).toMatchObject({ code: 'PropertyNotLoaded' });
    expect(await valuesOf(runtime, 'A1')).toEqual([[6]]);
  });

  test('a read-only sync after another writer is not stale', async () => {
    const runtime = await open();
    const values = await runtime.run(async (context) => {
      data(context).getRange('A1').load('address');
      await context.sync();
      await runtime.run(async (other) => {
        data(other).getRange('A1').values = [[5]];
        await other.sync();
      });
      const range = data(context).getRange('A1').load('values');
      await context.sync();
      return range.values;
    });
    expect(values).toEqual([[5]]);
  });

  test('a first sync that writes is not stale', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      await runtime.run(async (other) => {
        data(other).getRange('A1').values = [[5]];
        await other.sync();
      });
      data(context).getRange('A1').values = [[7]];
      await context.sync();
    });
    expect(await valuesOf(runtime, 'A1')).toEqual([[7]]);
  });
});

describe('the public code of each refusal', () => {
  test('a limit is ResourceLimitExceeded, at one past its value', async () => {
    const runtime = await open();
    const column = (n: number) => Array.from({ length: n }, (_, i) => [i]);
    await runtime.run(async (context) => {
      data(context).getRange('G1:G10000').values = column(10_000);
      await context.sync();
    });
    await expectRefusal(
      runtime.run(async (context) => {
        data(context).getRange('G1:G10000').values = column(10_000);
        data(context).getRange('H1').values = [[1]];
        await context.sync();
      }),
      'ResourceLimitExceeded',
      'sync.too-many-written-cells'
    );
    await runtime.run(async (context) => {
      const range = data(context).getRange('A1');
      // The getItem lookup and 9,999 loads: exactly 10,000 operations.
      for (let i = 0; i < 9_999; i += 1) range.load('address');
      await context.sync();
    });
    await expectRefusal(
      runtime.run(async (context) => {
        const range = data(context).getRange('A1');
        for (let i = 0; i <= 10_000; i += 1) range.load('address');
        await context.sync();
      }),
      'ResourceLimitExceeded',
      'sync.too-many-operations'
    );
    await expectRefusal(
      runtime.run(async (context) => {
        data(context).getRange('A1:B50001').load('address');
        await context.sync();
      }),
      'ResourceLimitExceeded',
      'range.too-many-cells'
    );
  });

  test('bytes over a cap of the bounded parse refuse with ResourceLimitExceeded', async () => {
    const entries: Record<string, Uint8Array> = {};
    for (let i = 0; i <= 10_000; i += 1) entries[`x/${i}.bin`] = new Uint8Array(0);
    const error = await expectRefusal(
      open(zipSync(entries, { level: 0 })),
      'ResourceLimitExceeded',
      'open.zip-entry-limit'
    );
    expect(error.limit).toBe('zip.maxEntries');
  });

  test('a protected cell, and a text longer than a cell holds, refuse with their codes', async () => {
    const runtime = await open();
    await expectRefusal(
      runtime.run(async (context) => {
        data(context).getRange('G1').values = [['x'.repeat(32_768)]];
        await context.sync();
      }),
      'InvalidArgument',
      'write.text-too-long'
    );
    const locked = await open(sampleWorkbook({ protectData: true }));
    await expectRefusal(
      locked.run(async (context) => {
        data(context).getRange('G1').values = [[1]];
        await context.sync();
      }),
      'NotSupported',
      'edit.locked'
    );
  });

  test('a locked cell refusal names the first locked cell and says what to do', async () => {
    // A1 and B1 take a format that is not locked; C1 is new, so it takes the locked default.
    const unlocked =
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">' +
      '<protection locked="0"/></xf>';
    const runtime = await open(
      buildXlsx({
        styles: { cellXfs: [unlocked] },
        sheets: [
          {
            name: 'Data',
            rows: row(1, [c('A1', 1, { s: 1 }), c('B1', 2, { s: 1 })]),
            beforeMerges: '<sheetProtection sheet="1"/>',
          },
        ],
      })
    );
    const error = await expectRefusal(
      runtime.run(async (context) => {
        data(context).getRange('A1:C1').values = [[5, 6, 7]];
        await context.sync();
      }),
      'NotSupported',
      'edit.locked'
    );
    expect(error.debugInfo.innerError?.message).toBe(
      'The worksheet protects the cell, so it cannot be changed. Remove the protection of the ' +
        'worksheet, or leave the locked cell out of the write. First refused cell: C1.'
    );
    // The cells that are not locked take the write.
    await runtime.run(async (context) => {
      data(context).getRange('A1:B1').values = [[5, 6]];
      await context.sync();
    });
    expect(await valuesOf(runtime, 'A1:B1')).toEqual([[5, 6]]);
  });

  test('a value in a cell that a merged range hides refuses, and names the cell', async () => {
    const runtime = await open(
      buildXlsx({
        sheets: [
          {
            name: 'Data',
            rows: row(1, [c('A1', 'title')]) + row(3, [c('B3', 1)]),
            merges: ['A1:C1', 'B3:C4'],
          },
        ],
      })
    );
    const write = (address: string, values: unknown[][]) =>
      runtime.run(async (context) => {
        data(context).getRange(address).values = values;
        await context.sync();
      });
    const hintOf = async (address: string, values: unknown[][]) =>
      (await expectRefusal(write(address, values), 'NotSupported', 'write.merged-cell')).debugInfo
        .innerError?.message;
    expect(await hintOf('A1:C1', [['new', 'hidden', null]])).toBe(
      'A merged range covers the cell and hides it, because the cell is not the top-left cell ' +
        'of the merged range. Write the value to the top-left cell, or give null or an empty ' +
        'string for this cell. First refused cell: B1.'
    );
    // A write that starts inside a merged range, and one that meets only its lower part.
    expect(await hintOf('C1', [[false]])).toEndWith(' First refused cell: C1.');
    expect(await hintOf('A4:D4', [[null, null, 2, null]])).toEndWith(' First refused cell: C4.');
    expect(await valuesOf(runtime, 'A1:C4')).toEqual([
      ['title', '', ''],
      ['', '', ''],
      ['', 1, ''],
      ['', '', ''],
    ]);
    // The top-left cell takes a value, and null or "" leaves the hidden cells empty.
    await write('A1:D4', [
      ['new', '', null, 5],
      [null, null, null, null],
      [null, 2, '', null],
      [null, null, null, null],
    ]);
    expect(await valuesOf(runtime, 'A1:D3')).toEqual([
      ['new', '', '', 5],
      ['', '', '', ''],
      ['', 2, '', ''],
    ]);
  });

  test('a refused text write names its cell, from the top-left cell of the address', async () => {
    const runtime = await open();
    const hintOf = async (address: string, values: unknown[][]) =>
      (
        await expectRefusal(
          runtime.run(async (context) => {
            data(context).getRange(address).values = values;
            await context.sync();
          }),
          'NotSupported',
          values.flat().includes('=1') ? 'write.formula-not-supported' : 'write.text-reads-as-value'
        )
      ).debugInfo.innerError?.message;
    // A write back of a loaded row with an error value.
    expect(await hintOf('A1:E1', [[1, 'a', '#DIV/0!', 2, 3]])).toEndWith(
      ' First refused cell: C1.'
    );
    expect(await hintOf('G1', [['=1']])).toEndWith(' First refused cell: G1.');
    // Either corner and `$` marks name the same rectangle.
    expect(
      await hintOf('$C$3:b2', [
        [1, 2],
        [3, 'TRUE'],
      ])
    ).toEndWith(' First refused cell: C3.');
    expect(await hintOf('XFD1048576', [['12%']])).toEndWith(' First refused cell: XFD1048576.');
    // An address that is not a plain cell or rectangle names no cell.
    const named = await hintOf('Data!A1', [['TRUE']]);
    expect(named).not.toContain('First refused cell');
  });

  test('a value that the API cannot write names its cell', async () => {
    const runtime = await open();
    const hintOf = async (address: string, values: unknown[][]) =>
      (
        await expectRefusal(
          runtime.run(async (context) => {
            data(context).getRange(address).values = values;
            await context.sync();
          }),
          'InvalidArgument',
          'argument.not-accepted'
        )
      ).debugInfo.innerError?.message;
    const accepted = 'Use a finite number, a boolean, a string, or null for each cell.';
    expect(
      await hintOf('F1:G2', [
        ['a', 'b'],
        ['c', Number.NaN],
      ])
    ).toBe(`${accepted} First refused cell: G2.`);
    expect(await hintOf('B5:D5', [[1, undefined, 2]])).toBe(`${accepted} First refused cell: C5.`);
    expect(await hintOf('Data!A1', [[new Date(0)]])).toBe(accepted);
  });
});

describe('cells that hold content in another form', () => {
  const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
  const book = () =>
    buildXlsx({
      sheets: [
        {
          name: 'Data',
          rows: row(1, [
            `<c r="A1"><mc:AlternateContent xmlns:mc="${MC}"><mc:Fallback><v>3</v>` +
              '</mc:Fallback></mc:AlternateContent></c>',
            '<c r="B1"><v>2</v><e:w xmlns:e="urn:example:e"><v>3</v></e:w></c>',
            c('C1', 1),
          ]),
        },
      ],
    });

  test('a value only in a wrapper refuses its read, so a read and write back keeps it', async () => {
    const runtime = await open(book());
    const error = await expectRefusal(
      valuesOf(runtime, 'A1:C1'),
      'NotSupported',
      'read.unreadable-value'
    );
    // The hint names the first cell that the read cannot give.
    expect(error.debugInfo.innerError?.message).toEndWith(' First refused cell: A1.');
    expect(error.message).not.toContain('A1.');
    expect(await valuesOf(runtime, 'C1')).toEqual([[1]]);
  });

  test('a string whose text is in a wrapper refuses its read, inline or shared', async () => {
    const wrapper = `<mc:AlternateContent xmlns:mc="${MC}"><mc:Fallback><t>b</t></mc:Fallback></mc:AlternateContent>`;
    const runtime = await open(
      buildXlsx({
        rawStrings: [`<si><t>a</t><r>${wrapper}</r></si>`],
        sheets: [
          {
            name: 'Data',
            rows: row(1, [
              `<c r="A1" t="inlineStr"><is><t>a</t>${wrapper}</is></c>`,
              c('B1', 0, { t: 's' }),
              c('C1', 1),
            ]),
          },
        ],
      })
    );
    for (const address of ['A1', 'B1'])
      await expectRefusal(valuesOf(runtime, address), 'NotSupported', 'read.unreadable-value');
    expect(await valuesOf(runtime, 'C1')).toEqual([[1]]);
  });

  test('a write to a cell whose content the edit cannot rewrite refuses with its code', async () => {
    const runtime = await open(book());
    await expectRefusal(
      runtime.run(async (context) => {
        data(context).getRange('B1:C1').values = [[5, 6]];
        await context.sync();
      }),
      'NotSupported',
      'edit.alternate-content'
    );
    expect(await valuesOf(runtime, 'C1')).toEqual([[1]]);
  });
});

describe('loaded values and the queue', () => {
  test('a set after a load makes the loaded value out of date, before and after the sync', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const range = data(context).getRange('A1').load('values');
      range.values = [[5]];
      await context.sync();
      expectThrows(() => range.values, 'PropertyNotLoaded');
      range.load('values');
      await context.sync();
      expect(range.values).toEqual([[5]]);
      range.values = [[6]];
      expectThrows(() => range.values, 'PropertyNotLoaded');
      await context.sync();
      expectThrows(() => range.values, 'PropertyNotLoaded');
    });
  });

  test('a refused sync empties the queue, so a later sync sends none of it', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      data(context).getRange('G1').values = [[1]];
      data(context).getRange('G2').values = [['=1']];
      await expectRefusal(context.sync(), 'NotSupported', 'write.formula-not-supported');
      await context.sync();
    });
    expect(await valuesOf(runtime, 'G1:G2')).toEqual([[''], ['']]);
  });

  test('writes that no sync sent do not apply, and a development build warns once', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const runtime = await open();
      await runtime.run(async (context) => {
        data(context).getRange('G1').values = [[1]];
        data(context).getRange('G2').values = [['=1']];
      });
      expect(await valuesOf(runtime, 'G1')).toEqual([['']]);
      expect(warn).toHaveBeenCalledTimes(1);
      const message = String(warn.mock.calls[0]![0]);
      expect(message).toContain('2 queued writes');
      expect(message).toContain('NotSupported');
      // A run that throws drops its writes with no warning.
      await runtime
        .run(async (context) => {
          data(context).getRange('G1').values = [[1]];
          throw new Error('stop');
        })
        .catch(() => undefined);
      expect(warn).toHaveBeenCalledTimes(1);
    } finally {
      warn.mockRestore();
      resetDevWarnings();
    }
  });

  test('numberFormat is read-only: a set fails to compile and throws, and keeps the value', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const range = data(context).getRange('A1').load('numberFormat');
      await context.sync();
      expect(range.numberFormat).toEqual([['General']]);
      expect(() => {
        // @ts-expect-error The package publishes numberFormat read-only.
        range.numberFormat = [['0.00']];
      }).toThrow(TypeError);
      expect(range.numberFormat).toEqual([['General']]);
      await context.sync();
    });
    const after = await runtime.run(async (context) => {
      const range = data(context).getRange('A1').load('numberFormat');
      await context.sync();
      return range.numberFormat;
    });
    expect(after).toEqual([['General']]);
  });

  test('a set copies its array, and the caller may change a loaded array', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const grid = [[1]];
      const range = data(context).getRange('G1');
      range.values = grid;
      grid[0]![0] = 2;
      await context.sync();
      range.load('values');
      await context.sync();
      expect(range.values).toEqual([[1]]);
      range.values[0]![0] = 9;
      expect(range.values).toEqual([[9]]);
      expectThrows(() => range.address, 'PropertyNotLoaded');
    });
    expect(await valuesOf(runtime, 'G1')).toEqual([[1]]);
  });

  test('a change to the bytes after createServer does not reach the workbook', async () => {
    const bytes = sampleWorkbook();
    const runtime = await open(bytes);
    bytes.fill(0);
    expect(await valuesOf(runtime, 'A1')).toEqual([[42]]);
    const again = await open(await runtime.save());
    expect(await valuesOf(again, 'A1')).toEqual([[42]]);
  });

  test('toJSON gives the loaded properties as plain data', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const range = data(context).getRange('A1:B1').load(['address', 'values']);
      const sheets = context.workbook.worksheets.load('items/name');
      await context.sync();
      const json = range.toJSON() as { address: string; values: unknown[][] };
      expect(json).toEqual({ address: 'Data!A1:B1', values: [[42, 'shared text']] });
      json.values[0]![0] = 0;
      expect(range.values[0]![0]).toBe(42);
      expect(sheets.toJSON()).toEqual({
        items: [{ name: 'Data' }, { name: 'My Sheet' }, { name: 'Hidden' }],
      });
    });
  });
});

describe('null objects', () => {
  test('a null object answers isNullObject after its sync, and has no name', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const sheet = context.workbook.worksheets.getItemOrNullObject('Nope');
      expectThrows(() => sheet.isNullObject, 'PropertyNotLoaded');
      await context.sync();
      expect(sheet.isNullObject).toBe(true);
      const error = (() => {
        try {
          return sheet.name;
        } catch (caught) {
          return caught;
        }
      })();
      expect(error).toMatchObject({
        code: 'PropertyNotLoaded',
        debugInfo: { innerError: { code: 'edit.null-object' } },
      });
      const found = context.workbook.worksheets.getItemOrNullObject('data').load('name');
      const missing = context.workbook.worksheets.getItemOrNullObject('Nope').load('name');
      await context.sync();
      expect(found.isNullObject).toBe(false);
      expect(found.name).toBe('Data');
      expect(missing.isNullObject).toBe(true);
      expectThrows(() => missing.name, 'PropertyNotLoaded');
    });
  });

  test('a range of a null object refuses, in the sync of the lookup and after it', async () => {
    const runtime = await open();
    await expectRefusal(
      runtime.run(async (context) => {
        context.workbook.worksheets.getItemOrNullObject('Nope').getRange('A1').values = [[1]];
        await context.sync();
      }),
      'InvalidObjectPath',
      'edit.null-object'
    );
    await runtime.run(async (context) => {
      const sheet = context.workbook.worksheets.getItemOrNullObject('Nope');
      await context.sync();
      sheet.getRange('A1').load('values');
      await expectRefusal(context.sync(), 'InvalidObjectPath', 'edit.null-object');
      sheet.getRange('A1').values = [[1]];
      await expectRefusal(context.sync(), 'InvalidObjectPath', 'edit.null-object');
    });
    expect(await valuesOf(runtime, 'A1')).toEqual([[42]]);
  });
});

describe('names that an error never echoes', () => {
  test('a worksheet name or an address from the call', async () => {
    const runtime = await open();
    const secret = 'Secret_Q4';
    const errors = [
      await expectRefusal(
        runtime.run(async (context) => {
          context.workbook.worksheets.getItem(secret).load('name');
          await context.sync();
        }),
        'ItemNotFound',
        'read.worksheet-not-found'
      ),
      await expectRefusal(
        runtime.run(async (context) => {
          data(context).getRange(`${secret}!A1`).load('values');
          await context.sync();
        }),
        'NotSupported'
      ),
      await expectRefusal(
        runtime.run(async (context) => {
          context.workbook.worksheets.getItem(secret).getRange('A1').values = [[1]];
          await context.sync();
        }),
        'ItemNotFound',
        'read.worksheet-not-found'
      ),
    ];
    for (const error of errors) expectNoEcho(error, [secret]);
  });

  test('a cell value and a worksheet name from the workbook', async () => {
    const runtime = await open();
    const error = await expectRefusal(
      runtime.run(async (context) => {
        context.workbook.worksheets.getItem('My Sheet').getRange('A1:A2').values = [['=x'], [1]];
        await context.sync();
      }),
      'NotSupported'
    );
    expectNoEcho(error, ['My Sheet', '=x']);
  });
});

/** The error that `call` throws at once. */
function thrown(call: () => unknown): XlsxEditorError {
  try {
    call();
  } catch (error) {
    return error as XlsxEditorError;
  }
  throw new Error('expected a throw');
}

describe('argument checks', () => {
  test('a load option outside the subset throws at the call', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const range = data(context).getRange('A1');
      const names = Array.from({ length: 65 }, () => 'address');
      const cases: [unknown, XlsxEditorErrorCode][] = [
        [{ select: 'values', expand: 'x' }, 'NotSupported'],
        [{ select: 'values', skip: 1 }, 'InvalidArgument'],
        [names, 'InvalidArgument'],
        [names.join(','), 'InvalidArgument'],
        ['a/b/c/d/e/f/g/h/i', 'InvalidArgument'],
        ['formulas', 'InvalidArgument'],
        ['items/name', 'InvalidArgument'],
      ];
      for (const [option, code] of cases) expectThrows(() => range.load(option as never), code);
      expectThrows(() => data(context).load('values'), 'InvalidArgument');
      expectThrows(() => context.workbook.worksheets.load('values'), 'InvalidArgument');
      expectThrows(() => context.workbook.worksheets.load('items/values'), 'InvalidArgument');
      let caught: unknown;
      try {
        context.workbook.load('worksheets');
      } catch (error) {
        caught = error;
      }
      expect(caught).toMatchObject({
        code: 'InvalidArgument',
        debugInfo: { innerError: { message: expect.stringContaining("load('items/name')") } },
      });
    });
  });

  test('a path through a name that is not a navigation property throws, and loads nothing', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const sheets = context.workbook.worksheets;
      const range = data(context).getRange('A1');
      const cases: [() => unknown, string, string][] = [
        [() => sheets.load('foo/bar'), 'workbook.worksheets.foo', "a path through 'items'"],
        [() => sheets.load('items/name/x'), 'workbook.worksheets.items.name', "'name'"],
        [
          () => range.load('format/font'),
          "workbook.worksheets.getItem().getRange('A1').format",
          "'address', 'values', or 'numberFormat'",
        ],
        [
          () => data(context).load('tables/items'),
          'workbook.worksheets.getItem().tables',
          "'name'",
        ],
      ];
      for (const [call, target, accepted] of cases) {
        const error = thrown(call);
        expect([error.code, error.target, error.debugInfo.errorLocation]).toEqual([
          'InvalidArgument',
          target,
          target,
        ]);
        expect(error.debugInfo.innerError?.code).toBe('argument.not-accepted');
        expect(error.message).toContain(accepted);
        expect(error.message).toContain(`(${target})`);
      }
      // A window of 0 is a window too: an object that is not a collection refuses it.
      for (const [option, key] of [
        [{ top: 0 }, 'top'],
        [{ skip: 0 }, 'skip'],
      ] as const) {
        const error = thrown(() => range.load(option as never));
        expect([error.code, error.target]).toEqual([
          'InvalidArgument',
          `workbook.worksheets.getItem().getRange('A1').${key}`,
        ]);
        expect(error.message).toContain(`${key} on a collection only`);
      }
      for (const object of [data(context), context.workbook])
        expectThrows(() => object.load({ top: 1 } as never), 'InvalidArgument');
      await context.sync();
      expectThrows(() => sheets.items, 'PropertyNotLoaded');
    });
  });

  test('top and skip load a window of the worksheet collection', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const windows = [
        [{ select: 'items/name', skip: 1, top: 1 }, ['My Sheet']],
        [{ select: 'items/name', skip: 1 }, ['My Sheet', 'Hidden']],
        [{ select: 'items/name', top: 2 }, ['Data', 'My Sheet']],
        [{ top: 0 }, []],
        [{ skip: 3 }, []],
        [{ skip: 9, top: 9 }, []],
        [{ select: 'name', expand: [] }, ['Data', 'My Sheet', 'Hidden']],
      ] as const;
      const sheets = context.workbook.worksheets;
      for (const [option, names] of windows) {
        sheets.load(option as never);
        await context.sync();
        expect([option, sheets.items.map((sheet) => sheet.name)]).toEqual([option, [...names]]);
      }
      // An item keeps its position in the workbook: the second worksheet is items[1].
      sheets.load({ skip: 1, top: 1 });
      await context.sync();
      const second = sheets.items[0]!;
      expect(thrown(() => second.load('nope')).target).toBe('workbook.worksheets.items[1].nope');
    });
  });

  test('a window that is not a whole number of 0 or more refuses at the call', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const sheets = context.workbook.worksheets;
      for (const [option, key] of [
        [{ top: -1 }, 'top'],
        [{ top: 1.5 }, 'top'],
        [{ top: '1' }, 'top'],
        [{ top: Number.POSITIVE_INFINITY }, 'top'],
        [{ skip: -1 }, 'skip'],
        [{ skip: Number.NaN }, 'skip'],
        [{ skip: 2 ** 53 }, 'skip'],
      ] as const) {
        const error = thrown(() => sheets.load(option as never));
        expect([error.code, error.target]).toEqual([
          'InvalidArgument',
          `workbook.worksheets.${key}`,
        ]);
        expect(error.message).toContain('a whole number of 0 or more');
      }
      const expand = thrown(() => sheets.load({ expand: ['items'] } as never));
      expect([expand.code, expand.target, expand.debugInfo.innerError?.code]).toEqual([
        'NotSupported',
        'workbook.worksheets.expand',
        'load.option-not-supported',
      ]);
    });
  });

  test("load('name') on the worksheet collection equals load('items/name')", async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const sheets = context.workbook.worksheets.load('name');
      await context.sync();
      expect(sheets.items.map((sheet) => sheet.name)).toEqual(['Data', 'My Sheet', 'Hidden']);
    });
  });

  test('a getItem worksheet is not a null object, also before the sync', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      expect(data(context).isNullObject).toBe(false);
    });
  });

  test('createServer keeps its own copy of the bytes', async () => {
    const bytes = sampleWorkbook();
    const runtime = await open(bytes);
    bytes.fill(0);
    expect(await valuesOf(runtime, 'A1')).toEqual([[42]]);
    expect((await runtime.save()).length).toBeGreaterThan(0);
  });

  test('load() with no argument loads every property', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const sheet = data(context).load();
      const range = data(context).getRange('A1:A3').load();
      await context.sync();
      expect(sheet.name).toBe('Data');
      expect(range.toJSON()).toEqual({
        address: 'Data!A1:A3',
        values: [[42], [0.25], [45000]],
        numberFormat: [['General'], ['0.00%'], ['mm-dd-yy']],
      });
    });
  });

  test('a written grid with a row count other than the range refuses', async () => {
    const runtime = await open();
    await expectRefusal(
      runtime.run(async (context) => {
        data(context).getRange('G1:H2').values = [[1, 2]];
        await context.sync();
      }),
      'InvalidArgument'
    );
    expect(await valuesOf(runtime, 'G1')).toEqual([['']]);
  });

  test("'items/$all' and $all load what they name", async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const sheets = context.workbook.worksheets.load('items/$all');
      const range = data(context).getRange('A1').load('$all');
      await context.sync();
      expect(sheets.items.map((s) => s.name)).toEqual(['Data', 'My Sheet', 'Hidden']);
      expect(range.toJSON()).toEqual({
        address: 'Data!A1',
        values: [[42]],
        numberFormat: [['General']],
      });
    });
  });

  test('text with a sign or an apostrophe mark refuses as text that reads as a value', async () => {
    const runtime = await open();
    for (const text of ['+1', "'abc", '-x', '+'])
      await expectRefusal(
        runtime.run(async (context) => {
          data(context).getRange('G1').values = [[text]];
          await context.sync();
        }),
        'NotSupported',
        'write.text-reads-as-value'
      );
  });

  test('an error value in any letter case refuses as text that reads as a value', async () => {
    // The engine reads `#n/a` as text, so only the write check refuses it.
    const runtime = await open();
    for (const text of ['#n/a', '#N/a', '#div/0!'])
      await expectRefusal(
        runtime.run(async (context) => {
          data(context).getRange('G1').values = [[text]];
          await context.sync();
        }),
        'NotSupported',
        'write.text-reads-as-value'
      );
  });

  test('a row that is not an array, and arguments of the wrong type, refuse', async () => {
    const runtime = await open();
    await expectRefusal(
      runtime.run(async (context) => {
        data(context).getRange('A1:A2').values = [[1], 2] as never;
        await context.sync();
      }),
      'InvalidArgument',
      'write.array-shape'
    );
    const address = await expectRefusal(
      runtime.run(async (context) => {
        data(context)
          .getRange(5 as never)
          .load('values');
        await context.sync();
      }),
      'InvalidArgument'
    );
    expect(address.target).toBe('workbook.worksheets.getItem().getRange');
    for (const call of [
      (context: Context) => context.workbook.worksheets.getItem(5 as never).load('name'),
      (context: Context) => context.workbook.worksheets.getItemOrNullObject({} as never),
      (context: Context) =>
        data(context)
          .getRange(5 as never)
          .load('values'),
    ])
      await expectRefusal(
        runtime.run(async (context) => {
          call(context);
          await context.sync();
        }),
        'InvalidArgument'
      );
    for (const whole of ['', undefined])
      await expectRefusal(
        runtime.run(async (context) => {
          data(context).getRange(whole).load('values');
          await context.sync();
        }),
        'NotSupported',
        'range.whole-sheet'
      );
    await expectRefusal(XlsxEditor.createServer('book.xlsx' as never), 'InvalidArgument');
    await expectRefusal(
      XlsxEditor.createServer(new Uint8Array(0)),
      'InvalidArgument',
      'argument.not-accepted'
    );
    await expectRefusal(runtime.run('callback' as never), 'InvalidArgument');
  });

  test('a context of an ended run, and an object of another run, refuse', async () => {
    const runtime = await open();
    const ended = await runtime.run(async (context) => context);
    expectThrows(() => ended.workbook, 'InvalidRequestContext');
    await expectRefusal(ended.sync(), 'InvalidRequestContext');
    await runtime.run(async (context) => {
      await runtime.run(async (other) => {
        const range = data(other).getRange('A1');
        expectThrows(() => context.load(range, 'values'), 'InvalidObjectPath');
        expectThrows(() => context.load({} as never, 'values'), 'InvalidArgument');
      });
    });
  });
});
