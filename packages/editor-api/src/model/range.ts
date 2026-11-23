/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See LICENSE.md in this package.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A range of cells, in the shape of the Office.js workbook object model: `address`, `values`,
// and `numberFormat`.
//
// VALUES AS STORED. A read gives the value that the file stores: a number, a string, a boolean,
// an error as its text (`#DIV/0!`), a date as its serial number, and "" for an empty cell. A
// cell that holds a formula, or a result of an array formula or a data table, refuses the read
// with `NotSupported`, because this API does not calculate formulas.
//
// A WRITE STORES WHAT IT IS GIVEN. A number, a boolean, `""` (empty the cell), `null` (leave the
// cell as it is), or text. Text that starts with `=` is a formula, which this API does not write,
// and text that a typed entry would read as another value (`12%`, `2026-01-31`, `TRUE`, `#N/A`,
// or text that starts with `+`, `-`, or `'`) is refused: this API never guesses a typed value.
// The check uses the typed entry rules of the engine (`parseCellInput`) in en-US, whatever the
// number format of the cell. `numberFormat` is read-only: an assignment fails to compile. In
// JavaScript strict-mode code (ES modules and classes) it throws a `TypeError`; in non-strict code
// it does nothing. DOCX Editor publishes a member that its subset does not write the same way.
//
// A SET MAKES THE LOADED VALUE OUT OF DATE. The setter forgets the loaded value of its property,
// so a read before the next load and sync is `PropertyNotLoaded`, never an old value.

import type {
  AutomationCellWrite,
  AutomationOperation,
  AutomationValue,
} from '@xlsx-editor.dev/core/automation';
import { EN_US } from '@xlsx-editor.dev/core/cascade';
import {
  MAX_COLS,
  MAX_ROWS,
  formatCellRef,
  parseCellInput,
  parseCellRef,
} from '@xlsx-editor.dev/core/editor';
import type { LoadRequest } from '../runtime/load-options.ts';
import { ClientObject, type RequestContext } from '../runtime/request-context.ts';
import { internals, type Worksheet } from './workbook.ts';

/**
 * An address that can name itself in a refusal target: letters, digits, `$`, and one `:`. Any
 * other caller text stays out of the target.
 */
const SHOWN_ADDRESS = /^\$?[A-Za-z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Za-z]{1,3}\$?[0-9]{1,7})?$/;

/**
 * The cell, as `A1`, at row `r` and column `c` of a write to `address`, or undefined when the
 * address is not a plain cell or rectangle.
 */
function cellAt(address: string, r: number, c: number): string | undefined {
  if (!SHOWN_ADDRESS.test(address)) return undefined;
  const [first, second = first] = address.toUpperCase().replaceAll('$', '').split(':');
  const a = parseCellRef(first!);
  const b = parseCellRef(second!);
  if (!a || !b) return undefined;
  const row = Math.min(a.row, b.row) + r;
  const col = Math.min(a.col, b.col) + c;
  return row < MAX_ROWS && col < MAX_COLS ? formatCellRef({ row, col }) : undefined;
}

/** The properties of a range in the subset. */
const PROPERTIES = ['address', 'values', 'numberFormat'] as const;

/** The write of one cell value, the detail of its refusal, or undefined for no cell value. */
type Classified = { readonly write: AutomationCellWrite } | { readonly refuse: string } | undefined;

/**
 * The error literals (ECMA-376 Part 1, 18.17.3) and the error values that workbooks store for
 * dynamic arrays, data types, and external data, as `packages/core/src/cascade/values.ts` reads
 * them. A typed entry reads each as an error value.
 */
const ERROR_TEXTS = new Set([
  '#NULL!',
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
  '#GETTING_DATA',
  '#SPILL!',
  '#CALC!',
  '#FIELD!',
  '#BLOCKED!',
  '#CONNECT!',
  '#BUSY!',
  '#UNKNOWN!',
]);

/**
 * A range of cells of one worksheet.
 *
 * @public
 */
export class Range extends ClientObject {
  readonly #sheet: Worksheet;
  readonly #address: string;

  /** @internal */
  constructor(context: RequestContext, sheet: Worksheet, address: string) {
    super(context);
    this.#sheet = sheet;
    this.#address = address;
  }

  /** The address of the range with its sheet name, such as `Sheet1!A1:B2`. */
  get address(): string {
    return this.loadedProperty<string>('address');
  }

  /**
   * The values of the cells as the file stores them, rows first. A range with a cell that holds
   * a formula, or a result of an array formula or a data table, refuses the read with
   * `NotSupported`, because this API does not calculate formulas.
   *
   * Set it to a two-dimensional array with one item for each cell: a number, a boolean, text,
   * `""` to empty a cell, or `null` to leave a cell as it is. A formula (text that starts with
   * `=`), text that a typed entry reads as another value, such as `12%` or `TRUE`, and a write to
   * a formula cell refuse the sync with `NotSupported`. So does a value in a cell that a merged
   * range hides: write to its top-left cell. A set makes the loaded value out of
   * date: load it again before you read it. Each read gives the same array.
   */
  // oxlint-disable-next-line typescript/no-explicit-any -- Office.js types it as any[][].
  get values(): any[][] {
    return this.loadedProperty<unknown[][]>('values');
  }

  // oxlint-disable-next-line typescript/no-explicit-any -- Office.js types it as any[][].
  set values(values: any[][]) {
    const grid = this.#grid('values', values);
    if (!grid) return;
    const writes: AutomationCellWrite[][] = [];
    for (const [r, row] of grid.entries()) {
      const out: AutomationCellWrite[] = [];
      for (const [c, value] of row.entries()) {
        const classified = classify(value);
        if (classified && 'refuse' in classified) {
          const cell = cellAt(this.#address, r, c);
          this.refuseWith('values', 'NotSupported', classified.refuse, true, cell);
          return;
        }
        if (!classified) {
          this.refuseArgument(
            'values',
            'a finite number, a boolean, a string, or null for each cell',
            true,
            cellAt(this.#address, r, c)
          );
          return;
        }
        out.push(classified.write);
      }
      writes.push(out);
    }
    const sheet = this.#sheet;
    const address = this.#address;
    this.#enqueueWrite('values', () => ({
      op: 'writeValues',
      sheet: internals(sheet).selector,
      address,
      values: writes,
    }));
  }

  /**
   * The number format code of each cell, such as `General` or `0.00%`, rows first. A read gives
   * the stored code, or the ECMA-376 code of a built-in format: id 14 reads as `mm-dd-yy`. A
   * built-in format whose code depends on the locale (ids 27 to 36 and 50 to 58) refuses with
   * `NotSupported`. Each read gives the same array.
   *
   * Read-only in this subset: a write of number formats is outside the subset, so an
   * assignment fails to compile instead of failing at the next sync.
   */
  // oxlint-disable-next-line typescript/no-explicit-any -- Office.js types it as any[][].
  get numberFormat(): any[][] {
    return this.loadedProperty<unknown[][]>('numberFormat');
  }

  /** A copy of a two-dimensional array argument, or undefined after a refusal was queued. */
  #grid(name: string, value: unknown): readonly (readonly unknown[])[] | undefined {
    this.requireUsable();
    // The loaded value no longer says what the cells hold after this set.
    this.forgetLoadedProperty(name);
    if (!Array.isArray(value)) {
      this.refuseWith(name, 'NotSupported', 'broadcast', true);
      return undefined;
    }
    if (!value.every((row) => Array.isArray(row))) {
      this.refuseWith(name, 'InvalidArgument', 'shape', true);
      return undefined;
    }
    return value.map((row: unknown[]) => [...row]);
  }

  #enqueueWrite(name: string, op: () => Extract<AutomationOperation, { op: 'writeValues' }>): void {
    const target = `${this.label}.${name}`;
    this.enqueue({
      target,
      write: true,
      op: () => {
        internals(this.#sheet).requireReal(target);
        return op();
      },
      mapError: (error) => internals(this.#sheet).mapError(error, target),
    });
  }

  /** @internal */
  protected get label(): string {
    const shown = SHOWN_ADDRESS.test(this.#address) ? `'${this.#address}'` : '';
    return `${internals(this.#sheet).path}.getRange(${shown})`;
  }

  /** @internal `address`, `values`, and `numberFormat`, read in one operation. */
  protected onLoad(request: LoadRequest): void {
    this.checkNames(request, PROPERTIES);
    const want = (name: string) => request.all || request.names.has(name);
    const values = want('values');
    const numberFormat = want('numberFormat');
    const sheet = this.#sheet;
    const address = this.#address;
    const target = `${this.label}.load`;
    // A set after this load, before its sync, makes the loaded value out of date.
    const valueSets = this.setsOf('values');
    const formatSets = this.setsOf('numberFormat');
    this.enqueue({
      target,
      op: () => {
        internals(sheet).requireReal(target);
        return {
          op: 'readRange',
          sheet: internals(sheet).selector,
          address,
          ...(values ? { values: true } : {}),
          ...(numberFormat ? { numberFormat: true } : {}),
        };
      },
      settle: (value: AutomationValue) => {
        const read = value as {
          address?: unknown;
          values?: unknown;
          numberFormat?: unknown;
        } | null;
        if (!read || typeof read.address !== 'string') return;
        if (want('address')) this.setLoadedProperty('address', read.address);
        // One copy that the caller may change, as an Office.js loaded array: each read gives
        // the same array, and a change to it never reaches the workbook.
        if (values && Array.isArray(read.values))
          this.setLoadedProperty('values', mutableGrid(read.values), valueSets);
        if (numberFormat && Array.isArray(read.numberFormat))
          this.setLoadedProperty('numberFormat', mutableGrid(read.numberFormat), formatSets);
      },
      mapError: (error) => internals(sheet).mapError(error, target),
    });
  }
}

/** A copy of a two-dimensional array of the host, whose rows the caller may change. */
function mutableGrid(grid: readonly unknown[]): unknown[][] {
  return grid.map((row) => (Array.isArray(row) ? [...row] : []));
}

/** One value of a `values` write, classified. */
function classify(value: unknown): Classified {
  if (value === null) return { write: null };
  if (typeof value === 'number')
    return Number.isFinite(value) ? { write: { type: 'number', value } } : undefined;
  if (typeof value === 'boolean') return { write: { type: 'boolean', value } };
  if (typeof value !== 'string') return undefined;
  if (value === '') return { write: { type: 'empty' } };
  if (value.startsWith('=')) return { refuse: 'formula' };
  return readsAsValue(value) ? { refuse: 'typed-text' } : { write: { type: 'string', value } };
}

/** Whether a typed entry would read `text` as something other than that text. */
function readsAsValue(text: string): boolean {
  // A typed entry drops the spaces around a value, so ` 12 ` reads as 12.
  const trimmed = text.trim();
  if (trimmed !== text) return trimmed.length > 0 && readsAsValue(trimmed);
  // A leading sign reads as a formula in the Office.js object model, and a leading apostrophe
  // as a text mark: neither is stored as the text it is.
  if (text.startsWith('+') || text.startsWith('-') || text.startsWith("'")) return true;
  if (text.startsWith('#') && ERROR_TEXTS.has(text.toUpperCase())) return true;
  const read = parseCellInput(text, EN_US);
  return read.type !== 'string' || read.value !== text;
}
