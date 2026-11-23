/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See LICENSE.md in this package.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The workbook, its worksheet collection, and a worksheet, in the shape of the Office.js
// workbook object model. The members are the subset that `compat/manifest.json` records.

import type {
  AutomationError,
  AutomationWorksheetRead,
  AutomationWorksheetSelector,
} from '@xlsx-editor.dev/core/automation';
import {
  diagnosticOf,
  fail,
  failArgument,
  refusal,
  type XlsxEditorError,
} from '../runtime/errors.ts';
import type { LoadRequest } from '../runtime/load-options.ts';
import {
  ClientObject,
  refuseName,
  refusePath,
  type RequestContext,
} from '../runtime/request-context.ts';
import { Range } from './range.ts';

/**
 * The workbook of a run: the top of the object model.
 *
 * @public
 */
export class Workbook extends ClientObject {
  #worksheets: WorksheetCollection | undefined;

  /** @internal */
  constructor(context: RequestContext) {
    super(context);
  }

  /** The worksheets of the workbook. */
  get worksheets(): WorksheetCollection {
    this.requireUsable();
    this.#worksheets ??= new WorksheetCollection(this.context);
    return this.#worksheets;
  }

  /** @internal */
  protected get label(): string {
    return 'workbook';
  }

  /**
   * @internal The workbook has no property of the subset to load. Its worksheets are a collection
   * with a load of its own.
   */
  protected onLoad(request: LoadRequest): void {
    if (request.names.has('worksheets') || request.paths.has('worksheets'))
      failArgument(
        'workbook.worksheets',
        "a load on the collection: context.workbook.worksheets.load('items/name')"
      );
    this.checkNames(request, []);
  }
}

/** Whether a host value is one worksheet. */
function isWorksheetRead(value: unknown): value is AutomationWorksheetRead {
  return (
    typeof value === 'object' &&
    value !== null &&
    typeof (value as { id?: unknown }).id === 'string' &&
    typeof (value as { name?: unknown }).name === 'string'
  );
}

/**
 * The worksheets of a workbook, in tab order. Chartsheets are not worksheets.
 *
 * @public
 */
export class WorksheetCollection extends ClientObject {
  /** @internal */
  constructor(context: RequestContext) {
    super(context);
  }

  /** The loaded worksheets. Load `items` (or `items/name`) first. Each read gives the same array. */
  get items(): Worksheet[] {
    return this.loadedProperty<Worksheet[]>('items');
  }

  /**
   * The worksheet named `key`, matched without letter case: each character that has a
   * one-character upper case compares in upper case. A name that no worksheet has refuses the
   * sync with `ItemNotFound`.
   */
  getItem(key: string): Worksheet {
    return this.#lookup('getItem', key, false);
  }

  /**
   * The worksheet named `key`, matched as `getItem` matches it, or a null object when no
   * worksheet has that name. Check `isNullObject` after a sync.
   */
  getItemOrNullObject(key: string): Worksheet {
    return this.#lookup('getItemOrNullObject', key, true);
  }

  /**
   * The active worksheet: the one the workbook opens on, `x:workbookView/@activeTab` when it is
   * a visible worksheet, else the first visible worksheet.
   */
  getActiveWorksheet(): Worksheet {
    this.requireUsable();
    const sheet = new Worksheet(this.context, { active: true }, 'getActiveWorksheet()', false);
    this.enqueue({
      target: `${this.label}.getActiveWorksheet`,
      op: () => ({ op: 'getWorksheet', sheet: { active: true } }),
      settle: (value) => internals(sheet).resolve(value),
    });
    return sheet;
  }

  #lookup(method: string, key: unknown, orNull: boolean): Worksheet {
    this.requireUsable();
    const name = typeof key === 'string' ? key : '';
    const sheet = new Worksheet(this.context, { name }, `${method}()`, orNull);
    if (typeof key !== 'string') {
      this.refuseArgument(method, 'a worksheet name as a string');
      return sheet;
    }
    this.enqueue({
      target: `${this.label}.${method}`,
      op: () => ({ op: 'getWorksheet', sheet: { name }, ...(orNull ? { orNull: true } : {}) }),
      settle: (value) => internals(sheet).resolve(value),
    });
    return sheet;
  }

  /** @internal */
  protected get label(): string {
    return 'workbook.worksheets';
  }

  /**
   * @internal `items` loads each worksheet with its name, the one property of the subset. A
   * name with no `items/` means the same on a collection, as in PPTX Editor: `load('name')`
   * equals `load('items/name')`. `skip` and `top` load a window of the worksheets.
   */
  protected onLoad(request: LoadRequest): void {
    for (const name of request.names)
      if (name !== 'items' && name !== 'name') refuseName(this.label, name, ['items', 'name']);
    for (const name of request.items)
      if (name !== 'name') refuseName(`${this.label}.items`, name, ['name']);
    for (const head of request.paths) refusePath(this.label, head, ['items'], ['items', 'name']);
    for (const head of request.itemPaths) refusePath(`${this.label}.items`, head, [], ['name']);
    this.enqueue({
      target: `${this.label}.items`,
      op: () => ({ op: 'listWorksheets' }),
      settle: (value) => {
        const list = Array.isArray(value) ? value.filter(isWorksheetRead) : [];
        // `skip` and `top` choose a window of the list, as in DOCX Editor and PPTX Editor.
        const skip = request.skip ?? 0;
        const end = request.top === undefined ? list.length : skip + request.top;
        // One array that the caller may change, as an Office.js loaded collection: each read
        // gives the same array.
        this.setLoadedProperty(
          'items',
          list.slice(skip, end).map((read, at) => {
            const index = skip + at;
            const sheet = new Worksheet(this.context, { id: read.id }, `items[${index}]`, false);
            internals(sheet).resolve(read, true);
            return sheet;
          })
        );
      },
    });
  }
}

/** Where a worksheet proxy stands: looked up, found, or found to be missing. */
type SheetState = 'pending' | 'found' | 'null';

/** What a range and the collection need from a worksheet, outside its public type. */
export interface WorksheetInternals {
  /** The consumer-facing path of the worksheet. */
  readonly path: string;
  /** The selector a range operation sends. */
  readonly selector: AutomationWorksheetSelector;
  /** Refuse an operation on a null object, before it is sent. */
  requireReal(target: string): void;
  /**
   * A refusal of a range operation on this worksheet. A worksheet of a `getItemOrNullObject`
   * lookup whose name the same batch did not find is a null object.
   */
  mapError(error: AutomationError, target: string): XlsxEditorError | undefined;
  /** Take the answer of the lookup: from then on, the worksheet is its id. */
  resolve(value: unknown, withName?: boolean): void;
}

const INTERNALS = new WeakMap<Worksheet, WorksheetInternals>();

/** The internals of a worksheet of this package. */
export function internals(sheet: Worksheet): WorksheetInternals {
  return INTERNALS.get(sheet)!;
}

/**
 * One worksheet.
 *
 * @public
 */
export class Worksheet extends ClientObject {
  #selector: AutomationWorksheetSelector;
  readonly #label: string;
  readonly #nullable: boolean;
  #state: SheetState;

  /** @internal */
  constructor(
    context: RequestContext,
    selector: AutomationWorksheetSelector,
    label: string,
    nullable: boolean
  ) {
    super(context);
    this.#selector = selector;
    this.#label = `workbook.worksheets.${label}`;
    this.#nullable = nullable;
    this.#state = 'id' in selector ? 'found' : 'pending';
    const sheet = this;
    INTERNALS.set(this, {
      path: this.#label,
      get selector() {
        return sheet.#selector;
      },
      requireReal: (target) => {
        if (this.#state === 'null') throw refusal('InvalidObjectPath', target, 'null-object');
      },
      mapError: (error, target) => this.#mapError(error, target),
      resolve: (value, withName = false) => this.#resolve(value, withName),
    });
  }

  /**
   * The name of the worksheet, as its tab shows it. Read-only: a rename is outside the subset.
   * A null object has no name: a read refuses with `PropertyNotLoaded`.
   */
  get name(): string {
    if (this.#state === 'null')
      fail({
        code: 'PropertyNotLoaded',
        target: `${this.#label}.name`,
        diagnostic: diagnosticOf('null-object')!,
      });
    return this.loadedProperty<string>('name');
  }

  /**
   * Whether the worksheet turned out not to exist: true only for a worksheet of
   * `getItemOrNullObject` whose name no worksheet has. Read it after the sync of the lookup; a
   * read before then is `PropertyNotLoaded`.
   */
  override get isNullObject(): boolean {
    if (!this.#nullable) return false;
    if (this.#state === 'pending')
      fail({ code: 'PropertyNotLoaded', target: `${this.#label}.isNullObject` });
    return this.#state === 'null';
  }

  /**
   * The range at `address`: one cell (`A1`) or a rectangle (`A1:B2`), in either letter case,
   * with optional `$` marks. One range reaches at most 100,000 cells. Every other form (a sheet
   * name, a whole row or column, a defined name, or no address) refuses the sync with
   * `NotSupported`.
   */
  getRange(address?: string): Range {
    this.requireUsable();
    const range = new Range(this.context, this, typeof address === 'string' ? address : '');
    if (address === undefined || address === '')
      this.refuseWith('getRange', 'NotSupported', 'whole-sheet');
    else if (typeof address !== 'string') this.refuseArgument('getRange', 'an address as a string');
    return range;
  }

  #mapError(error: AutomationError, target: string): XlsxEditorError | undefined {
    if (this.#nullable && this.#state === 'pending' && error.code === 'item-not-found')
      return refusal('InvalidObjectPath', target, 'null-object');
    return undefined;
  }

  #resolve(value: unknown, withName: boolean): void {
    if (value === null) {
      this.#state = 'null';
      return;
    }
    if (!isWorksheetRead(value)) return;
    this.#selector = Object.freeze({ id: value.id });
    this.#state = 'found';
    if (withName) this.setLoadedProperty('name', value.name);
  }

  /** @internal */
  protected get label(): string {
    return this.#label;
  }

  /** @internal `name` is the one property of the subset. */
  protected onLoad(request: LoadRequest): void {
    this.checkNames(request, ['name']);
    if (this.#state === 'null') return;
    this.enqueue({
      target: `${this.#label}.name`,
      // Built when the sync sends it: an earlier sync may have found the id.
      op: () => {
        const sheet = this.#selector;
        return {
          op: 'getWorksheet',
          sheet,
          ...(this.#nullable && 'name' in sheet ? { orNull: true } : {}),
        };
      },
      settle: (value) => this.#resolve(value, true),
      mapError: (error) => this.#mapError(error, `${this.#label}.name`),
    });
  }
}
