/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See LICENSE.md in this package.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// After PPTX Editor's `packages/editor-api/src/runtime/request-context.ts`,
// `client-object.ts`, and `batch.ts` at d4502552, which adapt DOCX Editor's runtime at
// 5993bd223945. Differences: no tracked objects, no trace messages, no client results, and no
// handles: a worksheet is addressed by an opaque id once a sync has found it, else by its name
// or as the active worksheet, and the host looks it up when the batch runs.
//
// ONE CONTEXT, ONE QUEUE. Proxies queue actions, and `sync()` sends the queue as one batch: all
// of it applies, or none of it. A refusal empties the queue, as in Office.js.
//
// ARGUMENT CHECKS GO TO THE SYNC. A write or a call that this API refuses for its argument does
// not throw at the call: the refusal is queued, and the next `sync()` refuses with it, so
// nothing of that batch applies. What throws at the call: a read before its load
// (`PropertyNotLoaded`), `load(...)` options, and a proxy or context used outside its run.
//
// STALE READS. After a sync, the context knows the revision it read at. A later sync that writes
// sends it as the expected revision: when another writer (the person at the editor) changed the
// workbook in between, the sync refuses with `StaleDocument` and nothing applies. A sync that only
// reads is never stale, and a context that has read nothing writes unconditionally, as in PPTX
// Editor.
//
// DROPPED WRITES. Actions that no sync sent are dropped at the end of the run. An Office.js host
// sends them with a final sync of its own, so this is a difference that the compatibility notes
// name. When a callback returns with writes still queued, a development build prints one warning
// that names them (`dev-warnings.ts`), as PPTX Editor does.

import {
  MAX_BATCH_OPERATIONS,
  type AutomationError,
  type AutomationHost,
  type AutomationOperation,
  type AutomationValue,
} from '@xlsx-editor.dev/core/automation';
import {
  argumentRefusal,
  fail,
  failArgument,
  listOf,
  refusal,
  XlsxEditorError,
  type XlsxEditorErrorCode,
} from './errors.ts';
import { warnUnsynced } from './dev-warnings.ts';
import { readLoadOption, type LoadOption, type LoadRequest } from './load-options.ts';
import type { Workbook } from '../model/workbook.ts';

/** One queued action: an operation to send, or a refusal to give at the sync. */
export interface Action {
  /** The consumer-facing path of the call, for an error. */
  readonly target: string;
  /** Builds the operation when the sync sends it. */
  readonly op?: () => AutomationOperation;
  /** Whether the action is a write: an operation that changes the workbook, or its refusal. */
  readonly write?: boolean;
  /** A refusal of the call's argument, which the sync gives. */
  readonly refusal?: XlsxEditorError;
  /** Takes the value of the operation after the batch applied. */
  readonly settle?: (value: AutomationValue) => void;
  /** Maps a host refusal of this action, when the call has a better public answer. */
  readonly mapError?: (error: AutomationError) => XlsxEditorError | undefined;
}

/** What a context needs from its runtime. */
export interface RuntimeSession {
  readonly host: AutomationHost;
  /** Fails with `RuntimeDisposed` once the runtime is disposed. */
  assertLive(target?: string): void;
  /** Makes the workbook proxy of a context. */
  openWorkbook(context: RequestContext): Workbook;
}

/** The internal surface of a context, for the proxies of this package. */
export interface ContextInternals {
  readonly session: RuntimeSession;
  push(action: Action): void;
  /** Refuse a finished context or a disposed runtime. */
  assertUsable(target: string): void;
  isFinished(): boolean;
}

/** The internal surface of each context. A map, so the surface stays out of the public types. */
const INTERNALS = new WeakMap<RequestContext, ContextInternals>();

/** The internal surface of `context`, for the proxies of this package. */
export function internalsOf(context: RequestContext): ContextInternals {
  const internals = INTERNALS.get(context);
  if (!internals) return fail({ code: 'InvalidRequestContext', target: 'context' });
  return internals;
}

/** The host refusal codes, as public codes. */
const CODES: Readonly<Record<AutomationError['code'], XlsxEditorErrorCode>> = Object.freeze({
  'stale-revision': 'StaleDocument',
  'invalid-argument': 'InvalidArgument',
  unsupported: 'NotSupported',
  'resource-limit': 'ResourceLimitExceeded',
  disposed: 'RuntimeDisposed',
  'transaction-refused': 'GeneralException',
  'unknown-operation': 'GeneralException',
  'item-not-found': 'ItemNotFound',
});

/** The public code of a store refusal (`transaction-refused`), by its rejection code. */
const STORE_CODES: ReadonlyMap<string, XlsxEditorErrorCode> = new Map([
  ['text-too-long', 'InvalidArgument'],
  ['invalid-character', 'InvalidArgument'],
  ['locked', 'NotSupported'],
  ['shared-formula', 'NotSupported'],
  ['cell-alternate-content', 'NotSupported'],
  ['unsupported-package', 'NotSupported'],
  ['range-too-large', 'ResourceLimitExceeded'],
  ['stale-revision', 'StaleDocument'],
]);

/** A host refusal as the error this runtime throws. */
export function hostRefusal(
  error: AutomationError,
  target: string | undefined,
  revisions: { expected?: number; actual?: number } = {}
): XlsxEditorError {
  let code = CODES[error.code] ?? 'GeneralException';
  let detail = error.detail;
  if (error.code === 'transaction-refused') code = STORE_CODES.get(detail ?? '') ?? code;
  if (error.code === 'invalid-argument') {
    if (detail === 'address') detail = 'address-outside';
    else if (detail === 'values') detail = 'shape';
  }
  if (error.code === 'item-not-found' && detail === 'gone') code = 'InvalidObjectPath';
  if (code === 'StaleDocument')
    return new XlsxEditorError({
      code,
      ...(target === undefined ? {} : { target }),
      ...(revisions.expected === undefined ? {} : { expectedRevision: revisions.expected }),
      ...(revisions.actual === undefined ? {} : { actualRevision: revisions.actual }),
    });
  return refusal(code, target, detail, error.cell);
}

/**
 * The context of one `run`: the workbook to address, and the queue that `sync()` sends.
 *
 * @public
 */
export class RequestContext {
  readonly #session: RuntimeSession;
  #queue: Action[] = [];
  #finished = false;
  #revision: number | undefined;
  #workbook: Workbook | undefined;

  private constructor(session: RuntimeSession) {
    this.#session = session;
    INTERNALS.set(
      this,
      Object.freeze({
        session,
        push: (action: Action) => {
          this.#assertUsable(action.target);
          this.#queue.push(action);
        },
        assertUsable: (target: string) => this.#assertUsable(target),
        isFinished: () => this.#finished,
      })
    );
  }

  /** @internal The context of a new run. Use `runtime.run(...)`. */
  static begin(session: RuntimeSession): RequestContext {
    return new RequestContext(session);
  }

  /** The workbook of this run. */
  get workbook(): Workbook {
    this.#assertUsable('workbook');
    this.#workbook ??= this.#session.openWorkbook(this);
    return this.#workbook;
  }

  /**
   * Queue the reads that fill in the selected properties of `object`, as `object.load(option)`
   * does. The values are readable after the next `sync()`. An object of another run refuses
   * with `InvalidObjectPath`.
   */
  load(object: ClientObject, option?: LoadOption | string | string[]): void {
    this.#assertUsable('load');
    if (!(object instanceof ClientObject)) fail({ code: 'InvalidArgument', target: 'load' });
    // An object of another run would queue its load there, as PPTX Editor's `loadRecursive`
    // refuses it.
    if (object.context !== this) fail({ code: 'InvalidObjectPath', target: 'load' });
    object.load(option);
  }

  /**
   * Send the queued actions as one batch. All of it applies, or none of it, and the queue is
   * empty afterward either way. Answers `passThroughValue`.
   *
   * One sync sends at most 10,000 operations, its range reads and writes reach at most
   * 1,000,000 cells in all, and its writes at most 10,000 cells: a larger batch refuses with
   * `ResourceLimitExceeded`. After a sync, a later sync that writes refuses with
   * `StaleDocument` when the workbook changed in between. A load of values or number formats
   * after a write in the same sync refuses with `NotSupported`: sync after the writes, then load.
   */
  sync<T>(passThroughValue?: T): Promise<T> {
    try {
      this.#send();
      return Promise.resolve(passThroughValue as T);
    } catch (error) {
      return Promise.reject(error);
    }
  }

  #send(): void {
    this.#assertUsable('sync');
    const actions = this.#queue;
    this.#queue = [];
    const refused = actions.find((action) => action.refusal);
    if (refused) throw refused.refusal;
    const sendable = actions.filter((action) => action.op);
    if (sendable.length === 0) return;
    if (sendable.length > MAX_BATCH_OPERATIONS)
      throw refusal('ResourceLimitExceeded', 'context.sync', 'operations');
    const operations = sendable.map((action) => action.op!());
    const expected = sendable.some((action) => action.write) ? this.#revision : undefined;
    const response = this.#session.host.execute({
      operations,
      ...(expected === undefined ? {} : { expectedRevision: expected }),
    });
    if (!response.ok) {
      const at = response.results.findIndex((r) => r.status === 'error');
      const result = response.results[at];
      const action = sendable[at];
      const error =
        result?.status === 'error'
          ? (action?.mapError?.(result.error) ??
            hostRefusal(result.error, action?.target, {
              ...(expected === undefined ? {} : { expected }),
              actual: response.revision,
            }))
          : new XlsxEditorError({ code: 'GeneralException' });
      throw error;
    }
    this.#revision = response.revision;
    response.results.forEach((result, at) => {
      if (result.status === 'ok') sendable[at]!.settle?.(result.value);
    });
  }

  #assertUsable(target: string): void {
    this.#session.assertLive(target);
    if (this.#finished) fail({ code: 'InvalidRequestContext', target });
  }

  /**
   * @internal End of the run. Queued actions are dropped, never sent. When the callback
   * `completed`, dropped writes print a development warning.
   */
  static finish(context: RequestContext, completed: boolean): void {
    context.#finished = true;
    const writes = context.#queue.filter((action) => action.write);
    context.#queue = [];
    if (completed && writes.length > 0)
      warnUnsynced(
        writes.length,
        writes.flatMap((action) => (action.refusal ? [action.refusal] : []))
      );
  }
}

/**
 * The context type of the Office.js `OfficeExtension` namespace. This package has one context
 * class, so it is {@link RequestContext}.
 *
 * @public
 */
export type ClientRequestContext = RequestContext;

/** Refuse a property name that an object does not have, with the list of its properties. */
export function refuseName(label: string, name: string, known: readonly string[]): never {
  return failArgument(
    `${label}.${name}`,
    known.length === 0
      ? 'load() with no property name, because this object has no property in the subset'
      : `a property of this object: ${listOf(known)}`
  );
}

/**
 * Refuse a path through `head`. The subset has no navigation property but `items` on a
 * collection (`navigation`), as PPTX Editor refuses a path through a name that is not one.
 */
export function refusePath(
  label: string,
  head: string,
  navigation: readonly string[],
  known: readonly string[]
): never {
  return failArgument(
    `${label}.${head}`,
    navigation.length > 0
      ? `a path through ${listOf(navigation)}, or a property name: ${listOf(known)}`
      : `a property name with no path, because this object has no navigation property in the subset: ${listOf(known)}`
  );
}

/**
 * A loaded value as plain data: a proxy becomes its own `toJSON()`, and an array (a row of a
 * two-dimensional value) becomes a copy, so a change to the result never reaches the proxy.
 */
function plain(value: unknown): unknown {
  if (value instanceof ClientObject) return value.toJSON();
  return Array.isArray(value) ? value.map(plain) : value;
}

/**
 * The base of every workbook proxy.
 *
 * A proxy that outlives its run still answers what it already loaded. Anything that would talk
 * to the workbook (`load`, a write, a method) refuses with `InvalidObjectPath`.
 *
 * @public
 */
export abstract class ClientObject {
  readonly #context: RequestContext;
  readonly #loaded = new Map<string, unknown>();
  /** How many sets each property has had: a load queued before a set does not settle it. */
  readonly #sets = new Map<string, number>();

  /** @internal */
  protected constructor(context: RequestContext) {
    this.#context = context;
  }

  /** The request context this object belongs to. */
  get context(): RequestContext {
    return this.#context;
  }

  /**
   * Whether this object turned out not to exist. Only an object of a `get...OrNullObject` call
   * can be a null object, and it has no answer until the sync that looked for it: a read before
   * then is `PropertyNotLoaded`.
   */
  get isNullObject(): boolean {
    return false;
  }

  /**
   * Queue the reads that fill in the selected properties. Returns `this`; the values are
   * readable after the next `sync()`. With no option, it loads every property of the subset.
   */
  load(option?: LoadOption | string | string[]): this {
    this.requireUsable();
    this.onLoad(readLoadOption(option, this.label));
    return this;
  }

  /** The loaded properties as a plain object, as Office.js `toJSON()` gives them. */
  toJSON(): Record<string, unknown> {
    const out: Record<string, unknown> = Object.create(null) as Record<string, unknown>;
    for (const [name, value] of this.#loaded) out[name] = plain(value);
    return { ...out };
  }

  /** @internal The consumer-facing path of this object. */
  protected abstract get label(): string;

  /** @internal What this kind of object does with a load request. */
  protected abstract onLoad(request: LoadRequest): void;

  /** @internal Refuse an object whose run ended, and a disposed runtime. */
  protected requireUsable(): void {
    internalsOf(this.#context).session.assertLive(this.label);
    if (internalsOf(this.#context).isFinished())
      fail({ code: 'InvalidObjectPath', target: this.label });
  }

  /** @internal Queue one action of this object. */
  protected enqueue(action: Action): void {
    this.requireUsable();
    internalsOf(this.#context).push(action);
  }

  /** @internal Queue the refusal of a call's argument, for the next sync. */
  protected refuse(name: string, error: XlsxEditorError, write = false): void {
    this.enqueue({ target: `${this.label}.${name}`, refusal: error, write });
  }

  /** @internal Queue an `InvalidArgument` refusal that names what the call accepts. */
  protected refuseArgument(name: string, accepted: string, write = false, cell?: string): void {
    this.refuse(name, argumentRefusal(`${this.label}.${name}`, accepted, cell), write);
  }

  /**
   * @internal Queue a refusal with the public diagnostic of `detail`. `cell` names the first
   * refused cell in the hint.
   */
  protected refuseWith(
    name: string,
    code: XlsxEditorErrorCode,
    detail: string,
    write = false,
    cell?: string
  ): void {
    this.refuse(name, refusal(code, `${this.label}.${name}`, detail, cell), write);
  }

  /** @internal Read a property that a completed `load` filled in; refuses if none did. */
  protected loadedProperty<T>(name: string): T {
    if (!this.#loaded.has(name)) {
      if (internalsOf(this.#context).isFinished())
        fail({ code: 'InvalidObjectPath', target: this.label });
      fail({ code: 'PropertyNotLoaded', target: `${this.label}.${name}` });
    }
    return this.#loaded.get(name) as T;
  }

  /**
   * @internal Record a value that a completed load produced. `since` is the `setsOf(name)` of
   * the load when it was queued: a set after that makes the loaded value out of date, so the
   * value is not recorded.
   */
  protected setLoadedProperty(name: string, value: unknown, since?: number): void {
    if (since !== undefined && since !== this.setsOf(name)) return;
    this.#loaded.set(name, value);
  }

  /** @internal How many sets the property has had, for `setLoadedProperty`. */
  protected setsOf(name: string): number {
    return this.#sets.get(name) ?? 0;
  }

  /** @internal Forget a loaded value: a set makes it out of date until the next load. */
  protected forgetLoadedProperty(name: string): void {
    this.#loaded.delete(name);
    this.#sets.set(name, this.setsOf(name) + 1);
  }

  /**
   * @internal Refuse a load of names that this object does not have, as PPTX Editor does:
   * `InvalidArgument` with every property of the object. An object that is not a collection
   * also refuses `top` and `skip`.
   */
  protected checkNames(request: LoadRequest, known: readonly string[]): void {
    if (request.items.size > 0 || request.itemPaths.size > 0)
      failArgument(`${this.label}.items`, "'items/' paths on a collection only");
    for (const key of ['top', 'skip'] as const)
      if (request[key] !== undefined)
        failArgument(`${this.label}.${key}`, `${key} on a collection only`);
    for (const name of request.names)
      if (!known.includes(name)) refuseName(this.label, name, known);
    for (const head of request.paths) refusePath(this.label, head, [], known);
  }
}
