/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See LICENSE.md in this package.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// After PPTX Editor's `packages/editor-api/src/runtime/load-options.ts` at d4502552, which adapts
// DOCX Editor's `packages/editor-api/src/runtime/load-options.ts` at 5993bd223945. Differences:
// the subset has no navigation property except `items` on a collection, so the request keeps the
// head of each deeper path for the object to refuse; and a non-empty `expand` and the property
// tree form (`{ values: true }`) are refused as `NotSupported` instead of read. `top` and `skip`
// are read as in DOCX Editor: a whole number of 0 or more, for a collection. An object that is
// not a collection refuses them. Each refusal gives a hint, as each PPTX Editor load refusal
// does; the target is the loaded object, as in PPTX Editor.
//
// THE FORMS. A name, a comma-separated list of names, an array of names, and
// `{ select, top, skip }`.
// `$all` loads every property of the subset; `items/$all` loads the items.
//
// REFUSAL IS THE POINT. A misspelled option key is the difference between "load selected
// properties" and "load nothing", so an unknown key or a value of the wrong type is
// `InvalidArgument` here. A name that is not a property of the object it reaches is refused by
// that object, with the list of its properties.
//
// THE REQUEST IS BOUNDED: at most 64 names, each at most 64 characters. Loaded values live in a
// `Map`, so a name such as `__proto__` is inert.

import { diagnosticOf, fail, failArgument } from './errors.ts';

/**
 * The object form of `load(...)`: `select`, and `top` and `skip` for a collection. A non-empty
 * `expand` is refused as `NotSupported`.
 *
 * @example
 * ```ts
 * const sheets = context.workbook.worksheets;
 * sheets.load({ select: 'items/name', skip: 1, top: 2 });
 * await context.sync();
 * ```
 *
 * @public
 */
export interface LoadOption {
  /** Which properties to load, as names, or as `items/name` on a collection. */
  readonly select?: string | readonly string[];
  /**
   * Reserved for Office.js source compatibility. Outside the subset: an empty list is accepted,
   * and any other value is refused as `NotSupported`.
   */
  readonly expand?: string | readonly string[];
  /**
   * For a collection: at most this many items, a whole number of 0 or more. An object that is not
   * a collection refuses it as `InvalidArgument`.
   */
  readonly top?: number;
  /**
   * For a collection: skip this many items first, a whole number of 0 or more. An object that is
   * not a collection refuses it as `InvalidArgument`.
   */
  readonly skip?: number;
}

/** What one `load(...)` asks of one object. */
export interface LoadRequest {
  /** The property names of the object. Empty with `all`: every property of the subset. */
  readonly names: ReadonlySet<string>;
  /** The names after `items/`, for a collection. */
  readonly items: ReadonlySet<string>;
  /** The first step of each deeper path, other than `items`. */
  readonly paths: ReadonlySet<string>;
  /** The first step after `items/` of each deeper path below `items`. */
  readonly itemPaths: ReadonlySet<string>;
  /** `$all` was named, or no names were given: load every property of the subset. */
  readonly all: boolean;
  /** For a collection: at most this many items. */
  readonly top?: number;
  /** For a collection: skip this many items first. */
  readonly skip?: number;
}

const MAX_NAMES = 64;
const MAX_STEPS = 8;
const NAME = /^[A-Za-z$][A-Za-z0-9]{0,63}$/;
const OPTION_KEYS = new Set(['select', 'expand', 'top', 'skip']);
const ALL = '$all';

function namesOf(value: unknown, target: string): string[] {
  if (typeof value !== 'string' && !Array.isArray(value))
    failArgument(target, 'a property name, a list of names, or { select }');
  const list: unknown[] = typeof value === 'string' ? [value] : value;
  if (list.length > MAX_NAMES) failArgument(target, `at most ${MAX_NAMES} property names`);
  const out: string[] = [];
  for (const item of list) {
    if (typeof item !== 'string') failArgument(target, 'property names as strings');
    if (item.length > MAX_NAMES * 66) failArgument(target, `at most ${MAX_NAMES} property names`);
    for (const part of item.split(',')) {
      const name = part.trim();
      if (name.length > 0) out.push(name);
      if (out.length > MAX_NAMES) failArgument(target, `at most ${MAX_NAMES} property names`);
    }
  }
  return out;
}

/** A `top` or `skip` value: a whole number of 0 or more, as in DOCX Editor. */
function count(value: unknown, target: string): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 0)
    failArgument(target, 'a whole number of 0 or more');
  return value;
}

/** Read the option of `load(...)`. Throws `InvalidArgument` or `NotSupported`. */
export function readLoadOption(option: unknown, target: string): LoadRequest {
  let names: string[] = [];
  let window: { top?: number; skip?: number } = {};
  if (option === undefined || option === null) names = [];
  else if (typeof option === 'string' || Array.isArray(option)) names = namesOf(option, target);
  else if (typeof option === 'object') {
    const keys = Object.keys(option);
    // A property tree, such as `{ values: true }`, is outside the subset.
    if (!keys.some((key) => OPTION_KEYS.has(key)))
      fail({ code: 'NotSupported', target, diagnostic: diagnosticOf('load-tree')! });
    for (const key of keys)
      if (!OPTION_KEYS.has(key))
        failArgument(
          target,
          "an options object with select, top, and skip only, such as { select: 'values' }"
        );
    const o = option as LoadOption;
    // An empty list asks for nothing, as in DOCX Editor.
    if (o.expand !== undefined && !(Array.isArray(o.expand) && o.expand.length === 0))
      fail({
        code: 'NotSupported',
        target: `${target}.expand`,
        diagnostic: diagnosticOf('load-option')!,
      });
    if (o.top !== undefined) window = { ...window, top: count(o.top, `${target}.top`) };
    if (o.skip !== undefined) window = { ...window, skip: count(o.skip, `${target}.skip`) };
    names = o.select === undefined ? [] : namesOf(o.select, `${target}.select`);
  } else failArgument(target, 'a property name, a list of names, or { select }');
  const out = {
    names: new Set<string>(),
    items: new Set<string>(),
    paths: new Set<string>(),
    itemPaths: new Set<string>(),
  };
  let all = false;
  for (const name of names) {
    const steps = name.split('/');
    if (steps.length > MAX_STEPS || !steps.every((step) => NAME.test(step)))
      failArgument(
        target,
        `property names of letters and digits, or paths of at most ${MAX_STEPS} such names`
      );
    const [head, next] = steps as [string, ...string[]];
    if (name === ALL) all = true;
    else if (name === `items/${ALL}`) out.names.add('items');
    else if (next === undefined) out.names.add(head);
    else if (head !== 'items') out.paths.add(head);
    else if (steps.length === 2) out.items.add(next);
    else out.itemPaths.add(next);
  }
  const named = out.names.size + out.items.size + out.paths.size + out.itemPaths.size;
  return { ...out, ...window, all: all || named === 0 };
}
