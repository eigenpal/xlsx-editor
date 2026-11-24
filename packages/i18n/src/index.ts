// Copied from docx-editor/packages/i18n/src/index.ts at a0825f1f (EigenPal). Differences: the
// package is `@xlsx-editor.dev/i18n` with English only. The writing direction helpers are in
// `direction.ts`, as in DOCX Editor; the platform-shortcut helpers are left out until chrome needs
// them. `formatMessage` is exported, where DOCX Editor keeps it private: the adapters'
// `useChromeTranslate` formats an override with it, because the messages here take values. It
// reads only the own values of `vars`, so a placeholder named after an object method stays as
// written. The disabled-reason helper is not here: the engine resolves its own reasons to
// catalog text (`disabledReasonText` in `@xlsx-editor.dev/core/editor`), so the adapters need no
// second copy. The adapter text names only what exists: the React adapter, its
// `useTranslation` hook, the `locale` prop of its root, and the engine's English fallback
// (`englishText` in the core editor lane).
/**
 * Shared locale data, types, and runtime helpers for the @xlsx-editor.dev
 * editor adapters.
 *
 * Import everything from the package root. `sideEffects: false` lets
 * consumer bundlers tree-shake unused locales.
 *
 * ```ts
 * import {
 *   en,                                 // typed locale data
 *   locales,                            // record keyed by BCP-47 tag
 *   deepMerge, createT,                 // build a t() for custom hosts
 *   formatMessage,                      // fill one message as t() does
 *   localeDirection, catalogDirection,  // the writing direction of a tag
 *   chromeDirection,                    //   or a catalog, and of the chrome
 *   type LocaleStrings,                 // shape of `en` (source of truth)
 *   type Translations,                  // shape of a community partial
 *   type TranslationKey,                // every valid `t()` key
 *   type LocaleCode,                    // 'en', the one locale shipped
 * } from '@xlsx-editor.dev/i18n';
 * ```
 *
 * The React adapter wraps `createT` in its own bindings: the `locale` prop
 * of `XlsxEditor.Root` and the `useTranslation` hook. Use those for app
 * code. Reach for `createT` directly when building another host.
 *
 * @packageDocumentation
 * @public
 */

// ─── GENERATED START — `bun run i18n:codegen` ───
// DO NOT EDIT — this block is rewritten from the on-disk
// `packages/i18n/*.json` filenames whenever `bun run i18n:codegen`
// runs (and `bun run i18n:new <lang>` runs it automatically).
// `bun run i18n:validate` fails CI if hand-edits drift from the
// JSON files. Edit the JSON, not this block.

import enJson from '../en.json';

/**
 * Full locale string set, auto-derived from `en.json` (the source of truth).
 * Every other locale is a `PartialLocaleStrings` against this shape.
 *
 * @public
 */
export type LocaleStrings = typeof enJson;

/**
 * Every locale code shipped from this package. Pass to `locales[code]`
 * for runtime lookup; assign to `_lang` to drive `Intl.PluralRules`.
 *
 * Custom codes are accepted at runtime ({@link PartialLocaleStrings._lang}
 * widens to any string), but the shipped union is the IDE-completion list.
 *
 * @public
 */
// prettier-ignore
export type LocaleCode = 'en';

/** English (`en`) — the source of truth, 100% covered. @public */
export const en: LocaleStrings = enJson;

/**
 * Every shipped locale, keyed by BCP-47 tag. Use for runtime locale
 * pickers and "look up the locale matching this user preference" code:
 *
 * ```ts
 * <XlsxEditor locale={locales[userLocale]} />
 * ```
 *
 * Importing `locales` defeats the per-locale tree-shake — the bundler
 * sees a static reference to every locale. If you only need one or two,
 * import them by name (`import { en } from '@xlsx-editor.dev/i18n'`) instead.
 *
 * @public
 */
export const locales: Record<LocaleCode, PartialLocaleStrings> = {
  en,
};
// ─── GENERATED END ───

/**
 * Recursive Partial that allows `null` at leaves to signal "not yet
 * translated, fall back to English." Community translations use this shape;
 * `bun run i18n:fix` keeps every locale aligned to `en.json` with `null`
 * placeholders for missing keys.
 *
 * @public
 */
export type DeepPartial<T> = {
  [K in keyof T]?: T[K] extends object ? DeepPartial<T[K]> : T[K] | null;
};

/**
 * Partial locale strings — what consumers pass to the `locale` prop of
 * `XlsxEditor.Root`.
 * Missing keys fall back to English. Optional `_lang` carries the BCP-47
 * tag used by `Intl.PluralRules`; shipped codes autocomplete but custom
 * strings are accepted.
 *
 * @public
 */
export type PartialLocaleStrings = DeepPartial<LocaleStrings> & {
  // `string & {}` keeps the literal-union completions for `LocaleCode`
  // while still accepting arbitrary strings at runtime.
  _lang?: LocaleCode | (string & {});
  /**
   * The writing direction of the interface strings. `'rtl'` mirrors the editor chrome, and
   * `'ltr'` keeps a left-to-right chrome on a right-to-left host page. Without it, a
   * right-to-left `_lang` mirrors the chrome ({@link chromeDirection}).
   */
  _dir?: 'ltr' | 'rtl' | (string & {});
};

/**
 * Alias for `PartialLocaleStrings`. Prefer this name when typing the
 * consumer-facing `locale` prop or a function parameter.
 *
 * @public
 */
export type Translations = PartialLocaleStrings;

type DotPath<T, Prefix extends string = ''> = {
  [K in keyof T & string]: T[K] extends Record<string, unknown>
    ? DotPath<T[K], `${Prefix}${K}.`>
    : `${Prefix}${K}`;
}[keyof T & string];

/**
 * Every valid dot-notation key into `LocaleStrings`, e.g. `'zoom.zoomIn'`
 * or `'grid.label'`. Pass to `t(key, vars?)` for
 * compile-time-checked translation lookup.
 *
 * @public
 */
export type TranslationKey = DotPath<LocaleStrings>;

// ─────────────────────────────────────────────────────────────────────────────
// Runtime helpers — used by the React adapter and the engine's English
// fallback; exported here so other hosts can build a typed `t()` without
// re-implementing the merge + ICU format logic.
// ─────────────────────────────────────────────────────────────────────────────

type AnyRecord = Record<string, unknown>;

function isRecord(v: unknown): v is AnyRecord {
  return v !== null && typeof v === 'object' && !Array.isArray(v);
}

/**
 * Deep-merge a partial locale over a base locale. Null leaves in the
 * override are treated as "not translated" and fall back to the base.
 * The adapter calls this once when the `locale` prop changes, then hands the
 * result to {@link createT}.
 *
 * @public
 */
export function deepMerge(base: AnyRecord, override: AnyRecord | undefined): AnyRecord {
  if (!override) return base;
  const result: AnyRecord = { ...base };
  for (const key of Object.keys(override)) {
    const overVal = override[key];
    if (overVal === null) continue;
    if (isRecord(base[key]) && isRecord(overVal)) {
      result[key] = deepMerge(base[key], overVal);
    } else if (overVal !== undefined) {
      result[key] = overVal;
    }
  }
  return result;
}

function lookupKey(obj: AnyRecord, path: string): string | undefined {
  let current: unknown = obj;
  for (const part of path.split('.')) {
    if (!isRecord(current)) return undefined;
    current = current[part];
  }
  return typeof current === 'string' ? current : undefined;
}

function parseBranches(branchStr: string): Record<string, string> {
  const parsed: Record<string, string> = {};
  const regex = /(=\d+|\w+)\s*\{([^}]*)\}/g;
  let match;
  while ((match = regex.exec(branchStr)) !== null) {
    parsed[match[1]] = match[2];
  }
  return parsed;
}

/**
 * Format one message as `t()` does: fill the `{name}` placeholders from `vars`, and pick the
 * branch of each ICU `{count, plural, ...}` with `Intl.PluralRules` for `lang` (`=N` matches
 * first). A placeholder with no value stays as written. Use it for a message that is not in a
 * catalog, such as an override of one key.
 *
 * @example
 * ```ts
 * formatMessage('{count, plural, one {# sheet} other {# sheets}}', { count: 2 }); // '2 sheets'
 * ```
 *
 * @public
 */
export function formatMessage(
  template: string,
  vars?: Record<string, string | number>,
  lang?: string
): string {
  if (!vars) return template;

  const result = template.replace(
    // Linear on hostile templates: the branch group keeps the single-char
    // `[^{}]` alternative (using `[^{}]+` would form `(X+)*`, an exponential
    // pattern), and there is no `\s*` before the group — that `\s*` overlapped
    // with the group's leading whitespace and let a run of spaces be
    // partitioned many ways (the polynomial-ReDoS source). parseBranches
    // already tolerates the leading whitespace now folded into the capture.
    /\{(\w+),\s*plural,((?:[^{}]|\{[^{}]*\})*)\}/g,
    (full, varName, branchStr) => {
      const count = Number(Object.hasOwn(vars, varName) ? vars[varName] : undefined);
      if (isNaN(count)) return full;
      const parsed = parseBranches(branchStr);
      const exact = parsed[`=${count}`];
      if (exact !== undefined) return exact.replace(/#/g, String(count));
      let category: string;
      try {
        category = new Intl.PluralRules(lang || 'en').select(count);
      } catch {
        category = count === 1 ? 'one' : 'other';
      }
      const text = parsed[category] ?? parsed['other'] ?? '';
      return text.replace(/#/g, String(count));
    }
  );

  return result.replace(/\{(\w+)\}/g, (_, key) => {
    // Own values only: a placeholder named `constructor` or `toString` never reads the
    // prototype of `vars`.
    const val = Object.hasOwn(vars, key) ? vars[key] : undefined;
    return val !== undefined ? String(val) : `{${key}}`;
  });
}

export { catalogDirection, chromeDirection, localeDirection } from './direction.ts';

/**
 * The signature of `t()`: look up a translation by dot-notation key,
 * interpolate `{vars}`, and resolve ICU plurals.
 *
 * @public
 */
export type TFunction = (key: TranslationKey, vars?: Record<string, string | number>) => string;

/**
 * Build a typed `t(key, vars?)` function from a merged locale.
 *
 * - **Lookup**: dot-notation paths against the locale tree
 *   (`'zoom.zoomIn'`, `'grid.label'`).
 * - **Interpolation**: `{name}` placeholders read from `vars`.
 * - **Plurals**: ICU `{count, plural, =0 {none} one {# item} other {# items}}`
 *   with `Intl.PluralRules` for CLDR categories and `=N` for exact matches.
 * - **Fallback**: missing keys return the key string itself, useful for
 *   spotting un-translated UI in development.
 *
 * The React adapter wraps this in `useTranslation()`; use it directly
 * when building another host (server-rendered docs, CLI, etc.).
 *
 * @example
 * ```ts
 * import { deepMerge, createT, en } from '@xlsx-editor.dev/i18n';
 * const t = createT(en);
 * t('zoom.zoomIn'); // → 'Zoom in'
 * t('grid.label', { sheet: 'Sales' });
 * ```
 *
 * @public
 */
export function createT(strings: LocaleStrings, lang = 'en'): TFunction {
  return (key, vars) => {
    const value = lookupKey(strings as AnyRecord, key);
    return formatMessage(value ?? key, vars, lang);
  };
}
