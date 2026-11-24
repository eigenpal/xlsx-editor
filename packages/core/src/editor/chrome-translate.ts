// The translators that the adapters give to host chrome: `useChromeTranslate(overrides)` and the
// `t` prop of each chrome part. After docx-editor/packages/react/src/i18n/useChromeTranslate.ts
// at 568ccf7c (EigenPal), which PPTX Editor copies with its `ChromeTranslate` type in its core
// editor lane (pptx-editor/packages/core/src/editor at 70052cf1). Here the logic is in core, so
// the React hook and the Vue composable share it.
//
// Intended differences from DOCX Editor:
// - An override's value is a message, formatted with the placeholder and plural rules of the
//   catalog (`formatMessage`). DOCX Editor returns it as a literal string, because its parts pass
//   no values for the keys a host renames. Here many messages take values, such as
//   `zoom.resetAt` with `{level}`, so a literal override would show its placeholders.
// - A part's `t` falls back to the inherited translator for each key that it leaves unresolved
//   (a result that is not a string), as DOCX Editor's `useToolbarLabelFor` does with `??`, so
//   the result type is `string | undefined`, where DOCX Editor writes `string`. A `t` that
//   throws falls back too.

import { formatMessage, type TFunction, type TranslationKey } from '@xlsx-editor.dev/i18n';

/**
 * A chrome label resolver: a key and the values of its placeholders. It is the type of each
 * chrome part's `t` prop and of `useChromeTranslate()`'s result. It takes `string`, not the
 * `TranslationKey` union, so a host can resolve its own keys with the same function. A part
 * reads its inherited catalog for a key that `t` returns `undefined` for.
 *
 * @public
 */
export type ChromeTranslate = (
  key: string,
  params?: Readonly<Record<string, string | number>>
) => string | undefined;

/**
 * The translator `base` with `overrides` first. It always returns a string, so it is a
 * `ChromeTranslate` that host code can also call for its own labels. An override's value is a
 * message: its placeholders and plural branches are filled from the values of the call, with
 * the plural rules of `lang`. A key that `overrides` does not hold reads `base`. `overrides` is a `Map`, because
 * its keys come from the caller: an object would answer `constructor` from its prototype.
 *
 * @internal
 */
export function chromeTranslator(
  base: TFunction,
  overrides: ReadonlyMap<string, string> | undefined,
  lang: string
): (key: string, params?: Readonly<Record<string, string | number>>) => string {
  return (key, params) => {
    const message = overrides?.get(key);
    const values = params ? { ...params } : undefined;
    return typeof message === 'string'
      ? formatMessage(message, values, lang)
      : base(key as TranslationKey, values);
  };
}

/**
 * The translator of a chrome part with a `t` prop: `t` first, and `inherited` for each key that
 * `t` leaves unresolved, so a raw key never reaches the screen. A `t` that throws reads
 * `inherited` too, so a host's mistake does not unmount the editor: `onError` receives the key
 * and the error, for a warning in development. Without `t`, `inherited`.
 *
 * @internal
 */
export function partTranslator(
  inherited: TFunction,
  t: ChromeTranslate | undefined,
  onError?: (key: string, error: unknown) => void
): TFunction {
  if (!t) return inherited;
  return (key, params) => {
    let text: unknown;
    try {
      text = t(key, params);
    } catch (error) {
      onError?.(key, error);
      return inherited(key, params);
    }
    return typeof text === 'string' ? text : inherited(key, params);
  };
}

/**
 * The development warning for a part whose `t` threw: once for each part name in `warned`, so a
 * `t` that throws for every key, or a new `t` on each render, prints one line. Each adapter
 * keeps its own `warned` set. Returns the `onError` of `partTranslator`.
 *
 * @internal
 */
export function warnPartTranslateError(
  part: string,
  warned: Set<string>,
  warn: (message: string) => void
): (key: string, error: unknown) => void {
  return (key, error) => {
    if (warned.has(part)) return;
    warned.add(part);
    const why = error instanceof Error ? error.message : String(error);
    warn(
      `[xlsx-editor] The t prop of ${part} threw for the key "${key}" (${why}), so the part ` +
        'shows the text of the catalog for each key that throws.'
    );
  };
}
