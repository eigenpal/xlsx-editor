/** @jsxImportSource react */
// The language of the chrome below a root: `LocaleProvider`, `useChromeTranslate`,
// `useChromeDir`, and the `t` prop of the chrome parts. After
// docx-editor/packages/react/src/i18n/LocaleContext.tsx and useChromeTranslate.ts at 0c907463
// and 568ccf7c (EigenPal), and the `t` prop wrapper of
// pptx-editor/packages/react/src/headless-hooks.tsx at 70052cf1 (`withChromeTranslate`). The
// wrapper is PPTX Editor's because DOCX Editor threads `t` through each part's own context by
// hand; the parts here read every string through `useTranslation()`, so one provider per part
// gives each part, and the parts and dialog inside it, the same rule. The Vue twin is
// `locale.ts` of the Vue adapter.
//
// Intended differences from DOCX Editor:
// - One context holds the merged catalog, its language, and both directions, where DOCX Editor
//   uses four. The translator stays in `TranslationContext`, the public context of this adapter.
// - `XlsxEditor.Root` is a provider too: its `locale` prop merges onto the catalog of an outer
//   `LocaleProvider`, as the `i18n` prop of DOCX Editor's packaged editor does. Root has no `dir`
//   prop; wrap it in `LocaleProvider` with `dir`, as in DOCX Editor.
// - A `LocaleProvider` with no `i18n` keeps the translator it inherits, also the `t` of a part
//   above it. One with `i18n` builds its translator from the merged catalog.
// - `useChromeTranslate` formats an override with the catalog's placeholder and plural rules
//   (`chromeTranslator` in core), and reads the nearest part's `t` first, as PPTX Editor's does.
// - `useTranslation()` keeps this adapter's `{ t }` shape. The direction is read with
//   `useChromeDir()`; DOCX Editor also returns `catalogue`, `lang`, and `dir`.
// - The parts set `dir` on their own element, and the grid gets none, so a right-to-left
//   catalog never mirrors the cells. The packaged editor also sets it on its footer.

import {
  createElement,
  useContext,
  useMemo,
  useRef,
  type ReactElement,
  type ReactNode,
} from 'react';
import {
  catalogDirection,
  chromeDirection,
  createT,
  deepMerge,
  type LocaleStrings,
  type Translations,
} from '@xlsx-editor.dev/i18n';
import {
  chromeTranslator,
  partTranslator,
  warnPartTranslateError,
  type ChromeTranslate,
} from '@xlsx-editor.dev/core/editor';
import { CatalogContext, TranslationContext, type CatalogScope } from './context.ts';
import { isDevelopment } from './is-development.ts';

/**
 * The catalog scope of a provider with `i18n` and `dir`, merged onto the inherited one. It
 * follows the content of `i18n`, not its identity, so an inline catalog builds no new
 * translator on each render. Internal.
 */
export function useCatalogScope(
  i18n: Translations | undefined,
  dir: 'ltr' | 'rtl' | undefined
): CatalogScope {
  const inherited = useContext(CatalogContext);
  const key = i18n ? JSON.stringify(i18n) : '';
  const memo = useRef<{
    key: string;
    inherited: CatalogScope;
    dir: 'ltr' | 'rtl' | undefined;
    scope: CatalogScope;
  } | null>(null);
  const last = memo.current;
  if (last && last.key === key && last.inherited === inherited && last.dir === dir)
    return last.scope;
  const scope = scopeOf(inherited, i18n, dir);
  memo.current = { key, inherited, dir, scope };
  return scope;
}

function scopeOf(
  inherited: CatalogScope,
  i18n: Translations | undefined,
  dir: 'ltr' | 'rtl' | undefined
): CatalogScope {
  // Merged onto the INHERITED catalog, not onto English, as in DOCX Editor: a provider in
  // another composes with it. At the top, the inherited catalog is English.
  const strings = i18n
    ? (deepMerge(
        inherited.strings as Record<string, unknown>,
        i18n as Record<string, unknown>
      ) as LocaleStrings)
    : inherited.strings;
  const lang = typeof i18n?._lang === 'string' ? i18n._lang : inherited.lang;
  const chromeDir = chromeDirection(dir, i18n, inherited.chromeDir);
  const direction = chromeDir ?? catalogDirection(i18n) ?? inherited.dir;
  const t = i18n ? createT(strings, lang) : inherited.t;
  if (
    strings === inherited.strings &&
    lang === inherited.lang &&
    chromeDir === inherited.chromeDir &&
    direction === inherited.dir
  )
    return inherited;
  return { strings, lang, dir: direction, chromeDir, t };
}

/** Props for `LocaleProvider`. @public */
export interface LocaleProviderProps {
  /**
   * A partial catalog over the catalog above: the root's `locale`, an outer provider, or
   * English. A missing or `null` entry inherits. `_lang` sets the plural rules.
   */
  readonly i18n?: Translations;
  /**
   * The writing direction of the chrome below. Defaults to the catalog's `_dir`, then to a
   * right-to-left `_lang`, then to the direction of an outer provider. `'rtl'` mirrors the
   * toolbar, the formula bar, the sheet tabs, the zoom control, the status message, and the
   * dialog. `'ltr'` keeps a left-to-right chrome on a right-to-left host page. Without any of
   * these, the chrome inherits the direction of the host page. The grid never mirrors.
   */
  readonly dir?: 'ltr' | 'rtl';
  /** The region whose chrome the provider changes. */
  readonly children?: ReactNode;
}

/**
 * Changes the strings or the direction of the chrome below it, as DOCX Editor's
 * `LocaleProvider`. Inside `XlsxEditor.Root`, it relabels a region, such as one toolbar.
 * Around the root, it sets the catalog that the root's `locale` merges onto, and your own
 * controls outside the root read it with `useTranslation()`.
 *
 * @example
 * ```tsx
 * <LocaleProvider i18n={{ formattingBar: { bold: 'Heavy' } }}>
 *   <XlsxEditor.Toolbar />
 * </LocaleProvider>
 * ```
 *
 * @public
 */
export function LocaleProvider({ i18n, dir, children }: LocaleProviderProps): ReactElement {
  const scope = useCatalogScope(i18n, dir);
  const inheritedT = useContext(TranslationContext);
  // Without a catalog, the translator above stays, also the `t` of a part.
  const t = i18n ? scope.t : inheritedT;
  return (
    <CatalogContext.Provider value={scope}>
      <TranslationContext.Provider value={t}>{children}</TranslationContext.Provider>
    </CatalogContext.Provider>
  );
}

/**
 * The chrome's translator, ready to pass as the `t` prop of any chrome part, as DOCX Editor's
 * `useChromeTranslate`: the `t` of the nearest part that has one, else the nearest provider's
 * catalog. `overrides` wins for the keys it holds. An override is a message, as in the catalog:
 * `{name}` placeholders and `{count, plural, ...}` branches are filled from the values that the
 * part passes.
 *
 * Keep the `overrides` map stable (module level or memoized): the returned function follows its
 * identity, and a new function renders every part that receives it again.
 *
 * @example
 * ```tsx
 * const LABELS = new Map([['zoom.resetAt', 'Zoom {level}. Press to reset.']]);
 *
 * function Footer() {
 *   const t = useChromeTranslate(LABELS);
 *   return <XlsxEditor.ZoomControl t={t} />;
 * }
 * ```
 *
 * @public
 */
export function useChromeTranslate(
  overrides?: ReadonlyMap<string, string>
): (key: string, params?: Readonly<Record<string, string | number>>) => string {
  const t = useContext(TranslationContext);
  const { lang } = useContext(CatalogContext);
  return useMemo(() => chromeTranslator(t, overrides, lang), [t, overrides, lang]);
}

/**
 * The `dir` that a chrome part sets on its root, as DOCX Editor's `useChromeDir`: the direction
 * a provider decided (its `dir` prop, the catalog's `_dir`, or a right-to-left `_lang`), or
 * `undefined`, so the part keeps the direction of the host page. Set it on your own controls to
 * follow the editor chrome.
 *
 * @public
 */
export function useChromeDir(): 'ltr' | 'rtl' | undefined {
  return useContext(CatalogContext).chromeDir;
}

/**
 * Give a chrome part DOCX Editor's `t` prop: with `t`, every string of the part and of the
 * parts inside it comes from `t` first, and from the inherited translator for each key that
 * `t` leaves unresolved. The provider always renders, so a change of `t` between set and unset
 * keeps the part mounted. Internal.
 */
export function withChromeTranslate<P extends { readonly t?: ChromeTranslate | undefined }>(
  Component: (props: Omit<P, 't'>) => ReactNode,
  name: string
): (props: P) => ReactElement {
  function Translated(props: P): ReactElement {
    const { t, ...rest } = props;
    const inherited = useContext(TranslationContext);
    const value = useMemo(
      () =>
        partTranslator(
          inherited,
          t,
          warnPartTranslateError(name, (message) => {
            if (isDevelopment()) console.warn(message);
          })
        ),
      [inherited, t]
    );
    return (
      <TranslationContext.Provider value={value}>
        {createElement(Component, rest as Omit<P, 't'>)}
      </TranslationContext.Provider>
    );
  }
  Object.defineProperty(Translated, 'name', { value: name });
  return Translated;
}

export type { ChromeTranslate };
