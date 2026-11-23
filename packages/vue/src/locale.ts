// The language of the chrome below a root: the Vue twin of the React adapter's `locale.tsx`,
// after docx-editor/packages/vue/src/i18n/LocaleContext.ts and useChromeTranslate.ts at
// 0c907463 (EigenPal). `LocaleProvider`, `useChromeTranslate`, `useChromeDir`, and the `t` prop
// of the chrome parts (`usePartTranslation`) act as in React; the rules are in core
// (`chromeTranslator`, `partTranslator`) and in `@xlsx-editor.dev/i18n` (`chromeDirection`).
//
// Intended differences from DOCX Editor, besides those of the React twin:
// - `useChromeTranslate` returns the named `UseChromeTranslateReturn` (a computed ref), and
//   takes its overrides as a value, a ref, or a getter, as `useEditorCommand` takes its slot.
// - `useChromeDir` returns the named `UseChromeDirReturn`.
// - A part calls `usePartTranslation` in its `setup`, where React wraps the part
//   (`withChromeTranslate`): a Vue component does not see what it provides itself, so the part
//   reads the returned translator, and the parts inside it inject it.

import {
  computed,
  defineComponent,
  inject,
  provide,
  shallowRef,
  toValue,
  watch,
  type ComputedRef,
  type ExtractPublicPropTypes,
  type MaybeRefOrGetter,
  type PropType,
  type ShallowRef,
  type VNodeChild,
} from 'vue';
import {
  catalogDirection,
  chromeDirection,
  createT,
  deepMerge,
  type LocaleStrings,
  type TFunction,
  type Translations,
} from '@xlsx-editor.dev/i18n';
import {
  chromeTranslator,
  partTranslator,
  warnPartTranslateError,
  type ChromeTranslate,
} from '@xlsx-editor.dev/core/editor';
import {
  catalogKey,
  englishCatalog as english,
  useTranslator,
  xlsxTranslationKey,
  type CatalogScope,
} from './context.ts';
import { isDevelopment } from './is-development.ts';

function scopeOf(
  inherited: CatalogScope,
  i18n: Translations | undefined,
  dir: 'ltr' | 'rtl' | undefined
): CatalogScope {
  // Merged onto the INHERITED catalog, not onto English, as in DOCX Editor.
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

/**
 * Provide the catalog scope of `i18n` and `dir`, merged onto the inherited one, and return it.
 * It follows the content of `i18n`, not its identity, so an inline catalog builds no new
 * translator on each parent render. Internal.
 */
export function provideCatalogScope(
  i18n: () => Translations | undefined,
  dir: () => 'ltr' | 'rtl' | undefined
): Readonly<ShallowRef<CatalogScope>> {
  const inherited = inject(catalogKey, english);
  const scope = shallowRef(scopeOf(inherited.value, i18n(), dir()));
  watch([() => (i18n() ? JSON.stringify(i18n()) : ''), inherited, dir], () => {
    scope.value = scopeOf(inherited.value, i18n(), dir());
  });
  provide(catalogKey, scope);
  return scope;
}

/**
 * The parts whose throwing `t` has printed its development warning: one line for each part
 * name, also when the host passes a new `t` on each render.
 */
const warnedParts = new Set<string>();

const localeProviderProps = {
  /**
   * A partial catalog over the catalog above: the root's `locale`, an outer provider, or
   * English. A missing or `null` entry inherits. `_lang` sets the plural rules.
   */
  i18n: { type: Object as PropType<Translations>, default: undefined },
  /**
   * The writing direction of the chrome below. Defaults to the catalog's `_dir`, then to a
   * right-to-left `_lang`, then to the direction of an outer provider. `'rtl'` mirrors the
   * toolbar, the formula bar, the sheet tabs, the zoom control, the status message, and the
   * dialog. `'ltr'` keeps a left-to-right chrome on a right-to-left host page. Without any of
   * these, the chrome inherits the direction of the host page. The grid never mirrors.
   */
  dir: { type: String as PropType<'ltr' | 'rtl'>, default: undefined },
} as const;

/** Props for `LocaleProvider`, the twin of React's `LocaleProviderProps`. @public */
export type LocaleProviderProps = ExtractPublicPropTypes<typeof localeProviderProps>;

/**
 * Changes the strings or the direction of the chrome below it, as DOCX Editor's
 * `LocaleProvider`. Inside `XlsxEditorRoot`, it relabels a region, such as one toolbar. Around
 * the root, it sets the catalog that the root's `locale` merges onto, and your own controls
 * outside the root read it with `useTranslation()`. @public
 */
export const LocaleProvider = defineComponent({
  name: 'LocaleProvider',
  props: localeProviderProps,
  setup(props, { slots }) {
    const scope = provideCatalogScope(
      () => props.i18n,
      () => props.dir
    );
    const inherited = useTranslator();
    // Without a catalog, the translator above stays, also the `t` of a part.
    const pick = (): TFunction => (props.i18n ? scope.value.t : inherited.value);
    const t = shallowRef<TFunction>(pick());
    watch([scope, inherited, () => props.i18n === undefined], () => {
      t.value = pick();
    });
    provide(xlsxTranslationKey, t);
    return (): VNodeChild => slots.default?.();
  },
});

/** What `useChromeTranslate` returns: the chrome's translator, as a computed ref. @public */
export type UseChromeTranslateReturn = ComputedRef<
  (key: string, params?: Readonly<Record<string, string | number>>) => string
>;

/**
 * The chrome's translator, ready to pass as the `t` prop of any chrome part, the twin of
 * React's `useChromeTranslate`: the `t` of the nearest part that has one, else the nearest
 * provider's catalog. `overrides` wins for the keys it holds. An override is a message, as in
 * the catalog: `{name}` placeholders and `{count, plural, ...}` branches are filled from the
 * values that the part passes. `overrides` can be a map, a ref, or a getter.
 *
 * @example
 * ```vue
 * <script setup lang="ts">
 * const t = useChromeTranslate(new Map([['zoom.resetAt', 'Zoom {level}. Press to reset.']]));
 * </script>
 * <template><XlsxEditorZoomControl :t="t" /></template>
 * ```
 *
 * @public
 */
export function useChromeTranslate(
  overrides?: MaybeRefOrGetter<ReadonlyMap<string, string> | undefined>
): UseChromeTranslateReturn {
  const base = useTranslator();
  const scope = inject(catalogKey, english);
  return computed(() => chromeTranslator(base.value, toValue(overrides), scope.value.lang));
}

/** What `useChromeDir` returns. @public */
export type UseChromeDirReturn = ComputedRef<'ltr' | 'rtl' | undefined>;

/**
 * The `dir` that a chrome part sets on its root, the twin of React's `useChromeDir`: the
 * direction a provider decided (its `dir` prop, the catalog's `_dir`, or a right-to-left
 * `_lang`), or `undefined`, so the part keeps the direction of the host page. @public
 */
export function useChromeDir(): UseChromeDirReturn {
  const scope = inject(catalogKey, english);
  return computed(() => scope.value.chromeDir);
}

/** The translator of a part with a `t` prop, and its direction. Internal. */
export interface PartTranslation {
  /** Translate a key: `t` first, then the inherited translator. Reads the ref at each call. */
  readonly t: TFunction;
  /** The translator ref, for a watcher that follows a change. */
  readonly translator: Readonly<ShallowRef<TFunction>>;
  /** The `dir` of the part's root element. */
  readonly dir: UseChromeDirReturn;
}

/**
 * Give a chrome part DOCX Editor's `t` prop: the part and the parts inside it read `t` first,
 * and the inherited translator for each key that `t` leaves unresolved or throws for. Call it in
 * the part's `setup`, with the part's name and a getter of its `t` prop. Internal.
 */
export function usePartTranslation(
  part: string,
  t: () => ChromeTranslate | undefined
): PartTranslation {
  const inherited = useTranslator();
  const build = (): TFunction =>
    partTranslator(
      inherited.value,
      t(),
      warnPartTranslateError(part, warnedParts, (message) => {
        if (isDevelopment()) console.warn(message);
      })
    );
  const translator = shallowRef<TFunction>(build());
  watch([inherited, t], () => {
    translator.value = build();
  });
  provide(xlsxTranslationKey, translator);
  return {
    t: ((key, params) => translator.value(key, params)) as TFunction,
    translator,
    dir: useChromeDir(),
  };
}

export type { ChromeTranslate };
