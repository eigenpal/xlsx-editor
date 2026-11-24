// The chrome catalog of the React and Vue demos, from the page's query: `?lang=<tag>` sets the language tag of
// the catalog, such as `?lang=ar` for a right-to-left chrome. The strings stay English (the
// fallback of a partial catalog); the direction and the plural rules follow the tag. A tag that
// does not parse is ignored, so the chrome stays English and left to right.

import type { Translations } from '@xlsx-editor.dev/i18n';

/** The catalog for the query `search`, or `undefined` when it names no valid language tag. */
export function chromeLocale(search: string): Translations | undefined {
  const asked = new URLSearchParams(search).get('lang');
  if (!asked) return undefined;
  try {
    const [tag] = Intl.getCanonicalLocales(asked);
    return tag ? { _lang: tag } : undefined;
  } catch {
    return undefined;
  }
}
