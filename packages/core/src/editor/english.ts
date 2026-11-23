// The translator an editor uses when the host passes none: the English catalog of the shared
// strings, as DOCX Editor's core falls back to `createT(en)`. So assistive technology reads
// English labels, never catalog keys.

import { createT, en, type TranslationKey } from '@xlsx-editor.dev/i18n';
import type { EditorTranslate } from './cell-input.ts';

const english = createT(en);

/** The English text of a catalog key, with its parameters. */
export const englishText: EditorTranslate = (key, params) =>
  english(key as TranslationKey, params ? { ...params } : undefined);
