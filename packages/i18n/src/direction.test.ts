// Adapted from docx-editor/packages/i18n/src/direction.test.ts at 0c907463 (EigenPal).
// Difference: only English ships here, so the right-to-left catalogs are written inline.
import { describe, expect, test } from 'bun:test';
import { catalogDirection, chromeDirection, en, localeDirection } from './index.ts';

const arabic = { _lang: 'ar', _dir: 'rtl' } as const;
const spanish = { _lang: 'es' } as const;

describe('catalogDirection', () => {
  test('a catalog that declares its direction keeps it', () => {
    expect(catalogDirection(arabic)).toBe('rtl');
    expect(catalogDirection(spanish)).toBe('ltr');
    expect(catalogDirection(en)).toBe('ltr');
  });

  test('_dir decides over _lang', () => {
    expect(catalogDirection({ _lang: 'en', _dir: 'rtl' })).toBe('rtl');
    expect(catalogDirection({ _lang: 'ar', _dir: 'ltr' })).toBe('ltr');
    expect(catalogDirection({ _lang: 'fa' })).toBe('rtl');
    expect(catalogDirection({ _dir: 'sideways' })).toBeUndefined();
  });

  test('a catalog without direction or language has none', () => {
    expect(catalogDirection({})).toBeUndefined();
    expect(catalogDirection(undefined)).toBeUndefined();
  });
});

describe('localeDirection', () => {
  test('right-to-left languages', () => {
    for (const tag of ['ar', 'ar-SA', 'he', 'he-IL', 'iw', 'fa-IR', 'ur', 'ps', 'ckb', 'yi']) {
      expect(localeDirection(tag)).toBe('rtl');
    }
  });

  test('left-to-right languages and missing tags', () => {
    for (const tag of ['en', 'en-US', 'de', 'pt-BR', 'zh-CN', 'ja', 'hi', '', 'arab']) {
      expect(localeDirection(tag)).toBe('ltr');
    }
    expect(localeDirection(undefined)).toBe('ltr');
    expect(localeDirection(null)).toBe('ltr');
  });

  test('a script subtag decides over the language', () => {
    expect(localeDirection('az-Arab')).toBe('rtl');
    expect(localeDirection('pa_Arab_PK')).toBe('rtl');
    expect(localeDirection('ku-Latn')).toBe('ltr');
    expect(localeDirection('uz-Cyrl')).toBe('ltr');
    expect(localeDirection('AR')).toBe('rtl');
  });
});

describe('chromeDirection', () => {
  test('the dir prop decides first, then _dir, then a right-to-left _lang', () => {
    expect(chromeDirection('ltr', arabic, 'rtl')).toBe('ltr');
    expect(chromeDirection(undefined, { _dir: 'ltr', _lang: 'ar' }, 'rtl')).toBe('ltr');
    expect(chromeDirection(undefined, arabic, undefined)).toBe('rtl');
    expect(chromeDirection(undefined, { _lang: 'fa' }, undefined)).toBe('rtl');
  });

  test('a left-to-right _lang alone keeps the inherited or host direction', () => {
    expect(chromeDirection(undefined, spanish, undefined)).toBeUndefined();
    expect(chromeDirection(undefined, spanish, 'rtl')).toBe('rtl');
    expect(chromeDirection(undefined, en, undefined)).toBeUndefined();
    expect(chromeDirection(undefined, undefined, undefined)).toBeUndefined();
  });
});
