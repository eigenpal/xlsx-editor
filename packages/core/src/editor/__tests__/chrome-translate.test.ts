// The translators of host chrome: overrides first with the catalog's placeholder and plural
// rules, a part's `t` with the inherited translator for each key that it leaves unresolved, and
// partial catalogs whose missing and `null` keys inherit English.

import { describe, expect, test } from 'bun:test';
import {
  createT,
  deepMerge,
  en,
  type LocaleStrings,
  type TFunction,
  type Translations,
} from '@xlsx-editor.dev/i18n';
import {
  chromeTranslator,
  partTranslator,
  warnPartTranslateError,
  type ChromeTranslate,
} from '../chrome-translate.ts';

const english = createT(en);
const catalogOf = (partial: Translations): TFunction =>
  createT(
    deepMerge(en as Record<string, unknown>, partial as Record<string, unknown>) as LocaleStrings,
    typeof partial._lang === 'string' ? partial._lang : 'en'
  );

describe('chromeTranslator', () => {
  test('an override wins, and every other key reads the base catalog', () => {
    const t = chromeTranslator(english, new Map([['zoom.zoomIn', 'Larger']]), 'en');
    expect(t('zoom.zoomIn')).toBe('Larger');
    expect(t('zoom.zoomOut')).toBe('Zoom out');
  });

  test('an override fills its placeholders from the values of the call', () => {
    const t = chromeTranslator(
      english,
      new Map([['zoom.resetAt', 'Back to 100 from {level}']]),
      'en'
    );
    expect(t('zoom.resetAt', { level: '150%' })).toBe('Back to 100 from 150%');
    // The base path fills placeholders too.
    expect(t('grid.label', { sheet: 'Sales' })).toBe('Worksheet Sales');
  });

  test('an override picks its plural branch with the rules of the language', () => {
    const message =
      '{count, plural, =0 {no sheet} one {# sheet} few {# sheets (few)} other {# sheets}}';
    const t = chromeTranslator(english, new Map([['host.sheets', message]]), 'en');
    expect(t('host.sheets', { count: 0 })).toBe('no sheet');
    expect(t('host.sheets', { count: 1 })).toBe('1 sheet');
    expect(t('host.sheets', { count: 3 })).toBe('3 sheets');
    // Polish has a `few` category: 3 is few, 5 is many, which falls back to `other`.
    const polish = chromeTranslator(english, new Map([['host.sheets', message]]), 'pl');
    expect(polish('host.sheets', { count: 3 })).toBe('3 sheets (few)');
    expect(polish('host.sheets', { count: 5 })).toBe('5 sheets');
  });

  test('a key on the prototype chain resolves to itself, never to a function', () => {
    const t = chromeTranslator(english, new Map([['zoom.zoomIn', 'Larger']]), 'en');
    for (const key of ['constructor', 'toString', '__proto__', 'constructor.name']) {
      expect(t(key)).toBe(key);
    }
  });

  test('without overrides it is the base catalog', () => {
    const t = chromeTranslator(english, undefined, 'en');
    expect(t('zoom.zoomIn')).toBe(english('zoom.zoomIn'));
  });

  test('the values of the call are not changed', () => {
    const params = Object.freeze({ level: '50%' });
    const t = chromeTranslator(english, new Map([['zoom.resetAt', 'At {level}']]), 'en');
    expect(t('zoom.resetAt', params)).toBe('At 50%');
  });
});

describe('partTranslator', () => {
  test('without t it is the inherited translator', () => {
    expect(partTranslator(english, undefined)).toBe(english);
  });

  test('t wins and receives the key and its values', () => {
    const seen: unknown[] = [];
    const t: ChromeTranslate = (key, params) => {
      seen.push([key, params]);
      return key === 'zoom.resetAt' ? `Reset (${params?.['level']})` : `[${key}]`;
    };
    const part = partTranslator(english, t);
    expect(part('zoom.resetAt', { level: '80%' })).toBe('Reset (80%)');
    expect(part('zoom.zoomIn')).toBe('[zoom.zoomIn]');
    expect(seen).toEqual([
      ['zoom.resetAt', { level: '80%' }],
      ['zoom.zoomIn', undefined],
    ]);
  });

  test('a key that t leaves unresolved reads the inherited translator', () => {
    const t: ChromeTranslate = (key) => (key === 'zoom.zoomIn' ? 'Larger' : undefined);
    const part = partTranslator(english, t);
    expect(part('zoom.zoomIn')).toBe('Larger');
    expect(part('zoom.zoomOut')).toBe('Zoom out');
    expect(part('grid.label', { sheet: 'A' })).toBe('Worksheet A');
  });

  test('a t that throws reads the inherited translator and reports the key', () => {
    const errors: [string, unknown][] = [];
    const part = partTranslator(
      english,
      (key) => {
        if (key === 'zoom.zoomIn') throw new Error('broken');
        return 'Fine';
      },
      (key, error) => errors.push([key, error])
    );
    expect(part('zoom.zoomIn')).toBe('Zoom in');
    expect(part('zoom.zoomOut')).toBe('Fine');
    expect(errors.map(([key]) => key)).toEqual(['zoom.zoomIn']);
    expect((errors[0]![1] as Error).message).toBe('broken');
  });

  test('the warning for a t that throws prints once for each part name', () => {
    const lines: string[] = [];
    const warned = new Set<string>();
    const push = (line: string) => void lines.push(line);
    warnPartTranslateError('ZoomControl', warned, push)('zoom.zoomIn', new Error('broken'));
    warnPartTranslateError('ZoomControl', warned, push)('zoom.zoomOut', new Error('broken'));
    warnPartTranslateError('SheetTabs', warned, push)('sheetTabs.add', 'text');
    expect(lines).toEqual([
      '[xlsx-editor] The t prop of ZoomControl threw for the key "zoom.zoomIn" (broken), so the ' +
        'part shows the text of the catalog for each key that throws.',
      '[xlsx-editor] The t prop of SheetTabs threw for the key "sheetTabs.add" (text), so the ' +
        'part shows the text of the catalog for each key that throws.',
    ]);
  });

  test('a part over chromeTranslator keeps the overrides and the catalog', () => {
    const german = catalogOf({ _lang: 'de', zoom: { zoomOut: 'Verkleinern' } });
    const part = partTranslator(
      german,
      chromeTranslator(german, new Map([['zoom.zoomIn', 'Vergrößern']]), 'de')
    );
    expect(part('zoom.zoomIn')).toBe('Vergrößern');
    expect(part('zoom.zoomOut')).toBe('Verkleinern');
    expect(part('zoom.reset')).toBe(english('zoom.reset'));
  });
});

describe('partial catalogs', () => {
  test('a missing key and a null key inherit English', () => {
    const t = catalogOf({
      _lang: 'de',
      zoom: { zoomIn: 'Vergrößern', zoomOut: null },
    } as Translations);
    expect(t('zoom.zoomIn')).toBe('Vergrößern');
    expect(t('zoom.zoomOut')).toBe('Zoom out');
    expect(t('zoom.reset')).toBe(english('zoom.reset'));
  });

  test('an unknown key returns the key', () => {
    expect(chromeTranslator(english, undefined, 'en')('host.unknown')).toBe('host.unknown');
  });
});
