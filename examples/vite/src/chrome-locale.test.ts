import { describe, expect, test } from 'bun:test';
import { chromeLocale } from './chrome-locale.ts';

describe('chromeLocale', () => {
  test('a valid tag gives a catalog with that tag', () => {
    expect(chromeLocale('?lang=ar')).toEqual({ _lang: 'ar' });
    expect(chromeLocale('?theme=dark&lang=he-il')).toEqual({ _lang: 'he-IL' });
  });

  test('no tag, or a tag that does not parse, gives no catalog', () => {
    expect(chromeLocale('')).toBeUndefined();
    expect(chromeLocale('?lang=')).toBeUndefined();
    expect(chromeLocale('?lang=not a tag')).toBeUndefined();
  });
});
