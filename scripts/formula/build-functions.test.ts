// The generated set of predefined function names equals the reference file.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { functionNames, generatedModule, OUT, SOURCE } from './build-functions.ts';

describe('predefined function names', () => {
  test('the generated module equals the reference file', () => {
    const text = readFileSync(SOURCE, 'utf8');
    expect(functionNames(text).length).toBe(356);
    expect(readFileSync(OUT, 'utf8')).toBe(generatedModule(text));
  });
});
