import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { newLocaleSteps } from '../lib/i18n-new-steps.mjs';

const readme = readFileSync(resolve(import.meta.dir, '../../packages/i18n/README.md'), 'utf8');

test('the printed steps name each command and file that the i18n README requires', () => {
  const steps = newLocaleSteps('de').join('\n');
  for (const item of [
    'scripts/lib-packages.mjs',
    'bun run build:packages',
    'node scripts/api-extractor.mjs --local --record-forgotten-exports',
    'scripts/api-forgotten-exports.baseline.json',
    'bun run i18n:status',
    'docs/api/i18n/',
  ]) {
    expect(readme).toContain(item);
    expect(steps).toContain(item);
  }
});

test('the printed steps name the new locale', () => {
  const steps = newLocaleSteps('pt-BR').join('\n');
  expect(steps).toContain('packages/i18n/pt-BR.json');
  expect(steps).toContain("Add 'pt-BR' to the entries");
});
