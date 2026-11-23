import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { FONT_SUBSTITUTES } from '../text/measurer.ts';

const readme = readFileSync(resolve(import.meta.dir, '../../../README.md'), 'utf8');

/** The rows of the table whose last column is "Stands in for", as family to lower-case names. */
function readmeTable(): Map<string, string[]> {
  const lines = readme.split(/\r?\n/);
  const header = lines.findIndex((line) =>
    /^\|\s*Files\s*\|\s*Family\s*\|\s*Stands in for/.test(line)
  );
  expect(header).toBeGreaterThan(-1);
  const rows = new Map<string, string[]>();
  for (const line of lines.slice(header + 2)) {
    if (!line.startsWith('|')) break;
    const [, , family, names] = line.split('|').map((cell) => cell.trim());
    rows.set(
      (family ?? '').replaceAll('`', ''),
      (names ?? '').split(',').map((name) => name.trim().toLowerCase())
    );
  }
  return rows;
}

describe('the font table of the core README', () => {
  test('lists each key of FONT_SUBSTITUTES under its first substitute', () => {
    const expected = new Map<string, string[]>();
    for (const [name, substitutes] of Object.entries(FONT_SUBSTITUTES)) {
      const family = substitutes[0] ?? '';
      expected.set(family, [...(expected.get(family) ?? []), name]);
    }
    const table = readmeTable();
    expect([...table.keys()].sort()).toEqual([...expected.keys()].sort());
    for (const [family, names] of expected) {
      expect([...(table.get(family) ?? [])].sort()).toEqual([...names].sort());
    }
  });
});
