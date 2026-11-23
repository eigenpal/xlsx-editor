// CI's half of the save check (design D5a): every fixture has a record whose copy hashes match
// the copies the writer emits now, and both copies opened with no repair prompt.

import { describe, expect, test } from 'bun:test';
import { readdirSync } from 'node:fs';
import { join } from 'node:path';
import { FIXTURES } from '../fixtures/catalog.ts';
import { fixtureCopies, readRecords, recordProblems } from './records.ts';

describe('save-verified records', () => {
  test('every fixture is save-verified by a current record', () => {
    expect(FIXTURES.flatMap((f) => recordProblems(f.name))).toEqual([]);
  });

  test('no record names a fixture that does not exist', () => {
    const names = new Set(FIXTURES.map((f) => f.name));
    expect(Object.keys(readRecords()).filter((n) => !names.has(n))).toEqual([]);
  });

  test('every workbook in fixtures/ is in the catalog', () => {
    const names = new Set(FIXTURES.map((f) => `${f.name}.xlsx`));
    const files = readdirSync(join(import.meta.dir, '..', '..', 'fixtures')).filter((f) =>
      f.endsWith('.xlsx')
    );
    expect(files.filter((f) => !names.has(f))).toEqual([]);
  });

  test('a missing, stale, or failed record is reported', () => {
    expect(recordProblems('basic-values', {})).toEqual(['basic-values: no save-check record']);
    const stale = {
      saved: { sha256: '0', result: 'verified' },
      edited: { sha256: '0', result: 'repair-prompt' },
    } as const;
    expect(recordProblems('basic-values', { 'basic-values': stale })).toHaveLength(3);
  });

  test('the copies are deterministic', () => {
    const fixture = FIXTURES[0]!;
    expect(fixtureCopies(fixture).edited).toEqual(fixtureCopies(fixture).edited);
  });
});
