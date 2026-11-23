// The save-verified records of design D5a: one record per fixture, committed in
// `fixtures/save-check.json`. A record holds the SHA-256 of the saved copy and of the edited copy
// that the local save check opened, and the result for each one.
//
// The writer is deterministic, so CI writes both copies again and compares their hashes with
// the record. Any change that changes the saved bytes makes the record stale: the writer, the
// reader, the store ops, the fixture, its edit, or a dependency such as the deflate library.
// A change that leaves the bytes as they were keeps the record current, because the check
// opened those same bytes.

import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { readXlsx, saveXlsx, WorkbookDocumentStore } from '../../packages/core/src/store/index.ts';
import { FIXTURES, fixtureEditOps, type Fixture } from '../fixtures/catalog.ts';

// The zip writer stores entry times as local-time fields, so the copy bytes depend on the time
// zone. The fixture build runs in UTC for the same reason (scripts/fixtures/build.ts).
process.env.TZ = 'UTC';

const ROOT = join(import.meta.dir, '..', '..');
export const RECORDS = join(ROOT, 'fixtures', 'save-check.json');

export const RESULTS = ['verified', 'repair-prompt', 'does-not-open', 'sheets-differ'] as const;
export type SaveResult = (typeof RESULTS)[number];

/** One opened copy: the SHA-256 of its bytes and what the check saw. */
export interface CopyRecord {
  readonly sha256: string;
  readonly result: SaveResult;
}

export interface SaveRecord {
  readonly saved: CopyRecord;
  readonly edited: CopyRecord;
}

export const sha256 = (bytes: Uint8Array): string =>
  createHash('sha256').update(bytes).digest('hex');

export function readRecords(): Record<string, SaveRecord> {
  if (!existsSync(RECORDS)) return {};
  return (JSON.parse(readFileSync(RECORDS, 'utf8')) as { records: Record<string, SaveRecord> })
    .records;
}

export function writeRecords(records: Record<string, SaveRecord>): void {
  const sorted = Object.fromEntries(Object.entries(records).sort(([a], [b]) => (a < b ? -1 : 1)));
  writeFileSync(RECORDS, `${JSON.stringify({ version: 2, records: sorted }, null, 2)}\n`);
}

/** The saved and the edited copy of a fixture, as the editor writes them. */
export function fixtureCopies(fixture: Fixture): { saved: Uint8Array; edited: Uint8Array } {
  const read = readXlsx(
    new Uint8Array(readFileSync(join(ROOT, 'fixtures', `${fixture.name}.xlsx`)))
  );
  if (!read.ok) throw new Error(`${fixture.name}: does not open (${read.error.reason})`);
  const store = new WorkbookDocumentStore(read.pkg);
  const r = store.transact(fixtureEditOps(fixture, read.pkg));
  if (!r.ok) throw new Error(`${fixture.name}: the edit is refused (${r.rejection.detail})`);
  return { saved: saveXlsx(read.pkg), edited: saveXlsx(store.package) };
}

/** Why a fixture is not save-verified, or an empty list. */
export function recordProblems(name: string, records = readRecords()): string[] {
  const record = Object.hasOwn(records, name) ? records[name] : undefined;
  if (!record) return [`${name}: no save-check record`];
  const fixture = FIXTURES.find((f) => f.name === name);
  if (!fixture) return [`${name}: no fixture`];
  const copies = fixtureCopies(fixture);
  const problems: string[] = [];
  for (const kind of ['saved', 'edited'] as const) {
    if (record[kind].sha256 !== sha256(copies[kind]))
      problems.push(`${name}: the ${kind} copy changed since the check opened it (stale)`);
    if (record[kind].result !== 'verified')
      problems.push(`${name}: ${kind} copy ${record[kind].result}`);
  }
  return problems;
}

/** Write the copies of every fixture to `dir`, for the local check: `<name>.saved.xlsx`, `<name>.edited.xlsx`. */
export function writeCopies(dir: string): string[] {
  mkdirSync(dir, { recursive: true });
  for (const fixture of FIXTURES) {
    const copies = fixtureCopies(fixture);
    writeFileSync(join(dir, `${fixture.name}.saved.xlsx`), copies.saved);
    writeFileSync(join(dir, `${fixture.name}.edited.xlsx`), copies.edited);
  }
  return FIXTURES.map((f) => f.name);
}
