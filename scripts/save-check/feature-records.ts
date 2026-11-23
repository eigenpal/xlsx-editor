// The save-verified records of the feature edits: one record per support matrix row with a
// recipe in `scripts/fidelity/feature-edits.ts`, committed in `fixtures/feature-edits.json`. A
// record holds the SHA-256 of the edited copy that the local save check opened, and the result.
//
// The recipes are deterministic, so CI runs them again and compares the copy hashes with the
// records, as `records.ts` does for the fixture copies. A change that changes the bytes of a
// copy makes its record stale: open the new copy, and record it again.
//
// Same role as SAVE_VERIFIED_FEATURE_EDITS in pptx-editor/docs/site/data/save-verified-fixtures.ts
// at 30dc48ea (EigenPal). Differences: the record holds the SHA-256 of the copy's bytes, not a
// canonical fingerprint of its parts, and also the result, as `records.ts` does.

import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { runFeatureEdit, type FeatureEditRun } from '../fidelity/feature-edit-kit.ts';
import { FEATURE_EDITS } from '../fidelity/feature-edits.ts';
import { sha256, type CopyRecord } from './records.ts';

const ROOT = join(import.meta.dir, '..', '..');
export const FEATURE_RECORDS = join(ROOT, 'fixtures', 'feature-edits.json');

export function readFeatureRecords(): Record<string, CopyRecord> {
  if (!existsSync(FEATURE_RECORDS)) return {};
  const parsed = JSON.parse(readFileSync(FEATURE_RECORDS, 'utf8')) as {
    records: Record<string, CopyRecord>;
  };
  // Row ids come from this repository, but keep the lookups free of inherited keys anyway.
  return Object.assign(Object.create(null) as Record<string, CopyRecord>, parsed.records);
}

export function writeFeatureRecords(records: Record<string, CopyRecord>): void {
  const sorted = Object.fromEntries(Object.entries(records).sort(([a], [b]) => (a < b ? -1 : 1)));
  writeFileSync(FEATURE_RECORDS, `${JSON.stringify({ version: 1, records: sorted }, null, 2)}\n`);
}

/** The problem of a row whose edit the save check has not opened. */
export const NO_RECORD = 'no save-check record';

/** A recipe that throws, as a run with that one problem. */
async function safeRun(row: string): Promise<FeatureEditRun | { readonly problems: string[] }> {
  try {
    return await runFeatureEdit(FEATURE_EDITS[row]!);
  } catch (e) {
    return { problems: [(e as Error).message] };
  }
}

/** Run every recipe once, in row order. */
export async function runAllFeatureEdits(): Promise<
  ReadonlyMap<string, FeatureEditRun | { readonly problems: readonly string[] }>
> {
  const out = new Map<string, FeatureEditRun | { readonly problems: readonly string[] }>();
  for (const row of Object.keys(FEATURE_EDITS).sort()) out.set(row, await safeRun(row));
  return out;
}

/**
 * Why the feature edit of `row` is not save-verified, or an empty list. `run` is the result of
 * its recipe. A missing record is one problem, so a caller can tell it from a stale one.
 */
export function featureRecordProblems(
  row: string,
  run: FeatureEditRun | { readonly problems: readonly string[] },
  records: Record<string, CopyRecord> = readFeatureRecords()
): string[] {
  if (run.problems.length > 0) return run.problems.map((p) => `the recipe fails: ${p}`);
  const record = Object.hasOwn(records, row) ? records[row] : undefined;
  if (!record) return [NO_RECORD];
  const problems: string[] = [];
  if (!('bytes' in run) || record.sha256 !== sha256(run.bytes))
    problems.push('the edited copy changed since the check opened it (stale)');
  if (record.result !== 'verified') problems.push(`the edited copy: ${record.result}`);
  return problems;
}
