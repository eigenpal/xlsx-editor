// The support matrix gate (task 7.7 of the active change). Every claim in
// `docs/site/data/support-matrix.ts` must stand on its evidence:
//
// - every cited fixture is in the catalog (`scripts/fixtures/catalog.ts`);
// - every save-verified fixture has a current record with both copies `verified` (design D5a,
//   `scripts/save-check/records.ts`), and a supported editing or save row cites one;
// - a supported row cites no save gap: an editing or save row with no save-verified fixture is
//   at most `partial`;
// - every fidelity result names a fixture with a current reference PDF (`referenceProblems`:
//   a record, an unchanged workbook, and an unchanged PDF) and a measurement, and its gaps equal
//   the known gaps of `scripts/fidelity/thresholds.ts` (design D8); a supported row declares
//   none, a supported or partial rendering row cites one, and a rendering row also lists each
//   fidelity fixture in `fixtures`;
// - every cited test file exists, and every row that is supported or partial cites one;
// - every cited approximation exists, a supported row cites none, a partial row cites one or a
//   known gap or a missing save check, and every approximation is cited by some row;
// - the row `save.round-trip` cites every fixture of the catalog in `saveVerified`;
// - every supported editing or save row has its own edit (`scripts/fidelity/feature-edits.ts`)
//   on a fixture that it cites, whose edited copy has a current `verified` record
//   (`fixtures/feature-edits.json`); a row without one is at most `partial`, as PPTX Editor
//   requires a recipe of each row that claims a full round trip. `save.round-trip` is the
//   exception: its own edits are the edits of the fixtures. Every recipe names a row.
//
// Each problem names its row and the remedy. Paths are compared with `/` separators, and the
// script sets its own time zone, so the gate gives the same result on every platform.
// The DOCX Editor matrix (docx-editor/docs/site/data/word-features.ts) has its checks in a unit
// test beside it; this gate adds the evidence ties, which that matrix does not type.

import { existsSync } from 'node:fs';
import { join } from 'node:path';
import {
  SUPPORT_CATEGORIES,
  SUPPORT_STATUSES,
  supportFeatures,
  type FidelityCheck,
  type SupportCategory,
  type SupportFeature,
} from '../docs/site/data/support-matrix.ts';
import { APPROXIMATIONS } from '../packages/core/src/layout/approximations.ts';
import { referenceProblems } from './fidelity/references.ts';
import { FAILURE_KINDS, KNOWN_GAPS, type FailureKind } from './fidelity/thresholds.ts';
import { FIXTURES } from './fixtures/catalog.ts';
import { recordProblems } from './save-check/records.ts';
import {
  featureRecordProblems,
  NO_RECORD,
  readFeatureRecords,
  runAllFeatureEdits,
} from './save-check/feature-records.ts';
import { FEATURE_EDITS } from './fidelity/feature-edits.ts';

// The matrix names the fidelity checks without importing the harness; keep the two equal.
type Same<A, B> = [A] extends [B] ? ([B] extends [A] ? true : false) : false;
const _sameChecks: Same<FidelityCheck, FailureKind> = true;
void _sameChecks;

// The save check rebuilds the saved copies, and the zip writer stores local-time entry times, so
// the copy hashes depend on the time zone. Set it here, not in a shell prefix that some shells
// cannot run. `scripts/save-check/records.ts` sets the same when it loads.
process.env.TZ = 'UTC';

const ROOT = join(import.meta.dir, '..');

export const RENDERING_CATEGORIES: ReadonlySet<SupportCategory> = new Set([
  'values',
  'formatting',
  'grid',
  'drawings',
]);
export const SAVED_CATEGORIES: ReadonlySet<SupportCategory> = new Set(['editing', 'save']);

/** What the gate compares the matrix with. Tests pass their own. */
export interface MatrixContext {
  readonly fixtures: ReadonlySet<string>;
  /** Why a fixture is not save-verified, or an empty list (`recordProblems`). */
  readonly saveProblems: (fixture: string) => readonly string[];
  /** Why a fixture's reference PDF is not current, or an empty list (`referenceProblems`). */
  readonly referenceProblems: (fixture: string) => readonly string[];
  /** Fixtures that each measured font environment checks. */
  readonly measured: ReadonlySet<string>;
  /** The known gaps of a fixture, over every measured font environment. */
  readonly knownGaps: (fixture: string) => ReadonlySet<string>;
  readonly approximations: ReadonlySet<string>;
  readonly fileExists: (relativePath: string) => boolean;
  /** The fixture of a row's feature edit (`scripts/fidelity/feature-edits.ts`), or undefined. */
  readonly featureEdit: (row: string) => string | undefined;
  /**
   * Why a row's feature edit is not save-verified (`featureRecordProblems`), or an empty list.
   * `NO_RECORD` alone means that the recipe passes and the save check has not opened its copy.
   */
  readonly featureEditProblems: (row: string) => readonly string[];
  /** The rows that a feature edit or a feature edit record names. */
  readonly featureEditRows: ReadonlySet<string>;
}

/** A repository-relative path with `/` separators, or undefined when it leaves the repository. */
export function normalizePath(path: string): string | undefined {
  const forward = path.replace(/\\/g, '/');
  if (forward.startsWith('/') || /^[A-Za-z]:\//.test(forward)) return undefined;
  const parts = forward.split('/').filter((p) => p !== '' && p !== '.');
  if (parts.includes('..') || parts.length === 0) return undefined;
  return parts.join('/');
}

/** The row that claims a save with no repair prompt for every fixture of the catalog. */
const ROUND_TRIP_ROW = 'save.round-trip';

const SAVE_REMEDY = 'run the save check again and record it (`bun run save-check record`)';
/** How to record the copy of the feature edit of `row`. */
const featureSaveRemedy = (row: string) =>
  `write the copy (\`bun scripts/fidelity/feature-edits.ts --out <dir>\`), open \`<dir>/${row}.xlsx\`, and record it (\`bun run save-check record-edit <dir> ${row} <result>\`); if you cannot run the save check, say so in your pull request, and a maintainer runs it`;
const RECIPE_REMEDY =
  'fix the edit or its recipe (`bun test scripts/fidelity/feature-edits.test.ts`)';
const REFERENCE_REMEDY =
  'export the reference PDF again and record it (`bun run fidelity:record-reference`)';
const GAPS_REMEDY =
  'make `gaps` equal the known gaps, or record them with `bun run fidelity --update-gaps`';

const sorted = (values: Iterable<string>): string[] => [...values].sort();
const sameSet = (a: ReadonlySet<string>, b: ReadonlySet<string>): boolean =>
  a.size === b.size && [...a].every((v) => b.has(v));

/** Every problem of the matrix, each naming its row. An empty list passes. */
export function supportMatrixProblems(
  rows: readonly SupportFeature[],
  ctx: MatrixContext
): string[] {
  const problems: string[] = [];
  const seen = new Set<string>();
  const cited = new Set<string>();
  const saveCache = new Map<string, readonly string[]>();
  const saveProblems = (name: string): readonly string[] => {
    let found = saveCache.get(name);
    if (!found) saveCache.set(name, (found = ctx.saveProblems(name)));
    return found;
  };
  const referenceCache = new Map<string, readonly string[]>();
  const referenceProblemsOf = (name: string): readonly string[] => {
    let found = referenceCache.get(name);
    if (!found) referenceCache.set(name, (found = ctx.referenceProblems(name)));
    return found;
  };

  for (const row of rows) {
    const at = (message: string) => problems.push(`${row.id}: ${message}`);
    if (seen.has(row.id)) at('the id is not unique: give each row its own id');
    seen.add(row.id);
    if (!(SUPPORT_STATUSES as readonly string[]).includes(row.status))
      at(`unknown status "${row.status}": use one of ${SUPPORT_STATUSES.join(', ')}`);
    if (!(SUPPORT_CATEGORIES as readonly string[]).includes(row.category))
      at(`unknown category "${row.category}": use one of ${SUPPORT_CATEGORIES.join(', ')}`);

    const ev = row.evidence;
    const claims = row.status !== 'unsupported';
    const fixturesOk = (names: readonly string[] | undefined, field: string) => {
      for (const name of names ?? [])
        if (!ctx.fixtures.has(name))
          at(
            `${field} cites "${name}", which is not in the catalog: add it to \`scripts/fixtures/catalog.ts\` or fix the name`
          );
    };
    fixturesOk(ev.fixtures, 'fixtures');
    fixturesOk(ev.saveVerified, 'saveVerified');
    fixturesOk(
      ev.fidelity?.map((f) => f.fixture),
      'fidelity'
    );

    // Save-verified records (design D5a).
    for (const name of ev.saveVerified ?? [])
      if (ctx.fixtures.has(name))
        for (const problem of saveProblems(name))
          at(`not save-verified: ${problem}: ${SAVE_REMEDY}`);
    const savesNothing = (ev.saveVerified ?? []).length === 0;
    if (claims && SAVED_CATEGORIES.has(row.category) && savesNothing && row.status === 'supported')
      at(
        'a supported editing or save row cites no save-verified fixture: cite one in `saveVerified`, or mark the row partial'
      );

    // The row's own edit (PPTX Editor's `FEATURE_EDITS`).
    let ownEditVerified = false;
    const editFixture = ctx.featureEdit(row.id);
    if (editFixture !== undefined) {
      const cites = new Set([
        ...(ev.fixtures ?? []),
        ...(ev.saveVerified ?? []),
        ...(ev.fidelity ?? []).map((f) => f.fixture),
      ]);
      if (!cites.has(editFixture))
        at(
          `the feature edit runs on "${editFixture}", which the row does not cite: cite it in \`saveVerified\` or \`fixtures\``
        );
      const problems = ctx.featureEditProblems(row.id);
      const unrecorded = problems.length === 1 && problems[0] === NO_RECORD;
      if (!unrecorded)
        for (const problem of problems)
          at(
            `the feature edit is not save-verified: ${problem}: ${problem.startsWith('the recipe fails') ? RECIPE_REMEDY : featureSaveRemedy(row.id)}`
          );
      ownEditVerified = problems.length === 0 && cites.has(editFixture);
    }
    const ownEditMissing =
      SAVED_CATEGORIES.has(row.category) && row.id !== ROUND_TRIP_ROW && !ownEditVerified;
    if (row.status === 'supported' && ownEditMissing)
      at(
        editFixture === undefined
          ? 'a supported editing or save row has no edit of its own: add a recipe to `FEATURE_EDITS` (`scripts/fidelity/feature-edits.ts`) on a fixture that the row cites, or mark the row partial'
          : `a supported editing or save row has no save-verified edit of its own: ${featureSaveRemedy(row.id)}; until then, mark the row partial`
      );

    // Fidelity results (design D8).
    let knownGap = false;
    const listed = new Set(ev.fixtures ?? []);
    for (const result of ev.fidelity ?? []) {
      const name = result.fixture;
      if (!ctx.fixtures.has(name)) continue;
      if (RENDERING_CATEGORIES.has(row.category) && !listed.has(name))
        at(`fidelity cites "${name}", which \`fixtures\` does not list: add it to \`fixtures\``);
      for (const problem of referenceProblemsOf(name))
        at(
          `fidelity cites "${name}", whose reference is not current: ${problem}: ${REFERENCE_REMEDY}`
        );
      if (!ctx.measured.has(name))
        at(
          `fidelity cites "${name}", which has no measurement: run \`bun run fidelity --update-gaps\` in each font environment`
        );
      const declared = new Set<string>(result.gaps ?? []);
      for (const check of declared)
        if (!(FAILURE_KINDS as readonly string[]).includes(check))
          at(
            `fidelity of "${name}" names unknown check "${check}": use one of ${FAILURE_KINDS.join(', ')}`
          );
      const recorded = ctx.knownGaps(name);
      if (!sameSet(declared, recorded))
        at(
          `fidelity of "${name}" declares gaps [${sorted(declared).join(', ')}], but the known gaps are [${sorted(recorded).join(', ')}]: ${GAPS_REMEDY}`
        );
      if (recorded.size > 0) knownGap = true;
      if (row.status === 'supported' && recorded.size > 0)
        at(
          `supported, but "${name}" has known gaps [${sorted(recorded).join(', ')}]: mark the row partial, or close the gaps`
        );
    }
    if (claims && RENDERING_CATEGORIES.has(row.category) && (ev.fidelity ?? []).length === 0)
      at(
        'a supported or partial rendering row cites no fidelity result: cite a measured fixture in `fidelity`, or move the row to the viewer category'
      );

    // Tests.
    for (const test of ev.tests ?? []) {
      const path = normalizePath(test);
      if (!path)
        at(`test "${test}" is not a repository-relative path: write it from the repository root`);
      else if (!/\.test\.tsx?$/.test(path))
        at(`"${test}" is not a test file: cite a \`.test.ts\` or \`.test.tsx\` file`);
      else if (!ctx.fileExists(path))
        at(`test "${path}" does not exist: fix the path, or add the test`);
    }
    if (claims && (ev.tests ?? []).length === 0)
      at('a supported or partial row cites no test: cite the tests that prove it in `tests`');

    // Known approximations.
    const approximations = ev.approximations ?? [];
    for (const id of approximations) {
      cited.add(id);
      if (!ctx.approximations.has(id))
        at(
          `approximation "${id}" does not exist: use an id of \`packages/core/src/layout/approximations.ts\``
        );
    }
    if (row.status === 'supported' && approximations.length > 0)
      at('supported, but it cites approximations: mark it partial');
    const missingSave = (SAVED_CATEGORIES.has(row.category) && savesNothing) || ownEditMissing;
    if (row.status === 'partial' && approximations.length === 0 && !knownGap && !missingSave)
      at(
        'partial, but it cites no approximation, no known gap, and no missing save check: cite the approximation that makes it partial, or mark it supported'
      );
  }

  // The round-trip row claims every fixture: it cites each fixture of the catalog.
  const roundTrip = rows.find((r) => r.id === ROUND_TRIP_ROW);
  if (roundTrip) {
    const saved = new Set(roundTrip.evidence.saveVerified ?? []);
    for (const name of sorted(ctx.fixtures))
      if (!saved.has(name))
        problems.push(
          `${ROUND_TRIP_ROW}: saveVerified does not cite "${name}" of the catalog: add it to \`EVERY_FIXTURE\``
        );
  }

  for (const id of [...ctx.featureEditRows].sort())
    if (!seen.has(id))
      problems.push(
        `feature edit "${id}" names no row of the matrix: rename it to the id of its row, or remove it and its record`
      );

  for (const id of ctx.approximations)
    if (!cited.has(id))
      problems.push(
        `approximation "${id}" is cited by no row: cite it in the row of each feature that it changes`
      );
  return problems;
}

/** The repository as it is. It runs every feature edit once. */
export async function repositoryContext(): Promise<MatrixContext> {
  const runs = await runAllFeatureEdits();
  const records = readFeatureRecords();
  const editProblems = new Map(
    [...runs].map(([row, run]) => [row, featureRecordProblems(row, run, records)])
  );
  // The list of each raster toolchain of each font environment.
  const gapLists = Object.values(KNOWN_GAPS).flatMap((lists) => lists ?? []);
  return {
    fixtures: new Set(FIXTURES.map((f) => f.name)),
    saveProblems: (name) => recordProblems(name),
    referenceProblems: (name) => referenceProblems(name),
    measured: new Set(
      gapLists.length === 0
        ? []
        : Object.keys(gapLists[0]!.fixtureLines).filter((name) =>
            gapLists.every((l) => Object.hasOwn(l.fixtureLines, name))
          )
    ),
    knownGaps: (name) =>
      new Set(
        gapLists.flatMap((l) =>
          Object.hasOwn(l.fixtures, name) ? Object.keys(l.fixtures[name]!) : []
        )
      ),
    approximations: new Set(APPROXIMATIONS.map((a) => a.id)),
    fileExists: (path) => existsSync(join(ROOT, path)),
    featureEdit: (row) =>
      Object.hasOwn(FEATURE_EDITS, row) ? FEATURE_EDITS[row]!.fixture : undefined,
    featureEditProblems: (row) => editProblems.get(row) ?? [],
    featureEditRows: new Set([...Object.keys(FEATURE_EDITS), ...Object.keys(records)]),
  };
}

if (import.meta.main) {
  const problems = supportMatrixProblems(supportFeatures, await repositoryContext());
  const counts = Object.fromEntries(
    SUPPORT_STATUSES.map((s) => [s, supportFeatures.filter((r) => r.status === s).length])
  );
  if (problems.length > 0) {
    console.error(
      `support matrix: ${problems.length} problem(s)\n${problems.map((p) => `  ${p}`).join('\n')}`
    );
    process.exit(1);
  }
  console.log(
    `support matrix: ${supportFeatures.length} rows (${SUPPORT_STATUSES.map((s) => `${counts[s]} ${s}`).join(', ')}), every claim has its evidence`
  );
}
