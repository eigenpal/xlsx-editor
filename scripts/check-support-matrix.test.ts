// The support matrix gate refuses each kind of claim without evidence, and each refusal names
// its remedy.

import { describe, expect, test } from 'bun:test';
import {
  supportFeatures,
  type FidelityEvidence,
  type SupportFeature,
} from '../docs/site/data/support-matrix.ts';
import { EDIT_REFUSAL_CODES } from '../packages/core/src/editor/cell-input.ts';
import {
  normalizePath,
  repositoryContext,
  supportMatrixProblems,
  type MatrixContext,
} from './check-support-matrix.ts';

const TEST = 'packages/core/src/cascade/__tests__/values.test.ts';
const SAVE_REMEDY = 'run the save check again and record it (`bun run save-check record`)';
const featureSaveRemedy = (row: string) =>
  `write the copy (\`bun scripts/fidelity/feature-edits.ts --out <dir>\`), open \`<dir>/${row}.xlsx\`, and record it (\`bun run save-check record-edit <dir> ${row} <result>\`); if you cannot run the save check, say so in your pull request, and a maintainer runs it`;
const NO_RECIPE =
  'a supported editing or save row has no edit of its own: add a recipe to `FEATURE_EDITS` (`scripts/fidelity/feature-edits.ts`) on a fixture that the row cites, or mark the row partial';
const REFERENCE_REMEDY =
  'export the reference PDF again and record it (`bun run fidelity:record-reference`)';
const GAPS_REMEDY =
  'make `gaps` equal the known gaps, or record them with `bun run fidelity --update-gaps`';

/**
 * A small repository: `good` passes everything, `gappy` has a baseline gap, `bad` has a stale
 * save-check record, `unprinted` has no reference PDF, and `stale` has a reference PDF that its
 * workbook no longer matches.
 */
function context(over: Partial<MatrixContext> = {}): MatrixContext {
  return {
    fixtures: new Set(['good', 'gappy', 'bad', 'unprinted', 'stale']),
    saveProblems: (name) => (name === 'bad' ? [`${name}: the edited copy changed (stale)`] : []),
    referenceProblems: (name) =>
      name === 'unprinted'
        ? [`${name}: no reference in fixtures/references.json`]
        : name === 'stale'
          ? [`${name}: the workbook changed after its reference was recorded (stale)`]
          : [],
    measured: new Set(['good', 'gappy', 'bad', 'stale']),
    knownGaps: (name) => new Set(name === 'gappy' ? ['baseline'] : []),
    approximations: new Set(['approx-a']),
    fileExists: (path) => path === TEST,
    // `edit.ok` has a verified edit of its own on `good`.
    featureEdit: (id) => (id === 'edit.ok' ? 'good' : undefined),
    featureEditProblems: () => [],
    featureEditRows: new Set(),
    ...over,
  };
}

const row = (over: Partial<SupportFeature> & Pick<SupportFeature, 'evidence'>): SupportFeature => ({
  id: 'row.under-test',
  name: 'Row under test',
  category: 'values',
  status: 'supported',
  ...over,
});

/** Rendering evidence: the fidelity results, with each fixture also listed in `fixtures`. */
const shows = (...fidelity: FidelityEvidence[]) => ({
  fixtures: fidelity.map((f) => f.fixture),
  fidelity,
});

/** A row that cites `approx-a`, so the "cited by no row" rule stays quiet. */
const citesApprox = row({
  id: 'row.approx',
  status: 'partial',
  evidence: { ...shows({ fixture: 'good' }), tests: [TEST], approximations: ['approx-a'] },
});

const problemsOf = (r: SupportFeature, ctx = context()) =>
  supportMatrixProblems([r, citesApprox], ctx).filter((p) => p.startsWith(`${r.id}:`));

describe('supportMatrixProblems', () => {
  test('a row with complete evidence passes', () => {
    expect(
      supportMatrixProblems(
        [
          row({ evidence: { ...shows({ fixture: 'good' }), tests: [TEST] } }),
          row({
            id: 'edit.ok',
            category: 'editing',
            evidence: { saveVerified: ['good'], tests: [TEST] },
          }),
          citesApprox,
        ],
        context()
      )
    ).toEqual([]);
  });

  test('a fixture that is not in the catalog fails', () => {
    expect(
      problemsOf(
        row({
          evidence: { fixtures: ['good', 'gone'], fidelity: [{ fixture: 'good' }], tests: [TEST] },
        })
      )
    ).toEqual([
      'row.under-test: fixtures cites "gone", which is not in the catalog: add it to `scripts/fixtures/catalog.ts` or fix the name',
    ]);
  });

  test('a stale, missing, or failed save-check record fails the row that cites it', () => {
    const editing = row({
      id: 'edit.ok',
      category: 'editing',
      evidence: { saveVerified: ['bad', 'good'], tests: [TEST] },
    });
    expect(problemsOf(editing).map((p) => p.replace('edit.ok', 'row.under-test'))).toEqual([
      `row.under-test: not save-verified: bad: the edited copy changed (stale): ${SAVE_REMEDY}`,
    ]);
  });

  test('a supported editing row needs a save-verified fixture; a partial one may name the gap', () => {
    const evidence = { tests: [TEST] };
    expect(problemsOf(row({ category: 'editing', evidence }))).toEqual([
      'row.under-test: a supported editing or save row cites no save-verified fixture: cite one in `saveVerified`, or mark the row partial',
      `row.under-test: ${NO_RECIPE}`,
    ]);
    expect(problemsOf(row({ category: 'editing', status: 'partial', evidence }))).toEqual([]);
  });

  test('a supported editing or save row needs a save-verified edit of its own', () => {
    const editing = (over: Partial<SupportFeature> = {}) =>
      row({
        id: 'edit.ok',
        category: 'editing',
        evidence: { saveVerified: ['good'], tests: [TEST] },
        ...over,
      });
    const own = (problems: string[], fixture = 'good') =>
      context({ featureEdit: () => fixture, featureEditProblems: () => problems });
    const noOwnEdit = `edit.ok: a supported editing or save row has no save-verified edit of its own: ${featureSaveRemedy('edit.ok')}; until then, mark the row partial`;

    expect(problemsOf(editing(), own([]))).toEqual([]);
    // A recipe that passes with no record: the row is at most partial, and partial names it.
    expect(problemsOf(editing(), own(['no save-check record']))).toEqual([noOwnEdit]);
    expect(problemsOf(editing({ status: 'partial' }), own(['no save-check record']))).toEqual([]);
    // No recipe at all: the same.
    expect(problemsOf(editing(), context({ featureEdit: () => undefined }))).toEqual([
      `edit.ok: ${NO_RECIPE}`,
    ]);
    // A stale or failed record, or a failing recipe, fails even a partial row.
    expect(
      problemsOf(
        editing({ status: 'partial' }),
        own(['the edited copy changed since the check opened it (stale)'])
      )
    ).toEqual([
      `edit.ok: the feature edit is not save-verified: the edited copy changed since the check opened it (stale): ${featureSaveRemedy('edit.ok')}`,
    ]);
    expect(
      problemsOf(editing({ status: 'partial' }), own(['the recipe fails: no edit found']))
    ).toEqual([
      'edit.ok: the feature edit is not save-verified: the recipe fails: no edit found: fix the edit or its recipe (`bun test scripts/fidelity/feature-edits.test.ts`)',
    ]);
    // The recipe's fixture must be one that the row cites.
    expect(problemsOf(editing(), own([], 'gappy'))).toEqual([
      'edit.ok: the feature edit runs on "gappy", which the row does not cite: cite it in `saveVerified` or `fixtures`',
      noOwnEdit,
    ]);
    // A fixture that the row cites only for fidelity counts.
    expect(
      problemsOf(
        editing({
          evidence: {
            ...shows({ fixture: 'gappy', gaps: ['baseline'] }),
            saveVerified: ['good'],
            tests: [TEST],
          },
          status: 'partial',
        }),
        own([], 'gappy')
      )
    ).toEqual([]);
    // A rendering row may have an edit too, and needs none.
    expect(
      problemsOf(row({ evidence: { ...shows({ fixture: 'good' }), tests: [TEST] } }), own([]))
    ).toEqual([]);
  });

  test('the round-trip row needs no edit of its own; every feature edit names a row', () => {
    const roundTrip = row({
      id: 'save.round-trip',
      category: 'save',
      evidence: { saveVerified: ['bad', 'good', 'gappy', 'stale', 'unprinted'], tests: [TEST] },
    });
    const ctx = context({
      saveProblems: () => [],
      approximations: new Set(),
      featureEditRows: new Set(['save.round-trip', 'edit.gone']),
    });
    expect(supportMatrixProblems([roundTrip], ctx)).toEqual([
      'feature edit "edit.gone" names no row of the matrix: rename it to the id of its row, or remove it and its record',
    ]);
  });

  test('a supported rendering row whose fixture has a known gap fails', () => {
    expect(
      problemsOf(
        row({ evidence: { ...shows({ fixture: 'gappy', gaps: ['baseline'] }), tests: [TEST] } })
      )
    ).toEqual([
      'row.under-test: supported, but "gappy" has known gaps [baseline]: mark the row partial, or close the gaps',
    ]);
  });

  test('declared gaps must equal the recorded gaps', () => {
    const partial = (gaps: ('baseline' | 'pixels')[] | undefined, fixture = 'gappy') =>
      row({ status: 'partial', evidence: { ...shows({ fixture, gaps }), tests: [TEST] } });
    expect(problemsOf(partial(['baseline']))).toEqual([]);
    expect(problemsOf(partial(undefined))).toEqual([
      `row.under-test: fidelity of "gappy" declares gaps [], but the known gaps are [baseline]: ${GAPS_REMEDY}`,
    ]);
    expect(problemsOf(partial(['pixels'], 'good'))).toContain(
      `row.under-test: fidelity of "good" declares gaps [pixels], but the known gaps are []: ${GAPS_REMEDY}`
    );
  });

  test('a fidelity fixture with no reference PDF or no measurement fails', () => {
    expect(
      problemsOf(row({ evidence: { ...shows({ fixture: 'unprinted' }), tests: [TEST] } }))
    ).toEqual([
      `row.under-test: fidelity cites "unprinted", whose reference is not current: unprinted: no reference in fixtures/references.json: ${REFERENCE_REMEDY}`,
      'row.under-test: fidelity cites "unprinted", which has no measurement: run `bun run fidelity --update-gaps` in each font environment',
    ]);
  });

  test('a stale reference fails every row that cites the fixture for fidelity', () => {
    const stale = `fidelity cites "stale", whose reference is not current: stale: the workbook changed after its reference was recorded (stale): ${REFERENCE_REMEDY}`;
    const a = row({ id: 'row.a', evidence: { ...shows({ fixture: 'stale' }), tests: [TEST] } });
    const b = row({ id: 'row.b', evidence: { ...shows({ fixture: 'stale' }), tests: [TEST] } });
    let calls = 0;
    const ctx = context({
      referenceProblems: (name) => {
        calls += 1;
        return context().referenceProblems(name);
      },
    });
    expect(supportMatrixProblems([a, b, citesApprox], ctx)).toEqual([
      `row.a: ${stale}`,
      `row.b: ${stale}`,
    ]);
    // Each fixture's reference is read once.
    expect(calls).toBe(2);
    // A fixture only listed in `fixtures`, with no fidelity claim, does not need a reference.
    const listed = row({
      status: 'unsupported',
      evidence: { fixtures: ['stale'] },
    });
    expect(problemsOf(listed)).toEqual([]);
  });

  test('a rendering row lists each fidelity fixture in `fixtures`', () => {
    expect(
      problemsOf(row({ evidence: { fidelity: [{ fixture: 'good' }], tests: [TEST] } }))
    ).toEqual([
      'row.under-test: fidelity cites "good", which `fixtures` does not list: add it to `fixtures`',
    ]);
    // A viewer row is not a rendering claim, so the rule does not apply.
    expect(
      problemsOf(
        row({ category: 'viewer', evidence: { fidelity: [{ fixture: 'good' }], tests: [TEST] } })
      )
    ).toEqual([]);
  });

  test('a supported rendering row needs a fidelity result', () => {
    expect(problemsOf(row({ evidence: { fixtures: ['good'], tests: [TEST] } }))).toEqual([
      'row.under-test: a supported or partial rendering row cites no fidelity result: cite a measured fixture in `fidelity`, or move the row to the viewer category',
    ]);
  });

  test('a missing test file fails, and backslashes are read as separators', () => {
    const windows = TEST.replace(/\//g, '\\');
    const good = shows({ fixture: 'good' });
    expect(problemsOf(row({ evidence: { ...good, tests: [windows] } }))).toEqual([]);
    expect(problemsOf(row({ evidence: { ...good, tests: ['scripts/gone.test.ts'] } }))).toEqual([
      'row.under-test: test "scripts/gone.test.ts" does not exist: fix the path, or add the test',
    ]);
    expect(problemsOf(row({ evidence: { ...good, tests: ['../outside.test.ts'] } }))).toEqual([
      'row.under-test: test "../outside.test.ts" is not a repository-relative path: write it from the repository root',
    ]);
    expect(problemsOf(row({ evidence: { ...good, tests: ['scripts/notes.md'] } }))).toEqual([
      'row.under-test: "scripts/notes.md" is not a test file: cite a `.test.ts` or `.test.tsx` file',
    ]);
  });

  test('a claim with no test fails; an unsupported row needs none', () => {
    expect(problemsOf(row({ evidence: shows({ fixture: 'good' }) }))).toEqual([
      'row.under-test: a supported or partial row cites no test: cite the tests that prove it in `tests`',
    ]);
    expect(problemsOf(row({ status: 'unsupported', evidence: {} }))).toEqual([]);
  });

  test('the round-trip row must cite every fixture of the catalog', () => {
    const roundTrip = (saveVerified: string[]) =>
      row({
        id: 'save.round-trip',
        category: 'save',
        evidence: { saveVerified, tests: [TEST] },
      });
    const all = ['bad', 'good', 'gappy', 'stale', 'unprinted'];
    const ctx = context({ saveProblems: () => [], approximations: new Set() });
    expect(supportMatrixProblems([roundTrip(all)], ctx)).toEqual([]);
    expect(supportMatrixProblems([roundTrip(all.filter((n) => n !== 'stale'))], ctx)).toEqual([
      'save.round-trip: saveVerified does not cite "stale" of the catalog: add it to `EVERY_FIXTURE`',
    ]);
  });

  test('approximation rules', () => {
    const base = { ...shows({ fixture: 'good' }), tests: [TEST] };
    expect(
      problemsOf(row({ status: 'partial', evidence: { ...base, approximations: ['nope'] } }))
    ).toEqual([
      'row.under-test: approximation "nope" does not exist: use an id of `packages/core/src/layout/approximations.ts`',
    ]);
    expect(problemsOf(row({ status: 'partial', evidence: base }))).toEqual([
      'row.under-test: partial, but it cites no approximation, no known gap, and no missing save check: cite the approximation that makes it partial, or mark it supported',
    ]);
    expect(problemsOf(row({ evidence: { ...base, approximations: ['approx-a'] } }))).toEqual([
      'row.under-test: supported, but it cites approximations: mark it partial',
    ]);
    expect(supportMatrixProblems([], context())).toEqual([
      'approximation "approx-a" is cited by no row: cite it in the row of each feature that it changes',
    ]);
  });

  test('duplicate ids, unknown statuses, and unknown categories fail', () => {
    const evidence = { ...shows({ fixture: 'good' }), tests: [TEST] };
    const twice = supportMatrixProblems(
      [row({ evidence }), row({ evidence }), citesApprox],
      context()
    );
    expect(twice).toEqual(['row.under-test: the id is not unique: give each row its own id']);
    const odd = { ...row({ evidence }), status: 'full' } as unknown as SupportFeature;
    expect(problemsOf(odd)).toContain(
      'row.under-test: unknown status "full": use one of supported, partial, unsupported'
    );
    const lost = { ...row({ evidence }), category: 'misc' } as unknown as SupportFeature;
    expect(problemsOf(lost)[0]).toStartWith(
      'row.under-test: unknown category "misc": use one of values, '
    );
  });
});

describe('normalizePath', () => {
  test('keeps repository-relative paths and refuses others', () => {
    expect(normalizePath('a\\b\\c.test.ts')).toBe('a/b/c.test.ts');
    expect(normalizePath('./a//b.test.ts')).toBe('a/b.test.ts');
    expect(normalizePath('/abs.test.ts')).toBeUndefined();
    expect(normalizePath('C:\\abs.test.ts')).toBeUndefined();
    expect(normalizePath('a/../../b.test.ts')).toBeUndefined();
  });
});

describe('the repository context', () => {
  test('reads the catalog, the references, the gaps, and the approximations', async () => {
    const ctx = await repositoryContext();
    expect(ctx.fixtures.has('basic-values')).toBe(true);
    expect(ctx.referenceProblems('basic-values')).toEqual([]);
    expect(ctx.referenceProblems('no-such-fixture')).toEqual([
      'no-such-fixture: no reference in fixtures/references.json',
    ]);
    expect(ctx.measured.has('basic-values')).toBe(true);
    expect([...ctx.knownGaps('basic-values')]).toEqual([]);
    expect(ctx.approximations.has('pixel-density')).toBe(true);
    expect(ctx.fileExists(TEST)).toBe(true);
  });

  test('the script sets its own time zone, so no shell prefix is needed', () => {
    expect(process.env.TZ).toBe('UTC');
  });
});

describe('the rows', () => {
  test('the refusal codes row names every editor refusal code, once', () => {
    const notes = supportFeatures.find((r) => r.id === 'refusals.codes')!.notes!;
    const named = [...notes.matchAll(/`([a-z0-9-]+)`/g)].map((m) => m[1]);
    expect([...named].sort()).toEqual([...EDIT_REFUSAL_CODES].sort());
  });

  test('each sentence of a note has 26 words or fewer', () => {
    const long = supportFeatures.flatMap((r) =>
      (r.notes ?? '')
        .split(/(?<=\.)\s+/)
        .filter((sentence) => sentence.split(/\s+/).length > 26)
        .map((sentence) => `${r.id}: ${sentence}`)
    );
    expect(long).toEqual([]);
  });
});
