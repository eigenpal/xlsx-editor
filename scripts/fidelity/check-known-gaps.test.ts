// The ratchet of `bun run fidelity:check-gaps` against the merge base (`check-known-gaps.ts`).
// The known-gap gate of a run: `harness.test.ts`.

import { describe, expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import {
  ACCEPT_CLIPPED,
  ACCEPT_NEW_TOOLCHAIN,
  EITHER_REMEDY,
  GAP_REMEDY,
  isCommit,
  isMissingPath,
  newFixtures,
  rereferencedFixtures,
  retiredNotices,
  listsOf,
  looserThan,
  type Ratcheted,
} from './check-known-gaps.ts';
import {
  checkGate,
  CLIPPED_BASE_REMEDY,
  CLIPPED_NEW_REMEDY,
  FEWER_LINES_REMEDY,
  CLIPPED_REFERENCE_REMEDY,
  CLIPPED_RUN_REMEDY,
  CLIPPED_RENAMED_REMEDY,
  RENAMED_LINES_REMEDY,
  unmeasuredProblem,
  type FixtureResult,
} from './known-gaps.ts';
import type { Manifest } from './references.ts';
import * as thresholds from './thresholds.ts';
import type { KnownGapList } from './thresholds.ts';

const tools = { poppler: 'poppler 1', chromium: 'chromium 1' };
const values = {
  pixels: 0,
  'line-breaks': 1,
  baseline: 0,
  'line-position': 0,
  'extra-text': 0,
};
const result = (patch: Partial<FixtureResult>): FixtureResult => ({
  name: 'basic-values',
  status: 'compared',
  failures: [],
  values,
  linesTotal: 3,
  linesMatched: 3,
  ...patch,
});
const list = (fixtures: KnownGapList['fixtures']): KnownGapList => ({
  ...tools,
  fixtureLines: { 'basic-values': 3 },
  fixtures,
});

describe('the base of the ratchet', () => {
  test('a --base that names no commit fails with exit code 2', () => {
    expect(isCommit('HEAD')).toBe(true);
    expect(isCommit('no-such-ref-of-the-fidelity-test')).toBe(false);
    const run = spawnSync(
      process.execPath,
      [join(import.meta.dir, 'cli.ts'), 'check-gaps', '--base', 'no-such-ref-of-the-fidelity-test'],
      { encoding: 'utf8' }
    );
    expect(run.status).toBe(2);
    expect(run.stderr.trim()).toBe(
      'fidelity: no-such-ref-of-the-fidelity-test is not a commit: fetch it, or name a commit'
    );
  });
});

describe('a checkout without git', () => {
  test('stops in one line with exit code 2', () => {
    for (const args of [[], ['--base', 'HEAD']]) {
      const run = spawnSync(
        process.execPath,
        [join(import.meta.dir, 'cli.ts'), 'check-gaps', ...args],
        {
          encoding: 'utf8',
          env: { ...process.env, PATH: '/nonexistent' },
        }
      );
      expect(run.status).toBe(2);
      expect(run.stderr.trim()).toStartWith('fidelity: git cannot run: ');
      expect(run.stderr.trim().split('\n')).toHaveLength(1);
    }
  });
});

describe('a base without thresholds', () => {
  test('only a path that the commit does not have counts as missing', () => {
    expect(
      isMissingPath("fatal: path 'scripts/fidelity/thresholds.ts' does not exist in 'HEAD'")
    ).toBe(true);
    expect(
      isMissingPath(
        "fatal: path 'scripts/fidelity/thresholds.ts' exists on disk, but not in 'HEAD'"
      )
    ).toBe(true);
    expect(isMissingPath('fatal: unable to read 0123456789abcdef')).toBe(false);
    expect(isMissingPath('')).toBe(false);
  });

  test('passes alone, and compared nothing with --require-thresholds', () => {
    // A repository of one empty commit: its HEAD has no thresholds.ts. A shallow CI checkout has
    // no such commit of this repository.
    const repo = mkdtempSync(join(tmpdir(), 'xlsx-gaps-base-'));
    try {
      const git = (...args: string[]) => spawnSync('git', args, { cwd: repo, encoding: 'utf8' });
      git('init', '-q');
      git(
        '-c',
        'user.name=test',
        '-c',
        'user.email=test@example.com',
        'commit',
        '-q',
        '--allow-empty',
        '-m',
        'empty'
      );
      const check = (...args: string[]) =>
        spawnSync(
          process.execPath,
          [join(import.meta.dir, 'cli.ts'), 'check-gaps', '--base', 'HEAD', ...args],
          { cwd: repo, encoding: 'utf8' }
        );
      const alone = check();
      expect([alone.status, alone.stdout.trim()]).toEqual([
        0,
        'No thresholds at HEAD: nothing to compare.',
      ]);
      const required = check('--require-thresholds');
      expect([required.status, required.stderr.trim()]).toEqual([
        2,
        'fidelity: no thresholds at HEAD: the check compared nothing',
      ]);
      // Only a missing path is a base without thresholds. A thresholds.ts whose object cannot be
      // read stops the check with exit code 2, with or without the flag.
      mkdirSync(join(repo, 'scripts', 'fidelity'), { recursive: true });
      writeFileSync(join(repo, 'scripts', 'fidelity', 'thresholds.ts'), 'export {};\n');
      git('add', '.');
      git('-c', 'user.name=test', '-c', 'user.email=test@example.com', 'commit', '-q', '-m', 'add');
      const blob = git('rev-parse', 'HEAD:scripts/fidelity/thresholds.ts').stdout.trim();
      rmSync(join(repo, '.git', 'objects', blob.slice(0, 2), blob.slice(2)), { force: true });
      for (const run of [check(), check('--require-thresholds')]) {
        expect(run.status).toBe(2);
        expect(run.stderr.trim()).toStartWith('fidelity: thresholds.ts at HEAD cannot be read: ');
      }
      // The same holds for the reference manifest of the base: only a missing one is empty.
      writeFileSync(join(repo, 'scripts', 'fidelity', 'thresholds.ts'), 'export const a = 1;\n');
      mkdirSync(join(repo, 'fixtures'), { recursive: true });
      writeFileSync(
        join(repo, 'fixtures', 'references.json'),
        '{ "version": 1, "references": {} }\n'
      );
      git('add', '.');
      git(
        '-c',
        'user.name=test',
        '-c',
        'user.email=test@example.com',
        'commit',
        '-q',
        '-m',
        'refs'
      );
      const refs = git('rev-parse', 'HEAD:fixtures/references.json').stdout.trim();
      rmSync(join(repo, '.git', 'objects', refs.slice(0, 2), refs.slice(2)), { force: true });
      const manifest = check();
      expect(manifest.status).toBe(2);
      expect(manifest.stderr.trim()).toStartWith(
        'fidelity: fixtures/references.json at HEAD cannot be read: '
      );
    } finally {
      rmSync(repo, { recursive: true, force: true });
    }
  });
});

describe('the ratchet against the merge base', () => {
  const base: Ratcheted = {
    THRESHOLDS: thresholds.THRESHOLDS,
    GAP_TOLERANCE: thresholds.GAP_TOLERANCE,
    KNOWN_GAPS: {
      reference: [list({ 'basic-values': { pixels: 0.01 } })],
      substitute: [list({})],
    },
  };
  const label = 'reference [poppler 1; chromium 1]';
  const names = new Set(['basic-values']);

  test('the same values are not looser', () => {
    expect(looserThan(base, base, names, names)).toEqual([]);
  });

  test('a per-fixture pixel limit only tightens, and stays while its fixture exists', () => {
    // No recorded gap, so that only the limits differ.
    const plain: Ratcheted = {
      ...base,
      KNOWN_GAPS: { reference: [list({})], substitute: [list({})] },
    };
    const limited = { ...plain, FIXTURE_PIXEL_MISMATCH: { 'basic-values': 0.0002 } };
    const tighter = { ...plain, FIXTURE_PIXEL_MISMATCH: { 'basic-values': 0.0001 } };
    const looser = { ...plain, FIXTURE_PIXEL_MISMATCH: { 'basic-values': 0.0003 } };
    expect(looserThan(plain, limited, names, names)).toEqual([]);
    expect(looserThan(limited, tighter, names, names)).toEqual([]);
    expect(looserThan(limited, looser, names, names)).toEqual([
      'pixel limit of basic-values: 0.0002 → 0.0003',
    ]);
    expect(looserThan(limited, plain, names, names)).toHaveLength(1);
    // A new entry for a fixture may not be looser than its old limit, the page limit.
    const newLoose = { ...plain, FIXTURE_PIXEL_MISMATCH: { 'basic-values': 0.05 } };
    expect(looserThan(plain, newLoose, names, names)).toEqual([
      'pixel limit of basic-values: 0.005 → 0.05',
    ]);
    // A fixture with its own limit records no pixels gap.
    const gapped = {
      ...limited,
      KNOWN_GAPS: {
        ...plain.KNOWN_GAPS,
        substitute: [list({ 'basic-values': { pixels: 0.0001 } })],
      },
    } as unknown as Ratcheted;
    expect(looserThan(limited, gapped, names, names)).toContain(
      'pixel limit of basic-values: a pixels gap is recorded in poppler 1'
    );
    // Removing or renaming a fixture with a limit fails, as for a known gap.
    expect(
      looserThan(limited, plain, names, new Set()).filter((m) => m.startsWith('pixel limit'))
    ).toEqual(['pixel limit of basic-values: the fixture was removed or renamed']);
  });

  test('a looser threshold, gap, or new gap fails', () => {
    const next = {
      ...base,
      THRESHOLDS: { ...base.THRESHOLDS, maxPixelMismatch: 0.01 },
      KNOWN_GAPS: {
        reference: [list({ 'basic-values': { pixels: 0.02 } })],
        substitute: [list({ 'basic-values': { baseline: 1 } })],
      },
    } as unknown as Ratcheted;
    expect(looserThan(base, next, names, names)).toHaveLength(3);
  });

  test('clipped lines only become fewer, and a new fixture needs an owner to clip', () => {
    const withClipped = (lines: Record<string, string[]> | undefined): Ratcheted => ({
      ...base,
      KNOWN_GAPS: {
        ...base.KNOWN_GAPS,
        substitute: [{ ...list({}), ...(lines ? { fixtureClipped: lines } : {}) }],
      },
    });
    const one = withClipped({ 'basic-values': ['Kept content'] });
    // A merge base without the record has nothing to compare with.
    expect(looserThan(withClipped(undefined), one, names, names)).toEqual([]);
    expect(looserThan(one, withClipped({}), names, names)).toEqual([]);
    const label = 'substitute [poppler 1; chromium 1]';
    expect(
      looserThan(one, withClipped({ 'basic-values': ['Kept content', 'Other'] }), names, names)
    ).toEqual([
      `${label}: basic-values names the clipped line \`Other\` 1 time, the merge base 0 times; ${CLIPPED_BASE_REMEDY}`,
    ]);
    // A second entry of the same line is looser too.
    expect(
      looserThan(
        one,
        withClipped({ 'basic-values': ['Kept content', 'Kept content'] }),
        names,
        names
      )
    ).toHaveLength(1);
    const head = new Set(['basic-values', 'new-fixture', 'other-fixture']);
    const added = withClipped({
      'basic-values': ['Kept content'],
      'new-fixture': ['Wide'],
      'other-fixture': ['Tall'],
    });
    expect(looserThan(one, added, names, head)).toEqual([
      `${label}: new-fixture is a new fixture with the clipped line \`Wide\`; ${CLIPPED_NEW_REMEDY}`,
      `${label}: other-fixture is a new fixture with the clipped line \`Tall\`; ${CLIPPED_NEW_REMEDY}`,
    ]);
    // A base fixture has no acceptance: its remedy fixes the layout or the fixture only.
    expect(CLIPPED_BASE_REMEDY).toBe('fix the layout or the fixture');
    // A new fixture names the owner's acceptance, which CI never gives.
    expect(CLIPPED_NEW_REMEDY).toContain('XLSX_FIDELITY_ACCEPT_CLIPPED=<fixture>');
    expect(looserThan(one, added, names, head)).toEqual([
      `${label}: new-fixture is a new fixture with the clipped line \`Wide\`; ${CLIPPED_NEW_REMEDY}`,
      `${label}: other-fixture is a new fixture with the clipped line \`Tall\`; ${CLIPPED_NEW_REMEDY}`,
    ]);
    expect(ACCEPT_CLIPPED).toBe('XLSX_FIDELITY_ACCEPT_CLIPPED');
    // The owner's acceptance is a comma-separated list of fixtures.
    expect(looserThan(one, added, names, head, undefined, [], 'new-fixture')).toHaveLength(1);
    expect(
      looserThan(one, added, names, head, undefined, [], 'new-fixture, other-fixture')
    ).toEqual([]);
  });

  test('a list cannot drop its clipped-line record, and a missing record of a second list is empty', () => {
    const lists = (a: KnownGapList, b: KnownGapList): Ratcheted => ({
      ...base,
      KNOWN_GAPS: { ...base.KNOWN_GAPS, substitute: [a, b] },
    });
    const dev = { ...list({}), fixtureClipped: { 'basic-values': ['Kept content'] } };
    const ci = { ...list({}), poppler: 'poppler 2', fixtureClipped: {} };
    const { fixtureClipped: _, ...devWithout } = dev;
    const was = lists(dev, ci);
    // Step 1: the list drops its record. That fails, also though it names no line.
    const step1 = lists(devWithout, ci);
    expect(looserThan(was, step1, names, names)).toEqual([
      'substitute [poppler 1; chromium 1]: the list drops its clipped-line record (`fixtureClipped`), which the merge base has; keep the record, empty when no line is clipped',
    ]);
    // Step 2, on a merge base with step 1: the other list keeps its record, so the list without
    // one counts as empty, and new clipped lines of a base fixture fail.
    const step2 = lists({ ...devWithout, fixtureClipped: { 'basic-values': ['Other'] } }, ci);
    expect(looserThan(step1, step2, names, names)).toEqual([
      `substitute [poppler 1; chromium 1]: basic-values names the clipped line \`Other\` 1 time, the merge base 0 times; ${CLIPPED_BASE_REMEDY}`,
    ]);
    // A merge base where no list has the record keeps the first-record rule.
    const { fixtureClipped: __, ...ciWithout } = ci;
    const none = lists(devWithout, ciWithout);
    expect(looserThan(none, step2, names, names)).toEqual([]);
  });

  test('a clipped-line record of no fixture fails, and a renamed fixture cannot clip', () => {
    const withClipped = (lines: Record<string, string[]>): Ratcheted => ({
      ...base,
      KNOWN_GAPS: { ...base.KNOWN_GAPS, substitute: [{ ...list({}), fixtureClipped: lines }] },
    });
    const label = 'substitute [poppler 1; chromium 1]';
    const was = withClipped({});
    // Change 1: a record under a name that is no fixture yet.
    const early = withClipped({ renamed: ['Wide'] });
    expect(looserThan(was, early, names, names)).toEqual([
      `${label}: renamed has a clipped-line record but is not a fixture; remove the record`,
    ]);
    // Change 2: the fixture `plain` of the merge base is renamed to it, so no fixture is new
    // (`newFixtures`). Its clipped lines compare with none, and no acceptance applies.
    const baseNames = new Set([...names, 'plain']);
    const renamedHead = new Set([...names, 'renamed']);
    const problem = `${label}: renamed names the clipped line \`Wide\` 1 time, the merge base 0 times; ${CLIPPED_RENAMED_REMEDY}`;
    expect(looserThan(was, early, baseNames, renamedHead)).toEqual([problem]);
    expect(looserThan(was, early, baseNames, renamedHead, undefined, [], 'renamed')).toEqual([
      problem,
    ]);
    // Also from a merge base with no clipped-line record at all, where a first record is not
    // compared.
    const { fixtureClipped: _, ...bare } = list({});
    const none = { ...base, KNOWN_GAPS: { ...base.KNOWN_GAPS, substitute: [bare] } } as Ratcheted;
    expect(looserThan(none, early, baseNames, renamedHead)).toEqual([problem]);
    // The remedy names the rename and the paths that work.
    expect(CLIPPED_RENAMED_REMEDY).toContain('a fixture of the merge base is gone in this change');
    expect(CLIPPED_RENAMED_REMEDY).toContain('keep the old fixture name');
    expect(CLIPPED_RENAMED_REMEDY).toContain('as an owner only, run `bun run fidelity:check-gaps`');
    expect(CLIPPED_RENAMED_REMEDY).toContain('XLSX_FIDELITY_ACCEPT_CLIPPED=<fixture>');
  });

  test('the checked lines of a fixture do not reset when it is renamed', () => {
    const lines = (fixtureLines: Record<string, number>): Ratcheted => ({
      ...base,
      KNOWN_GAPS: { ...base.KNOWN_GAPS, substitute: [{ ...list({}), fixtureLines }] },
    });
    const label = 'substitute [poppler 1; chromium 1]';
    const was = lines({ plain: 100 });
    const baseNames = new Set([...names, 'plain']);
    // The same drop under the same name fails.
    expect(looserThan(was, lines({ plain: 2 }), baseNames, baseNames)).toEqual([
      `${label}: plain checks 2 lines, fewer than 100; ${FEWER_LINES_REMEDY}`,
    ]);
    // A rename to `plain-2`, which checks 2 lines, fails too.
    const renamed = new Set([...names, 'plain-2']);
    expect(looserThan(was, lines({ 'plain-2': 2 }), baseNames, renamed)).toEqual([
      `${label}: plain checks 100 lines and was removed or renamed in a change that adds a fixture; ${RENAMED_LINES_REMEDY}`,
    ]);
    // A change that only removes the fixture passes, when the fixture has no known gap: a later
    // change that adds the name again adds a new fixture. The removal shows in review.
    expect(looserThan(was, lines({}), baseNames, names)).toEqual([]);
    // A new reference PDF does not let the fixture check fewer lines, and the remedy names the
    // paths that work.
    const plain = new Set(['plain']);
    expect(
      looserThan(was, lines({ plain: 2 }), baseNames, baseNames, undefined, [], 'plain', plain)
    ).toEqual([`${label}: plain checks 2 lines, fewer than 100; ${FEWER_LINES_REMEDY}`]);
    expect(FEWER_LINES_REMEDY).toBe(
      'fix the layout so that the fixture checks its lines again, or, as an owner only, remove the fixture and add it as a new one in separate changes; a fixture with a known gap cannot be removed, so fix its gap first, or keep the fixture'
    );
  });

  test('a gap of a check that the merge base does not have is refused', () => {
    const { 'extra-text': _, ...older } = thresholds.GAP_TOLERANCE;
    const was = { ...base, GAP_TOLERANCE: older } as unknown as Ratcheted;
    const next = {
      ...base,
      KNOWN_GAPS: {
        ...base.KNOWN_GAPS,
        substitute: [list({ 'basic-values': { 'extra-text': 1 } })],
      },
    };
    const refused = [
      `substitute [poppler 1; chromium 1]: basic-values has a new gap extra-text (1); ${GAP_REMEDY}`,
    ];
    expect(looserThan(was, next, names, names)).toEqual(refused);
    expect(looserThan(base, next, names, names)).toEqual(refused);
  });

  test('a threshold goes away only when RETIRED_LIMITS names it with the reason', () => {
    const { maxExtraLines: _, ...rest } = base.THRESHOLDS as Record<string, number>;
    const without = (retired?: Record<string, string>) =>
      ({
        ...base,
        THRESHOLDS: rest,
        ...(retired ? { RETIRED_LIMITS: retired } : {}),
      }) as unknown as Ratcheted;
    const refused = [
      'threshold maxExtraLines was removed (it was 0); keep it, or list it in `RETIRED_LIMITS` in thresholds.ts with the reason',
    ];
    expect(looserThan(base, without(), names, names)).toEqual(refused);
    expect(looserThan(base, without({ maxExtraLines: ' ' }), names, names)).toEqual(refused);
    expect(
      looserThan(
        base,
        without({ maxExtraLines: 'the line check counts extra lines' }),
        names,
        names
      )
    ).toEqual([]);
    expect(thresholds.RETIRED_LIMITS).toEqual({});
    // Each retired threshold prints a notice with its reason.
    const reason = { maxExtraLines: 'the line check counts extra lines' };
    expect(retiredNotices(base, without(reason))).toEqual([
      'NOTICE: maxExtraLines is retired: the line check counts extra lines',
    ]);
    expect(retiredNotices(base, base)).toEqual([]);
    // A reason of only spaces retires nothing, and prints no notice.
    expect(retiredNotices(base, without({ maxExtraLines: ' ' }))).toEqual([]);
  });

  test('the acceptance of a new toolchain never accepts a new clipped line of a base fixture', () => {
    const at = (poppler: string, clipped: Record<string, string[]>): KnownGapList => ({
      ...list({}),
      poppler,
      fixtureClipped: clipped,
    });
    const was = {
      ...base,
      KNOWN_GAPS: { ...base.KNOWN_GAPS, substitute: [at('poppler 1', {})] },
    } as Ratcheted;
    const next = {
      ...base,
      KNOWN_GAPS: {
        ...base.KNOWN_GAPS,
        substitute: [at('poppler 1', {}), at('poppler 3', { 'basic-values': ['Wide'] })],
      },
    } as Ratcheted;
    const accepted = looserThan(was, next, names, names, 'poppler 3; chromium 1');
    expect(accepted).toHaveLength(1);
    expect(accepted[0]).toEndWith(`; ${CLIPPED_BASE_REMEDY}`);
    // Nor a clipped line of a new fixture: only `XLSX_FIDELITY_ACCEPT_CLIPPED` accepts it.
    const fresh = {
      ...base,
      KNOWN_GAPS: {
        ...base.KNOWN_GAPS,
        substitute: [at('poppler 1', {}), at('poppler 3', { 'new-fixture': ['Wide'] })],
      },
    } as Ratcheted;
    const head = new Set([...names, 'new-fixture']);
    const notes: string[] = [];
    const refused = looserThan(was, fresh, names, head, 'poppler 3; chromium 1', notes);
    expect(refused).toHaveLength(1);
    expect(refused[0]).toEndWith(`; ${CLIPPED_NEW_REMEDY}`);
    expect(notes).toEqual([]);
    expect(
      looserThan(was, fresh, names, head, 'poppler 3; chromium 1', notes, 'new-fixture')
    ).toEqual([]);
    expect(notes).toHaveLength(1);
    expect(notes[0]).toStartWith('accepted by XLSX_FIDELITY_ACCEPT_CLIPPED: ');
  });

  test('a fixture with a new reference PDF takes new clipped lines only with an owner', () => {
    const identity = (pdf: string) => ({
      workbook: 'w',
      pdf,
      pages: 1,
      pageSizePt: [612, 792] as const,
      locale: { decimal: '.', group: ',' },
    });
    const manifest = (pdfs: Record<string, string>): Manifest => ({
      version: 1,
      references: Object.fromEntries(Object.entries(pdfs).map(([n, pdf]) => [n, identity(pdf)])),
    });
    // Only a fixture of both sides with another PDF hash has a new reference PDF.
    expect(
      rereferencedFixtures(
        manifest({ 'basic-values': 'a', formulas: 'b', gone: 'c' }),
        manifest({ 'basic-values': 'a2', formulas: 'b', added: 'd' })
      )
    ).toEqual(new Set(['basic-values']));
    // The new reference PDF moves the clip edge of a named line by 0.1 points.
    const at = (key: string): Ratcheted =>
      ({
        ...base,
        KNOWN_GAPS: {
          ...base.KNOWN_GAPS,
          substitute: [{ ...list({}), fixtureClipped: { 'basic-values': [key] } }],
        },
      }) as Ratcheted;
    const was = at('Kept content @ 47.0 | right clip 89.0');
    const moved = at('Kept content @ 47.0 | right clip 89.1');
    const problem =
      'substitute [poppler 1; chromium 1]: basic-values names the clipped line `Kept content @ 47.0 | right clip 89.1` 1 time, the merge base 0 times';
    // With the same reference PDF, no acceptance applies.
    expect(looserThan(was, moved, names, names, undefined, [], 'basic-values')).toEqual([
      `${problem}; ${CLIPPED_BASE_REMEDY}`,
    ]);
    // With a new reference PDF, the remedy names the owner's route, which CI never takes.
    const renewed = new Set(['basic-values']);
    expect(looserThan(was, moved, names, names, undefined, [], undefined, renewed)).toEqual([
      `${problem}; ${CLIPPED_REFERENCE_REMEDY}`,
    ]);
    expect(CLIPPED_REFERENCE_REMEDY).toContain('XLSX_FIDELITY_ACCEPT_CLIPPED=<fixture>');
    const notes: string[] = [];
    expect(looserThan(was, moved, names, names, undefined, notes, 'basic-values', renewed)).toEqual(
      []
    );
    expect(notes).toEqual([`accepted by XLSX_FIDELITY_ACCEPT_CLIPPED: ${problem}`]);
  });

  test('a change cannot flip the direction or the tightest value of a threshold', () => {
    const dirs = thresholds.THRESHOLD_DIRECTIONS;
    const was = { ...base, THRESHOLD_DIRECTIONS: dirs } as unknown as Ratcheted;
    // Flip `maxExtraLines` to smaller-is-looser and raise it: the merge base's entry holds.
    const flipped = {
      ...base,
      THRESHOLDS: { ...base.THRESHOLDS, maxExtraLines: 3 },
      THRESHOLD_DIRECTIONS: { ...dirs, maxExtraLines: { looser: 'smaller' } },
    } as unknown as Ratcheted;
    expect(looserThan(was, flipped, names, names)).toEqual([
      'threshold maxExtraLines changed its direction or tightest value in `THRESHOLD_DIRECTIONS`; keep the entry of the merge base',
    ]);
    const lowered = {
      ...base,
      THRESHOLD_DIRECTIONS: { ...dirs, minLineBreakMatch: { looser: 'smaller', tightest: 0.5 } },
    } as unknown as Ratcheted;
    expect(looserThan(was, lowered, names, names)).toHaveLength(1);
    // The same entries pass.
    expect(looserThan(was, { ...was }, names, names)).toEqual([]);
  });

  test('a new threshold or tolerance starts at its tightest value', () => {
    const b = base.THRESHOLDS as Record<string, number>;
    const without = (key: string) => {
      const { [key]: _, ...rest } = b;
      return { ...base, THRESHOLDS: rest } as unknown as Ratcheted;
    };
    const at = (key: string, value: number) =>
      ({ ...base, THRESHOLDS: { ...b, [key]: value } }) as unknown as Ratcheted;
    // A larger-is-looser threshold starts at 0.
    expect(looserThan(without('maxExtraLines'), at('maxExtraLines', 0), names, names)).toEqual([]);
    expect(looserThan(without('maxExtraLines'), at('maxExtraLines', 1), names, names)).toEqual([
      'threshold maxExtraLines is new at 1, looser than its tightest value 0; start it at 0',
    ]);
    // A smaller-is-looser threshold starts at its `tightest` value in `THRESHOLD_DIRECTIONS`.
    expect(
      looserThan(without('minLineBreakMatch'), at('minLineBreakMatch', 0.9), names, names)
    ).toEqual([
      'threshold minLineBreakMatch is new at 0.9, looser than its tightest value 1; start it at 1',
    ]);
    // A smaller-is-looser threshold without a tightest value is refused.
    expect(looserThan(without('compareDpi'), base, names, names)).toEqual([
      'threshold compareDpi is new and has no tightest value; give it `tightest` in `THRESHOLD_DIRECTIONS`',
    ]);
    // A threshold without a direction is refused, new or not.
    expect(looserThan(base, at('minSomething', 1), names, names)).toEqual([
      'threshold minSomething has no loosening direction; add it to `THRESHOLD_DIRECTIONS` in thresholds.ts',
    ]);
    // Every threshold of the harness has a direction.
    expect(
      Object.keys(thresholds.THRESHOLDS).filter((k) => !thresholds.THRESHOLD_DIRECTIONS[k])
    ).toEqual([]);
    // A looser threshold names its next step.
    expect(looserThan(base, at('maxExtraLines', 2), names, names)).toEqual([
      'threshold maxExtraLines: 0 → 2; thresholds only tighten, so fix the regression instead',
    ]);
    // The clip-match tolerances are fixed: a smaller value fails a correct render and a larger
    // value passes a wrong one, so the check refuses a change in either direction.
    const fixed = [
      'sameBaselinePt',
      'sameClipPt',
      'onItemPt',
      'sameEdgePt',
      'onRunPt',
      'sameRowPt',
      'clipEdgePt',
      'joinGapPt',
    ];
    for (const key of fixed) {
      expect(thresholds.THRESHOLD_DIRECTIONS[key]?.looser).toBe('either');
      const value = b[key]!;
      for (const changed of [value * 2, value / 2])
        expect(looserThan(base, at(key, changed), names, names)).toEqual([
          `threshold ${key}: ${value} → ${changed}; ${EITHER_REMEDY(key)}`,
        ]);
      // A merge base without it compares with the value that the reference PDFs need.
      for (const changed of [value * 2, value / 2])
        expect(looserThan(without(key), at(key, changed), names, names)).toEqual([
          `threshold ${key} is new at ${changed}, not its tightest value ${value}; start it at ${value}, or set its \`tightest\` in \`THRESHOLD_DIRECTIONS\` to the measured value`,
        ]);
      expect(looserThan(without(key), base, names, names)).toEqual([]);
      // The path after review: the old key retires with its reason, and a new key starts at its
      // measured tightest value.
      const { [key]: _old, ...kept } = b;
      const { [key]: _was, ...directions } = thresholds.THRESHOLD_DIRECTIONS;
      const renamed = {
        ...base,
        THRESHOLDS: { ...kept, [`${key}Next`]: value * 2 },
        THRESHOLD_DIRECTIONS: {
          ...directions,
          [`${key}Next`]: { looser: 'either', tightest: value * 2 },
        },
        RETIRED_LIMITS: { [key]: `measured again as ${key}Next` },
      } as unknown as Ratcheted;
      expect(looserThan(base, renamed, names, names)).toEqual([]);
    }
    expect(EITHER_REMEDY('onRunPt')).toContain(
      "retire onRunPt in `RETIRED_LIMITS` in thresholds.ts with the reason, and add a new threshold at its measured tightest value with `looser: 'either'`"
    );
    // A new tolerance takes at most the smallest tolerance of the merge base.
    const tolerance = base.GAP_TOLERANCE as Record<string, number>;
    const { pixels: _, ...older } = tolerance;
    const smallest = Math.min(...Object.values(older));
    const was = { ...base, GAP_TOLERANCE: older } as unknown as Ratcheted;
    const wide = { ...base, GAP_TOLERANCE: { ...tolerance, pixels: smallest + 1 } };
    expect(looserThan(was, wide as unknown as Ratcheted, names, names)).toEqual([
      `gap tolerance pixels is new at ${smallest + 1}, more than the smallest tolerance ${smallest}; start it at ${smallest} or less`,
    ]);
  });

  test('a recorded gap may not grow, even inside the run tolerance', () => {
    const grown = {
      ...base,
      KNOWN_GAPS: {
        ...base.KNOWN_GAPS,
        reference: [list({ 'basic-values': { pixels: 0.0105 } })],
      },
    };
    expect(thresholds.GAP_TOLERANCE.pixels).toBeGreaterThan(0.0005);
    expect(looserThan(base, grown, names, names)).toEqual([
      `${label}: basic-values pixels: 0.01 → 0.0105; ${GAP_REMEDY}`,
    ]);
    const tighter = {
      ...base,
      KNOWN_GAPS: {
        ...base.KNOWN_GAPS,
        reference: [list({ 'basic-values': { pixels: 0.0099 } })],
      },
    };
    expect(looserThan(base, tighter, names, names)).toEqual([]);
  });

  test('removing the list of an environment is not looser: its runs then fail', () => {
    const next = { ...base, KNOWN_GAPS: { substitute: base.KNOWN_GAPS.substitute } };
    expect(looserThan(base, next, names, names)).toEqual([]);
    expect(unmeasuredProblem('reference').kind).toBe('fix');
    expect(unmeasuredProblem('reference').text).toContain('no measured known-gap list');
  });

  test('the reference environment has no measured list', () => {
    expect(thresholds.KNOWN_GAPS.reference).toBeUndefined();
    expect(thresholds.KNOWN_GAPS.substitute?.length).toBeGreaterThan(0);
  });

  const other = { ...list({ 'basic-values': { pixels: 0.01 } }), poppler: 'poppler 2' };

  test('a list for a new toolchain is allowed when it is not looser than the tightest values', () => {
    const next = {
      ...base,
      KNOWN_GAPS: { ...base.KNOWN_GAPS, reference: [...listsOf(base.KNOWN_GAPS.reference), other] },
    };
    expect(looserThan(base, next, names, names)).toEqual([]);
    const tighter = { ...other, fixtures: { 'basic-values': { pixels: 0.005 } } };
    expect(
      looserThan(base, { ...base, KNOWN_GAPS: { reference: [tighter] } }, names, names)
    ).toEqual([]);
  });

  test('a list for a new toolchain may not be looser than any list of the merge base', () => {
    // The merge base has a local list with a gap and a CI list with a pass.
    const two: Ratcheted = {
      ...base,
      KNOWN_GAPS: {
        reference: [
          list({ 'basic-values': { pixels: 0.01 } }),
          { ...list({}), chromium: 'chromium 2' },
        ],
      },
    };
    const label3 =
      'reference [poppler 3; chromium 1] (new toolchain, against the tightest values of poppler 1)';
    const add = (l: KnownGapList): Ratcheted => ({
      ...base,
      KNOWN_GAPS: { reference: [...listsOf(two.KNOWN_GAPS.reference), l] },
    });
    const at3 = (fixtures: KnownGapList['fixtures'], lines = 3): KnownGapList => ({
      ...list(fixtures),
      poppler: 'poppler 3',
      fixtureLines: { 'basic-values': lines },
    });
    expect(looserThan(two, add(at3({})), names, names)).toEqual([]);
    expect(looserThan(two, add(at3({ 'basic-values': { pixels: 0.005 } })), names, names)).toEqual([
      `${label3}: basic-values has a new gap pixels (0.005); ${GAP_REMEDY}`,
    ]);
    expect(looserThan(two, add(at3({}, 2)), names, names)).toEqual([
      `${label3}: basic-values checks 2 lines, fewer than 3; ${FEWER_LINES_REMEDY}`,
    ]);
    // An owner accepts a looser new toolchain only through the environment variable.
    const notes: string[] = [];
    const looser = add(at3({ 'basic-values': { pixels: 0.005 } }));
    expect(looserThan(two, looser, names, names, 'poppler 3; chromium 1', notes)).toEqual([]);
    expect(notes).toEqual([
      `accepted by ${ACCEPT_NEW_TOOLCHAIN}: ${label3}: basic-values has a new gap pixels (0.005); ${GAP_REMEDY}`,
    ]);
    expect(looserThan(two, looser, names, names, 'poppler 4; chromium 1')).toHaveLength(1);
  });

  test('the poppler 24.02.0 list is not looser than the merge base of its pull request', () => {
    const merged = {
      ...base,
      KNOWN_GAPS: { substitute: thresholds.KNOWN_GAPS.substitute![0]! },
    } as unknown as Ratcheted;
    const all = new Set(Object.keys(thresholds.KNOWN_GAPS.substitute![0]!.fixtureLines));
    expect(looserThan(merged, thresholds as unknown as Ratcheted, all, all)).toEqual([]);
  });

  test('lists compare toolchain by toolchain, in any order', () => {
    const two: Ratcheted = { ...base, KNOWN_GAPS: { reference: [list({}), other] } };
    const swapped: Ratcheted = { ...base, KNOWN_GAPS: { reference: [other, list({})] } };
    expect(looserThan(two, swapped, names, names)).toEqual([]);
    const grown: Ratcheted = {
      ...base,
      KNOWN_GAPS: {
        reference: [list({}), { ...other, fixtures: { 'basic-values': { pixels: 0.02 } } }],
      },
    };
    expect(looserThan(two, grown, names, names)).toEqual([
      `reference [poppler 2; chromium 1]: basic-values pixels: 0.01 → 0.02; ${GAP_REMEDY}`,
    ]);
  });

  test('a removed list compares with the list that its runs then use', () => {
    const two: Ratcheted = { ...base, KNOWN_GAPS: { reference: [list({}), other] } };
    // Runs with poppler 2 fall back to the poppler 1 list, which has no gap: not looser.
    const tighter: Ratcheted = { ...base, KNOWN_GAPS: { reference: [list({})] } };
    expect(looserThan(two, tighter, names, names)).toEqual([]);
    // Runs with poppler 1 fall back to the poppler 2 list, which has a gap: looser.
    const looser: Ratcheted = { ...base, KNOWN_GAPS: { reference: [other] } };
    expect(looserThan(two, looser, names, names)).toEqual([
      `${label} (removed; runs use poppler 2): basic-values has a new gap pixels (0.01); ${GAP_REMEDY}`,
    ]);
    // No list left: runs in the environment fail.
    expect(looserThan(two, { ...base, KNOWN_GAPS: {} }, names, names)).toEqual([]);
  });

  test('a merge base with one list for each environment compares as one toolchain', () => {
    const old = {
      ...base,
      KNOWN_GAPS: { reference: list({ 'basic-values': { pixels: 0.01 } }) },
    } as unknown as Ratcheted;
    expect(listsOf(old.KNOWN_GAPS.reference)).toHaveLength(1);
    expect(looserThan(old, base, names, names)).toEqual([]);
    const grown: Ratcheted = {
      ...base,
      KNOWN_GAPS: { reference: [list({ 'basic-values': { pixels: 0.02 } })] },
    };
    expect(looserThan(old, grown, names, names)).toEqual([
      `${label}: basic-values pixels: 0.01 → 0.02; ${GAP_REMEDY}`,
    ]);
  });

  test('a removed fixture with a gap fails; a new fixture may bring a gap', () => {
    expect(looserThan(base, base, names, new Set())).toEqual([
      `${label}: basic-values has a known gap and was removed or renamed; keep the fixture, or fix its gap first`,
    ]);
    const next = {
      ...base,
      KNOWN_GAPS: { ...base.KNOWN_GAPS, substitute: [list({ 'new-fixture': { pixels: 1 } })] },
    };
    expect(looserThan(base, next, names, new Set([...names, 'new-fixture']))).toEqual([]);
  });

  test('the run remedy for a clipped line agrees with the merge-base check', () => {
    // A run clips a line that its list does not name, for a fixture of the merge base and for a
    // new fixture.
    const run = (name: string) =>
      checkGate([{ ...result({ clippedLines: ['Wide'] }), name }], list({}), tools, true);
    expect(run('basic-values')[0]!.text).toContain(CLIPPED_RUN_REMEDY);
    // The owner route of the remedy: name the line in each list, accept the new fixture.
    const named = (name: string): Ratcheted => ({
      ...base,
      KNOWN_GAPS: {
        ...base.KNOWN_GAPS,
        substitute: [{ ...list({}), fixtureClipped: { [name]: ['Wide'] } }],
      },
    });
    const was = named('none');
    // A fixture of the merge base: refused with "fix the layout or the fixture".
    expect(
      looserThan(was, named('basic-values'), names, names, undefined, [], 'basic-values')
    ).toEqual([
      `substitute [poppler 1; chromium 1]: basic-values names the clipped line \`Wide\` 1 time, the merge base 0 times; ${CLIPPED_BASE_REMEDY}`,
    ]);
    // A new fixture: the acceptance that the run remedy names lets it pass.
    const head = new Set([...names, 'new-fixture']);
    expect(
      looserThan(was, named('new-fixture'), names, head, undefined, [], 'new-fixture')
    ).toEqual([]);
    expect(CLIPPED_RUN_REMEDY.startsWith(CLIPPED_BASE_REMEDY)).toBe(true);
  });

  test('a renamed fixture is not new, so it cannot take a gap', () => {
    // `plain` has no gap at the merge base; the change renames it to `plain-2` and records a
    // gap for it. With a fixture of the merge base gone, no fixture is new.
    const was = new Set([...names, 'plain']);
    const now = new Set([...names, 'plain-2']);
    expect(newFixtures(was, now)).toEqual(new Set());
    expect(newFixtures(was, new Set([...was, 'plain-2']))).toEqual(new Set(['plain-2']));
    const next = {
      ...base,
      KNOWN_GAPS: { ...base.KNOWN_GAPS, substitute: [list({ 'plain-2': { pixels: 1 } })] },
    };
    expect(looserThan(base, next, was, now)).toEqual([
      `substitute [poppler 1; chromium 1]: plain-2 has a new gap pixels (1); ${GAP_REMEDY}`,
    ]);
    // Its clipped lines are not new either: no acceptance applies.
    const clipped = {
      ...base,
      KNOWN_GAPS: {
        ...base.KNOWN_GAPS,
        substitute: [{ ...list({}), fixtureClipped: { 'plain-2': ['Wide'] } }],
      },
    };
    const withRecord = {
      ...base,
      KNOWN_GAPS: { ...base.KNOWN_GAPS, substitute: [{ ...list({}), fixtureClipped: {} }] },
    };
    expect(looserThan(withRecord, clipped, was, now, undefined, [], 'plain-2')).toEqual([
      `substitute [poppler 1; chromium 1]: plain-2 names the clipped line \`Wide\` 1 time, the merge base 0 times; ${CLIPPED_RENAMED_REMEDY}`,
    ]);
  });
});
