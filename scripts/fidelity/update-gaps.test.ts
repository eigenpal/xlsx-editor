// `bun run fidelity --update-gaps` writes no worse list without `--accept-worse` (`update-gaps.ts`).

import { describe, expect, test } from 'bun:test';
import { FEWER_LINES_REMEDY, type FixtureResult } from './known-gaps.ts';
import type { KnownGapList } from './thresholds.ts';
import {
  listChanges,
  newListRefusal,
  NOTHING_WRITTEN,
  planUpdate,
  updateExit,
} from './update-gaps.ts';

const tools = { poppler: 'poppler 1.0', chromium: 'chromium 1.0' };
const values = {
  pixels: 0,
  'line-breaks': 1,
  baseline: 0,
  'line-position': 0,
  'extra-text': 0,
};
const result = (name: string, pixels: number, clipped: string[] = []): FixtureResult => ({
  name,
  status: 'compared',
  failures: pixels > 0.005 ? ['pixels'] : [],
  values: { ...values, pixels },
  linesTotal: 3,
  linesMatched: 3,
  clippedLines: clipped,
});
const old: KnownGapList = {
  ...tools,
  fixtureLines: { a: 3, b: 3 },
  fixtureClipped: { a: ['Kept @ 47.0 | right clip 89.0'] },
  fixtures: { a: { pixels: 0.009 }, b: { pixels: 0.008 } },
};
const plan = (results: FixtureResult[], complete = true, acceptWorse = false) =>
  planUpdate([old], results, tools, complete, acceptWorse);

describe('--update-gaps', () => {
  test('a run that is not worse writes its list', () => {
    const update = plan([result('a', 0.007, ['Kept @ 47.0 | right clip 89.0']), result('b', 0)]);
    if ('refused' in update) throw new Error(update.refused);
    expect(update.write).toBe(true);
    expect(update.next.fixtures).toEqual({ a: { pixels: 0.007 } });
    expect(update.changes).toEqual({ looser: [], added: [], removed: ['b pixels passes now'] });
  });

  test('a worse run writes nothing, and --accept-worse writes it', () => {
    const worse = [result('a', 0.0095, ['Kept @ 47.0 | right clip 89.0']), result('b', 0.008)];
    const refused = plan(worse);
    if ('refused' in refused) throw new Error(refused.refused);
    expect(refused.write).toBe(false);
    expect(refused.changes.looser).toEqual(['a pixels: 0.009 → 0.0095']);
    const accepted = plan(worse, true, true);
    if ('refused' in accepted) throw new Error(accepted.refused);
    expect(accepted.write).toBe(true);
    expect(accepted.next.fixtures.a).toEqual({ pixels: 0.0095 });
    // A new kind, a new fixture entry, and a clipped line named more often are new entries.
    const added = plan([
      result('a', 0.009, ['Kept @ 47.0 | right clip 89.0', 'Next @ 61.0 | left clip 20.0']),
      result('b', 0.008),
      result('c', 0.02),
    ]);
    if ('refused' in added) throw new Error(added.refused);
    expect(added.write).toBe(false);
    expect(added.changes.added).toEqual([
      'c: a new gap pixels (0.02)',
      'a: names the clipped line `Next @ 61.0 | left clip 20.0` 1 time, the old list 0 times',
    ]);
    expect(NOTHING_WRITTEN).toStartWith('Nothing was written');
    expect(NOTHING_WRITTEN).toContain('--accept-worse');
  });

  test('fewer checked lines are looser', () => {
    const next: KnownGapList = { ...old, fixtureLines: { a: 2, b: 3 } };
    expect(listChanges(old, next).looser).toEqual([
      `a: checks fewer lines (3 → 2); ${FEWER_LINES_REMEDY}`,
    ]);
    // A fixture with no line count checks 0 lines, unless it no longer exists.
    const gone: KnownGapList = { ...old, fixtureLines: { b: 3 } };
    expect(listChanges(old, gone).looser).toEqual([
      `a: checks fewer lines (3 → 0); ${FEWER_LINES_REMEDY}`,
    ]);
    expect(listChanges(old, gone, (name) => name !== 'a').looser).toEqual([]);
  });

  test('a run of some fixtures replaces only their entries', () => {
    const update = plan([result('b', 0.007)], false);
    if ('refused' in update) throw new Error(update.refused);
    expect(update.write).toBe(true);
    expect(update.next.fixtures).toEqual({ a: { pixels: 0.009 }, b: { pixels: 0.007 } });
    expect(update.next.fixtureClipped).toEqual(old.fixtureClipped);
  });

  test('a run with other tools starts a list only from every fixture', () => {
    const other = { poppler: 'poppler 2.0', chromium: 'chromium 1.0' };
    const partial = planUpdate([old], [result('a', 0.009)], other, false, false);
    expect('refused' in partial && partial.refused).toContain('without fixture names');
    // It compares with the tightest values of the environment.
    const complete = planUpdate([old], [result('a', 0.01), result('b', 0)], other, true, false);
    if ('refused' in complete) throw new Error(complete.refused);
    expect(complete.write).toBe(false);
    expect(complete.changes.looser).toEqual(['a pixels: 0.009 → 0.01']);
    expect(complete.next.poppler).toBe('poppler 2.0');
  });

  test('a new list starts only from a run of every fixture with no error', () => {
    const other = { poppler: 'poppler 2.0', chromium: 'chromium 1.0' };
    const error: FixtureResult = { name: 'a', status: 'error', error: 'page crashed' };
    const update = planUpdate([old], [error, result('b', 0)], other, true, false);
    expect('refused' in update && update.refused).toContain('these did not compare: a');
    // The refusal before the run measures needs only the tools and the fixture set.
    expect(newListRefusal([old], other, false)).toContain('without fixture names');
    expect(newListRefusal([old], other, true)).toBeUndefined();
    expect(newListRefusal([old], tools, false)).toBeUndefined();
    // With a list of its tools, a fixture with an error keeps its entries.
    const kept = plan([error, result('b', 0.008)]);
    if ('refused' in kept) throw new Error(kept.refused);
    expect([kept.write, kept.next.fixtures.a, kept.next.fixtureLines.a]).toEqual([
      true,
      { pixels: 0.009 },
      3,
    ]);
  });

  test('the exit of an update follows the write and the merge-base check', () => {
    const refused = updateExit(0, {});
    expect(refused).toEqual({ code: 1, line: NOTHING_WRITTEN });
    expect(updateExit(0, { refused: 'why' })).toEqual({ code: 1, line: 'fidelity: why' });
    const written = { written: true };
    expect(updateExit(0, written, { status: 0 })).toEqual({ code: 0 });
    expect(updateExit(1, written, { status: 0 })).toEqual({ code: 1 });
    expect(updateExit(0, written, { status: 1 })).toEqual({ code: 1 });
    const nothing = updateExit(0, written, { status: 2 });
    expect(nothing.code).toBe(2);
    expect(nothing.line).toContain('compared nothing');
    expect(nothing.line).not.toContain('fetch');
    const start = updateExit(0, written, { status: null, error: new Error('spawn ENOENT') });
    expect(start).toEqual({
      code: 2,
      line: 'fidelity: the merge-base check did not start: spawn ENOENT; thresholds.ts holds the new list unchecked; fix the cause, then run bun run fidelity:check-gaps',
    });
    // A failed gate keeps exit code 1, with the same line.
    expect(updateExit(1, written, { status: 2 }).code).toBe(1);
    // A signal passes through as a shell gives it, and the line names it.
    expect(updateExit(0, written, { status: null, signal: 'SIGINT' })).toEqual({
      code: 130,
      line: 'fidelity: SIGINT stopped the merge-base check; thresholds.ts holds the new list unchecked',
    });
    expect(updateExit(0, written, { status: null, signal: 'SIGTERM' }).code).toBe(143);
    // A write that fails says so; the report is written first (`run.ts`).
    expect(updateExit(0, { failed: 'prettier did not format thresholds.ts' })).toEqual({
      code: 2,
      line: 'fidelity: the known gaps could not be written: prettier did not format thresholds.ts; check thresholds.ts with git diff',
    });
    expect(NOTHING_WRITTEN).toContain('bun run fidelity:check-gaps --base <ref>');
  });
});
