// Each case of the hang tests runs somewhere: in the engine run of `bun run test`, or in the
// browser run with a reason. A new case in `cases.ts` takes its place with no other change.

import { describe, expect, test } from 'bun:test';
import { BROWSER_CASES, BROWSER_RUN, browserReason } from './browser-cases.ts';
import { HANG_CASES } from './cases.ts';
import { MAX_WINDOW } from '../../packages/core/src/layout/index.ts';
import { substituteMeasurer } from '../../packages/core/src/layout/__tests__/test-fonts.ts';
import {
  ENGINE_CASES,
  engineProblems,
  engineShare,
  runEngineCase,
  WIDE_WINDOW,
  WINDOW,
  windowCap,
  type EngineRun,
  type WindowWork,
} from './engine-run.ts';
import { EDIT_REFUSED } from './first-cells.ts';

/** A cap for the checks of the checker, with the size of the cap of a real window. */
const CAP = { looked: 1840, keySteps: 1840, searches: 160, ruleChecks: 1840 };

const SHARES = 6;

describe('the runs of the hang cases', () => {
  test('each case runs in the engine run or in the browser run with a reason', () => {
    const engine = new Set(ENGINE_CASES.map((c) => c.id));
    const nowhere = HANG_CASES.filter((c) => !engine.has(c.id) && !browserReason(c));
    expect(nowhere.map((c) => c.id)).toEqual([]);
  });

  test('a run past a bound of the engine run fails its case', () => {
    const testCase = ENGINE_CASES.find((c) => c.id === 'sheet-dense')!;
    const caps = { window: CAP, wide: CAP };
    const work: WindowWork = {
      looked: 800,
      keySteps: 880,
      searches: 80,
      ruleChecks: 800,
      searchSteps: 1500,
      mergesRead: 0,
      drawingsRead: 0,
      indexBuilds: 0,
      rowsMeasured: 0,
      rowsCached: 0,
    };
    const good: EngineRun = {
      outcome: 'open',
      firstCell: '0',
      lastCell: '474981',
      storedRows: 25_000,
      storedCells: 500_000,
      top: { ...work, rowsMeasured: 25_000 },
      scrolled: work,
      edited: true,
      reopenedFirst: 'x',
      openedContent: 100,
      savedContent: 100,
      saveReads: true,
      unreadParts: [],
      firstSheet: 'Hostile',
    };
    expect(engineProblems(testCase, good, caps)).toEqual([]);
    for (const bad of [
      { top: { ...work, looked: CAP.looked + 1 } },
      { scrolled: { ...work, keySteps: CAP.keySteps + 1 } },
      { scrolled: { ...work, searches: CAP.searches + 1 } },
      { top: { ...work, ruleChecks: CAP.ruleChecks + 1 } },
      { top: { ...work, searchSteps: 80 * 19 + 1 } },
      { top: { ...work, rowsMeasured: 25_001 } },
      { scrolled: { ...work, indexBuilds: 1 } },
      { scrolled: { ...work, rowsMeasured: 1 } },
      { scrolled: { ...work, rowsCached: 1 } },
      { edited: false, editRefusal: 'invalid-sheet: x' },
      { reopenedFirst: 'first' },
      { savedContent: 49 },
      { saveReads: false },
      { unreadParts: ['/xl/styles.xml'] },
      { firstSheet: 'Recovery' },
      { lastCell: '0' },
      { outcome: 'refuse' as const },
    ])
      expect(
        `${JSON.stringify(bad)}: ${engineProblems(testCase, { ...good, ...bad }, caps).length}`
      ).toBe(`${JSON.stringify(bad)}: 1`);
    const refusing = ENGINE_CASES.find((c) => EDIT_REFUSED.has(c.id))!;
    const tookIt = { ...good, lastCell: undefined, firstCell: undefined };
    expect(engineProblems(refusing, tookIt, caps)).toContain(
      'the store took the edit of A1, which it must refuse'
    );
    const otherCode = { ...good, edited: false, editRefusal: 'other: x', editRefusalCode: 'other' };
    expect(engineProblems(refusing, otherCode, caps).join()).toContain('not ');
  });

  test('a layout of every row that a pane allows passes the cap of the window', async () => {
    const measurer = await substituteMeasurer();
    const dense = ENGINE_CASES.find((c) => c.id === 'sheet-dense')!;
    // A window as tall as `MAX_WINDOW` rows of 20 points: a layout that ignores the window.
    const tall = { ...WINDOW, height: MAX_WINDOW * 20 };
    const problems = engineProblems(
      dense,
      runEngineCase(dense, measurer, tall),
      windowCap(measurer)
    );
    expect(problems.some((p) => p.startsWith('the window at the top: looked'))).toBe(true);
    expect(problems.some((p) => p.startsWith('the window at the top: searches'))).toBe(true);
  }, 120_000);

  test('name only engine cases in the lists of expected edit refusals and wide windows', () => {
    const engine = new Set(ENGINE_CASES.map((c) => c.id));
    for (const id of [...EDIT_REFUSED.keys(), ...WIDE_WINDOW.keys()])
      expect(`${id}: ${engine.has(id)}`).toBe(`${id}: true`);
  });

  test('the shares hold each case of the engine run once', () => {
    const shared = Array.from({ length: SHARES }, (_, i) => engineShare(i + 1, SHARES)).flat();
    expect(shared.map((c) => c.id).sort()).toEqual(ENGINE_CASES.map((c) => c.id).sort());
    for (let n = 1; n <= SHARES; n += 1)
      expect(Bun.file(new URL(`engine-cases-${n}.test.ts`, import.meta.url)).size).toBeGreaterThan(
        0
      );
  });

  test('the browser run keeps only the cases that need a browser, each with a reason', () => {
    for (const id of BROWSER_RUN.keys()) {
      const c = HANG_CASES.find((k) => k.id === id);
      expect(`${id}: ${c ? 'file case' : 'no case'}`).toBe(`${id}: file case`);
      expect(c?.network ?? c?.race).toBeUndefined();
    }
    expect(BROWSER_CASES.every((c) => browserReason(c))).toBe(true);
    const faults = HANG_CASES.filter((c) => c.network || c.race);
    expect(faults.every((c) => BROWSER_CASES.includes(c))).toBe(true);
  });
});
