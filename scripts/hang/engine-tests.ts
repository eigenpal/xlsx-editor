// The tests of one share of the engine run (`engine-run.ts`), for `engine-cases-*.test.ts`.

import { beforeAll, expect, test } from 'bun:test';
import { substituteMeasurer } from '../../packages/core/src/layout/__tests__/test-fonts.ts';
import type { TextMeasurer } from '../../packages/core/src/layout/text/measurer.ts';
import { engineProblems, engineShare, runEngineCase, windowCap } from './engine-run.ts';

/** The time of the largest case on a slow runner: it builds and reads 50 MB of XML. */
const CASE_TIMEOUT_MS = 120_000;

/** Defines one test for each case of share `part` of `of`. */
export function engineTests(part: number, of: number): void {
  let measurer: TextMeasurer;
  beforeAll(async () => {
    measurer = await substituteMeasurer();
  });
  for (const testCase of engineShare(part, of))
    test(
      `${testCase.id}: ${testCase.covers}`,
      () => {
        const run = runEngineCase(testCase, measurer);
        expect(engineProblems(testCase, run, windowCap(measurer))).toEqual([]);
      },
      CASE_TIMEOUT_MS
    );
}
