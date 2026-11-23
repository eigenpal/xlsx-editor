// The verdict of one case of the browser hang test, from what the run measured. The order is the
// order of the header of `run.ts`: the worst verdict that applies wins. `SLOW` is a mark, and it
// never fails the run.

export type Verdict = 'ok' | 'SLOW' | 'ERROR' | 'FETCH' | 'BLOCK' | 'HANG' | 'CRASH';

/** Whether a verdict fails the run. `ok` and `SLOW` do not. */
export const fails = (verdict: Verdict): boolean => verdict !== 'ok' && verdict !== 'SLOW';

export interface Measured {
  /** The verdict of the case steps: `ok`, or the first step that failed. */
  readonly stepVerdict: Verdict;
  readonly stepWhy?: string | undefined;
  /** Why the watchdog stopped the browser, or null. */
  readonly tripped: string | null;
  /** Whether the watchdog stopped the browser for its memory. */
  readonly trippedMemory: boolean;
  /** Whether the page or the browser closed during the case. */
  readonly closed: boolean;
  readonly peakMiB: number;
  readonly longestBlockMs: number;
  readonly blockLimitMs: number;
  /** The sum of the main-thread blocks of 100 ms or more during the case. */
  readonly repeatedBlockMs: number;
  readonly repeatedBlockLimitMs: number;
  readonly fetched: readonly string[];
  readonly errors: readonly string[];
}

/** Whether a refusal shows an exception of the engine, and not a reason of the editor. */
export const showsException = (refusal: string): boolean => /\b[A-Za-z]*Error: /.test(refusal);

export function verdictOf(m: Measured): { verdict: Verdict; why?: string } {
  // The watchdog stops the browser at the memory cap, which closes it as a crash would.
  if (m.tripped) return { verdict: m.trippedMemory ? 'CRASH' : 'HANG', why: m.tripped };
  if (m.closed)
    return { verdict: 'CRASH', why: `the page or the browser closed. ${m.stepWhy ?? ''}`.trim() };
  if (m.stepVerdict === 'HANG') return { verdict: 'HANG', why: m.stepWhy ?? 'a step did not end' };
  if (m.longestBlockMs > m.blockLimitMs)
    return {
      verdict: 'BLOCK',
      why: `the main thread ran no timer for ${m.longestBlockMs} ms (limit ${m.blockLimitMs})`,
    };
  if (m.fetched.length > 0) return { verdict: 'FETCH', why: `the page asked for ${m.fetched[0]}` };
  if (m.stepVerdict !== 'ok')
    return { verdict: m.stepVerdict, ...(m.stepWhy ? { why: m.stepWhy } : {}) };
  if (m.errors.length > 0) return { verdict: 'ERROR', why: m.errors[0]! };
  // A case with no memory reading proves nothing about the memory cap.
  if (m.peakMiB === 0)
    return { verdict: 'ERROR', why: 'the watchdog read no memory of the browser' };
  // Many short blocks are slow, but each one lets the page answer, so the mark does not fail.
  if (m.repeatedBlockMs > m.repeatedBlockLimitMs)
    return {
      verdict: 'SLOW',
      why: `the blocks of 100 ms or more add up to ${m.repeatedBlockMs} ms (mark ${m.repeatedBlockLimitMs})`,
    };
  return { verdict: 'ok' };
}
