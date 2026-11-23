import { describe, expect, test } from 'bun:test';
import { fails, showsException, verdictOf, type Measured } from './verdict.ts';

const clean: Measured = {
  stepVerdict: 'ok',
  tripped: null,
  trippedMemory: false,
  closed: false,
  peakMiB: 500,
  longestBlockMs: 100,
  blockLimitMs: 5000,
  repeatedBlockMs: 0,
  repeatedBlockLimitMs: 10_000,
  fetched: [],
  errors: [],
};

describe('the refusal of a hang case', () => {
  test.each([
    ['Could not open a.xlsx: RangeError: Invalid array length', true],
    ['Could not open a.xlsx: TypeError: x is not a function', true],
    ['Could not open a.xlsx: Error: boom', true],
    ['Could not open a.xlsx: the part is larger than the limit.', false],
    ['Could not open a.xlsx: the workbook has an error in its styles.', false],
  ])('%s', (text, shown) => {
    expect(showsException(text)).toBe(shown);
  });
});

describe('the verdict of a hang case', () => {
  test('passes a clean case', () => {
    expect(verdictOf(clean).verdict).toBe('ok');
  });

  test.each([
    ['CRASH', { tripped: 'the browser used 4000 MiB', trippedMemory: true }],
    ['HANG', { tripped: 'the case passed 180000 ms' }],
    ['CRASH', { closed: true }],
    ['HANG', { stepVerdict: 'HANG', stepWhy: 'no sheet' }],
    ['BLOCK', { longestBlockMs: 5001 }],
    ['SLOW', { repeatedBlockMs: 10_001 }],
    ['FETCH', { fetched: ['https://hang-test.invalid/remote.png'] }],
    ['ERROR', { errors: ['boom'] }],
    ['ERROR', { peakMiB: 0 }],
  ] as const)('gives %s', (verdict, change) => {
    expect(verdictOf({ ...clean, ...change }).verdict).toBe(verdict);
  });

  test('puts a block and a request to another origin before a failed step', () => {
    const failed = { ...clean, stepVerdict: 'ERROR', stepWhy: 'no recovery' } as const;
    expect(verdictOf({ ...failed, longestBlockMs: 9000 }).verdict).toBe('BLOCK');
    expect(verdictOf({ ...failed, fetched: ['https://hang-test.invalid/a'] }).verdict).toBe(
      'FETCH'
    );
    expect(verdictOf(failed)).toEqual({ verdict: 'ERROR', why: 'no recovery' });
  });

  test('keeps a block at the limit of the case', () => {
    expect(verdictOf({ ...clean, longestBlockMs: 5000 }).verdict).toBe('ok');
    expect(verdictOf({ ...clean, repeatedBlockMs: 10_000 }).verdict).toBe('ok');
  });

  test('fails the run on each verdict but ok and SLOW', () => {
    expect(fails('ok')).toBe(false);
    expect(fails('SLOW')).toBe(false);
    for (const verdict of ['ERROR', 'FETCH', 'BLOCK', 'HANG', 'CRASH'] as const)
      expect(fails(verdict)).toBe(true);
  });

  test('puts each failing verdict before the SLOW mark', () => {
    const slow = { ...clean, repeatedBlockMs: 12_000 };
    expect(verdictOf(slow)).toEqual({
      verdict: 'SLOW',
      why: 'the blocks of 100 ms or more add up to 12000 ms (mark 10000)',
    });
    expect(verdictOf({ ...slow, longestBlockMs: 5001 }).verdict).toBe('BLOCK');
    expect(verdictOf({ ...slow, errors: ['boom'] }).verdict).toBe('ERROR');
    expect(verdictOf({ ...slow, peakMiB: 0 }).verdict).toBe('ERROR');
    expect(verdictOf({ ...slow, stepVerdict: 'ERROR', stepWhy: 'no recovery' }).verdict).toBe(
      'ERROR'
    );
    expect(verdictOf({ ...slow, fetched: ['https://hang-test.invalid/a'] }).verdict).toBe('FETCH');
    expect(verdictOf({ ...slow, stepVerdict: 'HANG' }).verdict).toBe('HANG');
  });

  test('puts the watchdog before a crash that its own stop causes', () => {
    const m = { ...clean, tripped: 'the browser used 4000 MiB', trippedMemory: true, closed: true };
    expect(verdictOf(m)).toEqual({ verdict: 'CRASH', why: 'the browser used 4000 MiB' });
  });
});
