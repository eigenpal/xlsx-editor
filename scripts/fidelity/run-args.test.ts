// The command lines of the fidelity commands (`run-args.ts`).

import { describe, expect, test } from 'bun:test';
import {
  CHECK_GAPS_USAGE,
  checkGapsArgs,
  FIDELITY_USAGE,
  fidelityArgs,
  RECORD_REFERENCE_USAGE,
  wrapList,
  recordReferenceArgs,
} from './run-args.ts';

const fixtures = ['basic-values', 'number-formats'];

/** Whether `parsed` is a usage error that ends with `usage`, and holds `part`. */
const refused = (parsed: object, usage: string, part = ''): boolean =>
  'error' in parsed &&
  typeof parsed.error === 'string' &&
  parsed.error.endsWith(usage) &&
  parsed.error.includes(part);

describe('the command line of bun run fidelity', () => {
  test('fixture names, --update-gaps, and --accept-worse pass', () => {
    const none = { updateGaps: false, acceptWorse: false };
    expect(fidelityArgs([], fixtures)).toEqual({ names: [], ...none });
    expect(fidelityArgs(['number-formats'], fixtures)).toEqual({
      names: ['number-formats'],
      ...none,
    });
    expect(fidelityArgs(['--update-gaps'], fixtures)).toEqual({
      names: [],
      updateGaps: true,
      acceptWorse: false,
    });
    expect(fidelityArgs(['number-formats', '--update-gaps', '--accept-worse'], fixtures)).toEqual({
      names: ['number-formats'],
      updateGaps: true,
      acceptWorse: true,
    });
  });

  test('--accept-worse without --update-gaps is refused before any work', () => {
    expect(
      refused(fidelityArgs(['--accept-worse'], fixtures), FIDELITY_USAGE, '--update-gaps')
    ).toBe(true);
  });

  test('an unknown flag or fixture is refused with the usage text', () => {
    for (const argv of [['--update-gap'], ['-u'], ['--update-gaps=1']])
      expect(refused(fidelityArgs(argv, fixtures), FIDELITY_USAGE)).toBe(true);
    // The error lists the fixture names.
    expect(
      refused(fidelityArgs(['other'], fixtures), FIDELITY_USAGE, 'basic-values, number-formats')
    ).toBe(true);
  });

  test('--help and -h ask for the usage text, which names each flag and the exit codes', () => {
    expect(fidelityArgs(['--help'], fixtures)).toEqual({ help: true });
    expect(fidelityArgs(['-h', 'other'], fixtures)).toEqual({ help: true });
    for (const part of [
      '--update-gaps',
      '--accept-worse',
      '--help',
      '0    pass',
      '1    the gate fails',
      '2    a usage error',
      'a run lock that another run holds',
      'a report folder\n       that cannot be cleared',
      '130  SIGINT',
      '143  SIGTERM',
    ])
      expect(FIDELITY_USAGE).toContain(part);
    // An unknown option names itself and points at the help text.
    const bogus = fidelityArgs(['--bogus'], fixtures);
    expect('error' in bogus && bogus.error.split('\n')[0]).toBe("unknown option '--bogus'");
    // The messages of the parser start in lowercase, as the other usage errors do, and an
    // argument that check-gaps does not take points at --base.
    const extra = checkGapsArgs(['main']);
    expect('error' in extra && extra.error.split('\n')[0]).toBe(
      "unexpected argument 'main'; name the merge base with --base <ref>"
    );
    const missing = checkGapsArgs(['--base']);
    expect('error' in missing && missing.error.split('\n')[0]).toBe(
      "option '--base <value>' argument missing"
    );
    // The help of check-gaps names the variables with which an owner accepts a looser list.
    for (const name of ['XLSX_FIDELITY_ACCEPT_CLIPPED', 'XLSX_FIDELITY_ACCEPT_NEW_TOOLCHAIN'])
      expect(CHECK_GAPS_USAGE).toContain(`${name}=`);
    // Each line of the help texts fits 80 columns.
    for (const usage of [FIDELITY_USAGE, CHECK_GAPS_USAGE, RECORD_REFERENCE_USAGE])
      for (const line of usage.split('\n')) expect(line.length).toBeLessThanOrEqual(80);
    const names = Array.from({ length: 12 }, (_, i) => `fixture-name-${i}`);
    const listed = wrapList('the fixtures are', names);
    expect(listed.length).toBeGreaterThan(1);
    for (const line of listed) expect(line.length).toBeLessThanOrEqual(80);
    expect(listed.join(' ').replace(/\s+/g, ' ')).toBe(`the fixtures are ${names.join(', ')}`);
  });
});

describe('the command line of bun run fidelity:check-gaps', () => {
  test('no arguments, or --base with a ref, pass', () => {
    expect(checkGapsArgs([])).toEqual({});
    expect(checkGapsArgs(['--base', 'origin/main'])).toEqual({ base: 'origin/main' });
  });

  test('an unknown flag, a missing ref, or a positional is refused with the usage text', () => {
    for (const argv of [['--bse', 'main'], ['--base'], ['--base', ''], ['main']])
      expect(refused(checkGapsArgs(argv), CHECK_GAPS_USAGE)).toBe(true);
  });

  test('--help and -h ask for the usage text, which names each flag and the exit codes', () => {
    expect(checkGapsArgs(['--help'])).toEqual({ help: true });
    expect(checkGapsArgs(['-h'])).toEqual({ help: true });
    expect(checkGapsArgs(['--require-thresholds'])).toEqual({ requireThresholds: true });
    for (const part of ['--base <ref>', '--help', '0    not looser', '1    looser', 'not a commit'])
      expect(CHECK_GAPS_USAGE).toContain(part);
  });
});

describe('the command line of bun run fidelity:record-reference', () => {
  test('both separators and fixture names pass', () => {
    expect(
      recordReferenceArgs(['--decimal', ',', '--group', ' ', 'basic-values'], fixtures)
    ).toEqual({ decimal: ',', group: ' ', names: ['basic-values'] });
  });

  test('a missing separator, no fixture, an unknown fixture, or an unknown flag is refused', () => {
    for (const argv of [
      ['--decimal', ',', 'basic-values'],
      ['--group', ' ', 'basic-values'],
      ['--decimal', ',', '--group', ' '],
      ['--decimal', ',', '--group', ' ', 'other'],
      ['--decimal', ',', '--group', ' ', '--force', 'basic-values'],
    ])
      expect(refused(recordReferenceArgs(argv, fixtures), RECORD_REFERENCE_USAGE)).toBe(true);
    expect(recordReferenceArgs(['-h'], fixtures)).toEqual({ help: true });
  });
});
