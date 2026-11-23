// The command-line flags of the benches. An unknown flag, a flag with no value, or a value that is
// not a positive whole number stops the run before any work, as DOCX Editor's
// `scripts/bench/edit-bench.ts` at 1c9b6a29b refuses an unknown argument: a typo must not start a
// run at the full default sizes. `--help` prints the usage.

/** A wrong argument: the bench prints the message and stops, with no stack trace. */
export class UsageError extends Error {}

/** Prints a `UsageError` and exits with code 2; throws any other error again. */
export function stopOnUsageError(error: unknown): never {
  if (!(error instanceof UsageError)) throw error;
  console.error(error.message);
  process.exit(2);
}

/** The flags of a bench: each takes a value, or is a switch with no value. */
export interface FlagSpec {
  readonly values: readonly string[];
  readonly switches: readonly string[];
}

/** The flag values of `argv`, or undefined after `--help` printed `usage`. */
export function readFlags(
  argv: readonly string[],
  spec: FlagSpec,
  usage: string
): Map<string, string | true> | undefined {
  if (argv.includes('--help') || argv.includes('-h')) {
    console.log(usage);
    return undefined;
  }
  const out = new Map<string, string | true>();
  for (let i = 0; i < argv.length; i += 1) {
    const flag = argv[i]!;
    if (spec.switches.includes(flag)) out.set(flag, true);
    else if (spec.values.includes(flag)) {
      const value = argv[i + 1];
      if (value === undefined || value.startsWith('--'))
        throw new UsageError(`${flag} takes a value\n\n${usage}`);
      out.set(flag, value);
      i += 1;
    } else throw new UsageError(`unknown argument: ${flag}\n\n${usage}`);
  }
  return out;
}

/** The positive whole number of `flag`, or `fallback`. */
export function positive(
  flags: Map<string, string | true>,
  flag: string,
  fallback: number
): number {
  const value = flags.get(flag);
  const n = typeof value === 'string' ? Number(value) : fallback;
  if (!Number.isInteger(n) || n < 1) throw new UsageError(`${flag} takes a positive whole number`);
  return n;
}

/** The names of `flag`, split at commas, each one of `allowed`. */
export function names<T extends string>(
  flags: Map<string, string | true>,
  flag: string,
  allowed: readonly T[]
): T[] {
  const value = flags.get(flag);
  const list = typeof value === 'string' ? value.split(',') : [...allowed];
  for (const name of list)
    if (!allowed.includes(name as T))
      throw new UsageError(`${flag} takes ${allowed.join(', ')}, not ${name}`);
  return list as T[];
}
