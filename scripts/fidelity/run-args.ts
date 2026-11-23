// The command lines of `bun run fidelity` (`run.ts`), `bun run fidelity:check-gaps`
// (`check-known-gaps.ts`), and `bun run fidelity:record-reference` (`record-reference.ts`).
// Same role as pptx-editor/scripts/fidelity/run-args.ts at 1943d2c9 (EigenPal): flags are
// strict, and an unknown flag, a missing value, an unknown fixture, or flags that do not go
// together are a usage error (exit 2), never ignored; `--help` and `-h` print the usage text and
// exit 0. Difference: each script takes its own small set of arguments, so each has its own
// parser and usage text.

import { parseArgs } from 'node:util';

/** The usage text of `bun run fidelity`. */
export const FIDELITY_USAGE = `Usage: bun run fidelity [fixture names...]
                        [--update-gaps [--accept-worse]]
  Measures each fixture, or only the named ones, against its reference PDF
  and the known gaps.
  --update-gaps   record the known gaps of this toolchain in thresholds.ts,
                  then run bun run fidelity:check-gaps; with fixture names,
                  record only those fixtures. It writes nothing when the
                  run adds a gap, a check, or a clipped line, or when a
                  value is looser
  --accept-worse  with --update-gaps, write such a list too; then the
                  merge-base check applies its own rules, also in CI
  -h, --help      print this text
Exit codes:
  0    pass
  1    the gate fails, --update-gaps writes nothing, the list is looser
       than at the merge base, or the harness stops
  2    a usage error, a run lock that another run holds, a report folder
       that cannot be cleared, a run that measured nothing, an --update-gaps
       write that failed, or a merge-base check that could not compare (see
       its line)
  130  SIGINT stopped the run or the check against the merge base
  143  SIGTERM stopped the run or the check against the merge base`;

/** The usage text of `bun run fidelity:check-gaps`. */
export const CHECK_GAPS_USAGE = `Usage: bun run fidelity:check-gaps [--base <ref>]
                                   [--require-thresholds]
  Fails when a threshold or a known gap is looser than at the merge base.
  --base <ref>    compare with <ref> instead of the merge base of HEAD and
                  origin/main
  --require-thresholds
                  exit 2, not 0, when the base has no thresholds.ts;
                  bun run fidelity --update-gaps passes it
  -h, --help      print this text
Environment, for an owner only; CI never sets them:
  XLSX_FIDELITY_ACCEPT_CLIPPED=<fixture>[,<fixture>...]
                  accept the clipped lines of a new fixture, or of a fixture
                  with a new reference PDF
  XLSX_FIDELITY_ACCEPT_NEW_TOOLCHAIN=<poppler>; <chromium>
                  accept a looser list for a new toolchain
Exit codes:
  0    not looser, or a base without thresholds.ts
  1    looser
  2    a usage error, a --base that is not a commit, no merge base, a base
       without thresholds.ts with --require-thresholds, git that cannot
       run, or an error that stops the check`;

/** The usage text of `bun run fidelity:record-reference`. */
export const RECORD_REFERENCE_USAGE = `Usage: bun run fidelity:record-reference --decimal <sep> --group <sep>
                                         <fixture names...>
  Records the reference of each named fixture in fixtures/references.json.
  --decimal <sep> the decimal separator that the reference PDF prints
  --group <sep>   the group separator that the reference PDF prints
  -h, --help      print this text
Exit codes:
  0    recorded
  1    a refused reference PDF
  2    a usage error, a missing file, reference fonts that cannot be read,
       or an error that stops the recorder`;

/**
 * `label` and the comma-separated `items` as lines of at most 80 columns. The lines after the
 * first are indented by two spaces.
 */
export function wrapList(label: string, items: readonly string[]): string[] {
  const lines: string[] = [];
  let line = label;
  items.forEach((item, i) => {
    const word = `${item}${i < items.length - 1 ? ',' : ''}`;
    if (line.length + 1 + word.length > 80 && line.trim()) {
      lines.push(line);
      line = `  ${word}`;
    } else line = `${line} ${word}`;
  });
  lines.push(line);
  return lines;
}

/** A parse that asks for the usage text. */
export interface HelpArgs {
  readonly help: true;
}

/** A usage error, with the usage text after it. */
export interface UsageError {
  readonly error: string;
}

/**
 * The message of a usage error from `parseArgs`, or of any other value, in the style of the other
 * usage errors: its first sentence, starting in lowercase. The usage text follows it, so the
 * message does not point at `--help`. The message of `parseArgs` for an unknown option suggests
 * `--`, which no script here takes, and `positional` replaces the second sentence of its message
 * for an argument that the script does not take.
 */
function messageOf(error: unknown, positional?: string): string {
  const text = error instanceof Error ? error.message : String(error);
  const first = text.split('. ')[0]!;
  const lower = first.charAt(0).toLowerCase() + first.slice(1);
  const code = (error as { code?: string }).code;
  return positional && code === 'ERR_PARSE_ARGS_UNEXPECTED_POSITIONAL'
    ? `${lower}; ${positional}`
    : lower;
}

/** The names of `names` that are not in `fixtures`, as a usage error, or undefined. */
function unknownFixtures(
  names: readonly string[],
  fixtures: readonly string[],
  usage: string
): UsageError | undefined {
  const unknown = names.filter((n) => !fixtures.includes(n));
  if (!unknown.length) return undefined;
  return {
    error: [
      `unknown fixtures: ${unknown.join(', ')}`,
      ...wrapList('the fixtures are', fixtures),
      usage,
    ].join('\n'),
  };
}

/** The arguments of `bun run fidelity`, or the usage error. `fixtures`: the fixture names. */
export function fidelityArgs(
  argv: readonly string[],
  fixtures: readonly string[]
): { names: string[]; updateGaps: boolean; acceptWorse: boolean } | HelpArgs | UsageError {
  try {
    const { values, positionals } = parseArgs({
      args: [...argv],
      options: {
        'update-gaps': { type: 'boolean' },
        'accept-worse': { type: 'boolean' },
        help: { type: 'boolean', short: 'h' },
      },
      allowPositionals: true,
      strict: true,
    });
    if (values.help) return { help: true };
    const updateGaps = values['update-gaps'] === true;
    const acceptWorse = values['accept-worse'] === true;
    if (acceptWorse && !updateGaps)
      return { error: `--accept-worse needs --update-gaps\n${FIDELITY_USAGE}` };
    return (
      unknownFixtures(positionals, fixtures, FIDELITY_USAGE) ?? {
        names: positionals,
        updateGaps,
        acceptWorse,
      }
    );
  } catch (error) {
    return { error: `${messageOf(error)}\n${FIDELITY_USAGE}` };
  }
}

/** The arguments of `bun run fidelity:check-gaps`, or the usage error. */
export function checkGapsArgs(
  argv: readonly string[]
): { base?: string; requireThresholds?: true } | HelpArgs | UsageError {
  try {
    const { values } = parseArgs({
      args: [...argv],
      options: {
        base: { type: 'string' },
        'require-thresholds': { type: 'boolean' },
        help: { type: 'boolean', short: 'h' },
      },
      allowPositionals: false,
      strict: true,
    });
    if (values.help) return { help: true };
    const require =
      values['require-thresholds'] === true ? { requireThresholds: true as const } : {};
    if (values.base === undefined) return require;
    if (!values.base) return { error: `--base needs a ref\n${CHECK_GAPS_USAGE}` };
    return { base: values.base, ...require };
  } catch (error) {
    return {
      error: `${messageOf(error, 'name the merge base with --base <ref>')}\n${CHECK_GAPS_USAGE}`,
    };
  }
}

/** The arguments of `bun run fidelity:record-reference`, or the usage error. */
export function recordReferenceArgs(
  argv: readonly string[],
  fixtures: readonly string[]
): { decimal: string; group: string; names: string[] } | HelpArgs | UsageError {
  try {
    const { values, positionals } = parseArgs({
      args: [...argv],
      options: {
        decimal: { type: 'string' },
        group: { type: 'string' },
        help: { type: 'boolean', short: 'h' },
      },
      allowPositionals: true,
      strict: true,
    });
    if (values.help) return { help: true };
    for (const flag of ['decimal', 'group'] as const)
      if (!values[flag])
        return {
          error: [
            `--${flag} is required: the separator that the reference PDF prints`,
            RECORD_REFERENCE_USAGE,
          ].join('\n'),
        };
    if (!positionals.length)
      return { error: `name one or more fixtures\n${RECORD_REFERENCE_USAGE}` };
    return (
      unknownFixtures(positionals, fixtures, RECORD_REFERENCE_USAGE) ?? {
        decimal: values.decimal!,
        group: values.group!,
        names: positionals,
      }
    );
  } catch (error) {
    return { error: `${messageOf(error)}\n${RECORD_REFERENCE_USAGE}` };
  }
}
