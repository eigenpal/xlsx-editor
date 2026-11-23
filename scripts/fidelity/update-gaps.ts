// `bun run fidelity --update-gaps`: plan the list that a run records, and refuse a worse one.
//
// Same role as `planUpdate`, `listChanges`, and `mergeRecorded` of
// pptx-editor/scripts/fidelity/known-gaps.ts at 1943d2c9 (EigenPal). As there, the run writes
// nothing when the new list adds an entry or a kind, or is looser, unless `--accept-worse` is
// given, and a run of some fixtures replaces only the entries of the fixtures that it measured.
// Differences:
// - A list here records values, not only kinds. So a value that grows by any amount
//   (`looserRecord`) is looser too, as it is for the merge-base check (`check-known-gaps.ts`).
// - A list names its clipped lines (`fixtureClipped`). A line that the new list names more often
//   is a new entry.
// - A run with tools that have no list and a fixture error writes nothing (`newListRefusal`).
//   The source writes its one list then under the tools of the old list (`stamp` of
//   `mergeRecorded`). Here each toolchain has its own list, so a list under the old tools would
//   name values that those tools did not measure.
// - Each raster toolchain keeps its own list. The old list is the list of the run's tools. A run
//   with other tools compares with the tightest values of its font environment
//   (`tightestList`), and it must measure every fixture, because it starts a list.
// - After a write, the run starts the merge-base check (`bun run fidelity:check-gaps
//   --require-thresholds`), and `updateExit` gives its exit code. The source starts no check:
//   here a written list must still pass the rules of the merge base before the change lands.

import { constants } from 'node:os';
import {
  clippedOf,
  exactList,
  FEWER_LINES_REMEDY,
  lineCounts,
  looserRecord,
  recordedList,
  tightestList,
  timesText,
  type FixtureResult,
  type GapTools,
} from './known-gaps.ts';
import type { FailureKind, KnownGapList } from './thresholds.ts';

/** How a new list differs from the old one. */
export interface ListChanges {
  /** Values that grew, and fixtures that check fewer lines. */
  readonly looser: readonly string[];
  /** New entries: a gap of a new fixture, a new kind, or a clipped line named more often. */
  readonly added: readonly string[];
  /** Entries that are gone: checks that pass now, and clipped lines named less often. */
  readonly removed: readonly string[];
}

/** What `--update-gaps` does with a run. */
export type GapUpdate =
  | {
      /** The list to write, the list of the run's tools after the run. */
      readonly next: KnownGapList;
      readonly changes: ListChanges;
      /** False when the list is worse (`ListChanges.looser`, `added`) without `acceptWorse`. */
      readonly write: boolean;
    }
  | { readonly refused: string };

/** The list with no entry, to compare the first list of an environment with. */
const EMPTY: Omit<KnownGapList, 'poppler' | 'chromium'> = { fixtureLines: {}, fixtures: {} };

/**
 * The changes of `next` against `old` (`ListChanges`). A fixture of `old` that still exists
 * (`exists`, by default every fixture) and that `next` gives no line count checks 0 lines, as in
 * the merge-base check (`check-known-gaps.ts`).
 */
export function listChanges(
  old: KnownGapList,
  next: KnownGapList,
  exists: (name: string) => boolean = () => true
): ListChanges {
  const looser: string[] = [];
  const added: string[] = [];
  const removed: string[] = [];
  for (const [name, gap] of Object.entries(next.fixtures)) {
    const was = Object.hasOwn(old.fixtures, name) ? old.fixtures[name] : undefined;
    for (const [kind, value] of Object.entries(gap) as [FailureKind, number][]) {
      const recorded = was?.[kind];
      if (recorded === undefined) added.push(`${name}: a new gap ${kind} (${value})`);
      else if (looserRecord(kind, value, recorded))
        looser.push(`${name} ${kind}: ${recorded} → ${value}`);
    }
  }
  for (const [name, gap] of Object.entries(old.fixtures)) {
    const now = Object.hasOwn(next.fixtures, name) ? next.fixtures[name] : undefined;
    for (const kind of Object.keys(gap) as FailureKind[])
      if (now?.[kind] === undefined) removed.push(`${name} ${kind} passes now`);
  }
  for (const [name, lines] of Object.entries(old.fixtureLines)) {
    if (!exists(name)) continue;
    const now = Object.hasOwn(next.fixtureLines, name) ? next.fixtureLines[name] : 0;
    if (now < lines)
      looser.push(`${name}: checks fewer lines (${lines} → ${now}); ${FEWER_LINES_REMEDY}`);
  }
  for (const name of new Set([
    ...Object.keys(old.fixtureClipped ?? {}),
    ...Object.keys(next.fixtureClipped ?? {}),
  ])) {
    const before = lineCounts(clippedOf(old, name));
    const after = lineCounts(clippedOf(next, name));
    for (const [text, n] of after)
      if (n > (before.get(text) ?? 0))
        added.push(
          `${name}: names the clipped line \`${text}\` ${timesText(n)}, the old list ${timesText(before.get(text) ?? 0)}`
        );
    for (const [text, n] of before)
      if (n > (after.get(text) ?? 0))
        removed.push(
          `${name}: names the clipped line \`${text}\` ${timesText(after.get(text) ?? 0)}`
        );
  }
  return { looser, added, removed };
}

/**
 * The list after a run: the entries of the fixtures that the run compared come from the run.
 * The entries of fixtures that did not compare stay, and so do those of fixtures that did not
 * run in a run of some fixtures. A run of every fixture drops the entries of fixtures that no
 * longer exist.
 */
export function mergeRecorded(
  old: KnownGapList | undefined,
  results: readonly FixtureResult[],
  tools: GapTools,
  complete: boolean
): KnownGapList {
  const recorded = recordedList(results, tools);
  if (!old) return recorded;
  const compared = new Set(results.filter((r) => r.status === 'compared').map((r) => r.name));
  const exists = new Set(results.map((r) => r.name));
  const keep = (name: string): boolean => !compared.has(name) && (!complete || exists.has(name));
  const kept = <T>(entries: Readonly<Record<string, T>> | undefined): Record<string, T> =>
    Object.fromEntries(Object.entries(entries ?? {}).filter(([name]) => keep(name)));
  const sorted = <T>(entries: Record<string, T>): Record<string, T> =>
    Object.fromEntries(Object.entries(entries).sort(([a], [b]) => (a < b ? -1 : 1)));
  return {
    ...tools,
    fixtureLines: sorted({ ...kept(old.fixtureLines), ...recorded.fixtureLines }),
    fixtureClipped: sorted({ ...kept(old.fixtureClipped), ...recorded.fixtureClipped }),
    fixtures: sorted({ ...kept(old.fixtures), ...recorded.fixtures }),
  };
}

/**
 * Why `--update-gaps` cannot start a list for `tools`, or undefined. Without a list of the run's
 * tools, the run starts one, so it must measure every fixture: a run of some fixtures is refused
 * before it measures (`run.ts`), and a run with a fixture error is refused after it. The source
 * writes such a run under the old tools instead (see the header).
 */
export function newListRefusal(
  lists: readonly KnownGapList[],
  tools: GapTools,
  complete: boolean,
  results?: readonly FixtureResult[]
): string | undefined {
  if (exactList(lists, tools)) return undefined;
  if (!complete)
    return 'no known-gap list matches the tools of this run, so --update-gaps starts one and needs every fixture: run it without fixture names';
  const failed = (results ?? []).filter((r) => r.status !== 'compared').map((r) => r.name);
  if (failed.length)
    return `no known-gap list matches the tools of this run, so --update-gaps starts one and needs every fixture compared; these did not compare: ${failed.join(', ')}`;
  return undefined;
}

/**
 * Plan `--update-gaps` for a run with `tools` in an environment with `lists`. The list of the
 * run's tools is the old list. Without one, the run starts a list (`newListRefusal`), and it
 * compares with the tightest values of the environment. A worse list is written only with
 * `acceptWorse`.
 */
export function planUpdate(
  lists: readonly KnownGapList[],
  results: readonly FixtureResult[],
  tools: GapTools,
  complete: boolean,
  acceptWorse: boolean
): GapUpdate {
  const refused = newListRefusal(lists, tools, complete, results);
  if (refused) return { refused };
  const exact = exactList(lists, tools);
  const next = mergeRecorded(exact, results, tools, complete);
  const old = exact ?? tightestList(lists) ?? { ...tools, ...EMPTY };
  // A fixture that the run did not measure keeps its entries; one that no longer exists drops.
  const names = new Set(results.map((r) => r.name));
  const changes = listChanges(old, next, (name) => !complete || names.has(name));
  const worse = changes.looser.length > 0 || changes.added.length > 0;
  return { next, changes, write: acceptWorse || !worse };
}

/** The end of a `--update-gaps` run: the exit code, and the line to print, if any. */
export interface UpdateExit {
  readonly code: number;
  readonly line?: string;
}

/** What the run knows of the update: the written list, or why it wrote nothing or failed. */
export interface UpdateResult {
  readonly written?: unknown;
  /** The run starts a list that it may not start (`newListRefusal`). */
  readonly refused?: string;
  /** The write of `thresholds.ts` failed: the first line of its error. */
  readonly failed?: string;
}

/** The merge-base check that a run starts after a write, as `spawnSync` gives it. */
export interface CheckRun {
  readonly status: number | null;
  readonly signal?: NodeJS.Signals | null;
  readonly error?: Error;
}

/**
 * The exit of a `--update-gaps` run whose gate gave `gate` (0 or 1), with the line to print.
 * - Nothing written: 1, with the reason.
 * - The write failed, the check did not start, or it compared nothing (exit code 2): 2, and the
 *   line says that `thresholds.ts` holds an unchecked list. The check prints its own line first,
 *   such as a written `thresholds.ts` that does not load. A failed gate keeps 1.
 * - A signal stopped the check: 128 plus the signal number, as a shell gives it (130 for
 *   SIGINT, 143 for SIGTERM).
 * - The check finds a looser list (exit code 1): 1. Otherwise the code of the gate.
 */
export function updateExit(gate: number, update: UpdateResult, check?: CheckRun): UpdateExit {
  const unchecked = (cause: string): UpdateExit => ({
    code: gate === 1 ? 1 : 2,
    line: `fidelity: ${cause}; thresholds.ts holds the new list unchecked; fix the cause, then run bun run fidelity:check-gaps`,
  });
  if (update.failed)
    return {
      code: gate === 1 ? 1 : 2,
      line: `fidelity: the known gaps could not be written: ${update.failed}; check thresholds.ts with git diff`,
    };
  if (!update.written)
    return { code: 1, line: update.refused ? `fidelity: ${update.refused}` : NOTHING_WRITTEN };
  if (check?.signal) {
    const number = constants.signals[check.signal] ?? 0;
    return {
      code: 128 + number,
      line: `fidelity: ${check.signal} stopped the merge-base check; thresholds.ts holds the new list unchecked`,
    };
  }
  if (!check || check.error || check.status === null)
    return unchecked(
      `the merge-base check did not start: ${check?.error?.message ?? 'no exit code'}`
    );
  if (check.status === 2)
    return unchecked('the merge-base check compared nothing (its line is above)');
  return { code: check.status === 0 ? gate : 1 };
}

/** The line that a run prints when `--update-gaps` writes nothing. */
export const NOTHING_WRITTEN =
  'Nothing was written: the run adds a gap, a check, or a clipped line, or a value is looser (LOOSER, NEW above). Fix the regression. A new fixture or a new toolchain may record them with: bun run fidelity [fixture names...] --update-gaps --accept-worse. The merge-base check then applies its own rules, also in CI: bun run fidelity:check-gaps, or bun run fidelity:check-gaps --base <ref> in a checkout without origin/main.';
