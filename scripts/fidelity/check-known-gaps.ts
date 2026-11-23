// `bun run fidelity:check-gaps`: thresholds and known gaps only tighten (design D8). Compares
// `thresholds.ts` with its copy at the merge base and fails when a threshold or a tolerance
// becomes looser, when a new threshold or tolerance starts looser than its tightest value, when
// a recorded gap value grows by any amount (`GAP_TOLERANCE` applies only to a run against its
// record, never here), when a fixture of the merge base gets a gap of a new check, when a
// fixture that has a gap is renamed or removed, or when a list names a clipped line
// (`fixtureClipped`) more often than the merge base. A gap of a fixture that the merge base
// does not have is new, and allowed, unless a fixture of the merge base is gone in the change
// (`newFixtures`): then a renamed fixture cannot take a gap as new. Each threshold names the
// direction in which it becomes looser (`THRESHOLD_DIRECTIONS`), and a threshold without one
// fails; the merge base's entry of a threshold may not change. A threshold may go away only
// when `RETIRED_LIMITS` names it with the reason, as in the source, and the check prints a
// notice for each threshold that it retires. A new threshold declares its own tightest value
// (`THRESHOLD_DIRECTIONS`), which review checks: no measurement backs it. A new fixture with
// clipped lines, or a fixture of the merge base whose reference PDF changed
// (`rereferencedFixtures`) with new clipped lines, needs an owner's acceptance:
// `XLSX_FIDELITY_ACCEPT_CLIPPED=<fixture>[,<fixture>...]` for that run; CI never sets it. The
// acceptance of a new toolchain never accepts a new clipped line. A per-fixture pixel limit
// (`FIXTURE_PIXEL_MISMATCH`) may only tighten, may not be looser than the page limit, and stays
// while its fixture exists; removing or renaming a fixture with a limit fails.
// Removing every list of a font environment is not looser: a run in an environment with no list
// fails.
//
// Lists are compared toolchain by toolchain: each list of the merge base with the list of the
// same Poppler and Chromium. A removed list is compared with the tightest values of the lists
// that remain (`tightestList`), which a run with its tools then uses outside CI; in CI such a run
// fails. A list for a new toolchain is compared with the tightest recorded value of each fixture
// and check across all lists of its environment at the merge base, so a new toolchain cannot
// record a looser value. An owner who accepts a looser new toolchain sets
// `XLSX_FIDELITY_ACCEPT_NEW_TOOLCHAIN=<poppler>; <chromium>` (`toolchainText`) for that run. CI
// never sets it.
//
// Same role as pptx-editor/scripts/fidelity/check-known-gaps.ts at 1943d2c9 (EigenPal).
// Differences, in the order of design D8:
// 1. A new gap or a new check of a fixture of the merge base always fails here, also in a change
//    that touches no engine or harness source (the source allows it then, `nonSourceFiles`).
// 2. A new threshold and a new tolerance are compared with their tightest values
//    (`THRESHOLD_DIRECTIONS`, the smallest tolerance), where the source compares only the limits
//    of the merge base.
// 3. Each threshold names its loosening direction in `THRESHOLD_DIRECTIONS`, read from both
//    sides, where the source names the two that are looser when smaller in code. The fixed
//    matching tolerances of the line checks have the direction `either`: any change fails, and a
//    new key replaces a retired one. The source has no such tolerance as a threshold.
// 4. Known gaps are kept for each raster toolchain.
// 5. A list names its clipped lines, which only become fewer. The first record is not compared
//    for the fixtures of the merge base: a merge base where no list has `fixtureClipped` has
//    nothing to compare with. When any list of the merge base has the record, a list without it
//    counts as an empty record, and a list may not drop the record that its merge base list has.
//    A record of a name that is not a fixture fails. In a change that removes a fixture of the
//    merge base, a fixture that the merge base does not have can be a renamed one: its clipped
//    lines fail (`CLIPPED_RENAMED_REMEDY`), also before the first record. Such a change also
//    fails when a removed fixture had checked lines (`RENAMED_LINES_REMEDY`). A change that only
//    removes a fixture without a known gap passes, and a later change that adds the name again
//    adds a new fixture: an accepted limit, because the removal shows in review. A fixture with a
//    known gap cannot be removed until its gap is fixed. A fixture checks fewer lines only when
//    it is removed and added again (`FEWER_LINES_REMEDY`), also with a new reference PDF.
// 6. The head fixtures of `newFixtures` are the fixtures of the catalog and of the reference
//    manifest (`main`), where the source reads the manifest only: a fixture that the catalog
//    builds counts before its reference PDF lands.
// 7. `--base` is optional, and the merge base of `HEAD` and `origin/main` is its default, where
//    the source needs it. A `--base` that names no commit fails with exit code 2 (`isCommit`),
//    as in the source, so exit code 0 with no thresholds means a real commit without them.
// 8. Each error prints one `fidelity:` line. An error that stops the check, such as a file that
//    cannot be read, exits with code 2, so exit code 1 always means a looser list. The command
//    runs through `cli.ts`, which loads this module with a dynamic import: a head
//    `thresholds.ts` that does not load stops it in one line with code 2 too.
// 9. `--require-thresholds` makes a base without `thresholds.ts` exit with code 2: the check
//    compared nothing. `bun run fidelity --update-gaps` passes it, so a new list is never left
//    unchecked. Without it, such a base passes with exit code 0, as in the source. Only a base
//    that has no such file counts as a base without thresholds: any other read error exits with
//    code 2.
//
//   bun run fidelity:check-gaps [--base <ref>] [--require-thresholds]
//     (default base: the merge base with origin/main)

import { spawnSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import * as head from './thresholds.ts';
import { FIXTURES } from '../fixtures/catalog.ts';
import { firstLine } from './chromium.ts';
import { readManifest, type Manifest } from './references.ts';
import { CHECK_GAPS_USAGE, checkGapsArgs } from './run-args.ts';
import {
  clippedOf,
  CLIPPED_BASE_REMEDY,
  CLIPPED_NEW_REMEDY,
  CLIPPED_REFERENCE_REMEDY,
  CLIPPED_RENAMED_REMEDY,
  FEWER_LINES_REMEDY,
  RENAMED_LINES_REMEDY,
  lineCounts,
  looserRecord,
  timesText,
  sameTools,
  tightestList,
  toolchainText,
} from './known-gaps.ts';
import type { FailureKind, FontEnvironment, KnownGapList } from './thresholds.ts';

type Thresholds = typeof head;

/**
 * The known gaps of one side of the comparison. A merge base from before the toolchain lists
 * holds one list for each environment instead of an array (`listsOf`).
 */
type GapLists = Readonly<Partial<Record<FontEnvironment, KnownGapList | readonly KnownGapList[]>>>;

/** The thresholds, tolerances, and known gaps of one side of the comparison. */
export interface Ratcheted {
  readonly THRESHOLDS: Thresholds['THRESHOLDS'];
  readonly GAP_TOLERANCE: Thresholds['GAP_TOLERANCE'];
  readonly KNOWN_GAPS: GapLists;
  /** Absent at a merge base from before the record; the head's then apply to both sides. */
  readonly THRESHOLD_DIRECTIONS?: Thresholds['THRESHOLD_DIRECTIONS'];
  /** Absent at a merge base from before the record. */
  readonly RETIRED_LIMITS?: Thresholds['RETIRED_LIMITS'];
  /** Absent at a merge base from before the per-fixture limits. */
  readonly FIXTURE_PIXEL_MISMATCH?: Readonly<Record<string, number>>;
}

/** The lists of one environment, as an array. */
export function listsOf(
  value: KnownGapList | readonly KnownGapList[] | undefined
): readonly KnownGapList[] {
  if (value === undefined) return [];
  return Array.isArray(value) ? value : [value as KnownGapList];
}

/**
 * The fixtures that the change adds: those of the head that the merge base does not have. When a
 * fixture of the merge base is gone in the change, none is new, because a renamed fixture would
 * otherwise take any gap as new. Ported from pptx-editor/scripts/fidelity/check-known-gaps.ts
 * at 1943d2c9 (`newFixtures`).
 */
export function newFixtures(
  baseFixtures: ReadonlySet<string>,
  headFixtures: ReadonlySet<string>
): Set<string> {
  if ([...baseFixtures].some((n) => !headFixtures.has(n))) return new Set();
  return new Set([...headFixtures].filter((n) => !baseFixtures.has(n)));
}

/**
 * The fixtures of the merge base whose reference PDF changed in the change: the manifests of
 * both sides (`fixtures/references.json`) name a different PDF hash.
 */
export function rereferencedFixtures(base: Manifest, now: Manifest): Set<string> {
  return new Set(
    Object.keys(base.references).filter(
      (name) =>
        Object.hasOwn(now.references, name) &&
        now.references[name]!.pdf !== base.references[name]!.pdf
    )
  );
}

/**
 * One way in which a list is looser. `clipped` problems name a clipped line: only
 * `ACCEPT_CLIPPED` accepts one, never the acceptance of a new toolchain. `accepted` is a clipped
 * line that `ACCEPT_CLIPPED` accepts.
 */
interface Looser {
  readonly kind: 'gap' | 'clipped' | 'accepted';
  readonly text: string;
}

/** The remedy for a gap that a change records looser than the merge base. */
export const GAP_REMEDY =
  'fix the regression instead of recording it, or measure the gap on a new fixture';

/**
 * The ways in which the list `now` is looser than the list `was`, prefixed with `label`. A
 * fixture in `fresh` (`newFixtures`) may take any gap.
 */
function listLooser(
  label: string,
  was: KnownGapList,
  now: KnownGapList,
  fresh: ReadonlySet<string>,
  headFixtures: ReadonlySet<string>,
  scope: {
    readonly acceptClipped?: string;
    readonly rereferenced?: ReadonlySet<string>;
    /** Whether any list of the merge base has a clipped-line record (`fixtureClipped`). */
    readonly clippedRecorded?: boolean;
    /**
     * The head fixtures that the merge base manifest does not name and that no list of the merge
     * base clips. When a fixture of the merge base is gone, such a fixture can be a renamed one
     * (`newFixtures`): its clipped lines compare with an empty record, also on a merge base with
     * no clipped-line record, and no acceptance applies.
     */
    readonly clippedNew?: ReadonlySet<string>;
    /** The change removes a fixture of the merge base and adds one: it can rename a fixture. */
    readonly renaming?: boolean;
  }
): Looser[] {
  const acceptClipped = new Set(
    (scope.acceptClipped ?? '')
      .split(',')
      .map((n) => n.trim())
      .filter(Boolean)
  );
  const out: Looser[] = [];
  const looser = (text: string): void => void out.push({ kind: 'gap', text });
  const clipped = (text: string): void => void out.push({ kind: 'clipped', text });
  for (const name of Object.keys(was.fixtures))
    if (!headFixtures.has(name))
      looser(
        `${label}: ${name} has a known gap and was removed or renamed; keep the fixture, or fix its gap first`
      );
  for (const [name, lines] of Object.entries(was.fixtureLines))
    if (!headFixtures.has(name) && scope.renaming)
      looser(
        `${label}: ${name} checks ${lines} lines and was removed or renamed in a change that adds a fixture; ${RENAMED_LINES_REMEDY}`
      );
    else if (headFixtures.has(name) && (now.fixtureLines[name] ?? 0) < lines)
      looser(
        `${label}: ${name} checks ${now.fixtureLines[name] ?? 0} lines, fewer than ${lines}; ${FEWER_LINES_REMEDY}`
      );
  // Clipped lines (`fixtureClipped`) only become fewer: a list may name each line at most as
  // often as the merge base names it. Only a merge base where no list has the record has nothing
  // to compare with: its first record is not compared. When another list of the merge base has
  // the record, a list without it counts as an empty record. A list may not drop the record that
  // its merge base list has, so that a later change cannot add clipped lines as a first record.
  // A new fixture, or a fixture with a new reference PDF, may name more clipped lines only with
  // an owner's acceptance (`ACCEPT_CLIPPED`).
  if (was.fixtureClipped && !now.fixtureClipped)
    clipped(
      `${label}: the list drops its clipped-line record (\`fixtureClipped\`), which the merge base has; keep the record, empty when no line is clipped`
    );
  const compareClipped = was.fixtureClipped !== undefined || scope.clippedRecorded === true;
  for (const name of Object.keys(now.fixtureClipped ?? {})) {
    // A record of a name that is no fixture could take clipped lines for a later rename.
    if (!headFixtures.has(name)) {
      clipped(
        `${label}: ${name} has a clipped-line record but is not a fixture; remove the record`
      );
      continue;
    }
    const before = lineCounts(clippedOf(was, name));
    const accepted = acceptClipped.has(name);
    for (const [text, n] of lineCounts(clippedOf(now, name))) {
      if (fresh.has(name)) {
        const problem = `${label}: ${name} is a new fixture with the clipped line \`${text}\``;
        if (accepted) out.push({ kind: 'accepted', text: problem });
        else clipped(`${problem}; ${CLIPPED_NEW_REMEDY}`);
      } else if ((compareClipped || scope.clippedNew?.has(name)) && n > (before.get(text) ?? 0)) {
        const problem = `${label}: ${name} names the clipped line \`${text}\` ${timesText(n)}, the merge base ${timesText(before.get(text) ?? 0)}`;
        if (scope.clippedNew?.has(name)) clipped(`${problem}; ${CLIPPED_RENAMED_REMEDY}`);
        else if (!scope.rereferenced?.has(name)) clipped(`${problem}; ${CLIPPED_BASE_REMEDY}`);
        else if (accepted) out.push({ kind: 'accepted', text: problem });
        else clipped(`${problem}; ${CLIPPED_REFERENCE_REMEDY}`);
      }
    }
  }
  for (const [name, gap] of Object.entries(now.fixtures)) {
    if (fresh.has(name)) continue;
    const old = Object.hasOwn(was.fixtures, name) ? was.fixtures[name] : {};
    for (const [kind, value] of Object.entries(gap) as [FailureKind, number][]) {
      const recorded = old[kind];
      if (recorded === undefined)
        looser(`${label}: ${name} has a new gap ${kind} (${value}); ${GAP_REMEDY}`);
      else if (looserRecord(kind, value, recorded))
        looser(`${label}: ${name} ${kind}: ${recorded} → ${value}; ${GAP_REMEDY}`);
    }
  }
  return out;
}

/**
 * The environment variable with which an owner accepts the clipped lines of a new fixture: its
 * value is the fixture name.
 */
export const ACCEPT_CLIPPED = 'XLSX_FIDELITY_ACCEPT_CLIPPED';

/**
 * The remedy for a change to a threshold that is looser in either direction (`looser: 'either'`):
 * a smaller value can fail a correct render, and a larger one can pass a wrong render. The path
 * after review is a new threshold, and the old one retires with its reason.
 */
export const EITHER_REMEDY = (key: string): string =>
  `a smaller value can fail a correct render, and a larger one can pass a wrong render, so the check refuses a change in either direction; after review, retire ${key} in \`RETIRED_LIMITS\` in thresholds.ts with the reason, and add a new threshold at its measured tightest value with \`looser: 'either'\``;

/** The environment variable with which an owner accepts a looser list of a new toolchain. */
export const ACCEPT_NEW_TOOLCHAIN = 'XLSX_FIDELITY_ACCEPT_NEW_TOOLCHAIN';

/**
 * The ways in which `next` is looser than `base`. `accept` is the toolchain text of a new list
 * that an owner accepts although it is looser; each accepted problem goes to `notes`.
 * `acceptClipped` is the value of `ACCEPT_CLIPPED`, and `rereferenced` holds the fixtures of the
 * merge base with a new reference PDF (`rereferencedFixtures`).
 */
export function looserThan(
  base: Ratcheted,
  next: Ratcheted,
  baseFixtures: ReadonlySet<string>,
  headFixtures: ReadonlySet<string>,
  accept?: string,
  notes: string[] = [],
  acceptClipped?: string,
  rereferenced: ReadonlySet<string> = new Set()
): string[] {
  const out: string[] = [];
  const clippedRecorded = Object.values(base.KNOWN_GAPS).some((lists) =>
    listsOf(lists).some((l) => l.fixtureClipped !== undefined)
  );
  // A renamed fixture is not new (`newFixtures`). It has no reference and no clipped line at the
  // merge base, so its clipped lines compare with none, also before the first record.
  const baseClipped = new Set(
    Object.values(base.KNOWN_GAPS).flatMap((lists) =>
      listsOf(lists).flatMap((l) => Object.keys(l.fixtureClipped ?? {}))
    )
  );
  const clippedNew = new Set(
    [...headFixtures].filter((n) => !baseFixtures.has(n) && !baseClipped.has(n))
  );
  const renaming =
    [...baseFixtures].some((n) => !headFixtures.has(n)) &&
    [...headFixtures].some((n) => !baseFixtures.has(n));
  const scope = {
    ...(acceptClipped ? { acceptClipped } : {}),
    rereferenced,
    clippedRecorded,
    clippedNew,
    renaming,
  };
  const take = (found: readonly Looser[]): void => {
    for (const f of found)
      if (f.kind === 'accepted') notes.push(`accepted by ${ACCEPT_CLIPPED}: ${f.text}`);
      else out.push(f.text);
  };
  const fresh = newFixtures(baseFixtures, headFixtures);
  const b = base.THRESHOLDS as Readonly<Record<string, number>>;
  const n = next.THRESHOLDS as Readonly<Record<string, number>>;
  // The directions of the head, and of the merge base where it has them: a change cannot flip
  // the direction or the tightest value of a threshold to make a looser value pass.
  const nowDirections = next.THRESHOLD_DIRECTIONS ?? head.THRESHOLD_DIRECTIONS;
  const wasDirections = base.THRESHOLD_DIRECTIONS ?? nowDirections;
  const own = <T>(r: Readonly<Record<string, T>>, key: string): T | undefined =>
    Object.hasOwn(r, key) ? r[key] : undefined;
  const retired = next.RETIRED_LIMITS ?? {};
  for (const key of Object.keys(b))
    if (!(key in n) && !own(retired, key)?.trim())
      out.push(
        `threshold ${key} was removed (it was ${b[key]}); keep it, or list it in \`RETIRED_LIMITS\` in thresholds.ts with the reason`
      );
  for (const key of Object.keys(n)) {
    const direction = own(nowDirections, key);
    if (!direction) {
      out.push(
        `threshold ${key} has no loosening direction; add it to \`THRESHOLD_DIRECTIONS\` in thresholds.ts`
      );
      continue;
    }
    const was = own(wasDirections, key);
    if (was && (was.looser !== direction.looser || was.tightest !== direction.tightest)) {
      out.push(
        `threshold ${key} changed its direction or tightest value in \`THRESHOLD_DIRECTIONS\`; keep the entry of the merge base`
      );
      continue;
    }
    const looser = (value: number, than: number): boolean =>
      direction.looser === 'either'
        ? value !== than
        : direction.looser === 'larger'
          ? value > than
          : value < than;
    if (key in b) {
      if (direction.looser === 'either' && looser(n[key]!, b[key]!))
        out.push(`threshold ${key}: ${b[key]} → ${n[key]}; ${EITHER_REMEDY(key)}`);
      else if (looser(n[key]!, b[key]!))
        out.push(
          `threshold ${key}: ${b[key]} → ${n[key]}; thresholds only tighten, so fix the regression instead`
        );
      continue;
    }
    // A threshold that the merge base does not have starts at its tightest value.
    const start = direction.tightest ?? (direction.looser === 'larger' ? 0 : undefined);
    if (start === undefined)
      out.push(
        `threshold ${key} is new and has no tightest value; give it \`tightest\` in \`THRESHOLD_DIRECTIONS\``
      );
    else if (looser(n[key]!, start))
      out.push(
        direction.looser === 'either'
          ? `threshold ${key} is new at ${n[key]}, not its tightest value ${start}; start it at ${start}, or set its \`tightest\` in \`THRESHOLD_DIRECTIONS\` to the measured value`
          : `threshold ${key} is new at ${n[key]}, looser than its tightest value ${start}; start it at ${start}`
      );
  }
  // A check that the merge base does not have takes at most the smallest tolerance there.
  const smallest = Math.min(...Object.values(base.GAP_TOLERANCE));
  for (const [kind, value] of Object.entries(next.GAP_TOLERANCE) as [FailureKind, number][])
    if (!Object.hasOwn(base.GAP_TOLERANCE, kind) && value > smallest)
      out.push(
        `gap tolerance ${kind} is new at ${value}, more than the smallest tolerance ${smallest}; start it at ${smallest} or less`
      );
  // A per-fixture pixel limit only tightens, stays while its fixture exists, and is never looser
  // than the page limit. A fixture with no entry takes the page limit.
  const baseLimits = base.FIXTURE_PIXEL_MISMATCH ?? {};
  const nextLimits = next.FIXTURE_PIXEL_MISMATCH ?? {};
  const limitOf = (limits: Readonly<Record<string, number>>, name: string, page: number) =>
    Object.hasOwn(limits, name) ? limits[name]! : page;
  for (const name of new Set([...Object.keys(baseLimits), ...Object.keys(nextLimits)])) {
    if (!headFixtures.has(name)) {
      if (Object.hasOwn(baseLimits, name) && baseFixtures.has(name))
        out.push(`pixel limit of ${name}: the fixture was removed or renamed`);
      continue;
    }
    const was = limitOf(baseLimits, name, b.maxPixelMismatch!);
    const now = limitOf(nextLimits, name, n.maxPixelMismatch!);
    if (now > was) out.push(`pixel limit of ${name}: ${was} → ${now}`);
    else if (now > n.maxPixelMismatch!)
      out.push(`pixel limit of ${name}: ${now} is looser than the page limit`);
  }
  // A fixture with its own pixel limit records no pixels gap: the gap tolerance is larger than
  // such a limit, so a recorded gap would let its score grow past the limit.
  for (const lists of Object.values(next.KNOWN_GAPS))
    for (const list of listsOf(lists))
      for (const name of Object.keys(nextLimits))
        if (list.fixtures[name]?.pixels !== undefined)
          out.push(`pixel limit of ${name}: a pixels gap is recorded in ${list.poppler}`);
  for (const [kind, value] of Object.entries(base.GAP_TOLERANCE) as [FailureKind, number][])
    if ((next.GAP_TOLERANCE[kind] ?? Infinity) > value)
      out.push(
        `gap tolerance ${kind}: ${value} → ${next.GAP_TOLERANCE[kind]}; tolerances only tighten, so keep ${value} or less`
      );
  for (const env of Object.keys(base.KNOWN_GAPS) as FontEnvironment[]) {
    const wasLists = listsOf(base.KNOWN_GAPS[env]);
    const nowLists = listsOf(next.KNOWN_GAPS[env]);
    // No list left: a run in the environment has no measurement and fails.
    if (nowLists.length === 0) continue;
    for (const was of wasLists) {
      const label = `${env} [${was.poppler}; ${was.chromium}]`;
      const exact = nowLists.find((l) => sameTools(l, was));
      take(
        listLooser(
          exact ? label : `${label} (removed; runs use ${tightestList(nowLists)!.poppler})`,
          was,
          exact ?? tightestList(nowLists)!,
          fresh,
          headFixtures,
          scope
        )
      );
    }
    // A list of a new toolchain: against the tightest recorded values of the merge base.
    const tightest = tightestList(wasLists);
    if (!tightest) continue;
    for (const now of nowLists) {
      if (wasLists.some((w) => sameTools(w, now))) continue;
      const found = listLooser(
        `${env} [${now.poppler}; ${now.chromium}] (new toolchain, against ${tightest.poppler})`,
        tightest,
        now,
        fresh,
        headFixtures,
        scope
      );
      // The acceptance of a new toolchain accepts no clipped line: only `ACCEPT_CLIPPED` does.
      const gaps = found.filter((f) => f.kind === 'gap');
      take(found.filter((f) => f.kind !== 'gap'));
      if (gaps.length && accept === toolchainText(now))
        notes.push(...gaps.map((f) => `accepted by ${ACCEPT_NEW_TOOLCHAIN}: ${f.text}`));
      else out.push(...gaps.map((f) => f.text));
    }
  }
  return out;
}

/**
 * The notice for each threshold of `base` that `next` retires (`RETIRED_LIMITS`), as the
 * source prints it: `NOTICE: <key> is retired: <reason>`.
 */
export function retiredNotices(base: Ratcheted, next: Ratcheted): string[] {
  const retired = next.RETIRED_LIMITS ?? {};
  return Object.keys(base.THRESHOLDS)
    .filter((key) => !(key in next.THRESHOLDS) && Object.hasOwn(retired, key))
    .filter((key) => retired[key]?.trim())
    .map((key) => `NOTICE: ${key} is retired: ${retired[key]}`);
}

interface GitResult {
  readonly ok: boolean;
  readonly out: string;
  readonly err: string;
  /** Why git did not run, such as a missing executable. */
  readonly failed?: string;
}

function git(args: readonly string[]): GitResult {
  // `LC_ALL=C` keeps the messages of git in English, so that `isMissingPath` can read them.
  const r = spawnSync('git', args, { encoding: 'utf8', env: { ...process.env, LC_ALL: 'C' } });
  if (r.error) return { ok: false, out: '', err: '', failed: firstLine(r.error) };
  return { ok: r.status === 0, out: r.stdout ?? '', err: r.stderr ?? '' };
}

/**
 * Whether a failed `git show <commit>:<path>` failed because the commit has no file at the path.
 * Any other failure, such as an object that cannot be read, is an error that stops the check.
 */
export function isMissingPath(stderr: string): boolean {
  return /\bdoes not exist in '|\bexists on disk, but not in '/.test(stderr);
}

/** The first line of a git error, without its `fatal: ` prefix. */
function gitError(stderr: string): string {
  return firstLine(stderr.trim()).replace(/^fatal: /, '') || 'git failed';
}

/** Whether `ref` names a commit of this repository (`git rev-parse --verify`), as in the source. */
export function isCommit(ref: string): boolean {
  return git(['rev-parse', '--verify', '--quiet', `${ref}^{commit}`]).ok;
}

async function main(): Promise<number> {
  const parsed = checkGapsArgs(process.argv.slice(2));
  if ('help' in parsed) {
    console.log(CHECK_GAPS_USAGE);
    return 0;
  }
  if ('error' in parsed) {
    console.error(`fidelity: ${parsed.error}`);
    return 2;
  }
  // Without git, no other answer of git is true: a ref is not "not a commit".
  const version = git(['--version']);
  if (version.failed) {
    console.error(`fidelity: git cannot run: ${version.failed}`);
    return 2;
  }
  const ref = parsed.base;
  if (ref !== undefined && !isCommit(ref)) {
    console.error(`fidelity: ${ref} is not a commit: fetch it, or name a commit`);
    return 2;
  }
  const base = ref ?? git(['merge-base', 'HEAD', 'origin/main']).out.trim();
  if (!base) {
    console.error('fidelity: no merge base: pass --base <ref>');
    return 2;
  }
  const old = git(['show', `${base}:scripts/fidelity/thresholds.ts`]);
  if (!old.ok && !isMissingPath(old.err)) {
    console.error(`fidelity: thresholds.ts at ${base} cannot be read: ${gitError(old.err)}`);
    return 2;
  }
  if (!old.ok) {
    // `--update-gaps` needs a comparison: there, a base without thresholds compared nothing.
    if (parsed.requireThresholds) {
      console.error(`fidelity: no thresholds at ${base}: the check compared nothing`);
      return 2;
    }
    console.log(`No thresholds at ${base}: nothing to compare.`);
    return 0;
  }
  const dir = mkdtempSync(join(tmpdir(), 'xlsx-gaps-'));
  try {
    const file = join(dir, 'thresholds.ts');
    writeFileSync(file, old.out);
    const was = (await import(file)) as Ratcheted;
    const manifest = git(['show', `${base}:fixtures/references.json`]);
    if (!manifest.ok && !isMissingPath(manifest.err)) {
      console.error(
        `fidelity: fixtures/references.json at ${base} cannot be read: ${gitError(manifest.err)}`
      );
      return 2;
    }
    const baseManifest: Manifest = manifest.ok
      ? JSON.parse(manifest.out)
      : { version: 1, references: {} };
    const notes: string[] = [];
    const problems = looserThan(
      was,
      head,
      new Set(Object.keys(baseManifest.references)),
      new Set([...FIXTURES.map((f) => f.name), ...Object.keys(readManifest().references)]),
      process.env[ACCEPT_NEW_TOOLCHAIN] || undefined,
      notes,
      process.env[ACCEPT_CLIPPED] || undefined,
      rereferencedFixtures(baseManifest, readManifest())
    );
    for (const p of retiredNotices(was, head)) console.log(p);
    for (const p of notes) console.log(`note: ${p}`);
    for (const p of problems) console.error(`looser: ${p}`);
    if (problems.length === 0)
      console.log(`Thresholds and known gaps are not looser than ${base}.`);
    return problems.length ? 1 : 0;
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

/**
 * `bun run fidelity:check-gaps` (`cli.ts check-gaps`). An error that stops the check, such as a
 * file that cannot be read or parsed, prints one line and exits with code 2: the check compared
 * nothing, so it never reports "looser" (exit code 1).
 */
export function command(): Promise<number> {
  return main().catch((error: unknown) => {
    console.error(`fidelity: the check stopped: ${firstLine(error)}`);
    return 2;
  });
}

if (import.meta.main) process.exit(await command());
