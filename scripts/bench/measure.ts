// The values of the browser bench (`browser-bench.ts`): the statistics of its samples, the
// values of its table, and the leak check of the memory mode. It has no browser code, so
// `measure.test.ts` and `edit-bench.ts` read it.
//
// `p95` and `windowsOf` follow DOCX Editor's browser bench at 1c9b6a29b (EigenPal):
// `summarize` in `e2e/edit-browser-bench-harness.ts`, and the first and last windows of
// `runSustained` in `e2e/edit-browser-bench-probe.ts`, here as values of the table only. That
// window is 10 samples, or a third of the samples when there are fewer than 30. Here it is always
// 10, so the values of a window always hold 10 samples, and a step with fewer than 20 samples
// gets no window row. The default runs give 60 or more samples of each edit step.
// `growthOf` and `leaksOf` follow the verdict of PPTX Editor's `e2e/memory.bench.ts` at d4502552
// (EigenPal).

import type { Resources } from './memory-probe.ts';

export type Mode = 'open' | 'edit' | 'scroll' | 'memory';
export type WorkbookName = 'grid' | 'sheets' | 'strings';

/** The workbooks of the browser bench, in the order that it runs them. */
export const WORKBOOKS: readonly WorkbookName[] = ['grid', 'sheets', 'strings'];

/** The values of one workbook in one mode. */
export interface Row {
  workbook: WorkbookName;
  cells: number;
  /** The SHA-256 of the workbook bytes, so two saved runs show that they used the same bytes. */
  workbookSha256: string;
  samples: Record<string, number[]>;
  /**
   * The JS heap and its array buffers in MiB: after the open, or in the memory mode after the
   * warmup commits.
   */
  heapMiB?: number;
  /** The heap growth for each commit over the second half of the readings, in KiB (`growthOf`). */
  growthKiBPerCommit?: number;
  /** The memory readings after the warmup commits and after each group of commits (memory mode). */
  readings?: Reading[];
  longestTaskMs: number;
  peakMiB: number;
}

export const median = (xs: readonly number[]): number => {
  const s = [...xs].sort((a, b) => a - b);
  return s[Math.floor(s.length / 2)] ?? 0;
};
/** The p95, at the rank of DOCX Editor's `summarize`. */
export const p95 = (xs: readonly number[]): number => {
  const s = [...xs].sort((a, b) => a - b);
  return s[Math.max(0, Math.ceil(s.length * 0.95) - 1)] ?? 0;
};

/** One memory reading after garbage collection, with the resources of the page probe. */
export interface Reading extends Resources {
  /** The JS heap and the array buffers behind it, in bytes. */
  heapBytes: number;
  /** The DOM nodes, detached ones too. */
  nodes: number;
  listeners: number;
  documents: number;
}

/** The samples of each window of the first and last edits (`windowsOf`). */
export const WINDOW = 10;

/**
 * The median of the first and of the last `WINDOW` samples of a step, as DOCX Editor's browser
 * bench compares its first and last ten edits. For the `type` step, 10 samples are the keys of
 * three to four edits. A step that grows slower with each edit, for example
 * with the undo history, shows a last median far above the first. Undefined when there are fewer
 * than two windows of samples.
 */
export function windowsOf(xs: readonly number[]): { first: number; last: number } | undefined {
  if (xs.length < 2 * WINDOW) return undefined;
  return { first: median(xs.slice(0, WINDOW)), last: median(xs.slice(-WINDOW)) };
}

/** The least-squares slope of `ys` for each step. */
function slope(ys: readonly number[]): number {
  const n = ys.length;
  const mx = (n - 1) / 2;
  const my = ys.reduce((a, b) => a + b, 0) / n;
  let num = 0;
  let den = 0;
  for (let i = 0; i < n; i += 1) {
    num += (i - mx) * (ys[i]! - my);
    den += (i - mx) ** 2;
  }
  return den === 0 ? 0 : num / den;
}

/**
 * The growth for each commit over the second half of `readings`, taken every `per` commits. As
 * in the verdict of PPTX Editor's `e2e/memory.bench.ts`, a heap grows when its least-squares
 * slope passes the budget and its last reading is above the first one of the half by at least
 * half of what the budget allows: one high reading in the middle is not growth. The value is the
 * smaller of the slope and twice the mean rise, so it stays within a budget when either one does,
 * and goes over it exactly when both do.
 */
export function growthOf(readings: readonly number[], per: number): number {
  const half = readings.slice(Math.floor(readings.length / 2));
  if (half.length < 2) return 0;
  const commits = (half.length - 1) * per;
  const rise = (half[half.length - 1]! - half[0]!) / commits;
  return Math.min(slope(half) / per, 2 * rise);
}

/**
 * The growth for each commit that counts as a leak, as the budget of a scenario of PPTX Editor's
 * memory verdict:
 * - Heap: 10 KiB. PPTX Editor's typing scenario allows 2 KB of heap for each key, so 10 KiB for
 *   a commit of five keys. The warmup fills the undo history (`MEMORY_WARMUP_COMMITS`), so each
 *   later commit drops the oldest undo step. Then a commit keeps 1 to 4 KiB on each bench
 *   workbook, in local runs of `bench:memory:browser` at `BROWSER_SIZES` and at 250,000 strings
 *   (2 to 3 KiB on the strings workbook at both sizes), where each commit writes a new text and
 *   adds a shared string item. The limit is far below DOCX Editor's browser bench ceiling
 *   (`SUSTAINED_HEAP_CEILING_BYTES` at 1c9b6a29b: 50 MiB over 180 edits, about 284 KiB for each
 *   edit). A cache that keeps the shared string list of each commit on a row element stays
 *   bounded by the undo history, so this verdict cannot see it: the row height test checks
 *   directly that a commit keeps no shared string list.
 * - DOM nodes: 2. PPTX Editor's typing scenario allows 0.2 nodes for each key, so 1 node for a
 *   commit of five keys. `bench:memory:browser` measures 1 DOM node for each commit on each bench
 *   workbook, with a full undo history. The node comes from the typing into the cell editor's
 *   `textarea`, not from the commit: the same bench with each edit canceled by Escape, so that
 *   the store does not change, also grows 1 node for each edit. Chromium keeps the nodes of each
 *   typed edit of a text field in its own edit history, and that history is bounded: a plain
 *   `textarea` that is typed into and cleared 1500 times grows 1 node for each typing and 1 for
 *   each clear up to 1000 of each, and then stays level (local runs). So the limit is that node
 *   and PPTX Editor's 1 node for each commit.
 * - Listeners: 0.05, PPTX Editor's 0.01 for each key, for five keys.
 * As in PPTX Editor's verdict, a growth above the limit is a leak, and exactly the limit passes.
 */
export const LEAK_LIMITS = { heapKiB: 10, nodes: 2, listeners: 0.05 } as const;

/**
 * The untimed commits before the first memory reading: the undo history of the store keeps 200
 * steps (`HISTORY_LIMIT` in `packages/core/src/store/store/document-store.ts`). As PPTX Editor's
 * typing scenario warms up 200 keys because its history keeps 200 entries, the warmup fills the
 * history, so each later commit drops the oldest undo step, and the growth is what a commit keeps
 * for good.
 */
export const MEMORY_WARMUP_COMMITS = 200;

/** The resources that the editor must release: any steady growth of one is a leak. */
export const RESOURCES: readonly (keyof Resources | 'documents')[] = [
  'blobs',
  'fonts',
  'observers',
  'workers',
  'bitmaps',
  'intervals',
  'windowListeners',
  'documentListeners',
  'documents',
];

/**
 * The leaks of `readings`, taken every `per` commits, as PPTX Editor's verdict: a heap, DOM nodes,
 * or listeners that grow past `LEAK_LIMITS` (`growthOf`), and each of `RESOURCES` that grows, from
 * the first reading to the last, by at least 2 and at least one for each four commits, with a
 * slope above zero over the second half of the readings.
 */
export function leaksOf(readings: readonly Reading[], per: number): string[] {
  const leaks: string[] = [];
  const heap = growthOf(
    readings.map((r) => r.heapBytes),
    per
  );
  if (heap > LEAK_LIMITS.heapKiB * 1024)
    leaks.push(
      `heap +${Math.round(heap / 1024)} KiB for each commit (limit ${LEAK_LIMITS.heapKiB} KiB)`
    );
  const nodes = growthOf(
    readings.map((r) => r.nodes),
    per
  );
  if (nodes > LEAK_LIMITS.nodes)
    leaks.push(`DOM nodes +${nodes.toFixed(1)} for each commit (limit ${LEAK_LIMITS.nodes})`);
  const listeners = growthOf(
    readings.map((r) => r.listeners),
    per
  );
  if (listeners > LEAK_LIMITS.listeners)
    leaks.push(
      `listeners +${listeners.toFixed(2)} for each commit (limit ${LEAK_LIMITS.listeners})`
    );
  const base = readings[0];
  const end = readings[readings.length - 1];
  const commits = (readings.length - 1) * per;
  if (base && end) {
    const late = readings.slice(Math.floor(readings.length / 2));
    for (const k of RESOURCES) {
      const grew = end[k] - base[k];
      const steady = slope(late.map((r) => r[k])) > 0;
      if (grew > 0 && grew >= Math.max(2, commits / 4) && steady)
        leaks.push(`${k} ${base[k]} -> ${end[k]}`);
    }
  }
  return leaks;
}

/** The steps of an edit, which give the windows of `windowsOf`. */
const EDIT_STEPS: readonly string[] = ['begin', 'type', 'commit'];

/** One value of the table of the browser bench. */
export interface Value {
  metric: string;
  value: number;
}

/** The values of a row, in the order of the table. */
export function valuesOf(mode: Mode, row: Row): Value[] {
  const out: Value[] = [];
  // The memory mode times its edits only to wait for them: the edit mode measures them.
  const timed = mode === 'memory' ? [] : Object.entries(row.samples);
  for (const [kind, xs] of timed) {
    out.push(
      { metric: `${kind} median ms`, value: median(xs) },
      { metric: `${kind} p95 ms`, value: p95(xs) },
      { metric: `${kind} max ms`, value: Math.max(...xs) }
    );
    const windows = EDIT_STEPS.includes(kind) ? windowsOf(xs) : undefined;
    if (windows)
      out.push(
        { metric: `${kind} first ${WINDOW} median ms`, value: windows.first },
        { metric: `${kind} last ${WINDOW} median ms`, value: windows.last }
      );
  }
  if (row.heapMiB !== undefined) out.push({ metric: 'heap MiB', value: row.heapMiB });
  const last = row.readings?.[row.readings.length - 1];
  if (last && row.readings!.length > 1)
    out.push({ metric: 'heap after edits MiB', value: Math.round(last.heapBytes / 1048576) });
  if (row.growthKiBPerCommit !== undefined)
    out.push({ metric: 'growth KiB per commit', value: row.growthKiBPerCommit });
  out.push({ metric: 'peak MiB', value: row.peakMiB });
  return out;
}
