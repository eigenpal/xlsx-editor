// The watchdog of the browser hang test and of the corpus runner
// (`scripts/corpus-runner/runner.ts`). It runs outside the watched process, so a page or a child
// that does not answer cannot stop it. It reads the process tree of the one browser or child that
// the caller started, adds up its resident memory, and stops only that tree.

/** One line of the process table. */
export interface ProcessRow {
  readonly pid: number;
  readonly ppid: number;
  /** Resident memory, in KiB. */
  readonly rssKb: number;
}

/** The memory that one case may use: more than this is a defect, never a budget to raise. */
export const MEMORY_CAP_BYTES = 3 * 1024 * 1024 * 1024;

/** The rows of `ps -A -o pid=,ppid=,rss=`. A line that does not read is left out. */
export function parseProcessTable(text: string): ProcessRow[] {
  const rows: ProcessRow[] = [];
  for (const line of text.split('\n')) {
    const fields = line.trim().split(/\s+/);
    if (fields.length !== 3) continue;
    const [pid, ppid, rssKb] = fields.map(Number) as [number, number, number];
    if (Number.isInteger(pid) && Number.isInteger(ppid) && Number.isFinite(rssKb))
      rows.push({ pid, ppid, rssKb });
  }
  return rows;
}

/** `root` and all its descendants. Empty when `root` is not in the table. */
export function treeOf(rows: readonly ProcessRow[], root: number): ProcessRow[] {
  const children = new Map<number, ProcessRow[]>();
  for (const row of rows) {
    const list = children.get(row.ppid);
    if (list) list.push(row);
    else children.set(row.ppid, [row]);
  }
  const top = rows.find((row) => row.pid === root);
  if (!top) return [];
  const tree: ProcessRow[] = [];
  const seen = new Set<number>();
  const queue = [top];
  while (queue.length > 0) {
    const row = queue.pop()!;
    if (seen.has(row.pid)) continue;
    seen.add(row.pid);
    tree.push(row);
    queue.push(...(children.get(row.pid) ?? []));
  }
  return tree;
}

/**
 * The resident memory of `tree`, in bytes. Shared pages count once per process, which reads high.
 * Pages that the system compressed or swapped out are not resident, which reads low under memory
 * pressure.
 */
export const treeBytes = (tree: readonly ProcessRow[]): number =>
  tree.reduce((sum, row) => sum + row.rssKb * 1024, 0);

async function processTable(): Promise<ProcessRow[]> {
  const ps = Bun.spawn(['ps', '-A', '-o', 'pid=,ppid=,rss='], { stdout: 'pipe', stderr: 'ignore' });
  const text = await new Response(ps.stdout).text();
  await ps.exited;
  return parseProcessTable(text);
}

/** Stops `root` and its descendants. It signals no other process. */
export async function killTree(root: number): Promise<void> {
  for (const row of treeOf(await processTable(), root).reverse()) {
    try {
      process.kill(row.pid, 'SIGKILL');
    } catch {
      // The process ended between the read of the table and the signal.
    }
  }
}

/** Stops the process group that `root` leads, if it still exists. It signals no other process. */
export function signalGroup(root: number): void {
  try {
    process.kill(-root, 'SIGKILL');
  } catch {
    // The group ended already.
  }
}

/**
 * Stops the process group of `root` and then its tree. `root` must lead its own group (a child
 * spawned with `detached: true`), so a descendant that starts while the table is read stops too.
 */
export async function killGroup(root: number): Promise<void> {
  signalGroup(root);
  await killTree(root);
}

/** A running watch over one process tree. */
export interface Watch {
  /** The highest memory of the tree so far, in bytes. */
  readonly peakBytes: () => number;
  /** Why the watch stopped the tree, or null while it did not. */
  readonly tripped: () => string | null;
  /** Whether the watch stopped the tree for its memory. */
  readonly trippedMemory: () => boolean;
  /** Starts the time limit again from now, for a process that starts its next unit of work. */
  readonly restart: () => void;
  readonly stop: () => void;
}

/** The limits of {@link watchProcess}. */
export interface WatchLimits {
  /** The time after which the watch stops the tree, from the start or the last restart. */
  readonly deadlineMs: number;
  /** The time between two reads of the process table. */
  readonly everyMs?: number;
  /** The resident memory of the tree above which the watch stops it. */
  readonly capBytes?: number;
  /** What the reason calls the tree, for example `the browser`. */
  readonly name?: string;
  /** What the reason calls the timed unit of work, for example `the case`. */
  readonly unit?: string;
  /** Stop the process group of `root` too ({@link killGroup}). `root` must lead its group. */
  readonly group?: boolean;
  /** The clock in ms. Tests give their own; the default is `Date.now`. */
  readonly now?: () => number;
  /** The read of the process table. Tests give their own; the default runs `ps`. */
  readonly readTable?: () => Promise<ProcessRow[]>;
}

/**
 * A read that comes this much later than planned means that the machine slept or the watch was
 * paused. That time does not count against the time limit: the work did not run either.
 */
export const LATE_READ_MS = 5000;

/**
 * Reads the tree of `root` every `everyMs`. When the tree passes the memory cap, or the time
 * since the start or the last restart passes `deadlineMs`, it stops the tree and records why.
 * A read more than {@link LATE_READ_MS} late does not count its delay as time of the work.
 */
export function watchProcess(root: number, limits: WatchLimits): Watch {
  const {
    deadlineMs,
    everyMs = 500,
    capBytes = MEMORY_CAP_BYTES,
    name = 'the process',
    unit = 'the work',
    group = false,
    now = Date.now,
    readTable = processTable,
  } = limits;
  let started = now();
  // When the next read was planned.
  let planned = started;
  let peak = 0;
  let why: string | null = null;
  let memory = false;
  let stopped = false;
  const trip = async (reason: string, forMemory: boolean) => {
    // One stop only: a later signal could reach a process that took over the id.
    if (why !== null) return;
    why = reason;
    memory = forMemory;
    await (group ? killGroup(root) : killTree(root));
  };
  const tick = async () => {
    if (stopped) return;
    // The time since the planned read, or since a restart after it, did not run the work.
    if (now() - planned > LATE_READ_MS) started += Math.max(0, now() - Math.max(planned, started));
    try {
      const table = await readTable();
      // A stop during the read: the process can be gone, and its id can be another process now.
      if (stopped) return;
      const bytes = treeBytes(treeOf(table, root));
      peak = Math.max(peak, bytes);
      if (bytes > capBytes) await trip(`${name} used ${Math.round(bytes / 1048576)} MiB`, true);
      else if (now() - started > deadlineMs) await trip(`${unit} passed ${deadlineMs} ms`, false);
    } catch {
      // A failed read of the process table must not end the watch: the next tick reads again.
    }
    if (stopped) return;
    planned = now() + everyMs;
    timer = setTimeout(() => void tick(), everyMs);
  };
  // The first read is at once, so a short case also has a memory reading.
  let timer = setTimeout(() => void tick(), 0);
  return {
    peakBytes: () => peak,
    tripped: () => why,
    trippedMemory: () => memory,
    restart: () => {
      started = now();
    },
    stop: () => {
      stopped = true;
      clearTimeout(timer);
    },
  };
}

/** {@link watchProcess} over the browser of one hang case. */
export function watchBrowser(
  root: number,
  deadlineMs: number,
  everyMs = 500,
  capBytes = MEMORY_CAP_BYTES
): Watch {
  return watchProcess(root, {
    deadlineMs,
    everyMs,
    capBytes,
    name: 'the browser',
    unit: 'the case',
  });
}
