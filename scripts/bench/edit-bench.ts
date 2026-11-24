// Editing benchmark: opening a workbook, one keystroke, one committed number, one committed text,
// one scroll step, and one sheet switch on the generated workbooks of `workbooks.ts`, run headless
// through the editor.
//
// The editor runs in happy-dom with the production text measurer and the packaged substitute
// font, and renders after each step as the browser does on its next frame (`flushEditor`): the
// layout and the DOM of the grid. happy-dom has no canvas, so no time holds the canvas paint, as
// DOCX Editor's bench at 1c9b6a29b leaves out the browser paint. The browser benches
// (`browser-bench.ts`) time the paint.
//
// Timings depend on the machine; the counters do not. Read the counters first:
// - `Rows measured`: the rows whose height layout measured. A keystroke measures none, and a
//   committed value measures the row that it changed and each row with a wrapped cell. The
//   bench workbooks have no wrapped cell.
// - `Rows cached`: the rows whose height layout took from its cache. A keystroke, a scroll step,
//   and a switch to a sheet seen before take none: layout reuses the row geometry. A committed
//   value walks the rows of its sheet again, from the cache.
// - `Sheet indexes`: the builds of the cell index of a worksheet, and `Cells indexed` the cells
//   that they read. A keystroke builds none. A committed value builds the index of its sheet
//   again.
// The counters show only the row heights and the sheet indexes. A commit also does other work in
// proportion to its sheet, which no counter shows: a CPU profile of commits on 20,000 rows shows
// the edit state of the sheet (`sheetState`, with the scans of `clearRangeBlocker`:
// `blockedCells`, `masterCells`, and the array formula scan `arrayRangesOf`), the formula scan of
// the workbook, `sheetFacts`, and the check of the touched parts (`touchedPartViolation`). Each
// new sheet index also sorts the cell keys of the sheet again (`sortedKeys` in `sheet-layout.ts`)
// and reads the geometry of each row for the hidden lines (`hiddenLines` in `sheet-facts.ts`).
// A commit of a new text also copies the text index and the item list of the shared string table.
// The open also reads each worksheet before the first sheet paints, and each commit runs the
// structure check again over the other worksheets.
// The run fails when a keystroke, a scroll step, or a sheet switch reads a row height or builds a
// sheet index, or when a committed value measures another number of rows than one: more is work
// in proportion to the sheet, which the grid must not do (AGENTS.md, "Layout"), and none is a
// stale height.
// `scripts/bench/edit-bench-gates.test.ts` pins the same counters inside `bun run test`.
//
// The bench reads no memory. As in PPTX Editor and DOCX Editor, memory is measured in the
// browser (`browser-bench.ts memory`), where the heap is the heap of the user's page.
//
// Model: PPTX Editor's `scripts/bench/edit-bench.ts` at d4502552 (EigenPal), which
// follows DOCX Editor's `scripts/bench/edit-bench.ts` at adf0e178. Differences, each for a
// reason:
// - The sizes are workbooks, not decks: `--rows`, `--sheets`, and `--strings` set the size of
//   each workbook, and `--workbooks` picks the workbooks. A workbook has no rail of thumbnails,
//   so there is no thumbnail counter.
// - An unknown argument stops the run (`args.ts`), as in DOCX Editor's bench, and `--help`
//   prints the usage.
// - Each run of a step is timed from the input to the end of the render. The editor renders on
//   the next frame, and the bench renders at once.
// - `--compare <before.json>` compares with a report of `--json`, as in DOCX Editor's bench at
//   1c9b6a29b. It refuses a report of other sizes, and a report whose workbook bytes have another
//   SHA-256, as DOCX Editor's refuses another fixture. PPTX Editor's bench has no `--compare`;
//   this bench takes it from DOCX Editor, because its timings have no budget, and a comparison is
//   how a contributor shows a change in speed.
// - The defaults are 40 timed runs after 5 untimed ones, as in PPTX Editor's bench.
// - The p95 is the sample at rank `ceil(0.95 n)` (`p95` in `measure.ts`), as in DOCX Editor's
//   bench at 1c9b6a29b, so both benches of this repository use one rank. PPTX Editor's bench
//   takes the sample after it, `floor(0.95 n) + 1`, so at 40 runs it gives the 39th sample, not
//   the 38th.
// - Each timed run of a step must do the same work, as in DOCX Editor's bench: a run that does
//   other work stops the bench. Each commit must write its value, as DOCX Editor's bench checks
//   that each edit commits. The commits stay inside the rows of the data, and each one writes a
//   value that no cell of the data holds, so each one does the same work. `--runs` and
//   `--warmup` that need more rows for a selected workbook stop the run before any work.
// - DOCX Editor's bench at 1c9b6a29b also compares its layout after an edit with a clean layout.
//   This bench does not: no editor function that it calls returns a layout. The editor lays out
//   each window with `layoutSheet` (`surface-render.ts`), which holds both the row geometry cache
//   and the row height cache, and the editor keeps no row geometry of its own. So
//   `packages/core/src/layout/__tests__/row-height-cache.test.ts` compares on that function:
//   after a number, a new text, a font size, undo and redo, a wrapped cell, and a new measurer,
//   the heights must equal the heights of the saved bytes read again.
// - PPTX Editor's bench registers the happy-dom globals. This bench makes one happy-dom `Window`
//   for each workbook and gives the editor a node of it, because the editor finds its window
//   from a node that it owns (AGENTS.md, "Host window").
// - PPTX Editor's bench reads the state of each toolbar slot after each change. This bench does
//   not: `toolbarCommandState` reads only the snapshot, and each timed step builds the snapshot,
//   so the step already holds that work.
//
// Usage: see `USAGE`.

import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { createXlsxEditor, flushEditor } from '../../packages/core/src/editor/create-editor.ts';
import { names, positive, readFlags, stopOnUsageError, UsageError } from './args.ts';
import { median, p95, WORKBOOKS, type WorkbookName } from './measure.ts';
import { rowHeightTestRecorder } from '../../packages/core/src/layout/row-height.ts';
import { sheetIndexTestRecorder } from '../../packages/core/src/store/workbook/sheet-index.ts';
import {
  BENCH_SIZES,
  gridWorkbook,
  SHEET_COLS,
  SHEET_ROWS,
  sheetsWorkbook,
  stringsWorkbook,
} from './workbooks.ts';

type Editor = ReturnType<typeof createXlsxEditor>;

export interface Args {
  workbooks: WorkbookName[];
  /** The rows of the grid workbook. */
  rows: number;
  /** The worksheets of the sheets workbook. */
  sheets: number;
  /** The shared strings of the strings workbook. */
  strings: number;
  runs: number;
  warmup: number;
  json: boolean;
  /** A report of `--json` to compare with. */
  compare?: string;
}

const USAGE = `Usage: bun run bench:edit -- [options]

  --workbooks <names>  grid, sheets, strings, or a list (default: all three)
  --rows <n>           rows of the grid workbook (default: ${BENCH_SIZES.gridRows})
  --sheets <n>         worksheets of the sheets workbook (default: ${BENCH_SIZES.sheets})
  --strings <n>        shared strings of the strings workbook (default: ${BENCH_SIZES.strings})
  --runs <n>           timed runs of each step (default: 40)
  --warmup <n>         untimed runs before them (default: 5)
  --json               print the results as JSON, not as a table
  --compare <file>     also compare with a report of --json at the same sizes
  --help               print this text

The commits of a workbook need 2 x (runs + warmup) rows of data. The sheets
workbook has 100 rows a sheet.`;

/** The arguments, or undefined after `--help`. */
function parseArgs(argv: readonly string[]): Args | undefined {
  const flags = readFlags(
    argv,
    {
      values: ['--workbooks', '--rows', '--sheets', '--strings', '--runs', '--warmup', '--compare'],
      switches: ['--json'],
    },
    USAGE
  );
  if (!flags) return undefined;
  const compare = flags.get('--compare');
  const sheets = positive(flags, '--sheets', BENCH_SIZES.sheets);
  if (sheets < 2) throw new UsageError('--sheets takes 2 or more: the sheet switch needs 2 sheets');
  return {
    ...(typeof compare === 'string' ? { compare } : {}),
    workbooks: names(flags, '--workbooks', WORKBOOKS),
    rows: positive(flags, '--rows', BENCH_SIZES.gridRows),
    sheets,
    strings: positive(flags, '--strings', BENCH_SIZES.strings),
    runs: positive(flags, '--runs', 40),
    warmup: positive(flags, '--warmup', 5),
    json: flags.has('--json'),
  };
}

interface Timing {
  medianMs: number;
  p95Ms: number;
}

/** The work of one step, the same in each timed run. */
interface Counters {
  rowsMeasured: number;
  rowsCached: number;
  sheetIndexes: number;
  cellsIndexed: number;
}

export interface ScenarioResult {
  name: string;
  workbook: WorkbookName;
  /** The cells of the workbook. */
  cells: number;
  /** The SHA-256 of the workbook bytes, so `--compare` refuses a report of other bytes. */
  workbookSha256: string;
  total: Timing;
  perOperation: Counters;
}

const timing = (xs: readonly number[]): Timing => ({ medianMs: median(xs), p95Ms: p95(xs) });

/** The substitute font that the tests use, from the packaged fonts. */
function substituteFont() {
  const pkg = Bun.resolveSync(
    '@docx-editor.dev/fonts/package.json',
    join(import.meta.dir, '../../packages/core/src')
  );
  const bytes = new Uint8Array(readFileSync(join(dirname(pkg), 'assets/Carlito-Regular.ttf')));
  return [{ family: 'Carlito', bold: false, italic: false, bytes, origin: 'substitute' as const }];
}

const rowWork = rowHeightTestRecorder();
const indexWork = sheetIndexTestRecorder();

/** An editor in a happy-dom window, with a scroller of a fixed size. */
class Harness {
  readonly win = new Window();
  readonly host: HTMLElement;
  readonly editor: Editor;

  constructor() {
    const doc = this.win.document as unknown as Document;
    this.host = doc.createElement('div');
    doc.body.append(this.host);
    this.editor = createXlsxEditor({ fonts: substituteFont() });
    this.editor.attach(this.host);
    const scroller = this.scroller;
    Object.defineProperty(scroller, 'clientWidth', { get: () => 1400 });
    Object.defineProperty(scroller, 'clientHeight', { get: () => 800 });
  }

  get scroller(): HTMLElement {
    return this.host.querySelector('.xlsx-scroller') as HTMLElement;
  }

  get input(): HTMLTextAreaElement {
    return this.host.querySelector('.xlsx-cell-editor') as HTMLTextAreaElement;
  }

  key(target: EventTarget, key: string): void {
    const KeyboardEvent = this.win.KeyboardEvent as unknown as typeof globalThis.KeyboardEvent;
    target.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true }));
  }

  paint(): void {
    flushEditor(this.editor);
  }

  close(): void {
    this.editor.destroy();
    void this.win.happyDOM.close();
  }
}

/** The counters since the last reset of the recorders. */
const counters = (): Counters => ({
  rowsMeasured: rowWork.measured,
  rowsCached: rowWork.cached,
  sheetIndexes: indexWork.builds,
  cellsIndexed: indexWork.cells,
});

/** The workbook of a step: its name, its cells, and the SHA-256 of its bytes. */
interface Subject {
  workbook: WorkbookName;
  cells: number;
  workbookSha256: string;
}

/**
 * Times `step` over the warmup and the timed runs. `check` runs after each run, untimed, and
 * throws when the run did not do its edit. Each timed run must do the same work, as in DOCX
 * Editor's bench: a run that does other work stops the bench.
 */
function measure(
  h: Harness,
  name: string,
  subject: Subject,
  args: Args,
  step: (i: number) => void,
  check?: (i: number) => void
): ScenarioResult {
  for (let i = 0; i < args.warmup; i += 1) {
    step(i);
    h.paint();
    check?.(i);
  }
  // As in PPTX Editor's bench, the garbage of the warmup is collected before the timed runs.
  Bun.gc(true);
  const xs: number[] = [];
  let work: Counters | undefined;
  for (let i = 0; i < args.runs; i += 1) {
    rowWork.reset();
    indexWork.reset();
    const t0 = performance.now();
    step(args.warmup + i);
    h.paint();
    xs.push(performance.now() - t0);
    const run = counters();
    check?.(args.warmup + i);
    work ??= run;
    if (JSON.stringify(run) !== JSON.stringify(work))
      throw new Error(
        `${name} on ${subject.workbook}: the work changed between runs: ` +
          `${JSON.stringify(work)}, then ${JSON.stringify(run)}`
      );
  }
  return { name, ...subject, total: timing(xs), perOperation: work! };
}

/** The bytes and the cell count of a bench workbook. */
function workbookOf(name: WorkbookName, args: Args): { bytes: Uint8Array; cells: number } {
  switch (name) {
    case 'grid':
      return {
        bytes: gridWorkbook(args.rows, BENCH_SIZES.gridCols),
        cells: args.rows * BENCH_SIZES.gridCols,
      };
    case 'sheets':
      return { bytes: sheetsWorkbook(args.sheets), cells: args.sheets * SHEET_ROWS * SHEET_COLS };
    case 'strings':
      return { bytes: stringsWorkbook(args.strings), cells: args.strings };
  }
}

const sha256 = (bytes: Uint8Array): string =>
  new Bun.CryptoHasher('sha256').update(bytes).digest('hex');

/** The message for a report whose workbook has other bytes. */
const otherBytes = (name: WorkbookName): string =>
  `--compare: the ${name} workbook has other bytes than in the report ` +
  '(its SHA-256 differs), so the steps did other work';

/**
 * The full data rows of the first sheet of a bench workbook. The last row of the strings workbook
 * can be part of a row, with no cell in the columns that the commits write, so it does not count.
 */
function dataRows(name: WorkbookName, args: Args): number {
  if (name === 'grid') return args.rows;
  if (name === 'sheets') return SHEET_ROWS;
  return Math.floor(args.strings / BENCH_SIZES.gridCols);
}

/** The rows that the commits of one workbook write: one for each run of the two commit steps. */
const commitRows = (args: Args): number => 2 * (args.warmup + args.runs);

/**
 * Stops the run when a selected workbook has fewer data rows than its commits need. A row past
 * the data adds a row to the sheet, and each later commit then walks one more row.
 */
export function checkRuns(args: Args): void {
  for (const name of args.workbooks) {
    const rows = dataRows(name, args);
    if (commitRows(args) <= rows) continue;
    const more = name === 'grid' ? ' or more rows' : name === 'strings' ? ' or more strings' : '';
    throw new UsageError(
      `the ${name} workbook has ${rows} data rows, and ${args.warmup} warmup and ${args.runs} ` +
        `timed runs need ${commitRows(args)} rows for their commits: ` +
        `use fewer runs or warmup runs${more}`
    );
  }
}

/** The number that the commit step writes: no cell of a bench workbook holds it. */
const COMMIT_NUMBER = '7123';

export async function benchWorkbook(name: WorkbookName, args: Args): Promise<ScenarioResult[]> {
  checkRuns({ ...args, workbooks: [name] });
  const { bytes, cells } = workbookOf(name, args);
  const subject: Subject = {
    workbook: name,
    cells,
    workbookSha256: sha256(bytes),
  };
  Bun.gc(true);
  const h = new Harness();
  try {
    return await steps(h, name, args, subject, bytes);
  } finally {
    h.close();
  }
}

/** The steps of `benchWorkbook` in the editor of `h`. */
async function steps(
  h: Harness,
  name: WorkbookName,
  args: Args,
  subject: Subject,
  bytes: Uint8Array
): Promise<ScenarioResult[]> {
  const rows = dataRows(name, args);
  const commits = commitRows(args);
  const { editor } = h;
  // The fonts load with the first workbook, so a warm editor opens the bench workbook.
  await editor.load(new Uint8Array(0));
  // Open: from the bytes to the first sheet painted.
  rowWork.reset();
  indexWork.reset();
  const t0 = performance.now();
  const opened = await editor.load(bytes);
  h.paint();
  const openMs = performance.now() - t0;
  if (opened.status !== 'ready') {
    const why = opened.failure;
    // A size above the default can make a part larger than the XML part limit (64 MiB).
    throw new UsageError(
      `the ${name} workbook did not open: ${why?.code ?? 'unknown'}` +
        (why?.detail ? ` (${why.detail})` : '') +
        '. A size above the default can pass a part limit: use a smaller size.'
    );
  }
  const out: ScenarioResult[] = [
    {
      name: 'open',
      ...subject,
      total: { medianMs: openMs, p95Ms: openMs },
      perOperation: counters(),
    },
  ];
  // The middle of the first sheet, as a user who scrolled there.
  const middle = Math.floor(rows / 2);

  // (a) One keystroke in the cell editor: the first key starts the edit, the rest type.
  editor.select({ row: middle, col: 3 });
  h.paint();
  h.key(h.scroller, '4');
  h.paint();
  out.push(
    measure(
      h,
      'keystroke',
      subject,
      args,
      () => {
        const input = h.input;
        input.value += '4';
        input.dispatchEvent(new h.win.Event('input', { bubbles: true }) as unknown as Event);
      },
      // The first key and each run typed one 4.
      (i) => {
        const text = editor.snapshot().editText;
        const want = '4'.repeat(i + 2);
        if (text !== want)
          throw new Error(`keystroke on ${name}: the cell editor holds ${text}, not ${want}`);
      }
    )
  );
  editor.cancelEdit();
  h.paint();

  // Each commit must have written its cell. The check selects the cell, untimed; the next run
  // selects its own cell first.
  const wrote = (step: string, col: number, value: string) => () => {
    editor.select({ row: row - 1, col });
    const got = editor.snapshot().activeCell?.content;
    if (got !== value)
      throw new Error(`${step} on ${name}: row ${row} holds ${got}, not the value ${value}`);
  };

  // (b) One committed number: a key starts the edit, and Enter writes the cell. The commits start
  // in the middle, or higher when the rows below it are too few for them.
  let row = Math.min(middle, rows - commits);
  out.push(
    measure(
      h,
      'commit',
      subject,
      args,
      () => {
        editor.select({ row: row++, col: 3 });
        h.key(h.scroller, '7');
        const input = h.input;
        input.value = COMMIT_NUMBER;
        input.dispatchEvent(new h.win.Event('input', { bubbles: true }) as unknown as Event);
        h.key(input, 'Enter');
      },
      wrote('commit', 3, COMMIT_NUMBER)
    )
  );

  // (c) One committed text that is new to the workbook, so the shared string table grows.
  out.push(
    measure(
      h,
      'commit-text',
      subject,
      args,
      (i) => {
        editor.select({ row: row++, col: 4 });
        h.key(h.scroller, 'x');
        const input = h.input;
        input.value = `bench text ${i}`;
        input.dispatchEvent(new h.win.Event('input', { bubbles: true }) as unknown as Event);
        h.key(input, 'Enter');
      },
      (i) => wrote('commit-text', 4, `bench text ${i}`)()
    )
  );

  // (d) One scroll step. The steps stay inside the rows of the sheet, at about 20 pixels a row.
  // Each step must not show a row above the row of the step before, and the last step must show
  // a row below the row of the first step. A step can be shorter than a row.
  const step = Math.min(2000, Math.floor((rows * 20) / (args.warmup + args.runs + 1)));
  const topRow = () =>
    Number(h.host.querySelector('[aria-rowindex]')?.getAttribute('aria-rowindex') ?? 0);
  let y = 0;
  let firstTop = 0;
  let top = 0;
  out.push(
    measure(
      h,
      'scroll',
      subject,
      args,
      () => {
        y += step;
        h.scroller.scrollTop = y;
        h.scroller.dispatchEvent(new h.win.Event('scroll') as unknown as Event);
      },
      (i) => {
        const shown = topRow();
        if (i > 0 && shown < top)
          throw new Error(`scroll on ${name}: the grid shows row ${shown} after row ${top}`);
        if (i === 0) firstTop = shown;
        top = shown;
      }
    )
  );
  if (top <= firstTop)
    throw new Error(`scroll on ${name}: the grid did not scroll from row ${firstTop}`);

  // (e) A sheet switch between two neighbors that the user saw before. The first visit of a
  // sheet lays it out once, before the timed runs.
  if (name === 'sheets' && args.sheets > 1) {
    const first = Math.floor((args.sheets - 1) / 2);
    const show = (sheet: number) => {
      if (!editor.setActiveSheet(sheet))
        throw new Error(`sheet-switch on ${name}: sheet ${sheet} did not show`);
    };
    // The grid names its sheet (`grid.label`), so a run checks that the grid shows its sheet.
    const label = () => h.scroller.getAttribute('aria-label');
    const labels = new Map<number, string | null>();
    for (const sheet of [first, first + 1]) {
      show(sheet);
      h.paint();
      labels.set(sheet, label());
    }
    if (labels.get(first) === labels.get(first + 1))
      throw new Error(
        `sheet-switch on ${name}: the grid shows sheets ${first} and ${first + 1} alike`
      );
    const sheetOf = (i: number) => (i % 2 ? first + 1 : first);
    out.push(
      measure(
        h,
        'sheet-switch',
        subject,
        args,
        (i) => show(sheetOf(i)),
        (i) => {
          if (label() !== labels.get(sheetOf(i)))
            throw new Error(`sheet-switch on ${name}: the grid does not show sheet ${sheetOf(i)}`);
        }
      )
    );
  }
  return out;
}

/** The steps that must do no work in proportion to the sheet, with the reason. */
export function regressions(results: readonly ScenarioResult[]): string[] {
  const out: string[] = [];
  for (const r of results) {
    const c = r.perOperation;
    const at = `${r.name} on ${r.workbook}`;
    if (r.name === 'open') continue;
    if (r.name.startsWith('commit')) {
      if (c.rowsMeasured !== 1) out.push(`${at} measured ${c.rowsMeasured} rows, not 1`);
      continue;
    }
    if (c.rowsMeasured > 0) out.push(`${at} measured ${c.rowsMeasured} rows`);
    if (c.rowsCached > 0) out.push(`${at} read ${c.rowsCached} cached rows`);
    if (c.sheetIndexes > 0)
      out.push(`${at} built ${c.sheetIndexes} sheet ${c.sheetIndexes === 1 ? 'index' : 'indexes'}`);
  }
  return out;
}

/** One step against the same step of an earlier report. */
export interface Comparison {
  name: string;
  workbook: WorkbookName;
  /** The change of the median time, in percent. */
  medianChangePct: number;
  rowsMeasuredChange: number;
  rowsCachedChange: number;
  sheetIndexesChange: number;
  cellsIndexedChange: number;
}

/** The report that `--json` prints. */
export interface Report {
  schema: 1;
  /** The runtime, for example `bun 1.3.0`. */
  runtime?: string;
  args: Args;
  results: ScenarioResult[];
  /** The comparison of `--compare`. */
  comparison?: Comparison[];
}

/**
 * Each step of `results` against the same step of `before`. A report of other sizes is refused:
 * its steps did other work.
 */
export function compareReports(
  before: Report,
  args: Args,
  results: readonly ScenarioResult[]
): Comparison[] {
  const sizes = (a: Args) => sizesOf(a, args.workbooks);
  if (
    before?.schema !== 1 ||
    !Array.isArray(before.results) ||
    !before.results.every(validResult) ||
    typeof before.args !== 'object' ||
    before.args === null ||
    !(['rows', 'sheets', 'strings'] as const).every((k) => Number.isInteger(before.args[k]))
  )
    throw new UsageError('--compare takes a report of --json');
  if (sizes(before.args) !== sizes(args))
    throw new UsageError(
      `--compare: the report has ${sizes(before.args)}, not ${sizes(args)}. ` +
        'Run with the sizes of the report.'
    );
  const pct = (next: number, prior: number) => (prior === 0 ? 0 : ((next - prior) / prior) * 100);
  return results.flatMap((r) => {
    const b = before.results.find((x) => x?.name === r.name && x.workbook === r.workbook);
    if (!b) return [];
    if (b.workbookSha256 !== r.workbookSha256) throw new UsageError(otherBytes(r.workbook));
    return [
      {
        name: r.name,
        workbook: r.workbook,
        medianChangePct: pct(r.total.medianMs, b.total.medianMs),
        rowsMeasuredChange: r.perOperation.rowsMeasured - b.perOperation.rowsMeasured,
        rowsCachedChange: r.perOperation.rowsCached - b.perOperation.rowsCached,
        sheetIndexesChange: r.perOperation.sheetIndexes - b.perOperation.sheetIndexes,
        cellsIndexedChange: r.perOperation.cellsIndexed - b.perOperation.cellsIndexed,
      },
    ];
  });
}

/** A result of a `--json` report: a report changed by hand can lack a value. */
function validResult(x: ScenarioResult | null | undefined): boolean {
  const work = x?.perOperation;
  return (
    typeof x?.name === 'string' &&
    typeof x.workbook === 'string' &&
    typeof x.workbookSha256 === 'string' &&
    typeof x.total?.medianMs === 'number' &&
    typeof work === 'object' &&
    work !== null &&
    [work.rowsMeasured, work.rowsCached, work.sheetIndexes, work.cellsIndexed].every(
      (n) => typeof n === 'number'
    )
  );
}

/**
 * The sizes in `a` of the workbooks in `workbooks`: another size does not change their steps.
 */
function sizesOf(a: Args, workbooks: readonly WorkbookName[]): string {
  return [
    workbooks.includes('grid') ? `${a.rows} rows` : '',
    workbooks.includes('sheets') ? `${a.sheets} sheets` : '',
    workbooks.includes('strings') ? `${a.strings} strings` : '',
  ]
    .filter(Boolean)
    .join(', ');
}

/** The report in `file`; a file that cannot be read stops the run with its reason. */
function readReport(file: string): Report {
  try {
    return JSON.parse(readFileSync(file, 'utf8')) as Report;
  } catch (error) {
    throw new UsageError(`--compare: cannot read ${file}: ${(error as Error).message}`);
  }
}

const signed = (n: number, digits = 0): string => `${n >= 0 ? '+' : ''}${n.toFixed(digits)}`;

const fmt = (t: Timing): string => `${t.medianMs.toFixed(2)} / ${t.p95Ms.toFixed(2)}`;

async function main(): Promise<void> {
  const args = parseArgs(process.argv.slice(2));
  if (!args) return;
  // Read the earlier report first: a wrong file stops the run before any work.
  const before = args.compare ? readReport(args.compare) : undefined;
  checkRuns(args);
  if (args.compare) {
    const missing = Array.isArray(before?.results)
      ? args.workbooks.filter((w) => !before!.results.some((r) => r?.workbook === w))
      : [];
    if (missing.length > 0)
      throw new UsageError(
        `--compare: the report has no step of ${missing.join(', ')}. ` +
          'Run with --workbooks of the report.'
      );
    compareReports(before!, args, []);
    // The bytes of each workbook too: a change to `workbooks.ts` gives other bytes.
    for (const name of args.workbooks) {
      const sha = sha256(workbookOf(name, args).bytes);
      if (before!.results.some((r) => r?.workbook === name && r.workbookSha256 !== sha))
        throw new UsageError(otherBytes(name));
    }
  }
  const results: ScenarioResult[] = [];
  for (const name of args.workbooks) {
    // Progress goes to stderr, so the JSON of `--json` on stdout stays one document.
    console.error(`run    ${name}`);
    results.push(...(await benchWorkbook(name, args)));
  }
  const failed = regressions(results);
  if (failed.length > 0) process.exitCode = 1;
  const lines = [
    `Sizes: ${sizesOf(args, args.workbooks)}. ` +
      `Runs: ${args.runs} timed after ${args.warmup} untimed, on bun ${Bun.version}. ` +
      'The open runs once.',
    '',
    '| Step | Workbook | Cells | Total ms (median / p95) | Rows measured | Rows cached |' +
      ' Sheet indexes | Cells indexed |',
    '| --- | --- | --: | --: | --: | --: | --: | --: |',
  ];
  for (const r of results) {
    const c = r.perOperation;
    const counts = [c.rowsMeasured, c.rowsCached, c.sheetIndexes, c.cellsIndexed];
    lines.push(
      `| ${r.name} | ${r.workbook} | ${r.cells} | ${fmt(r.total)} | ${counts.join(' | ')} |`
    );
  }
  lines.push(
    '',
    failed.length === 0
      ? 'No keystroke, scroll step, or sheet switch read a row height or built a sheet index, ' +
          'and each commit measured only its row. A commit still takes the other rows from ' +
          'the cache and builds its sheet index again.'
      : 'Failed work checks:'
  );
  for (const why of failed) lines.push(`- ${why}`);
  const comparison = before ? compareReports(before, args, results) : undefined;
  if (comparison) {
    lines.push(
      '',
      `Compared with ${args.compare}:`,
      '',
      '| Step | Workbook | Median time | Rows measured | Rows cached | Sheet indexes |' +
        ' Cells indexed |',
      '| --- | --- | --: | --: | --: | --: | --: |'
    );
    for (const c of comparison)
      lines.push(
        `| ${c.name} | ${c.workbook} | ${signed(c.medianChangePct, 1)}% | ` +
          `${signed(c.rowsMeasuredChange)} | ${signed(c.rowsCachedChange)} | ` +
          `${signed(c.sheetIndexesChange)} | ${signed(c.cellsIndexedChange)} |`
      );
  }
  if (args.json) {
    for (const why of failed) console.error(why);
    const report: Report = {
      schema: 1,
      runtime: `bun ${Bun.version}`,
      args,
      results,
      ...(comparison ? { comparison } : {}),
    };
    console.log(JSON.stringify(report));
  } else {
    if (failed.length > 0) lines.push('', 'The run fails.');
    console.log(lines.join('\n'));
  }
}

if (import.meta.main) await main().catch(stopOnUsageError);
