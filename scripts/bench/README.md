# Speed benches

The speed benches open generated workbooks in the editor and report the time and the work of each step. No time from them fails CI. The `Speed bench` job compares the edit bench times of each pull request with its base commit, as information only. The gate is the work counter test, `edit-bench-gates.test.ts`, which runs inside `bun run test`. A count of work is exact on every machine, and a time is not.

To run the gate alone, enter this command:

```bash
bun test scripts/bench/edit-bench-gates.test.ts
```

If the gate fails, its message names the step, the workbook, and the counter. To see the work of each step at the sizes of the gate, run the edit bench at those sizes:

```bash
bun run bench:edit -- --rows 2000 --sheets 20 --strings 20000 \
  --runs 3 --warmup 1
```

If a change lowers the work on purpose, update the pins in `EXPECTED` in the gate test.

The generated workbooks are in `workbooks.ts`:

| Workbook | Content |
| --- | --- |
| `grid` | One worksheet of 20 columns: shared strings in column A and numbers in the other columns. It has 20,000 rows in the edit bench |
| `sheets` | 200 worksheets of 100 rows and 10 columns |
| `strings` | A shared string table and one worksheet that names each item once, 20 to a row. It has 200,000 items in the edit bench |

## Edit bench

The edit bench opens each workbook in the editor under happy-dom, with the production text measurer. It times six steps: `open`, `keystroke`, `commit` (a committed number), `commit-text` (a committed new text), `scroll`, and `sheet-switch` (on the `sheets` workbook only, between sheets that the grid showed before). The pins name each step with its workbook, for example `commit grid`, and the table and a failure line give the step and the workbook. Each time holds the layout and the DOM of the grid, but no canvas paint: happy-dom has no canvas. The browser benches time the paint. The edit bench also counts the work of each step:

- Rows measured: the rows whose height layout measured. A keystroke measures none, and a commit measures the row that it changed and each row with a wrapped cell. The bench workbooks have no wrapped cell.
- Rows cached: the rows whose height layout took from its cache. A commit takes each other row of its sheet from the cache, for example 1999 rows on a sheet of 2000 rows.
- Sheet indexes and Cells indexed: the builds of the cell index of a worksheet, and the cells that they read.

The open measures each row of the first sheet, so its time grows with the rows. The open also builds the index of each worksheet, so on the `sheets` workbook it builds 200 indexes.

The run fails when a keystroke, a scroll step, or a sheet switch reads a row height or builds a sheet index. It also fails when a commit measures more or fewer rows than the one row that it changed. The run stops when two timed runs of a step do different work. It also stops when a step does not do its edit. For example, a key does not type, a commit does not write its value, a scroll does not move, or a sheet does not show.

To run the edit bench at the default sizes, enter this command:

```bash
bun run bench:edit
```

At the default sizes and runs, local runs of the bench took 3.2 to 3.3 GiB of memory at their peak and 30 to 31 s. The memory grows with the sizes. If you pass larger sizes with `--rows` or `--strings`, use a machine with much free memory.

The commits of each workbook need two rows of data for each run. So the `sheets` workbook takes at most 50 runs and warmup runs together. The bench checks this for each selected workbook before any work.

To compare a change with the code before it, save a report before the change, and compare with it after the change:

```bash
bun run bench:edit -- --json > /tmp/edit-before.json
# Make the change.
bun run bench:edit -- --compare /tmp/edit-before.json
```

A comparison refuses a report of other sizes, and a report of other workbook bytes, for example after a change to `workbooks.ts`. `--help` lists the options, for example `--rows` and `--workbooks`.

## Browser benches

The browser benches open each workbook in the demo, in Playwright Chromium, on the Vite development server. So the times include the development build of React. They need Playwright's Chromium (`bunx playwright install chromium`).

| Command | What it measures |
| --- | --- |
| `bun run bench:open:browser` | The open, until the grid shows the first sheet, and the heap after it |
| `bun run bench:edit:browser` | The first key of an edit (`begin`), each typed key (`type`), each commit (`commit`), each arrow key (`move`), and on the `sheets` workbook each switch to the next sheet (`sheet-switch`) |
| `bun run bench:scroll:browser` | Each wheel step of the grid, from its `scroll` event. On the `sheets` workbook, it only opens the workbook and gives its peak memory |
| `bun run bench:memory:browser` | The heap, the DOM nodes, the listeners, and the documents after groups of commits |

Each time sample ends after the frame that paints the change, so it holds the paint that the user sees. Each browser bench also prints these values:

- The longest task, from the browser's long task entries. It covers the whole page, from the load of the demo and the open of the workbook to the last step. A value of 0 ms means that no task took 50 ms or more.
- The peak memory of the browser processes, in MiB.
- For `bench:edit:browser`, the median of the first and of the last 10 samples of each edit step. A step that grows slower with each edit shows a last median far above the first.

The browser `sheet-switch` goes to a sheet that the grid did not show before, so each time holds the first layout of that sheet. The `sheet-switch` of the edit bench goes between sheets that the grid showed before.

The browser benches use 15,000 rows and 125,000 shared strings. At these sizes, two local runs of the edit and memory modes took up to 1977 MiB (`workbooks.ts`). These sizes keep a margin under the 3 GiB cap of the hang watchdog. At larger sizes, the browser comes near the cap or goes above it.

To run one workbook, or to save the values, add options after `--`. `--help` lists all the options:

```bash
bun run bench:edit:browser -- --workbooks grid --out /tmp/edit-browser.json
```

The edit and memory modes refuse a `--rows` or `--strings` value that is too small for their edits, before they start the browser.

To compare two browser runs, save a file with `--out` before the change and another after it, and compare the values of the same workbook. Each file gives the values of the table, the samples, the browser version, and the SHA-256 of each workbook. A file of the memory bench also gives each memory reading, the leaks, and the workbooks that stopped. So you can check that both runs used the same browser and the same bytes.

A browser bench fails only when a workbook stops, or when the memory bench finds a leak. A workbook stops, for example, on a page error or on an edit that does not land. A step that adds no time sample in 120 s stops its workbook with the step, the key, and the element that has the focus.

The memory bench repeats commits only. Repeated opens and sheet switches have no memory scenario. The memory bench commits a new text each time, so each commit adds a shared string. First, 200 commits fill the undo history of the editor, with a garbage collection after each 5 commits. After that, each commit drops the oldest undo step. The bench then commits 20 groups of 5 texts, and it reads the memory after each group. A leak is one of these:

- A heap that grows by more than 10 KiB for each commit. A commit keeps about 1 to 4 KiB.
- More than 2 DOM nodes or 0.05 listeners for each commit. Each edit adds about 1 DOM node to the browser's own edit history of the cell editor. In local runs of a plain text field, that history stopped growing after about 1000 edits. The bench does not show this for the cell editor, because its 300 edits do not reach that count.
- A count that keeps growing of documents, object URLs, font faces, observers, workers, image bitmaps, intervals, or listeners on `window` or `document`.

During the run, the memory bench prints each reading after its group of commits. Each browser bench also prints each console error and warning of the page, with its count. So a warning that repeats on each edit shows next to a leak.

To see which objects grow, run the memory bench with heap snapshots. It prints the constructors that grew most from the end of the warmup to the end of the run. A snapshot of a large heap takes much memory, so use small workbooks:

```bash
MEMORY_SNAPSHOTS=1 bun run bench:memory:browser -- --workbooks grid --rows 2000
```

## Speed bench in pull requests

The `Speed bench` workflow (`.github/workflows/bench.yml`) runs `bench:edit` on each pull request and on its base commit. The base commit is the tip of `main` that the merge ref of the pull request is built on, and it runs its own copy of the bench. The rounds are interleaved: head, base, base, head, two times, each with 10 timed runs. Then `edit-bench-comment.ts` writes one pull request comment, and the job updates that comment after each push. The comment gives these values:

- The median time of each step on both commits, as the mean of the round medians, and the change in percent.
- The p95 time of each step on the head, as the mean of the round p95 values.
- The work counters that changed.

The `open` step runs once in each round, so its value comes from one sample in each round.

The times are information only, because a shared runner is noisy. So no time fails the job. A change inside 5%, or inside the spread of the round medians of a step, shows as within noise. A slower change of more than 20% that is not within noise gets a warning. A head round fails the job when it fails as a local run does: for example, when a step does work in proportion to the sheet. A base round that fails is left out of the comment. If the install or the worktree of the base commit fails, the job fails.

If the base commit has no edit bench, every base round fails, its rounds use another setup, or its workbooks have other bytes, the comment gives the head only. A pull request from a fork gets no comment. The `edit-bench-report` artifact of the workflow run holds the reports and the comment.

## Compare times

A time depends on the machine, its load, and its power mode. Compare times only between runs on the same machine. Run one bench at a time, and never with `bun run fidelity`, `bun run test:e2e`, or `bun run test:hang:browser`.

## Sources

These benches follow the benches of DOCX Editor and PPTX Editor:

- `edit-bench.ts` follows PPTX Editor's `scripts/bench/edit-bench.ts` at d4502552, which follows DOCX Editor's at adf0e178. `--compare` comes from DOCX Editor's at 1c9b6a29b.
- `edit-bench-gates.test.ts` follows DOCX Editor's `scripts/bench/edit-bench-gates.test.ts` at 1c9b6a29b, the commit that added that test with DOCX Editor's bench README. Its scaling counts follow PPTX Editor's `packages/core/src/editor/__tests__/edit-scaling.test.ts` at d4502552.
- `browser-bench.ts` follows PPTX Editor's `e2e/open.bench.ts`, `e2e/edit.bench.ts`, and `e2e/memory.bench.ts` at d4502552. The warmup edits and the first and last 10 edits come from DOCX Editor's `e2e/edit-browser.bench.spec.ts` and `e2e/edit-browser-bench-probe.ts` at 1c9b6a29b.
- `workbooks.ts` follows PPTX Editor's `scripts/bench/text-deck.ts` at d4502552.
- `measure.ts` follows `summarize` of DOCX Editor's `e2e/edit-browser-bench-harness.ts` at 1c9b6a29b for the p95, and `runSustained` of its `e2e/edit-browser-bench-probe.ts` for the first and last 10 edits. Its leak check follows the memory verdict of PPTX Editor's `e2e/memory.bench.ts` at d4502552.
- `.github/workflows/bench.yml` and `edit-bench-comment.ts` follow DOCX Editor's `.github/workflows/bench.yml` and `scripts/bench/edit-bench-comment.mjs` at 0eb629bcd, with the interleaved rounds of 66819cf1a and the four rounds a side of 2b9941d98.
- `args.ts` refuses an unknown argument, as DOCX Editor's `scripts/bench/edit-bench.ts` at 1c9b6a29b does.
- `memory-probe.ts` is a copy of the page probe of PPTX Editor's `e2e/memory.bench.ts` at d4502552, and `heap-summary.ts` is a copy of its heap snapshot summary.

The header of each file records its differences from its source.
