# Contributing to xlsx-editor.dev

This guide covers setup, checks, and pull requests for the monorepo.

## Prerequisites

- [Bun](https://bun.sh/), the version in `.bun-version`, which CI uses.
- [Node.js](https://nodejs.org/) 24, the CI version.
- `xmllint` from libxml2, for `bun run test` and `bun run validate:xsd`.

These versions apply to repository development. Published packages declare their supported Node.js versions in their `engines` fields.

## Development setup

Clone the repository, install dependencies, and run the tests:

```bash
git clone https://github.com/eigenpal/xlsx-editor.git
cd xlsx-editor
bun install
bun run reference:fetch
bun run test
```

`bun run reference:fetch` downloads ECMA-376 Parts 1 to 4 and writes a text copy and a subclause index of each part (`reference/README.md`). Git ignores these files, so run it in each new clone and each new git worktree. It needs Poppler (`pdftotext`).

`bun run validate:xsd` and `bun run test` need `xmllint` from libxml2. macOS includes it. On Debian or Ubuntu, install the `libxml2-utils` package. Without it, both commands fail; neither passes without validating.

`bun install` also installs the pre-commit hook.

## Architecture overview

The engine design reads a workbook into one canonical tree per part, edits it only through typed ops, and lays out and paints the grid from that tree. For the pipeline and its rules, see "Architecture: one pipeline" in [Agent instructions](AGENTS.md).

### Core lanes

Each directory under `packages/core/src/` is a lane with a declared dependency edge and environment. `packages/core/src/__tests__/core-lane-graph.ts` is the machine-readable graph. If a lane needs a new dependency, edit that graph in the same change. Inside this repository, the demo, the tests, and `bun run typecheck` map the package names to the sources in `src/`, so they need no build.

These guards hold the lanes:

- `packages/core/src/__tests__/core-lane-graph.test.ts` proves that the graph has no cycle and that each lane is a package subpath.
- `packages/core/src/__tests__/core-lane-imports.test.ts` checks every import statement of every lane against the graph.
- `packages/core/src/__tests__/browser-bundle-graph.test.ts` walks what a browser import reaches, and refuses Node built-in modules and Node-only packages.
- `bun run check:lane-boundaries` compiles each DOM-free lane without the DOM library.

The guards are copied from PPTX Editor, which adapted them from DOCX Editor. Each file header names its source and its differences.

## Agent instructions

[Agent instructions](AGENTS.md) holds the rules for every contributor and coding agent. A local `local/AGENTS.md` file, which git ignores, can add steps for one machine.

## Collaboration compatibility

Base features come first. Collaboration starts in its own change after core editing works, and from then on every editing feature supports it. For the checks, see "Collaboration verification" in [Agent instructions](AGENTS.md).

## Make changes

1. Fork the repository and create a branch from `main`.
2. Read [Agent instructions](AGENTS.md) and the [repository map](docs/repo-map.md).
3. Read the requirements of the [active change](openspec/changes/typed-ooxml-workbook-editor/) that your work belongs to.
4. Keep each change focused, with tests. A rendering change also needs a synthetic fixture and its reference PDF, built by the fixture builder (task 2.2 of the active change).
5. Verify the change:
   ```bash
   bun run typecheck
   bun run lint
   bun run format:check
   bun run test
   bun run check:security
   bun run check:css-scope
   bun run check:terminology
   bun run check:docs
   bun run check:parity
   bun run check:lane-boundaries
   bun run fixtures:check
   bun run validate:xsd
   bun run i18n:validate
   bun run check:tsconfig-lib-floor
   bun run check:shipped-advisories
   bun run build:packages
   bun run notices:generate
   bun run api:check
   bun run check:consumer-types
   bun run check:package-artifacts
   bun run check:shipped-shaper
   bun run check:package-size
   bun run check:consumer-install
   bun run check:consumer-floors
   bun run check:api-changeset --base origin/main
   bun run fidelity
   bun run fidelity:check-gaps
   bun run test:e2e
   bun run test:hang:browser
   ```
   `bun run fidelity` needs Chromium and Poppler (see [Fidelity harness](#fidelity-harness)). `bun run test:e2e` runs the demo smoke check in Chromium, and needs Playwright's Chromium (`bunx playwright install chromium`). CI runs it in the `e2e` check. To run another engine, add `--browser firefox` or `--browser webkit`; these runs are not part of the gate. Run `bun run test:e2e` when you change the demo, the editor chrome, input, focus, the clipboard, paint, or the sample workbook (see [Demo sample and icons](#demo-sample-and-icons)). If a step fails, fix the first failed step first, and look at its screenshot in `screenshots/`. To watch the run in a browser window, set `DEMO_HEADED=1`. `bun run test:hang:browser` needs Chromium, and CI runs it in the `hang` job on each push to `main`. Run it when you change parsing, layout, paint, font loading, saving, or the open, save, and headings flows of the demo. A SLOW case does not fail the run. The run prints it, and the CI job summary lists it. Its report and the screenshot of each case that is not `ok` are in `test-results/hang/`. To run some cases, set `HANG_ONLY=id,id`. To run every case, set `HANG_ALL=1`. To watch the run in a browser window, set `HANG_HEADED=1`. `bun run build:packages` builds the publishable packages into `dist/`, and the package checks after it read that build. `bun run notices:generate` writes the `THIRD_PARTY_NOTICES.md` of each package from that build, and `bun run notices:check` fails when one is not current. `bun run check:shipped-advisories` needs network access to the npm registry. `bun run check:consumer-install` needs Chromium and network access to the npm registry. `bun run check:consumer-floors` needs the same, and its `--node` mode runs only on a floor Node.js version. If you change a public export, run `bun run build:packages`, then `bun run api:extract`, and commit the changed reports in `docs/api/`. If a public signature reaches a type that its entry does not export, export the type from that entry, or review it and run `node scripts/api-extractor.mjs --local --record-forgotten-exports`, and commit `scripts/api-forgotten-exports.baseline.json`. If you changed `openspec/`, also run `openspec validate typed-ooxml-workbook-editor --strict`. Install the OpenSpec CLI on your machine first (`@fission-ai/openspec`).
6. Add a changeset for each code change to a published package with `bun changeset`. Name the exact package, use `patch` by default, `minor` for additive public API, and `major` for breaks. Documentation, test, and CI-only changes do not need one. See [Releases](#releases).
7. Submit a pull request against `main` with a conventional-commit title. Fill in the sections of the pull request template.

## Releases

The `Release` workflow (`.github/workflows/release.yml`) publishes `@xlsx-editor.dev/core`, `@xlsx-editor.dev/react`, `@xlsx-editor.dev/vue`, and `@xlsx-editor.dev/i18n` as one fixed group, with one version.

1. A push to `main` with pending changesets opens or updates the `chore: release` pull request. That pull request applies the version bumps and the changelog entries.
2. A maintainer merges the `chore: release` pull request. The workflow then runs the prepublish checks, builds the packages, checks the package artifacts, and publishes the checked build.

The workflow publishes through npm trusted publishing. No npm token exists. Each package trusts only `release.yml` in this repository.

The packages are in prerelease mode. While `.changeset/pre.json` exists, each release is a beta version, for example `0.1.0-beta.0`, then `0.1.0-beta.1`. The workflow publishes a beta version to the `beta` dist-tag, never to `latest`. Install a beta with `npm install @xlsx-editor.dev/react@beta @xlsx-editor.dev/core@beta`. To make the next release stable, run `bun changeset pre exit` in a pull request.

## Run tests

Run one test file during development:

```bash
bun test path/to/file.test.ts
```

A test file in `packages/react` or `packages/vue` needs the DOM preload:

```bash
bun test --preload ./scripts/test/preload-dom.ts packages/react/src/__tests__/file.test.tsx
```

Other test commands:

- `bun run test` runs one `bun test` process per test file over a pool of workers, as CI does. CI splits the files over three `unit` jobs (`--shard 1/3` to `3/3`), balanced by the durations that each push to `main` records, and runs each source check (`lint`, `typecheck`, `policy`, `xsd`, `build`, `fidelity`, `e2e`, `changeset`) as a job of its own. It includes the adapter tests in `packages/react` and `packages/vue`, and preloads happy-dom for those files only.

  CI runs on the standard GitHub-hosted `ubuntu-24.04` runners. Each job has 4 vCPUs, and `bun run test` starts one worker for each CPU (`availableParallelism()`), up to the memory cap. Pull requests from forks run the same jobs with no repository secrets; only the sticky comments of the `package-size` job and the speed bench need a branch of this repository, and a fork pull request finds those numbers in the job artifacts. The durations cache key names the runner (`test-durations-ubuntu-2404-`), so a change of runner starts a new balance instead of a split measured on other machines. Until the first push to `main` on a new runner saves durations, each shard splits by file count.

- `bun run test:changed` runs only the test files that your change can affect. Add `--list` to print them.
- `bun run test:serial` runs the suite serially in two processes: the DOM-free packages and scripts, then `packages/react` and `packages/vue` with the DOM preload. Use it to find a test that depends on file order or leaves module state behind.
- `bun run test:dom` runs the adapter tests in one process with the DOM preload, as PPTX Editor does. Use it to find a test that leaves state on `document`.
- `bun run test:docx-shared` runs DOCX Editor's tests of the shared modules against `packages/ooxml-core`, and reports how far each module drifts from its source. It needs a local DOCX Editor checkout.

  The script looks in `~/Developer/docx-editor` and `~/docx-editor`. Set `DOCX_EDITOR_DIR`, or pass `--docx <path>`, for another place. The script reads `origin/main` of the checkout by default (`--ref <ref>` picks another commit, and `--ref worktree` reads the working tree). Run `git fetch` there first. The script never writes to the checkout.

### Browser tests

The flows of the demo (`examples/vite`) run as happy-dom tests in `bun run test`, as the editor tests do: `examples/vite/src/__tests__/demo-*.test.tsx` render the demo's page with stand-ins for the network, the file save picker, the download link, and the question before unsaved edits are lost (`demo-harness.tsx`). They cover the start, the focus regions, view only, the theme, each Save, the number formats, typed entry, the headings switch, a file that does not open, the questions before unsaved edits are lost, the work that overlaps a load, the logic of a save with no file save picker, and each fixture's open, edit, save, and reopen. Add a test there for a new flow of the demo. `bun test examples/vite/src/__tests__/<file>` runs one file with a DOM.

`bun run test:e2e` runs a short smoke check of the demo in Playwright Chromium: the sample opens and every sheet paints, and a typed value saves and opens again from the saved copy. It also keeps the steps that need a real browser engine: the paint and the loaded fonts, the system clipboard and its `copy`, `cut`, and `paste` events, focus that the engine moves in an editable host page, the layout of the formula bar, computed styles of the chrome in a host page, and a real download. `--browser firefox` and `--browser webkit` run the same steps in another engine, outside the default gate, and `--help` lists the options. It fails on a console error, a page error, an alert that is still on the page when a step ends, a sheet that does not show, or a saved copy that loses an edit or a sheet. The header of `scripts/verify-demo.ts` lists each check. A step that one engine cannot pass with no browser window is listed in `HEADLESS_SKIPS` with its cause, and its result line names the cause; never skip a step without one. A headless Firefox run skips the cut step, because the headless Firefox clipboard drops the private clipboard type; `DEMO_HEADED=1` runs it. `bun run test:e2e` runs in the CI `e2e` check and not in the pre-commit hook. Fidelity stays Chromium only. Run `bun run test:e2e` before you push a change to `examples/vite/`, the editor chrome, input, focus, the clipboard, paint, or the sample workbook. Firefox and WebKit have no file save picker, so the demo downloads the saved copy there: the `save with no picker` step reads the bytes of a download in every engine, with no stand-in picker. The other steps save through a stand-in picker.

The hostile workbooks of `scripts/hang/cases.ts` run in two places. Each workbook is built at test time. `bun run test` runs each file case with no network fault and no race through the engine, with no DOM (`scripts/hang/engine-run.ts`): the open check of the demo, the store, the layout of the visible window, a save, an edit, and a reopen. Work counters of the engine, not a clock, bound the work of the window: cell lookups, key reads, key searches, and rule checks stay under a cap from a reference workbook, and a scroll builds no sheet index and reads no row height. The first layout still reads each stored row once for its height, and no bound covers the merged ranges and drawings that a window reads; the header of `engine-run.ts` lists what each bound checks. A new file case runs there with no other change, and `scripts/hang/engine-coverage.test.ts` fails on a case that runs nowhere. `bun run test:hang:browser` keeps the cases that need a real browser (`scripts/hang/browser-cases.ts`): stalled and blocked font downloads, a lost network, races in the demo, the cases whose checks need the page (requests to other origins, the paint, and the chrome), and the large workbooks whose browser memory and main-thread blocks it measures. It opens each one in the demo, in its own Playwright Chromium. A watchdog outside the page stops the browser at 3 GiB of memory or at the time limit of the case. The run fails on a hang, a main-thread block over 5 s, a crash, a request to another origin, a page error, or a page that opens no workbook afterward. Main-thread blocks of 100 ms or more that add up to more than 10 s give a SLOW mark: the run prints it and the job summary lists it, but it does not fail the job. The header of `scripts/hang/run.ts` lists the verdicts. `HANG_ONLY=id,id` runs some cases, and `HANG_ALL=1` runs every case of `cases.ts`. It runs in the CI `hang` job and not in the pre-commit hook. Run it before you push a change to parsing, layout, paint, font loading, saving, or the open, save, and headings flows of the demo. A hang that it finds is a defect: fix its cause. Never raise a limit to pass. Pull request CI skips it, and each push to `main` runs it, so run it yourself when you change the open, the parse limits, or the save.

## Reference PDFs

Each fixture in `fixtures/` has a reference PDF next to it, and `fixtures/references.json` records its identity: the SHA-256 of the workbook and of the PDF, the page count, the page size, and the number locale that the PDF shows: its separators, and the built-in format codes that it replaces. The fidelity harness measures the render against these PDFs.

Maintainers make the reference PDFs. If you change a fixture, say so in your pull request, and a maintainer makes its reference PDF.

To add or update a reference:

1. In a fixture, name only the open faces that the harness draws: Carlito, Liberation Sans, Liberation Serif, and Liberation Mono. A cell format may name another family only when its `x:scheme` element replaces that name with a theme font. The font files are in `packages/core/node_modules/@docx-editor.dev/fonts/assets/` after `bun install`. The pattern `OPEN_FACES` in `scripts/fidelity/references.test.ts` is the rule: the test fails on any other embedded face.
2. Build the fixtures with `bun run fixtures:build`.
3. Put the reference PDF of each changed fixture at `fixtures/<name>.pdf`. Check with `pdffonts fixtures/<name>.pdf` that the PDF embeds only those faces; a test in `scripts/fidelity/references.test.ts` fails on any other face.
4. Record each reference with the decimal and group separators that the PDF prints:
   ```bash
   bun run fidelity:record-reference --decimal , --group ' ' basic-values
   ```

The reference PDFs print recalculated formula results, so the cached value of each fixture formula must equal its true result. The test in `scripts/fidelity/references.test.ts` fails when a fixture changes and its reference is stale. A missing or stale reference is a gap, never a pass.

## Fidelity harness

Common cases:

- To check your change, run `bun run fidelity`. Exit code 0 is a pass.
- If your change improves a score, run `bun run fidelity --update-gaps` and review the diff of `thresholds.ts`.
- If a run says that the lock is held, wait for the other run. Remove the lock folder only when the message gives the `rm -r` command.
- If the run stops with exit code 2, read its one `fidelity:` line. Run `bun run fidelity --help` for each exit code.

`bun run fidelity` measures the render of each fixture against its reference PDF. The harness page lays out and paints the first worksheet with the print profile. Chromium takes a screenshot of the page for the pixel score. Chromium also prints the page to a Letter PDF for the line scores. Poppler draws the reference PDF, and pdf.js reads the text lines of both PDFs. The harness compares pixels and text lines with the thresholds in `scripts/fidelity/thresholds.ts`.

The harness writes `fidelity-report/report.md`. The report folder also holds the rasters, the diff, and the printed PDF (`ours.pdf`) of each fixture.

Only one run at a time uses `fidelity-report/`. A run holds the lock folder `fidelity-report/.run.lock`, which names its process id. The run releases the lock when it ends, or when SIGINT or SIGTERM stops it. A second run does not start while the process of the lock runs. A run takes over a lock whose process is gone.

Remove the lock folder by hand only when no `bun run fidelity` runs and the run message tells you to. That message gives the `rm -r` command.

A run clears the entries of `fidelity-report/` that it writes again, all or none. If it cannot move an entry, it moves back the entries that it moved and exits with code 2. The message names the entry: give write access to it or remove it, then run again. If an entry cannot move back either, it stays in a folder `fidelity-report/.run.lock.cleared.*`, which no run removes. The message names that folder, and each later run prints a notice with the command that removes it. Move its entries back into `fidelity-report/`, or remove the folder. If the run cleared every entry but cannot remove the folder, it exits with code 2 and gives the command that removes that folder. A later run removes the folder when it can. Otherwise the run prints a notice with that command, and the folder stays until you run it.

To measure some fixtures only, name them: `bun run fidelity basic-values number-formats`. Such a run clears and writes only the folders of those fixtures and the report files. Its report and gate cover only those fixtures. With `--update-gaps`, it records only those fixtures in the list of your tools. Run `bun run fidelity --help` for the flags and exit codes.

Before you run it, install the tools:

```bash
bunx playwright install chromium
brew install poppler   # or: sudo apt-get install poppler-utils
```

The harness measures text with the reference fonts when the `XLSX_REFERENCE_FONTS` environment variable names one or more folders of font files. Separate folders with the path list separator. Never commit these fonts. Without the variable, the harness uses the `substitute` font environment. It draws the open faces of `@docx-editor.dev/fonts`, which the fixtures name, and says so in the report. With the variable, the harness loads only the faces in those folders and no substitutes. So the folders must hold Carlito and Liberation Sans, the faces that the fixtures name. Each font environment has its own known gaps.

Only the substitute environment has a measured list. A run with the reference fonts fails, and the report names the missing measurement. A missing record is a gap, never a pass.

The harness renders the print profile (`PRINT_PROFILE`) only. The screen profile (`SCREEN_PROFILE`) is the default of the editor. It has no fidelity measurement, and the report says so.

The line scores read the text of the PDF that Chromium prints of the painted canvas. So they measure the glyphs that the paint drew, at their painted places. The harness refuses that PDF in these cases:

- The PDF has more than one page.
- The PDF draws with a font that the harness page did not load.
- The print paint did not run exactly once.

The print paint uses a device pixel ratio of 1, and the screenshot for the pixel score uses 3. Both draw from the same layout in points.

The harness and `bun run fidelity:record-reference` also refuse a reference PDF in these cases:

- The PDF draws with a font that it does not embed.
- The PDF draws with a family that the workbook does not name.

The family comes from the name table of a reference font file with the same PostScript name. Without such a file, the family comes from the PostScript name itself.

### Clipped lines

Chromium leaves out a glyph that a cell clip hides entirely. The reference PDFs keep such glyphs as text (approximation "Line scores of clipped text"). So a painted line that a clip cuts matches its reference line only by the clip rectangles of both PDFs. Each of these rules uses `clipEdgePt`, 0.5 points, as its tolerance. Lines pair in this order. First come the same text and then the bullet form on its row, then the joins. Then the clipped pass and the extra-cut pass run. Then comes the same text on another row. Last, a clipped line that is left pairs with a cut of it on another row, by the rules of the clipped pass. Each step that pairs by text takes the nearest pair first. The clipped pass and the extra-cut pass take the reference lines in page order. Pieces of one page join to hold a line of the other page when each piece lies within `maxBaselineDeltaPt` of that line. A painted piece needs a key, so a cell of one character other than a digit joins nothing. The baseline of the worst piece is scored. Where the items of that line split at the pieces, its gap there is at most 1 em. Then the gap of the pieces is at most 1 em plus `joinGapPt` at each end. Elsewhere the gap of the pieces is at most 1 em plus `joinGapPt`. No piece width counts toward a gap. Each piece scores the nearer of its start and end distances from the characters that it holds. A painted piece scores the nearer of its edges that lie on reference item edges. When a piece has no edge on an item edge, the width of the group is scored too. The edges of the group count only for a line without items. So a join scores as the same pieces do as separate lines, apart from the width of the group. An inner end of a painted piece inside a reference item is named as `painted split inside a reference item`, and the run fails. The run prints `painted split inside a reference item: "<key>"`. The cause can be the layout or the paint, which splits text that the reference PDF draws as one item. It can also be the reference PDF, which draws the text of two cells as one item (approximation "Reference items across cells"). To tell them apart, read the reference runs of that item with `textRuns` in `scripts/fidelity/pdf-clips.ts`. An inner end on a boundary of two reference items gets the clip rules of a line end. A join piece off the row or past the gap limit leaves the line missing and the pieces extra (approximation "Join gap limit"). The gate does not guard the em shares `JOIN_GAP_EM`, `LINE_GAP_EM`, and `SPACE_GAP_EM`.

A painted line matches a cut reference line when all of these hold:

- The reference text runs past the reference clip on the cut side by more than the threshold.
- Without white space, the painted text starts the reference text for a cut on the right. For a cut on the left, it ends the reference text.
- The painted line holds each reference glyph whose advance starts before the clip edge (on a left cut: ends past it), and no glyph more. The threshold applies here too. This also holds for a painted line with the whole reference text. When the reference glyphs do not account for the text, such as a ligature, no rule checks the glyphs. Then the gate fails that end as a skipped check, `<text> @ <baseline> | <side> reference glyphs unknown`. The run prints it as `unknown reference glyphs: "<key>"`.
- The painted clip on that side is within the threshold of the reference clip.
- The painted line reaches the painted clip edge within the threshold.

A painted line of one character can match a cut reference line by these rules too.

Only the cut side is compared. The clip of an end is the clip of the run that draws the last glyph there. A run that starts a line of the next cell is not part of the line, also where it overlaps it. Neither is a run inside an item of the next cell, under the clip of the run that starts that item. When runs with different clips start that item, the end is unknown.

Paint draws no glyph for a space. So the clip edge can lie in a space of the reference text, and the painted line can end at the glyph before that space. Such a line is cut when all of these hold:

- The painted line ends within the threshold of the advance end of that reference word.
- The clip edge lies past that end, within the threshold.
- The clip edge lies before the advance start of the next reference word, within the threshold.

The run fails the gate when a clip differs in these ways:

- The reference clip cuts a line, and the painted line is not cut at the same edge. The run prints `uncut: "<key>"`. This rule also holds when the texts match in full.
- The reference clip holds an end of a line whole, and the painted line runs past the painted clip there by more than the threshold. The run prints `extra cut: "<key>"`.
- The painted clip hides whole glyphs at such an end, or its edge lies in a space there. The painted line then holds only the start or the end of the reference text. The run also prints `extra cut: "<key>"`.
- The painted clip is unknown at such an end. No painted run lies there, or its clip is not one rectangle that holds the baseline. Or runs with different clips start an item of the next cell there. The run prints `extra cut: "<text> @ <baseline> | <side> clip unknown"`.
- Several painted lines hold one reference line, and an inner end of one of them breaks one of these rules. The reference clip at that end is the clip of the characters that the painted line holds. At an end on a boundary of two reference items, the run prints the key of the reference item at that end. Elsewhere it prints the key of the painted line.

The reference PDF clips only text that overflows its cell. The other text of a row shares the clip of the print area, so the texts of two cells can form one reference line. The painted page clips the text of each cell, so it shows two lines there.

A lost or an extra clip is a regression to fix, not to record. The rules need each reference clip as one axis-aligned rectangle. Otherwise a cut line stays missing, and no rule checks that end of a matched line. The run prints `unknown reference clip: "<key>"` and fails the gate: a skipped check is a gap, never a pass. A missing line that the reference clip cuts on both sides prints `missing, cut on both sides: "<text>"`. A painted clip that is not one rectangle fails the gate as an uncut or an extra cut line.

Each list in `thresholds.ts` names the clipped lines of each fixture (`fixtureClipped`). A key holds the whole text, the baseline, and the clip edge of the reference PDF. A list names a key once for each time a clip cuts the line. The gate fails when a clip cuts a line more often than the list of the run names it. The gate asks for a new record when a clip cuts a named line less often.

Each run prints the key of each clipped line under its fixture, as `clipped: "<key>"`. In CI, the log of the **Fidelity against reference PDFs** step of the `fidelity` job shows the same lines. Paste the quoted key as it is into `fixtureClipped` of the list of that toolchain.

### Known gaps

A fixture that fails a threshold today is a known gap: `thresholds.ts` records the measured value of each check that it fails. The gate fails when a fixture fails a check that has no gap. It also fails when a score moves beyond the tolerance of its check (`GAP_TOLERANCE`). Each font environment keeps one list for each raster toolchain, the Poppler and Chromium versions. Other tools rasterize the same page differently. A run uses the list of its own tools, and in CI, a run with no such list fails.

Outside CI, the gate compares with the tightest recorded value of each check across all lists of the font environment. It reports the tool difference, a better score, and a check that passes now as notes. It fails on a worse score, an unlisted failure, a missing reference, and fewer reference lines. If your change improves a score, record the gaps again and review the diff. `--update-gaps` rewrites only the list of your tools, or adds one. It then runs `bun run fidelity:check-gaps` against the merge base:

```bash
bun run fidelity --update-gaps
```

A new Poppler, pdf.js, or Chromium version in CI, for example from a dependency update, has no list yet, so the `fidelity` job fails. The job then records the list of its tools with `--update-gaps` and uploads `thresholds.ts` as the `fidelity-thresholds` artifact. Download it, review its diff, and commit it in the pull request. You do not need the CI tools on your machine.

`--update-gaps` prints each change of the list: `LOOSER`, `NEW`, and `REMOVED`. It writes nothing when the run adds a gap, a check, or a clipped line, or when a value is looser. The run then prints `Nothing was written` and exits with code 1. Fix the regression instead.

The check against the merge base needs `origin/main`, and a merge base that has `thresholds.ts`. If the checkout has no `origin/main`, or the merge base has no `thresholds.ts`, `--update-gaps` exits with code 2, or 1 when the gate fails. It says that `thresholds.ts` holds an unchecked list. Then run the check with a base commit:

```bash
bun run fidelity:check-gaps --base <ref>
```

A new fixture or a new toolchain can need a worse list. Then add `--accept-worse`:

```bash
bun run fidelity new-fixture --update-gaps --accept-worse
```

The merge-base check (`bun run fidelity:check-gaps`) then applies its own rules, also in CI:

- A gap of a new fixture passes.
- A looser value or a new gap of a fixture of the merge base fails.
- A list for a new toolchain that is looser anywhere fails, unless an owner sets `XLSX_FIDELITY_ACCEPT_NEW_TOOLCHAIN`.
- A new clipped line of a new fixture, or of a fixture with a new reference PDF, fails unless an owner sets `XLSX_FIDELITY_ACCEPT_CLIPPED`. CI never sets either variable.
- A new clipped line of any other fixture fails.

A run with tools that have no list starts one. So it needs every fixture with no fixture error, and it takes no fixture names.

`bun run fidelity:check-gaps` compares with the merge base of `HEAD` and `origin/main`. `--base <ref>` compares with another commit. The check fails in these cases:

- A threshold becomes looser than at the merge base.
- A recorded gap value grows by any amount. The tolerance applies only to a run against its record.
- A fixture that has a gap is renamed or removed.
- A fixture of the merge base checks fewer lines, also with a new reference PDF. Fix the layout, or, as an owner, remove the fixture and add it as a new one in separate changes.
- A change that adds a fixture removes a fixture that has checked lines. A change that only removes a fixture without a known gap passes, and a later change that adds the name again adds a new fixture.
- A list names a clipped line more often than the merge base names it.
- A list drops its `fixtureClipped` record, or names clipped lines of a name that is not a fixture.
- A renamed fixture names a clipped line. When a fixture of the merge base is gone, a fixture that the merge base does not have takes no clipped line.
- A fixture of the merge base gets a gap of a check that the merge base does not have. A new check lands without a gap.

A new threshold starts at its tightest value, and review checks that `tightest` value. A new tolerance is at most the smallest tolerance of the merge base. Each threshold names the direction in which it becomes looser in `THRESHOLD_DIRECTIONS` in `thresholds.ts`. The eight matching tolerances of the line checks, from `sameBaselinePt` to `joinGapPt`, are thresholds with the direction `either`. A smaller value can fail a correct render, and a larger one can pass a wrong render. So the check refuses any change to them. After review, retire the old key in `RETIRED_LIMITS` with the reason. Then add a new key at its measured tightest value with `looser: 'either'`. A threshold without a direction fails, and a change cannot change the entry of a threshold of the merge base. A threshold goes away only when `RETIRED_LIMITS` names it with the reason. When a fixture of the merge base is gone in the change, no fixture counts as new, so a renamed fixture cannot take a gap.

The check compares the lists toolchain by toolchain. A removed list is compared with the tightest values of the lists that remain. A list for a new toolchain is compared with the tightest recorded value of each check across all lists of its font environment. It fails when it is looser anywhere. If a new toolchain measures looser, accepting it is an owner decision. Then run the check with `XLSX_FIDELITY_ACCEPT_NEW_TOOLCHAIN` set to the toolchain text of the list.

The toolchain text is `<poppler>; <chromium>`, for example `poppler 24.02.0, pdf.js 5.4.624; chromium 158.0.8127.18`. The CI workflow never sets this variable. A new gap needs a new fixture.

A new fixture that clips text needs an owner's acceptance of its clipped lines. So does a new reference PDF that adds or moves a clipped line. The line scores do not check the glyphs that a clip hides. A reference PDF is new when its hash in `fixtures/references.json` differs from the merge base. Add the lines to `fixtureClipped` of each list, with the values that each toolchain measured.

Then run the check with `XLSX_FIDELITY_ACCEPT_CLIPPED` set to the fixture name, or to a comma-separated list of fixture names. The CI workflow never sets this variable. So the **Known gaps only tighten** step of CI fails for such a pull request, and the owner decides whether to merge it. Prefer a fixture whose text fits its cells.

### Corpus checks

The corpus checks run over a private folder of workbooks on your machine. They do not run in CI, and the repository holds no corpus. Name the folder with `--dir <folder>` or the `XLSX_CORPUS_DIR` environment variable. The checks take the `.xlsx`, `.xltx`, `.xlsm`, and `.xltm` files of the folder and its subfolders, and skip hidden entries, links, and owner files such as `~$name.xlsx`. If you set both, `--dir` wins. No folder is assumed. To try a check, give it the fixtures: `--dir fixtures`.

| Command | Use |
| --- | --- |
| `bun run fidelity:census -- --dir <folder>` | Counts what the engine does with each workbook: open refusals, refused edits, saved copies that do not reopen or lose content, sheet problems, and every element of the workbook and worksheet parts, the content types, and the cell types in use. It writes `results.jsonl`, `run.json`, and `census.md` to `local/corpus/census/`. The runner stops a workbook after 120 seconds. |
| `bun run validate:xsd:corpus -- --dir <folder>` | Runs the `validate:xsd` gate, and then validates the original, saved, and edited copies of each workbook in the folder. It writes `results.jsonl` and `run.json` to `local/corpus/xsd/`. The runner stops a workbook after 300 seconds. |

`--help` lists the flags of each command. The commands check their flags and the output folder before they start.

`fidelity:census` reports and does not fail: it exits with code 0 when it finishes, also when workbooks are refused at open or stopped. Read `census.md` for the counts.

`validate:xsd:corpus` fails when a saved or edited copy has a schema message that the original does not have. It compares the messages with their values. It also fails in these cases:

- A workbook is refused at open, is stopped, or fails its task.
- A saved or edited copy does not reopen.
- No worksheet takes the edits, so the edit check was skipped. The check edits the first worksheet that takes at least one edit.
- The store refuses the row and column session for a reason other than `structure-blocked`.

Both commands run on the corpus runner, `scripts/corpus-runner/runner.ts`:

- One child process runs at a time. Each child takes a batch of 25 workbooks, so a leak stays in one batch.
- A watchdog outside the child reads the memory of its process tree. It stops the tree above 3 GiB, or when one workbook takes longer than the time of its task. Time while the machine sleeps does not count. The result line of that workbook says `stopped` with the cause. The memory that earlier workbooks of a child keep counts against the workbook that runs. So a workbook that crosses the memory cap after other workbooks runs once more, alone in a new child, before it is recorded. The workbooks of the batch that did not start run in the next child.
- Each child leads its own process group, so a stop also reaches the processes that it started. If the runner ends, a thread in the child stops the group. A child that cannot write to its runner stops the same way. The thread also stops the child at 1.25 times the memory cap, in case the runner cannot watch it: for example, a background run that the terminal paused. The result line of that workbook says `stopped` with the cause `memory`.
- The work files of a child stay in a temporary folder that the runner removes when the child ends. If the runner ends, the thread in the child removes the folder before it stops the group. If the runner cannot remove the folder after a stop signal, it names the folder: remove it yourself.
- Ctrl+Z ends the running child, removes its work folder, prints a line, and pauses the runner. `fg` continues with a new child for the workbooks that did not finish.
- `--out <folder>` takes a folder outside every git work tree or a folder that git ignores, and refuses any other. It also refuses a folder that it cannot make or write. Git ignores `local/`.
- The output holds the ids `workbook1` to `workbookN`, by sorted path, with counts and codes. It holds no file name, sheet name, or workbook text. It keeps an element or attribute name only when the ECMA-376 schemas declare it, and a content type only when ECMA-376 defines it or it is a common type. Any other name becomes `(name)`, and any other content type becomes `(other)`. Schema messages keep no value. The headers of `scripts/fidelity/census.ts` and `scripts/validate-xsd.ts` give the details.
- A run refuses to replace the results of an earlier run in its output folder. Add `--fresh` to replace them. A run with `--fresh` removes every output file of an earlier run, also one of the other command.
- After SIGHUP, SIGINT, or SIGTERM, a run stops its child, writes nothing more, and exits with code 129, 130, or 143. Its `results.jsonl` holds the results so far, and it has no `run.json`.

A census and schema validation cannot show a repair prompt. Neither command writes saved copies. The [save check](#save-check) covers repair prompts.

## Measured values

Some rules of rendering are not settled by the standard. The engine checks them against measured values in `fixtures/measured.json`: the displayed text, the fill and font colors, the font name, the bold state, and the bottom border of each case cell.

1. Add a case to `scripts/fixtures/cases.ts`. A number format case goes in a format case list, and a color or font case goes in `STYLE_CASES`.
2. Build the fixtures with `bun run fixtures:build`.
3. Record the measured values of the cells. `bun scripts/fixtures/measured.ts refs` lists the cells, and `bun scripts/fixtures/measured.ts record` reads the measured values, one tab-separated line per cell: fixture, cell, text, fill as `R,G,B`, font color as `R,G,B`, font name, bold, and bottom border as `R,G,B|style`, where style is an `ST_BorderStyle` value.
4. Make the reference PDF, record it, and run the save check, as the next sections describe.

A record holds the hash of the fixture that was measured. A test that reads a stale record fails.

## Save check

A saved workbook must open with no repair prompt, and schema validation cannot show a repair prompt. The save check opens the saved copy and the edited copy of each fixture, and records whether each copy opens with no repair prompt. The edited copy applies the `edit` that `scripts/fixtures/catalog.ts` defines for the fixture. The check also opens the copy of each feature edit (see [Feature edits](#feature-edits)).

`fixtures/save-check.json` holds one record per fixture: the SHA-256 of each copy that the check opened, and the result. The writer is deterministic, so the test in `scripts/save-check/records.test.ts` writes the copies again and fails when a hash differs, a result is not `verified`, or a record is missing. A change that alters the saved bytes therefore needs a new check. If you cannot run the check, say so in your pull request, and a maintainer runs it.

A `verified` result means that the copy opens with no repair prompt, with its sheets and its edited cell. It does not mean that formula results are current: the edited copy keeps the cached results of formulas that depend on the edited cell.

- `bun run save-check copies <dir>` writes the saved and edited copies to `<dir>`.
- `bun run save-check record <dir> <fixture> <saved-result> <edited-result>` records the result of opening the copies in `<dir>`. It refuses copies that differ from what the writer emits. A result is `verified`, `repair-prompt`, `does-not-open`, or `sheets-differ`.
- `bun run save-check record-edit <dir> <row> <result>` records the result of opening `<dir>/<row>.xlsx`, the copy of the feature edit of a support matrix row. It refuses a copy that differs from what the recipe saves.

### Feature edits

Each supported editing or save row of the support matrix proves its own round trip. `scripts/fidelity/feature-edits.ts` holds one recipe per row, with the kit in `scripts/fidelity/feature-edit-kit.ts`. A recipe edits a fixture through the public editor API, saves, reopens, and finds the edit. The run fails when the check also passes on the unedited copy, when a second save gives another package, or when the edit changes a part that it does not own. `scripts/fidelity/feature-edits.test.ts` runs every recipe in `bun run test`.

`fixtures/feature-edits.json` holds the save check record of each recipe's copy, in the form of `fixtures/save-check.json`. `bun run check:support-matrix` requires a current `verified` record for each supported editing or save row, other than `save.round-trip`, and the recipe's fixture must be one that the row cites. A row without one is `partial`, and the missing record is its shortfall.

To add a recipe, add it under the id of its row, and run `bun test scripts/fidelity/feature-edits.test.ts`. `bun scripts/fidelity/feature-edits.ts --out <dir>` runs every recipe and writes one edited copy per row as `<row>.xlsx`, with the sheet count of each copy. Open the copy with the save check, and record it with `record-edit`. If you cannot run the save check, keep the row `partial`, say so in your pull request, and a maintainer runs it. A change that alters the bytes of a recorded copy makes its record stale, and the test fails until the new copy is recorded.

The demo's sample workbook has its own record, `scripts/sample-workbook/save-check.json`, with the same form. Its edited copy makes the Bold edit of the demo save tests (`examples/vite/src/__tests__/demo-save.test.tsx`) and one value edit. A change to the writer or to `scripts/fixtures/builder.ts` can make both the fixture records and the sample record stale, so check both.

- `bun run sample:save-check copies <dir>` writes the saved and edited copies of the sample to `<dir>`.
- `bun run sample:save-check record <dir> <saved-result> <edited-result>` records the result of opening them.

## Demo sample and icons

The demo opens `examples/vite/public/sample.xlsx` at start. `scripts/sample-workbook/` builds it, and `scripts/sample-workbook/showcase.test.ts` fails when the committed file differs from its build or when its save record is stale.

- If you change `scripts/sample-workbook/` or `scripts/fixtures/builder.ts`, run `bun run sample:build`. Then run the sample save check, and commit `examples/vite/public/sample.xlsx` and `scripts/sample-workbook/save-check.json`.
- If you change the package writer (`packages/core/src/store/` or the zip writer, `packages/ooxml-core/src/zip.ts`), the sample does not change, but its saved and edited copies can. Run the sample save check, and commit `scripts/sample-workbook/save-check.json`.
- The sample save check is manual, as the fixture save check is. Make the copies, open both, and record what you see, with the `sample:save-check` commands in [Save check](#save-check).
- If you change the lockup, update `examples/vite/src/BrandLogo.tsx` and both README logos, `.github/assets/readme-logo-light.svg` and `.github/assets/readme-logo-dark.svg`. No check compares them.
- If you change `examples/vite/public/icon.svg`, run `bun run icons:build`, and commit `favicon.ico`, `icon.png`, `web-app-manifest-192x192.png`, and `web-app-manifest-512x512.png` from `examples/vite/public/`. It needs Playwright's Chromium. Run it only when `icon.svg` changes: the image bytes depend on the Chromium version.
- Run `bun run test:e2e` before you push a change to the demo, the editor chrome, input, focus, the clipboard, paint, or the sample workbook.

## Code style

The project uses [oxlint](https://oxc.rs/docs/guide/usage/linter.html) and Prettier. The lint configuration is `.oxlintrc.json`, and repository-specific rules live in `scripts/oxlint/xlsx-rules.mjs`.

```bash
bun run lint:fix
bun run format
```

The pre-commit hook formats and lints staged files. Then it runs the type check, lint, i18n, security, terminology, docs, lane boundary, fixture, schema, formatting, and lib floor gates. Last, it builds the packages and checks the API reports.

## Public terminology

Committed text uses ECMA-376 terms and never names a proprietary office app, its vendor, or a trademarked feature name. For the rule and its exceptions, see "Naming" in [Agent instructions](AGENTS.md). Run `bun run check:terminology` after you edit any text. The check reads every file that git tracks or would track, and every file name. It does not catch every name, so read your diff too.

## Shared code

`packages/ooxml-core` is a copy of PPTX Editor's package, and most of its modules come from DOCX Editor. The first lines of each file name the source path and commit. Keep a copied file equal to its source, and record each intended difference in its header. A fix in a shared module is a fix for every product: name it in the file header and the commit message.

## Write documentation

Follow the [Google developer documentation style guide](https://developers.google.com/style/highlights) for READMEs, guides, feature descriptions, and release notes.

- Lead with the task or behavior. Use active voice and address the reader as "you."
- Keep sentences to 26 words or fewer and paragraphs to six sentences. Give each sentence one idea.
- Use sentence-case headings, descriptive links, and exact identifiers in code formatting.
- State prerequisites, defaults, units, and limits near the relevant example.
- Keep implementation history in design and review records.

Every package needs a README with a title and a purpose paragraph. Public packages also need an installation command.

## Contributor license agreement

You must sign the [Contributor License Agreement](CLA.md). The CLA assistant adds signing instructions to your first pull request.

## License

Contributions to `packages/editor-api/`, `packages/pro/`, and `packages/xlsx-to-pdf/` use the EigenPal Pro License. Other code contributions use [Apache 2.0](LICENSE). Third-party material keeps its own terms.

## Report bugs

Open a [GitHub issue](https://github.com/eigenpal/xlsx-editor/issues) with:

- Steps to reproduce.
- Expected and actual behavior.
- A synthetic or sanitized `.xlsx` file, when relevant. Never attach a confidential workbook.

### Issue conventions

Every issue has a type (`Bug`, `Feature`, or `Task`), exactly one `area:*` label, and exactly one `priority:*` label. A `component:*` label is optional.

| Label | Use it for |
| --- | --- |
| `area:fidelity` | The render or the saved file differs from the reference PDF or from ECMA-376 |
| `area:performance` | Load, layout, scroll, edit, or save time, and memory use |
| `area:collaboration` | Shared editing, convergence, presence, and reconnect |
| `area:ux` | The editor chrome, keyboard, focus, and accessibility |
| `area:developer-experience` | The public API, adapters, docs, examples, and errors |
| `area:agent-ready` | The automation API and agent workflows |
| `priority:high` | Data loss, a security issue, a broken save, or a common path that fails |
| `priority:medium` | A wrong result on a common path with a workaround |
| `priority:low` | An edge case or a cosmetic difference |
| `component:<package>` | The package that holds the fix, for example `component:core` or `component:ooxml-core` |

Follow [Security policy](SECURITY.md) for vulnerabilities.
