# Repository map

This page tells you where everything is in XLSX Editor. Read [`AGENTS.md`](../AGENTS.md) first: it holds the rules. This page holds the map. Keep it current when you add, move, or remove a module.

## Start here

Read these in order before you change anything:

1. [`AGENTS.md`](../AGENTS.md): the repository entry point and shared rules.
2. This page.
3. The OpenSpec changes, each with its proposal, design, tasks, and requirements:
   - [`openspec/changes/typed-ooxml-workbook-editor/`](../openspec/changes/typed-ooxml-workbook-editor/): the production authority. It also holds row and column insert and delete.
   - [`openspec/changes/workbook-structure-editing/`](../openspec/changes/workbook-structure-editing/): sheet, clipboard, and resize edits.
   - [`openspec/changes/workbook-formulas/`](../openspec/changes/workbook-formulas/): the formula engine, planned.
   - [`openspec/specs/`](../openspec/specs/): the requirements of finished changes, such as conditional formatting and fills and styles. Their changes are in `openspec/changes/archive/`.
4. For parsing, cascade, or layout work: [`reference/quick-ref/`](../reference/quick-ref/) and the schemas in [`reference/ecma-376/`](../reference/ecma-376/).

## Top-level layout

| Path | What it holds |
| --- | --- |
| `packages/core/` | `@xlsx-editor.dev/core`, the engine. Its lanes and their import rules are in its [README](../packages/core/README.md) |
| `packages/i18n/` | `@xlsx-editor.dev/i18n`, the strings of the chrome. `en.json` is the source of truth. See its [README](../packages/i18n/README.md) |
| `packages/react/` | `@xlsx-editor.dev/react`, the React adapter: the root, the viewport, the grid, and the chrome parts. Core is its peer. See its [README](../packages/react/README.md) |
| `packages/vue/` | `@xlsx-editor.dev/vue`, the Vue 3 adapter: the twin of the React adapter, with the same props, ref members, and chrome parts. Core is its peer. See its [README](../packages/vue/README.md) |
| `examples/vite/` | The demo: a Vite page with the React editor. It opens the sample workbook (`public/sample.xlsx`) at start, and also the committed fixtures and a picked file. `bun run demo` serves it on port 5180 |
| `examples/vue/` | The Vue demo: a Vite page with the Vue editor. It opens the same sample workbook and a picked file. `bun run dev:vue` serves it on port 5181. It is for local use only |
| `packages/ooxml-core/` | `@eigenpal/ooxml-core`, the format-neutral OOXML layer. Private. See its [README](../packages/ooxml-core/README.md) for each module and where it comes from |
| `reference/ecma-376/` | ECMA-376 files, unmodified except one derived name list: the Transitional schemas (Part 4), seven Strict schemas (Part 1), the DrawingML preset geometry, and the predefined function names of Part 1, 18.17.7. See [`reference/README.md`](../reference/README.md) |
| `reference/quick-ref/` | Short SpreadsheetML and theme color summaries, written for this project |
| `openspec/changes/` | Active OpenSpec changes. The production authority is `typed-ooxml-workbook-editor` |
| `scripts/` | The repository gates and generators (see "Scripts") |
| `docs/api/` | The API Extractor reports, one for each published entry point (`<pkg-slug>/<entry>.api.md`, for example `core/store.api.md`). `bun run api:check` fails when the public API drifts from them |
| `.github/` | CI, CLA, label sync, Dependabot, issue forms, and the pull request template |
| `.changeset/` | Changesets configuration and pending release notes |
| `signatures/` | Signed CLA records, written by the CLA workflow |
| `docs/preserved-content.md` | The parts and elements that the engine keeps without support |
| `docs/site/data/support-matrix.ts` | The typed support matrix: one row for each user-visible feature, with its status and evidence. `bun run check:support-matrix` checks it |
| `docs/site/content/` | The pages of the docs site, in MDX, with the navigation in each `meta.json`. A separate site syncs `docs/site/content/` and `docs/site/data/` at build time (`docs/site/README.md`) |
| `docs/site/content/known-approximations.mdx` | The public list of known approximations. `packages/core/src/layout/approximations.ts` holds the same list in code, and a test keeps the two equal |
| `docs/site/content/workbook-limits.mdx` | The limits of a workbook read. A test keeps its values equal to the constants of `packages/ooxml-core` |
| `fixtures/` | The synthetic fixture workbooks, built by `scripts/fixtures/`, their reference PDFs with `references.json`, the save-check records of the fixtures in `save-check.json`, and of the feature edits in `feature-edits.json` |
| `local/` | Private files for one checkout. Git ignores it |

## Scripts

| Script | Command | What it does |
| --- | --- | --- |
| `scripts/test/run-parallel.mjs` | `bun run test` | Runs one `bun test` process per test file over a memory-aware worker pool. Adapted from DOCX Editor |
| `scripts/test/changed.mjs` | `bun run test:changed` | Selects the test files that import a changed file |
| `scripts/test/xlsx-suite.mjs` |  | The facts about this repository that the copied runner reads: workspace scopes and the DOM preload rule |
| `scripts/check-max-lines-caps.mjs` | part of `bun run lint` | Keeps the `max-lines` caps in `.oxlintrc.json` honest |
| `scripts/oxlint/xlsx-rules.mjs` | part of `bun run lint` | The repository lint rules: HTML sinks, global keyframe names, varargs spread in DOM-free lanes |
| `scripts/validate-i18n.mjs` | `bun run i18n:validate`, `bun run i18n:new`, `bun run i18n:status` | Keeps each locale file in step with `en.json` and generates the locale exports. Adapted from DOCX Editor |
| `scripts/check-i18n-unused.mjs` | part of `bun run i18n:validate` | Fails when no shipping code uses a key of `en.json`. Copied from DOCX Editor |
| `scripts/check-security-grep.mjs` | `bun run check:security` | The security grep of `AGENTS.md`, "Security" |
| `scripts/check-css-scope.mjs` | `bun run check:css-scope` | Checks that each rule of the chrome stylesheet selects an `xlsx-` class first, so the stylesheet styles nothing outside the editor. Copied from PPTX Editor |
| `scripts/check-docs-style.mjs` | `bun run check:docs-style` | The documentation style gate for READMEs and contributor docs, adapted from DOCX Editor |
| `scripts/check-docs-mdx.mjs` | `bun run check:docs-mdx`, part of `bun run check:docs` | Fails on a bare `{name}` in the prose of a site page, which MDX reads as JavaScript. Copied from DOCX Editor |
| `scripts/check-support-matrix.ts` | `bun run check:support-matrix`, part of `bun run check:docs` | Ties each row of the support matrix to its evidence: fixtures, save-check records, the feature edit of each editing and save row, known fidelity gaps, test files, and known approximations |
| `scripts/check-public-terminology.mjs` | `bun run check:terminology` | The naming gate of `AGENTS.md`, "Naming". Exceptions are in `scripts/public-terminology-allow.json` |
| `scripts/shared/docx-module-tests.ts` | `bun run test:docx-shared` | Runs DOCX Editor's tests of the shared modules against `packages/ooxml-core` |
| `scripts/check-lane-boundaries.mjs` | `bun run check:lane-boundaries` | Compiles each DOM-free lane of `packages/core` and `packages/ooxml-core` without the DOM library. `bun run typecheck` runs it too, as in PPTX Editor |
| `scripts/check-lockfile-workspaces.mjs` | `bun run check:lockfile-workspaces`, part of `bun run lint` | Checks that the workspace entries of `bun.lock` match each `package.json`. Copied from DOCX Editor |
| `scripts/check-parity-contract.mjs`, `scripts/check-composable-parity.mjs`, `scripts/check-adapter-css-thin.mjs`, `scripts/check-adapter-mirror.mjs`, `scripts/parity/` | `bun run check:parity` | Keeps the Vue adapter in step with the React adapter: each React prop, ref member, and export has its Vue twin in `parity/parity.contract.json`, each hook has a composable, both adapter stylesheets only import core, and each adapter file has its twin. Copied from PPTX Editor |
| `scripts/fixtures/` | `bun run fixtures:build`, `bun run fixtures:check` | Builds the synthetic fixtures in `fixtures/` from `catalog.ts`, deterministically, and checks that the committed ones match their build |
| `scripts/fidelity/` | `bun run fidelity`, `bun run fidelity:check-gaps`, `bun run fidelity:record-reference`, `bun run fidelity:census` | The fidelity harness: the harness page (`page/`) paints each fixture, Chromium takes a screenshot of it for the pixel score and prints it to PDF for the line scores, and the harness compares both with the reference PDF (drawn by Poppler, read by pdf.js) against `thresholds.ts` and its known gaps. Known gaps are kept for each font environment and raster toolchain (Poppler and Chromium versions). CI needs the list of exactly its tools; a list for a new toolchain may not be looser than the tightest recorded values at the merge base. Only the substitute font environment has measured gap lists. It prints its coverage gaps, `COVERAGE_GAPS` in `report.ts` (see `docs/site/content/known-approximations.mdx`, "Coverage gaps"). Also records the identity and number locale of each reference PDF in `fixtures/references.json`; a test fails on a stale or missing reference. `census.ts` is the local corpus census: open refusals, save and edit results, and feature counts, on the corpus runner. `feature-edits.ts` holds one recipe per support matrix row that edits or keeps content, with the kit in `feature-edit-kit.ts`; `bun scripts/fidelity/feature-edits.ts --out <dir>` writes their copies for the save check |
| `scripts/save-check/` | `bun run save-check` | Writes the saved and edited fixture copies for the local save check and records its results, also for the feature edits (`record-edit`, `feature-records.ts`); its test fails on a stale, failed, or missing record |
| `scripts/validate-xsd.ts` | `bun run validate:xsd`, `bun run validate:xsd:corpus` | Validates the modeled parts of each fixture, original, saved, and edited, and of the demo sample, original, saved, and edited, against the Transitional schemas with `xmllint`. Lists the known schema exceptions. Its corpus mode validates each corpus workbook on the corpus runner and fails on a schema message that the saved or edited copy adds |
| `scripts/sample-workbook/` | `bun run sample:build`, `bun run sample:save-check` | Builds the demo's sample workbook, `examples/vite/public/sample.xlsx`, deterministically. Its test fails when the committed file differs from its build, or when its save record (`save-check.ts`, `save-check.json`) is stale |
| `scripts/verify-demo.ts` | `bun run test:e2e` | The smoke check of the demo in Playwright Chromium: the sample opens and paints, and a typed value saves and opens again. It keeps the steps that need a real browser engine: paint and fonts, the system clipboard (`verify-demo-engine.ts`, `verify-demo-clipboard-mark.ts`), the layout of the formula bar (`verify-demo-editing.ts`), an editable host page (`verify-demo-host.ts`), the editor in a host page (`verify-demo-embedding.ts`), and a real download (`verify-demo-download.ts`). The engines, options, and skip list are in `verify-demo-browser.ts`; `--browser` adds Firefox or WebKit. Its header lists each check. Runs in the CI `demo` job |
| `scripts/bench/` | `bun run bench:edit`, `bun run bench:open:browser`, `bun run bench:edit:browser`, `bun run bench:scroll:browser`, `bun run bench:memory:browser` | Local speed benches on generated workbooks (`workbooks.ts`), with the usage in `README.md`. `edit-bench.ts` times and counts the work of each edit step in happy-dom, and `edit-bench-gates.test.ts` pins those counts inside `bun run test`. `browser-bench.ts` reports the times and the memory of the demo in Playwright Chromium, with the page probe (`memory-probe.ts`) and the heap summary (`heap-summary.ts`) of PPTX Editor. Adapted from PPTX Editor and DOCX Editor; each file header records its differences |
| `scripts/hang/` | `bun run test:hang:browser` | Opens each hostile or large workbook of `cases.ts` in the demo, in its own Playwright Chromium, under a memory and time watchdog. Its header lists the verdicts. Runs in the CI `hang` job. The corpus runner uses the same watchdog (`watchdog.ts`) |
| `scripts/corpus-runner/` | `bun run fidelity:census`, `bun run validate:xsd:corpus` | The corpus checks: they run only on your machine, and their tests run in `bun run test`. Runs a task over a folder of workbooks that `--dir` or `XLSX_CORPUS_DIR` names, one child process at a time, a new child for each batch, under the watchdog. Writes only counts, codes, feature names, content types, schema messages without their values, and opaque workbook ids to a folder that git ignores (`local/corpus/`) or that is outside every git work tree. `CONTRIBUTING.md`, "Corpus checks", gives the usage |
| `scripts/brand/build-icons.ts` | `bun run icons:build` | Builds the demo's favicons from `examples/vite/public/icon.svg`. Adapted from PPTX Editor |
| `scripts/build-core-declarations.mjs` | `bun run build:packages` (core) | Builds the declarations of `packages/core` with TypeScript 7 and bundles them with rollup-plugin-dts, with the types of `@eigenpal/ooxml-core` inlined. Adapted from DOCX Editor |
| `scripts/build-declarations.mjs`, `scripts/declaration-options.mjs`, `scripts/lib/declaration-files.mjs` | `bun run build:packages` (react, vue, and i18n) | Builds the declarations of the other packages the same way, from each `tsup.config.ts`, through `withDeclarations`. Adapted from DOCX Editor; `scripts/lib/declaration-files.mjs` is copied with no difference |
| `scripts/check-built-siblings.mjs` | Each package `build` script | Fails a package build when a sibling package that its declarations read has no current build. Adapted from DOCX Editor |
| `scripts/lib-packages.mjs` | None (a library) | The publishable packages and their entry points, which the API and package checks read. `scripts/__tests__/lib-packages.test.ts` fails when the list and the non-private packages differ. Adapted from PPTX Editor |
| `scripts/api-extractor.mjs`, `scripts/lib/api-entries.mjs` | `bun run api:check`, `bun run api:extract` | Fails on a code export with no listed entry. Compares each built entry point with its report in `docs/api/`, or rewrites the reports with `--local`. `scripts/api-extractor.mjs` is adapted from PPTX Editor; `scripts/lib/api-entries.mjs` has no sister source |
| `scripts/check-api-changeset.mjs` | `bun run check:api-changeset` | Requires a changeset when a report in `docs/api/` changes, once a package is released. Adapted from PPTX Editor |
| `scripts/check-package-artifacts.mjs` | `bun run check:package-artifacts` | Checks the built `exports` targets, the declarations, the bare imports, and what each tarball would publish: the README, the license, `THIRD_PARTY_NOTICES.md`, `licenses/`, no source maps, and the Ecma notice with each bundled ECMA-376 table. Adapted from PPTX Editor |
| `scripts/generate-third-party-notices.mjs`, `scripts/vendored-notices.mjs`, `scripts/harfbuzz-notices.mjs` | `bun run notices:generate`, `bun run notices:check` | Writes the `THIRD_PARTY_NOTICES.md` of each package from the build metafiles, the data copied into source files, and the HarfBuzz component notices in `packages/core/licenses/`. `--check` fails when a file is not current. Adapted from PPTX Editor and DOCX Editor; each file header records its differences |
| `scripts/check-shipped-advisories.mjs`, `scripts/bundled-dependencies.mjs` (the code a build inlines), `scripts/shipped-advisories-ignore.json` | `bun run check:shipped-advisories` | Fails on an advisory against a dependency that a package ships or that a consumer installs with it. Copied from PPTX Editor; each script header records its differences (`bundled-dependencies.mjs` has none), and `shipped-advisories-ignore.json` is equal to its PPTX Editor source |
| `scripts/lib/ecma-legal-comment.ts` | None (a library) | The Ecma notice as a legal comment (`/*!`), which `scripts/geometry/build-presets.ts` and `scripts/formula/build-functions.ts` write into the generated ECMA-376 tables. Copied from DOCX Editor; the file header records its differences |
| `scripts/check-shipped-shaper.mjs` | `bun run check:shipped-shaper` | Reads the built core: no bare `module` import, `harfbuzz.wasm` present, and `setHarfBuzzWasmUrl` wired into the text shaper. Adapted from DOCX Editor |
| `scripts/check-package-size.mjs`, `scripts/package-size-baseline.mjs` | `bun run check:package-size` | Measures each tarball. CI compares the sizes with the base commit of the pull request and comments on the pull request. Copied from PPTX Editor |
| `scripts/check-consumer-types.mjs` | `bun run check:consumer-types` | Compiles `packages/core/src/__tests__/consumer.test-d.ts` against the built declarations. Adapted from PPTX Editor |
| `scripts/check-consumer-install.mjs`, `scripts/lib/readme-samples.mjs` | `bun run check:consumer-install` | Installs the packed tarballs into an empty app, imports each entry under Node.js, builds with Vite and Bun, and shows the Vite build and the Vite dev server in Chromium. `scripts/check-consumer-install.mjs` is adapted from PPTX Editor; `scripts/lib/readme-samples.mjs` has no sister source. Builds the marked README samples as written and checks each grid in Chromium. Type-checks the marked React guide samples. Resolves each documented asset subpath through the installed `exports` maps, and checks that a webpack 5 build emits the text shaper and the fonts and keeps both stylesheets |
| `scripts/check-consumer-floors.mjs`, `scripts/lib/consumer-floors.mjs` | `bun run check:consumer-floors` | Fails when a stated consumer floor differs from the tested floor, and runs the packed packages at each floor |
| `scripts/lib/consumer-packs.mjs` | None (a library) | Packs each published package as `changeset publish` does, for `check:consumer-install` and `check:consumer-floors`. Adapted from PPTX Editor |
| `scripts/lib/package-doc-links.mjs` | None (a library) | Reads the links of the installed READMEs and guides. It also reads the repository links of each site page that they link. So `check:consumer-install` proves that each link reaches a file and a heading |
| `scripts/lib/package-artifact-rules.mjs` | None (a library) | The rules of `check:package-artifacts` that read one string or one manifest, so a unit test holds each rule without a build |
| `scripts/lib/notices-errors.mjs` | None (a library) | The texts of the module load, thrown error, write, and delete failures of `notices:generate` and the fix that each names, so a unit test holds each fix without a build |
| `scripts/lib/api-extractor-forgotten-exports.mjs` | None (a library) | The forgotten-export policy of `api:check` and its baseline, `scripts/api-forgotten-exports.baseline.json`. Adapted from DOCX Editor |
| `scripts/lib/api-report-config.mjs` | None (a library) | The API Extractor configuration of one published entry, for `api:check` and `api:extract`. Adapted from PPTX Editor |
| `scripts/lib/i18n-new-steps.mjs` | None (a library) | The steps that `bun run i18n:new` prints, which a test keeps equal to "Add a locale" in `packages/i18n/README.md` |
| `scripts/check-tsconfig-lib-floor.mjs` | `bun run check:tsconfig-lib-floor` | Keeps the ES `lib` of every program that compiles core's source at ES2022 or later. Adapted from DOCX Editor |
| `scripts/geometry/build-presets.ts` | `bun run geometry:build` | Generates the preset geometry table from `presetShapeDefinitions.xml` |
| `scripts/formula/build-functions.ts` | `bun run functions:build` | Generates the set of predefined function names from `reference/ecma-376/part1/predefined-functions.txt` |

## Tests

| Where | What |
| --- | --- |
| `packages/core/src/__tests__/` | The lane guards: the lane graph, every import against it, and what a browser import reaches |
| `packages/core/src/<lane>/__tests__/` | The unit tests of each lane |
| `packages/ooxml-core/src/__tests__/` | One test file for each shared module |
| `packages/react/src/__tests__/`, `packages/vue/test/` | The adapter tests. `bun run test` runs them with the happy-dom preload, and `bun run test:dom` runs only them. `packages/vue/test/cross-adapter.test.ts` mounts both adapters and compares their markup and focus step by step |
| `scripts/**/*.test.*` | Tests of the gates and the sample workbook |
| `examples/vite/src/*.test.ts` | Tests of the demo's modules: the headings switch, the open check, the unsaved-edit state, the save, and the start theme |
| `examples/vite/src/__tests__/demo-*.test.tsx` | The flows of the demo in happy-dom, over the stand-ins of `demo-harness.tsx`: the start, the focus regions, view only, the theme, the saves, the number formats, typed entry, the headings switch, the questions before unsaved edits are lost, the work that overlaps a load, the download logic, and each fixture's open, edit, save, and reopen. They run with a DOM (`scripts/test/xlsx-suite.mjs`) |
| `.github/cla/cla.test.sh` | The CLA logic, run by CI |

## Where to look

| Task | Start in |
| --- | --- |
| A file does not open, or the zip or XML read refuses it | `packages/ooxml-core/src/` (`zip.ts`, `opc-names.ts`, `xml-reader.ts`, `limits.ts`) |
| Security caps and untrusted input | `packages/ooxml-core/src/` (`limits.ts`, `sinks.ts`, `zip.ts`, `xml-reader.ts`) and `AGENTS.md`, "Security". A fix there goes to DOCX Editor and PPTX Editor too |
| Licenses and notices | `LICENSE` and `packages/core/LICENSE` (their scope notices), `scripts/vendored-notices.mjs`, `scripts/generate-third-party-notices.mjs`, `scripts/lib/notices-errors.mjs`, `scripts/harfbuzz-notices.mjs`, `scripts/check-shipped-advisories.mjs`, `scripts/bundled-dependencies.mjs`, `scripts/shipped-advisories-ignore.json`, `packages/core/licenses/`, `reference/ecma-376/NOTICE`, the Ecma legal comment of `scripts/lib/ecma-legal-comment.ts`, and the tarball checks in `scripts/check-package-artifacts.mjs` |
| A SpreadsheetML part or element | `reference/quick-ref/spreadsheetml.md`, then `reference/ecma-376/part4/transitional-schemas/sml.xsd` |
| Theme and cell colors | `reference/quick-ref/themes-colors.md` |
| The demo, its sample workbook, or its icons | `examples/vite/src/`, `scripts/sample-workbook/`, and `CONTRIBUTING.md`, "Demo sample and icons" |
| A new gate | Add its script under `scripts/`, its command to `package.json`, its step to `.github/workflows/ci.yml` and `.husky/pre-commit` (not the hook when it needs a browser or system tools, as `fidelity` and `test:e2e` do, or a packed tarball, the npm registry, or a base ref, as `check:package-artifacts`, `check:package-size`, `check:consumer-install`, `check:consumer-floors`, `check:shipped-advisories`, and `check:api-changeset` do; `check:consumer-types`, `check:shipped-shaper`, and `notices:generate` need a build in `dist/` and stay out of the hook, as `check:consumer-types` does in PPTX Editor and `check:shipped-shaper` does in DOCX Editor, its source), and its line to `AGENTS.md`, "Verify", and `CONTRIBUTING.md` |

## Before you finish a change

Follow the verification steps in [`CONTRIBUTING.md`](../CONTRIBUTING.md). After the first publish, add a changeset for code changes to a published package.
