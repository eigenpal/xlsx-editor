# xlsx-editor.dev

WYSIWYG editor and rendering engine for XLSX, based on ECMA-376 (SpreadsheetML). One engine parses, edits, lays out, paints, and saves. React and Vue are adapters over it.

XLSX Editor is the third product of the EigenPal editor family. DOCX Editor (`eigenpal/docx-editor`) is the golden standard for architecture, gates, security, docs, and SDK shape. PPTX Editor (`eigenpal/pptx-editor`) is the second product. When this file does not settle a question, do it the DOCX Editor way, and take PPTX Editor's adaptation where it solved the same problem.

## Start here

**Before any work, have the standard in your checkout.** Check that `reference/ecma-376/part1/text/index.tsv` exists. If it does not, run `bun run reference:fetch` first. Git ignores these files, so each new clone and each new git worktree starts without them. Do not start a task, a review, or an investigation without them.

When you start work on an issue, find the ECMA-376 subclauses and schema types that it touches, and read them in the local text. Post them in a comment on the issue before you change code, and cite them in the code comments and the pull request.

1. [Repository map](docs/repo-map.md): package locations and task-specific reading.
2. [Contribution guide](CONTRIBUTING.md): setup, tests, fidelity, save checks, and pull requests.
3. [Production authority](openspec/changes/typed-ooxml-workbook-editor/): read before changing behavior. Superseded proposals are not requirements.
4. [OOXML references](reference/README.md): how to find a subclause in the local text of ECMA-376 and in the schemas. When a quick reference and the schema disagree, the schema wins.

Read `local/AGENTS.md` if it exists. It can add steps for one machine. It cannot relax a rule in this file. The `local/` directory is in `.gitignore`. Never copy its contents into committed files or pull requests.

## Communication

Write all replies in ASD-STE100 Simplified Technical English.

- Use the active voice and simple tenses.
- Keep sentences to 20 words or fewer, with one idea each.
- Use the same word for the same idea each time.
- Do not use idioms, slang, or jargon.
- Keep paragraphs to 6 sentences or fewer.
- Keep file paths, function names, flags, and numbers exact.

In user-facing text, always use the name `EigenPal Pro License`. Do not rename the legal license files or `LicenseRef-EigenPal-Pro-Evaluation-1.0` metadata.

## Priorities

Work in this order. Do not start a lower item while a higher one has a known defect in the same area.

1. **Round-trip safety.** Open, edit, save, and reopen lose nothing. A saved file opens with no repair prompt.
2. **Measured fidelity.** The render matches the reference PDF within the thresholds.
3. **Editing.** Each edit writes valid SpreadsheetML.
4. **Collaboration.** It starts only after core editing passes its hands-on checkpoint.
5. **Chrome.** The toolbar, dialogs, and panes.

## Engineering rules

- Preserve unknown content through open, edit, save, and reopen.
- Write workbook changes only through `WorkbookDocumentStore.transact` over typed ops. A transaction that leaves a touched part invalid is refused.
- Keep `store`, `cascade`, `layout`, and `automation` free of DOM dependencies. A lane that takes a new dependency edits `packages/core/src/__tests__/core-lane-graph.ts`.
- **Values.** Show the value that the file stores. Never evaluate a formula on open. Never present a value as current when an edit can have changed it.
- **Layout.** DOM-free, with an injected text measurer, in points. Column widths come from the maximum digit width (ECMA-376 Part 1, 18.3.1.13). Layout and paint cost depend on the visible cells, never on the sheet size.
- **Editing projections.** A cell editor emits the smallest ops that explain the edit, or refuses. Never drop an edit silently.
- **Host window.** Resolve the document and window from a node that the editor owns, at use time (`editor/host-realm.ts`).
- **Chrome registry.** `CHROME_GROUPS` is the toolbar taxonomy. `ChromeSlotId` is public API. `toolbarCommandState` is the only source of enabled state.
- Keep React and Vue behavior aligned. Put shared logic in `core`. `scripts/parity/parity.contract.json` lists the paired members.
- Keep feature claims in the [support matrix](docs/site/data/support-matrix.ts), with evidence.

## Packages

| Package | What | Status |
| --- | --- | --- |
| `ooxml-core` | The format-neutral OOXML layer, copied from PPTX Editor | private, inlined into `core` |
| `core` | **The engine**: `store/`, `cascade/`, `layout/`, `output/`, `editor/`, `contracts/`, `automation/` | publishable |
| `react` | The React adapter: provider and hooks, no editing state | publishable |
| `vue` | The Vue 3 adapter twin, parity-gated against `react` | publishable |
| `i18n` | Shared strings | publishable |
| `editor-api` | `XlsxEditor` automation object model. EigenPal Pro License | private |
| `pro` | Collaboration and review modules. EigenPal Pro License | not yet created |
| `xlsx-to-pdf` | XLSX to PDF on Node.js. EigenPal Pro License | not yet created |

**The engine must resolve to ONE copy.** It holds the text shaper and identity-keyed layout caches. So `core` is a **peer** of `react`, `vue`, and `pro`, never a regular `dependency`.

## Public API

Public workbook automation APIs keep the Office.js shape: `Workbook`, `Worksheet`, `Range`, `Table`, and their collections, with their member names, signatures, enum values, and `load()` / `context.sync()` behavior. The entry point is `XlsxEditor.createServer(bytes)` and `XlsxEditor.createBrowser(editor)`, then `runtime.run(async (context) => { ... })`.

- Use existing Office.js members. Do not add competing convenience methods.
- Refuse unsupported modes and operations explicitly. Never approximate them silently.
- Record supported subsets and differences in `packages/editor-api/compat/manifest.json` and the public documentation.
- Do not claim full Office.js compatibility from subset conformance.

API Extractor snapshots are in `docs/api/<pkg-slug>/<entry>.api.md`. On drift, run `bun run build:packages`, `bun run api:extract`, and commit. Vue composables declare a named `Use<Name>Return` interface. Exported names describe capabilities, never engine internals (no "tree").

## Fidelity

- Measure rendering against the committed reference PDFs. Never loosen a threshold to pass a check. Thresholds only tighten.
- Treat a missing reference, a stale reference, or a skipped check as a gap, never a pass.
- Use the same layout and paint pipeline in the editor and the fidelity harness.
- Every rendering feature lands with a synthetic fixture, built by a script, and its reference PDF.
- Name every approximation in `packages/core/src/layout/approximations.ts` and in the public known approximations. A test keeps the two lists equal.
- A save check records any repair prompt for each saved and edited fixture copy. Schema validation alone cannot show one.

## Evidence and fonts

- Cite ECMA-376 for each behavior that you add or change: the element, its schema type (`CT_*`, `ST_*`), and the subclause, in the code comment and in the pull request body. Read the subclause in the local text before you write the code.
- Never cite a subclause number from memory. Find it in the local index, and run `bun run reference:check-citations` before you push. A wrong citation is worse than none.
- Where the standard is silent, cite committed evidence: a reference PDF, a record in `fixtures/` (`measured.json`, `save-check.json`), or a fixture.
- Use only open fonts in fixtures. `scripts/fidelity/references.test.ts` refuses any other embedded face.
- Respect font embedding permissions.
- Never commit private workbooks, their contents, or reference fonts.

## Security

**Every value from an XLSX, pasted HTML, or an embedded part is attacker-controlled.** Sanitize at the bounded parse boundary, never at render time.

- **No HTML from strings.** No `innerHTML`, `outerHTML`, `insertAdjacentHTML`, or `document.write` on file-derived values.
- **URLs through `sanitizeHref`.** Allowlist `http`, `https`, `mailto`, `tel`, and `ftp`. Strip tab, LF, and CR.
- **CSS.** Escape file-derived `@font-face` family names and inline `style`.
- **XML.** No DTD, no external entities, no nested entity expansion.
- **Zip.** Cap the decompression ratio and size. Reject paths with `..` or a leading `/`.
- **No zero-click external fetch.** Never load external relationships, external workbook links, data connections, query tables, or remote images automatically.
- **Resource limits.** Cap recursion depth and element counts. Never feed a file-supplied number into an allocation, `.repeat()`, or a loop bound. Iterate the cells that exist, never the grid that a reference names.
- **Escape on save.** `escapeXml` every attacker-derived string.
- **Prototype pollution.** Use `Map` or null-prototype records for file-keyed data. Guard against `__proto__`, `constructor`, and `prototype`.
- **Formulas, macros, and embedded content.** Never evaluate a formula on open. Never execute `HYPERLINK`, DDE, web queries, macros, OLE objects, or embedded controls. Keep them inert.
- **Formula injection on export.** A CSV or TSV export escapes a cell value that starts with `=`, `+`, `-`, `@`, tab, or CR.

`bun run check:security` runs the sink grep (`scripts/check-security-grep.mjs`). Fix sibling sinks when you fix one. A fix in `packages/ooxml-core` is a security fix for every product: name it in the commit message, and send it to the DOCX Editor and PPTX Editor teams.

## Collaboration verification

Collaboration is its own OpenSpec change. It starts after core editing passes its hands-on checkpoint. Until then, every edit is a typed op through `WorkbookDocumentStore.transact`, the tree holds no session state, and no module keeps document state in a global. From the collaboration change on, verify each feature through two editor instances on real sessions, as DOCX Editor does.

## Verify

Run the checks in [the contribution guide](CONTRIBUTING.md) that apply to the change. Report failures and skipped checks accurately. A bypassed failing gate is not a passing gate.

```bash
bun run typecheck
bun run lint
bun run format:check
bun run test
bun run check:security
bun run check:css-scope
bun run check:terminology
bun run check:license-headers
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
bun run test:e2e:chat
bun run test:hang:browser
```

Also run `bun run reference:check-citations` when you add or change a citation. It reads the local index, so CI does not run it. Add each new gate to this list in the change that adds its script. Run `openspec validate <change> --strict` when you change `openspec/`. Run `bun run format` before you push.

## i18n

`packages/i18n/en.json` is the source of truth. Other locales mirror its shape; `null` falls back to English. Never hardcode user-facing English in shipped library components. Files under `examples/` must not add keys to `packages/i18n`.

## Docs site

Author pages in `docs/site/content/` (MDX). Keep feature status in `docs/site/data/support-matrix.ts`, never in prose. A change to user-visible behavior updates its matrix rows and pages in the same pull request. Follow the [Google developer documentation style guide](https://developers.google.com/style/highlights). Diagrams are mermaid.

## Releasing

Every code change to a published package gets a changeset: the exact package name, `patch` by default, `minor` for additive public API, `major` for breaks, and a one-sentence summary for consumers. The packages are in prerelease mode (`.changeset/pre.json`), so each release is a beta; see `CONTRIBUTING.md`, "Releases". Never push the `chore: release` commit by hand or hand-edit `CHANGELOG.md` or `package.json#version`. Each publishable package ships its generated `THIRD_PARTY_NOTICES.md`.

## Naming

Use ECMA-376 terminology for spreadsheet features: workbook, worksheet, shared string table, cell format (`x:xf`), number format (`x:numFmt`), defined name, conditional formatting, data validation, table, and pivot table.

Trademark policy: do not use third-party product names in public text. Public text includes code, comments, tests, fixtures, docs, commit messages, pull requests, issues, and labels.

Preserve format identifiers, namespace URIs, content types, npm package names, Office.js identifiers, font family names stored in workbooks, and required legal attribution. `bun run check:terminology` enforces what it can match; `scripts/public-terminology-allow.json` lists each exact exception. A pass does not prove the rule. Read your diff.

## Conventions

- **History.** The git history becomes public. Private plans, investigation notes, and review records stay in `local/`. Fold review rounds into the commit that they fix; do not file them as separate commits or issues.
- **PRs.** Short title with a conventional-commit prefix. The body is the minimum the diff does not show, in the sections of `.github/pull_request_template.md`. A PR that resolves an issue ends its body with `Fixes #N`. No `@`-mentions, file lists, tooling footers, or emojis.
- **Commits.** Conventional commits: `fix(core): ...`, `feat(react): ...`, `docs: ...`. A fix that closes an issue ends its subject with `(fixes #N)`.
- **Issues.** Exactly one `area:*` label, one `priority:*` label, and a native type: `gh issue create --type Bug|Feature|Task --label "area:...,priority:..."`. The taxonomy is in `CONTRIBUTING.md`, "Issue conventions".
- **Styling.** All chrome CSS and color tokens live in the core stylesheet. Use `--doc-*` tokens, never hex or rgba. The workbook canvas is not themed.
- **Shared code.** Code taken from DOCX Editor or PPTX Editor keeps a header with its source path and commit, and lists each intended difference.
- **TypeScript.** `typecheck` runs `bun check`. The `typescript` dependency stays on version 6. Declaration builds use `typescript7`.
- **ESM only.** No `require()`.
- **Focus stealing.** Chrome `mousedown` must `preventDefault()`, except on `INPUT`, `SELECT`, and `TEXTAREA`.
- **Icons.** Inline SVG (Material Symbol paths), not a font.
