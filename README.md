<p align="center">
  <a href="https://xlsx-editor.dev/">
    <picture>
      <source media="(prefers-color-scheme: dark)" srcset="./.github/assets/readme-logo-dark.svg" />
      <source media="(prefers-color-scheme: light)" srcset="./.github/assets/readme-logo-light.svg" />
      <img src="./.github/assets/readme-logo-light.svg" alt="XLSX Editor by EigenPal" width="320" height="90" />
    </picture>
  </a>
</p>

# xlsx-editor.dev

An open-source editor and rendering engine for `.xlsx` workbooks, based on ECMA-376 (SpreadsheetML). It is part of the EigenPal editor family with [DOCX Editor](https://github.com/eigenpal/docx-editor) and [PPTX Editor](https://github.com/eigenpal/pptx-editor), and follows their architecture and security rules. Its packages take the same SDK shape: `@xlsx-editor.dev/core`, `@xlsx-editor.dev/react`, `@xlsx-editor.dev/vue`, and `@xlsx-editor.dev/i18n`.

To use the editor, read the [documentation](docs/site/content/index.mdx) and [XLSX fidelity](docs/site/content/xlsx-fidelity.mdx). The docs site renders the [support matrix](docs/site/data/support-matrix.ts) on that page: what the editor shows, edits, and saves.

The repository holds the engine (`packages/core`), its React and Vue adapters, the shared string catalog, the format-neutral OOXML layer (`packages/ooxml-core`), the SpreadsheetML reference material, the fixtures with their reference PDFs, and the repository tooling. The production plan is the OpenSpec change in [`openspec/changes/typed-ooxml-workbook-editor/`](openspec/changes/typed-ooxml-workbook-editor/). The OpenSpec change in [`openspec/changes/workbook-structure-editing/`](openspec/changes/workbook-structure-editing/) adds the sheet operations, the clipboard of a cell range, and the resize of columns and rows. The OpenSpec change in [`openspec/changes/workbook-fills-and-styles/`](openspec/changes/workbook-fills-and-styles/) paints pattern and gradient fills, and measures cell formats under cell style formats.

## Packages

If you build an app, install the React adapter, the engine, and React:

```bash
npm install @xlsx-editor.dev/react @xlsx-editor.dev/core react react-dom
```

| Package | What it holds |
| --- | --- |
| [`@xlsx-editor.dev/core`](packages/core/README.md) | The engine: read, layout, paint, edit, and save of a workbook in a browser. Read and save also run on a server. |
| [`@xlsx-editor.dev/react`](packages/react/README.md) | The React editor component, with its toolbar, and the hooks |
| [`@xlsx-editor.dev/vue`](packages/vue/README.md) | The Vue 3 editor component, with its toolbar, and the composables |
| [`@xlsx-editor.dev/i18n`](packages/i18n/README.md) | The UI strings of the editor, with the English catalog |

## Development

You need [Bun](https://bun.sh/) (the version in `.bun-version`) and [Node.js](https://nodejs.org/) 24.

```bash
git clone https://github.com/eigenpal/xlsx-editor.git
cd xlsx-editor
bun install
bun run test
```

## Run the demo

The demo in `examples/vite/` shows a workbook in the React editor, with its toolbar. It opens a sample workbook at start. From the **Workbook** list, you can also open the committed fixtures, or you can open a file that you pick. You can edit values and formats, and insert and delete rows and columns.

- To write the edited workbook to a file, select **Save** in the page header or in the toolbar, or press Control+S (Command+S on macOS). In a browser with a file save picker, you pick the file. In another browser, the demo downloads the file. The demo cannot confirm a download, so if the workbook has unsaved edits, it still asks before you leave the page.
- In the grid, Tab moves between the cells. To leave the grid from any cell, press Escape, then Tab. If Escape ends a pending copy or cut or cancels a cell edit, press Escape again before Tab. After that Escape, a cue over the grid says to press Tab or Shift+Tab. To move the focus among the page header, the toolbar, the grid, and the sheet tabs, press F6, or Shift+F6 to go back. To close a message, press Escape.
- **Show all headings** is on by default. It shows the row and column headings of every sheet, also in a file that hides them, so you can select whole rows and columns. It changes only the view, not the saved file.
- **View only** switches the editor to view mode, where edits are refused.
- The sun and moon buttons switch the page between the light and the dark theme. The workbook keeps its own colors. The page starts light, or in the theme that you picked before, or in the theme of a `?theme=dark` or `?theme=light` address.
- Before you open another workbook or leave the page with unsaved edits, the demo asks first.

After `bun install`, run this command:

```bash
bun run demo
```

Then open `http://localhost:5180/` in a browser.

To run the smoke check of the demo in Playwright Chromium, run `bun run test:e2e`. If you do not have Chromium for Playwright, install it first with `bunx playwright install chromium`. The flows of the demo run as happy-dom tests in `bun run test`. To rebuild the sample workbook or the demo icons, see [Demo sample and icons](CONTRIBUTING.md#demo-sample-and-icons).

For setup, checks, and pull requests, see [Contributing to xlsx-editor.dev](CONTRIBUTING.md). For the rules that every contributor and agent follows, see [Agent instructions](AGENTS.md).

## Repository layout

| Path | What it holds |
| --- | --- |
| `packages/ooxml-core/` | `@eigenpal/ooxml-core`: bounded zip and XML reading, OPC names, sinks, limits, text units, and DrawingML color and geometry math. Private |
| `packages/core/` | `@xlsx-editor.dev/core`, the engine: the store, the cascade, the layout, the paint, and the workbook editor |
| `packages/react/` | `@xlsx-editor.dev/react`: the React adapter of the workbook editor |
| `packages/vue/` | `@xlsx-editor.dev/vue`: the Vue 3 adapter of the workbook editor, the twin of the React adapter |
| `packages/i18n/` | `@xlsx-editor.dev/i18n`: the shared string catalog |
| `examples/vite/` | The demo that `bun run demo` starts |
| `examples/vue/` | The Vue demo that `bun run dev:vue` starts |
| `reference/` | ECMA-376 schemas and preset geometry, and quick references for SpreadsheetML and theme colors |
| `openspec/` | The active OpenSpec changes and their requirements |
| `fixtures/` | The synthetic workbooks that scripts build, their reference PDFs, and the save-check records |
| `scripts/` | Repository gates and tools: the test runner, the line caps, lint rules, the security grep, the terminology check, the adapter parity checks, the fixture builder, schema validation, the fidelity harness, the save check, the support-matrix check, the sample workbook, and the demo check |
| `docs/` | The repository map, the preserved content, the API reports, and the docs site source (`docs/site/`). The site pages include the [known approximations](docs/site/content/known-approximations.mdx) and the [workbook limits](docs/site/content/workbook-limits.mdx). |

## Issue conventions

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

## Contributing

Read [Contributing to xlsx-editor.dev](CONTRIBUTING.md). You must sign the [Contributor License Agreement](CLA.md).

## License

The repository uses [Apache 2.0](LICENSE), except for the material that the license scope notice in [`LICENSE`](LICENSE) excludes. The core package bundles ECMA-376 tables, and the Ecma copyright notice covers them. The files under `reference/ecma-376/` are Ecma International material; see [`reference/README.md`](reference/README.md). A package directory with its own `LICENSE.md` file uses the terms in that file. Third-party material keeps its own terms.

## Security

To report a vulnerability, see [Security policy](SECURITY.md).
