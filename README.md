<p align="center">
  <a href="https://www.xlsx-editor.dev/">
    <picture>
      <source media="(prefers-color-scheme: dark)" srcset="./.github/assets/readme-logo-dark.svg" />
      <source media="(prefers-color-scheme: light)" srcset="./.github/assets/readme-logo-light.svg" />
      <img src="./.github/assets/readme-logo-light.svg" alt="XLSX Editor by EigenPal" width="320" height="90" />
    </picture>
  </a>
</p>

<p align="center">
  <a href="https://www.npmjs.com/package/@xlsx-editor.dev/core"><img src="https://img.shields.io/npm/v/@xlsx-editor.dev/core.svg?style=flat-square&color=107C41" alt="npm version" /></a>
  <a href="https://www.npmjs.com/package/@xlsx-editor.dev/core"><img src="https://img.shields.io/npm/dm/@xlsx-editor.dev/core.svg?style=flat-square&color=107C41" alt="npm downloads" /></a>
  <a href="https://github.com/eigenpal/xlsx-editor/blob/main/LICENSE"><img src="https://img.shields.io/badge/license-Apache_2.0_%2B_Pro-blue.svg?style=flat-square&color=107C41" alt="license" /></a>
  <a href="https://preview.xlsx-editor.dev/"><img src="https://img.shields.io/badge/Live_Demo-107C41?style=flat-square" alt="Demo" /></a>
</p>

> **Beta.** XLSX Editor is not ready for production use. The API can change in each release.

**XLSX Editor** is a WYSIWYG editor and rendering engine for `.xlsx` workbooks. Use it to open, edit, render, and save workbooks in the browser. It uses the ECMA-376 (Office Open XML) format. It is part of the EigenPal editor family, with [DOCX Editor](https://www.docx-editor.dev/) and [PPTX Editor](https://www.pptx-editor.dev/).

<p align="center">
  <a href="https://preview.xlsx-editor.dev/">Live demo</a> | <a href="docs/site/content/index.mdx">Documentation</a> | <a href="https://www.docx-editor.dev/"><img src="./.github/assets/docx-editor-icon.svg" alt="" width="14" height="20" align="middle" /></a> <a href="https://www.docx-editor.dev/">DOCX Editor</a> | <a href="https://www.pptx-editor.dev/"><img src="./.github/assets/pptx-editor-icon.svg" alt="" width="24" height="20" align="middle" /></a> <a href="https://www.pptx-editor.dev/">PPTX Editor</a> | <a href="https://www.xlsx-editor.dev/"><img src="./.github/assets/xlsx-editor-icon.svg" alt="" width="24" height="20" align="middle" /></a> <a href="https://www.xlsx-editor.dev/">XLSX Editor</a>
</p>

## Quick start

Install the adapter for your framework and its required engine peer:

```bash
npm install @xlsx-editor.dev/react @xlsx-editor.dev/core   # React
npm install @xlsx-editor.dev/vue @xlsx-editor.dev/core     # Vue
```

See the [React](#react) or [Vue](#vue) quick start.

For server use on Node.js, use `^20.16.0 || >=22.3.0`. Browser applications can use the packages with their build tool.

## Features

| Feature | Description |
| --- | --- |
| [Rendering](docs/site/content/xlsx-fidelity.mdx) | Show values, number formats, fonts, fills, borders, and merged cells. |
| [Cell editing](packages/react/docs/editing-cells.md) | Edit values and formulas, apply formats, and insert or delete rows and columns, with undo and redo. |
| [Save and reopen](docs/site/content/xlsx-fidelity.mdx#round-trip-contract) | Save the workbook with no loss of the content that the editor does not change. |
| [React and Vue](#react) | Use either adapter with the same engine. |

XLSX Editor is in early development. The [XLSX fidelity](docs/site/content/xlsx-fidelity.mdx) page lists what is supported.

## Packages

| Package | Description | Docs |
| --- | --- | --- |
| [`@xlsx-editor.dev/react`](https://www.npmjs.com/package/@xlsx-editor.dev/react) | React editor components and hooks. | [React adapter](packages/react/README.md) |
| [`@xlsx-editor.dev/vue`](https://www.npmjs.com/package/@xlsx-editor.dev/vue) | Vue 3 editor components and composables. | [Vue adapter](packages/vue/README.md) |
| [`@xlsx-editor.dev/core`](https://www.npmjs.com/package/@xlsx-editor.dev/core) | The engine: read, lay out, paint, edit, and save. | [Core engine](packages/core/README.md) |
| [`@xlsx-editor.dev/i18n`](https://www.npmjs.com/package/@xlsx-editor.dev/i18n) | Interface strings. | [Translations](packages/i18n/README.md) |
| [`@xlsx-editor.dev/editor-api`](packages/editor-api/README.md) | A supported subset of the Office.js workbook API for server editing. | [Workbook automation](packages/editor-api/README.md) |

`@xlsx-editor.dev/editor-api` uses the EigenPal Pro License. See the [license terms](packages/editor-api/LICENSE.md).

## React

Import the stylesheet once, pass the workbook bytes to the editor, and give the editor a height:

```tsx
import { XlsxEditor } from '@xlsx-editor.dev/react';
import '@xlsx-editor.dev/react/styles.css';

export function App({ workbook }: { workbook: Uint8Array }) {
  return <XlsxEditor document={workbook} style={{ height: '100vh' }} />;
}
```

Set the height on the editor itself. A height on a parent element is not sufficient. For Next.js, load the editor with `dynamic(..., { ssr: false })`, as [Use with Next.js](packages/react/README.md#use-with-nextjs) shows.

Full docs: [React adapter](packages/react/README.md) · [Layout and props](packages/react/docs/layout-and-props.md).

## Vue

Pass the same workbook bytes through the Vue adapter:

```vue
<script setup lang="ts">
import { XlsxEditor } from '@xlsx-editor.dev/vue';
import '@xlsx-editor.dev/vue/styles.css';

defineProps<{ workbook: Uint8Array }>();
</script>

<template>
  <XlsxEditor :document="workbook" style="height: 100vh" />
</template>
```

For Nuxt and server-side rendering, show the editor only in the browser. The editor requires browser APIs.

Full docs: [Vue adapter](packages/vue/README.md) · [Props and events](packages/vue/README.md#props-and-events).

## Documentation

- [Introduction](docs/site/content/index.mdx) and [quickstart](docs/site/content/quickstart.mdx): install the editor, open a workbook, and save it.
- [Installation](docs/site/content/installation.mdx): packages, peer dependencies, styles, bundlers, and fonts.
- [Customization](docs/site/content/customization/index.mdx): layout, toolbar, color mode, number locale, fonts, and language.
- [Editor API](docs/site/content/editor-api/index.mdx): read and write cell values from code on a server.
- [XLSX fidelity](docs/site/content/xlsx-fidelity.mdx): what the editor shows, edits, and saves.
- [Known approximations](docs/site/content/known-approximations.mdx) and [workbook limits](docs/site/content/workbook-limits.mdx): where the render differs, and the limits of a file that opens.
- [Accessibility](docs/site/content/accessibility.mdx): the keyboard, the screen reader, and the host page.

The [`react`](packages/react/README.md) and [`vue`](packages/vue/README.md) READMEs list every part.

## Development

Install the toolchain in [Contributing](CONTRIBUTING.md#prerequisites), then run these commands from the repository root:

```bash
bun install
bun run demo         # React demo: localhost:5180
bun run dev:vue      # Vue demo: localhost:5181
bun run dev:chat     # Chat example: localhost:5182
bun run test
```

The demo opens a sample workbook. You can also open a file of your own, edit it, and save it. A switcher in its top bar moves between the React demo and the Vue demo. In development, its links go to the other dev server, so run both `bun run demo` and `bun run dev:vue`. The [repository map](docs/repo-map.md) lists the packages and what to read for each task.

Examples: [Vite](examples/vite) | [Vue](examples/vue) | [Chat panel](examples/chat)

## Contributing

See [Contributing](CONTRIBUTING.md) for setup, tests, and the [Contributor License Agreement](CLA.md) (CLA). The rules for every contributor are in [Agent instructions](AGENTS.md).

## License

The repository uses [Apache 2.0](LICENSE), except for the files that the license scope notice in [`LICENSE`](LICENSE) excludes. [`@xlsx-editor.dev/editor-api`](packages/editor-api/LICENSE.md) uses the EigenPal Pro License: production use requires a commercial agreement. Contact [licensing@eigenpal.com](mailto:licensing@eigenpal.com). Third-party material keeps its own terms.

## Security and feedback

Report vulnerabilities through the [security policy](SECURITY.md). For other bugs, open a [GitHub issue](https://github.com/eigenpal/xlsx-editor/issues) with the steps to reproduce and a synthetic or sanitized workbook. Never attach confidential workbooks.

---

<p align="center">
  <a href="https://www.docx-editor.dev/"><img src="./.github/assets/docx-editor-icon.svg" alt="" width="21" height="30" align="middle" /></a> <a href="https://www.docx-editor.dev/"><b>DOCX Editor</b></a>
  &nbsp;&nbsp;&nbsp;&nbsp;
  <a href="https://www.pptx-editor.dev/"><img src="./.github/assets/pptx-editor-icon.svg" alt="" width="36" height="30" align="middle" /></a> <a href="https://www.pptx-editor.dev/"><b>PPTX Editor</b></a>
  &nbsp;&nbsp;&nbsp;&nbsp;
  <a href="https://www.xlsx-editor.dev/"><img src="./.github/assets/xlsx-editor-icon.svg" alt="" width="36" height="30" align="middle" /></a> <a href="https://www.xlsx-editor.dev/"><b>XLSX Editor</b></a>
</p>
