# Site documentation source

This directory holds the public developer documentation in the DOCX Editor site structure (DOCX Editor `docs/site/README.md`, commit 68ad4957eeb9, and PPTX Editor's copy of it, commit 412d750f). A docs site syncs `content/` and `data/` at build time and renders the MDX pages. The site application is not in this repository. This file is for contributors; the site does not show it. Update these docs in the pull request that changes a user-facing feature.

## Layout

- `content/<slug>.mdx` is the page `/docs/<slug>`.
- `content/<dir>/index.mdx` is the page `/docs/<dir>`.
- `content/**/meta.json` sets the sidebar order and groups (the `pages` array; `---Label---` entries are group separators). The root `content/meta.json` lists every page by its path without `.mdx`. A folder index is the folder name, such as `customization`, or the path of its index, such as `react/index`, as in PPTX Editor.
- `data/support-matrix.ts` is the support matrix, with the rows of the `data/support-matrix-*.ts` files that it imports. The `<FeatureMatrix>`, `<FeatureSummary>`, and `<FeatureBadge>` components render it.

## Where the text lives

Each topic has one source. Do not copy a section or a sample between a page and a guide or a README; link to it:

- The site pages hold the topics that span the packages. These are installation, the quickstart, customization, the language, the support matrix, the known approximations, the workbook limits, and accessibility. On the sister sites, the known approximations and the workbook limits are sections of the fidelity page. Here they are their own pages, because tests read each page and keep it equal to the code. `approximations.test.ts` and `harness.test.ts` read the approximations and the coverage gaps, and `open-check.test.ts` reads the limits.
- The React task guides in `packages/react/docs/` stay in the package, because npm ships them with the package and `bun run check:consumer-install` type-checks their marked samples. The Vue parts behave the same, so the Vue README links to the same guides. Site pages link to a guide with its GitHub address, for example `https://github.com/eigenpal/xlsx-editor/blob/main/packages/react/docs/editing-cells.md`.
- Each adapter README holds the first steps, because npm shows the README on the package page. The first steps are the install command, the stylesheet and height step, and the first sample. They also include the open of a picked file, the edit keys, the open failure, and the default fonts. The introduction, the installation page, the quickstart, and the two adapter pages repeat them for both adapters. So a reader can start from either place. The i18n page repeats the install step, the translator sample, and the note on whole messages of the i18n README, as on the DOCX Editor site. These are the only overlaps. A longer topic stays in one place, and the other places link to it.

## Required frontmatter

```yaml
---
title: 'Installation' # Up to 60 characters; the site appends the brand.
description: 'Install the XLSX editor packages for React or Vue.' # Summarizes the page for search results.
category: 'Getting started' # Badge text and llms.txt group.
---
```

`seoTitle` (optional) is the long title for the HTML `<title>` and Open Graph tags. Keep `title` short: it is the page heading and the sidebar label. Start the body with text, not with a heading of level 1.

## MDX components

The pages use these components without imports. A site that renders this directory must provide each of them.

| Component | Props | What it renders |
| --- | --- | --- |
| `FrameworkTabs`, `Framework` | `variant`, `value` | A React and Vue switch, as on the DOCX Editor site |
| `FeatureMatrix` | `category`: a key of `SUPPORT_CATEGORY_LABELS` | One table of the rows of that category: the name (linked to `docsLink` when set), the status, and the notes |
| `FeatureSummary` | None | The number of rows, and the count of each status |
| `FeatureBadge` | `id`: a row `id` | The name and status of one row, inline |

The three feature components are the DOCX Editor site components (`src/components/docs/FeatureMatrix.tsx` in the DOCX Editor website repository). Here they read `supportFeatures`, `supportFeatureById`, and `SUPPORT_CATEGORY_LABELS` from `data/support-matrix.ts`, in place of the `word-features.ts` exports. A row of this matrix has one `status`: `supported`, `partial`, or `unsupported`. A DOCX Editor row has an editing, a rendering, and a round trip axis in its place. So the table shows one status column. The components show no `evidence`: the evidence is for `bun run check:support-matrix`. The rows are in `data/support-matrix.ts` and the `data/support-matrix-*.ts` files that it imports, so the site copies each of these files. It does not copy the test files.

Update `data/support-matrix.ts` when the status of a feature changes. Never write support claims in prose. The notes show on the page, so they describe behavior only. A row can name the page that covers it in `docsLink`, such as `/docs/workbook-limits`. `data/fidelity-page.test.ts` checks that each component tag names a category or a row that exists, and that each `docsLink` names a page.

### Framework switch

A page that shows the same example in both adapters wraps the two versions in `FrameworkTabs`, React first. Introduce the sample with text before `FrameworkTabs`:

````mdx
<FrameworkTabs>
<Framework value="react">

```tsx
<XlsxEditor document={bytes} />
```

</Framework>
<Framework value="vue">

```vue
<XlsxEditor :document="bytes" />
```

</Framework>
</FrameworkTabs>
````

Blank lines around the fences are required, or MDX treats the block as JSX. Use the switch only when both versions exist.

The switch shares a row with the copy button of the code block. If a panel starts with prose or a table, pass `variant="block"` to give the switch a separate row:

```mdx
<FrameworkTabs variant="block">
```

If each framework needs a separate introduction, keep it inside its panel and use `variant="block"`.

## Conventions

Follow the [Google developer documentation style guide](https://developers.google.com/style/highlights). Lead with the task or behavior. Use the active voice, address the reader as "you," and use sentence case for headings. Keep paragraphs short. Format identifiers as code and UI labels in bold.

- Describe only what the published packages ship.
- Link between pages with root-relative paths: `[Workbook limits](/docs/workbook-limits)`.
- End each page with a short "Next steps" or "See also" section.
- Register each page in `content/meta.json` and in its folder's `meta.json`.
- Introduce each code sample, and keep it consistent with the public API.
- MDX reads `{` and `<` as code. In prose, put braces and angle brackets in code spans.
- Use Mermaid for diagrams. Add alt text to images and descriptive text to links.

## Validate changes

Run the docs checks from the repository root:

```bash
bun run check:docs
bun test docs
```

`bun run check:docs` checks the style of the pages, refuses a bare `{name}` in prose, and ties each matrix row to its evidence. To show braces in prose, put them in a code span or escape them as `\{name\}`. `bun test docs` checks the frontmatter, the navigation, the `/docs/` links, the component tags, and the `docsLink` of each row. The checks do not build or render the site.
