// Copied from pptx-editor/scripts/lib-packages.mjs at e22f3f3b (EigenPal). Differences: the list
// names the XLSX Editor packages that exist: i18n, core, react, vue, and editor-api; it adds
// `PUBLISHED_PACKAGES`; and the comment says that DOCX Editor keeps its package list without
// entries, where the source says "the same list".
//
// The packages with public entry points: each publishable package, and the private
// `@xlsx-editor.dev/editor-api`, whose API is reviewed in the same snapshots. One list for
// api-extractor, notices, and release checks; `PUBLISHED_PACKAGES` leaves out the private one
// for the checks that install the packed tarballs (DOCX Editor keeps its package list, without
// entries, in `scripts/lib/packages.mjs` at adf0e178).
export const PACKAGES = [
  {
    name: '@xlsx-editor.dev/i18n',
    root: 'packages/i18n',
    slug: 'i18n',
    entries: ['index', 'en'],
  },
  {
    name: '@xlsx-editor.dev/core',
    root: 'packages/core',
    slug: 'core',
    entries: ['index', 'contracts', 'store', 'cascade', 'layout', 'automation', 'output', 'editor'],
  },
  { name: '@xlsx-editor.dev/react', root: 'packages/react', slug: 'react', entries: ['index'] },
  { name: '@xlsx-editor.dev/vue', root: 'packages/vue', slug: 'vue', entries: ['index'] },
  {
    name: '@xlsx-editor.dev/editor-api',
    root: 'packages/editor-api',
    slug: 'editor-api',
    entries: ['index'],
    private: true,
  },
];

/**
 * The packages that npm publishes: each entry of `PACKAGES` that is not `private`. The consumer
 * checks install only these. `scripts/__tests__/lib-packages.test.ts` keeps `private` equal to
 * each manifest.
 */
export const PUBLISHED_PACKAGES = PACKAGES.filter((pkg) => pkg.private !== true);
