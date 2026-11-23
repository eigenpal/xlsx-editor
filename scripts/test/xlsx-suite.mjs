// XLSX Editor only (after pptx-editor/scripts/test/pptx-suite.mjs at 8dc88c4d): the facts about
// this repository that the runner copied from DOCX Editor (`run-parallel.mjs`, `changed.mjs`)
// needs. Kept apart, so the copied files stay equal to their source apart from the lines that
// read these values.

/**
 * The npm scopes of the workspace packages. Each maps `<scope><name>` to `packages/<name>`.
 * `@eigenpal/ooxml-core` lives in `packages/ooxml-core`, so a change there selects the core
 * tests that import it.
 */
export const WORKSPACE_SCOPES = ['@xlsx-editor.dev/', '@eigenpal/'];

/**
 * The tests that run with a global DOM: the adapter tests, and the tests of the demo's flows,
 * which render the demo's React page (`examples/vite/src/__tests__/`).
 */
export const DOM_TEST_ROOTS = ['packages/react/', 'packages/vue/', 'examples/vite/src/__tests__/'];

/** Registers happy-dom before a DOM test module is evaluated. */
export const DOM_PRELOAD = './scripts/test/preload-dom.ts';

/**
 * The extra `bun test` arguments for one file.
 *
 * DOCX Editor preloads happy-dom for the whole suite from `bunfig.toml`. XLSX Editor does not,
 * as PPTX Editor does not: every test outside the adapters runs with no global `document`, so
 * it takes the same code paths as a server or a worker, and a DOM-only dependency in a DOM-free
 * lane fails there.
 * A test that needs a DOM makes its own happy-dom window. Only the files under
 * `DOM_TEST_ROOTS` preload it.
 */
export function bunArgumentsFor(file) {
  return DOM_TEST_ROOTS.some((root) => file.startsWith(root)) ? ['--preload', DOM_PRELOAD] : [];
}
