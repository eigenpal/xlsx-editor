// Registers happy-dom for a `bun test` run that names DOM test files only, so
// `bun test packages/react/src/__tests__/<file>` runs one adapter file with no flag, and
// `bun test examples/vite/src/__tests__/<file>` runs one test of the demo's flows. Loaded for
// every run by `bunfig.toml`. A run that names any other path, or none, gets no DOM: the engine
// tests stay without one (`preload-dom.ts`), and `bun run test` and `bun run test:dom` pass
// the DOM preload to the DOM test files themselves (`xlsx-suite.mjs`).
export {};

const ADAPTER =
  /(^|[\\/])(packages[\\/](react|vue)|examples[\\/]vite[\\/]src[\\/]__tests__)([\\/]|$)/;
const paths = process.argv.slice(1).filter((a) => !a.startsWith('-'));
if (paths.length > 0 && paths.every((a) => ADAPTER.test(a))) await import('./preload-dom.ts');
