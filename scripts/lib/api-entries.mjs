// The code entry points of a package's `exports` map that the list in `scripts/lib-packages.mjs`
// leaves out. `api:check` fails on each one, so a published entry of a package in that list cannot
// ship without an API report.

import { readdirSync } from 'node:fs';
import { join } from 'node:path';

const ASSET_SUBPATH = /\.(?:css|json|wasm)$/;

/**
 * @param {{ entries: readonly string[] }} pkg The package record of `scripts/lib-packages.mjs`.
 * @param {Record<string, unknown>} exportsMap The `exports` field of its `package.json`.
 * @returns {string[]} The `exports` subpaths that have no entry.
 */
export function unlistedEntries(pkg, exportsMap) {
  return Object.keys(exportsMap).filter((subpath) => {
    if (ASSET_SUBPATH.test(subpath)) return false;
    return !pkg.entries.includes(subpath === '.' ? 'index' : subpath.slice(2));
  });
}

/**
 * @param {readonly string[]} args The command-line arguments.
 * @param {readonly { name: string }[]} packages The package records.
 * @returns {string | null} Why the `--package` value selects no package, or `null`.
 */
export function packageArgumentProblem(args, packages) {
  const at = args.indexOf('--package');
  if (at === -1) return null;
  const value = args[at + 1];
  const valid = `Valid values: ${packages.map(({ name }) => name).join(', ')}.`;
  if (!value || value.startsWith('--')) return `--package requires a value. ${valid}`;
  if (packages.some(({ name }) => name === value)) return null;
  return `Unknown package '${value}'. ${valid}`;
}

/**
 * @param {string} apiRoot The `docs/api` directory.
 * @returns {Map<string, string[]>} The file names of each subdirectory. The key `''` holds the
 * files directly in `apiRoot`.
 */
export function reportsBelow(apiRoot) {
  const listing = readdirSync(apiRoot, { withFileTypes: true });
  const reports = new Map([
    ['', listing.filter((entry) => !entry.isDirectory()).map((entry) => entry.name)],
  ]);
  for (const entry of listing)
    if (entry.isDirectory()) reports.set(entry.name, readdirSync(join(apiRoot, entry.name)));
  return reports;
}

/**
 * @param {readonly { slug: string, entries: readonly string[] }[]} packages The package records.
 * @param {ReadonlyMap<string, readonly string[]>} reports The file names in each `docs/api/<slug>`.
 * The key `''` holds the files directly in `docs/api/`.
 * @returns {string[]} Each report, as `<slug>/<file>` or `<file>`, that no published entry
 * produces.
 */
export function orphanReports(packages, reports) {
  const expected = new Set(
    packages.flatMap(({ slug, entries }) => entries.map((entry) => `${slug}/${entry}.api.md`))
  );
  return [...reports].flatMap(([slug, files]) =>
    files
      .map((file) => (slug === '' ? file : `${slug}/${file}`))
      .filter((report) => !expected.has(report))
  );
}
