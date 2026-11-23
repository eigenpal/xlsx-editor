#!/usr/bin/env node
/**
 * Copied from pptx-editor/scripts/check-api-changeset.mjs at e22f3f3b (EigenPal). Difference:
 * while every published package is at version 0.0.0, nothing has been released, and AGENTS.md
 * ("Releasing") says that no PR has a changeset until the first package is published. The check
 * then reports the API change and passes. `changeset version` raises the versions at the first
 * release, and from then on the check requires a changeset, as in PPTX Editor. A private package
 * (`private: true` in `scripts/lib-packages.mjs`) is never published, so a changeset cannot name
 * it: its API snapshots are checked by `api:check` and need no changeset, as an ignored package.
 *
 * A PR that changes the public API surface must carry a changeset
 * (docx-editor/scripts/check-api-changeset.mjs). Adapted from DOCX Editor at commit adf0e178.
 *
 * `docs/api/*.api.md` are the API Extractor snapshots: a diff there means a `@public`
 * symbol moved. The snapshots themselves gate nothing about VERSIONING — regenerating and
 * committing them is green — so without this check a PR can delete a public export and
 * ship it as an unversioned change. AGENTS.md's rule is "every code PR gets a changeset";
 * an API-surface change is never a test/docs/CI-only PR, so here the rule is enforceable.
 *
 * Usage: node scripts/check-api-changeset.mjs --base origin/main
 * Runs on pull_request only (the workflow guards merge commits and the
 * changeset-release/* branches, where consumed changesets legitimately disappear).
 */
import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { PACKAGES } from './lib-packages.mjs';

const baseIndex = process.argv.indexOf('--base');
const base = baseIndex > -1 ? process.argv[baseIndex + 1] : null;
if (!base) {
  console.error('usage: check-api-changeset.mjs --base <ref>');
  process.exit(2);
}

const diff = execFileSync('git', ['diff', '--name-only', `${base}...HEAD`], {
  encoding: 'utf8',
})
  .split('\n')
  .filter(Boolean);

// Ignored and private packages do not participate in Changesets releases. Keep their API
// snapshots checked without requiring a changeset that cannot version them.
const config = JSON.parse(
  readFileSync(new URL('../.changeset/config.json', import.meta.url), 'utf8')
);
const ignoredReports = PACKAGES.filter(
  (pkg) => pkg.private === true || config.ignore?.includes(pkg.name)
).map((pkg) => `docs/api/${pkg.slug}/`);
const apiChanges = diff.filter(
  (path) =>
    path.startsWith('docs/api/') && !ignoredReports.some((prefix) => path.startsWith(prefix))
);
if (apiChanges.length === 0) {
  console.log('api-changeset: no versioned package API change in this PR — nothing to require.');
  process.exit(0);
}

const unreleased = PACKAGES.every(
  (pkg) =>
    JSON.parse(readFileSync(new URL(`../${pkg.root}/package.json`, import.meta.url), 'utf8'))
      .version === '0.0.0'
);
if (unreleased) {
  console.log(
    `api-changeset: ${apiChanges.length} docs/api change(s), but no package has been released ` +
      '(every version is 0.0.0), so no changeset is required (AGENTS.md, "Releasing").'
  );
  process.exit(0);
}

const changesets = diff.filter(
  (path) => /^\.changeset\/[^/]+\.md$/.test(path) && path !== '.changeset/README.md'
);
if (changesets.length > 0) {
  console.log(
    `api-changeset: ${apiChanges.length} docs/api change(s) covered by ${changesets.length} changeset(s).`
  );
  process.exit(0);
}

console.error('This PR changes the public API surface without a changeset:');
for (const path of apiChanges.slice(0, 20)) console.error(`  - ${path}`);
if (apiChanges.length > 20) console.error(`  … and ${apiChanges.length - 20} more`);
console.error('');
console.error(
  'Run `bun changeset` (or hand-write .changeset/<name>.md) so the release pipeline versions this change.'
);
process.exit(1);
