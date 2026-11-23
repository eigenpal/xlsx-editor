// Copied from pptx-editor/scripts/api-extractor.mjs at e22f3f3b (EigenPal). Differences: paths
// resolve from the repository root, as in DOCX Editor, so the `api:check` script of each package
// works from the package folder too; the drift message names the build step, and only drift prints
// it, while another failure prints its error and warning counts; and a code subpath
// of `exports` with no entry in `scripts/lib-packages.mjs` fails, because it would have no report,
// and a report in `docs/api/` that no published entry produces fails. Each report holds the types
// that a public signature reaches and the entry does not export, so a change to their shape is drift.
// Forgotten exports follow DOCX Editor's baseline: a new one fails, and only
// `--local --record-forgotten-exports` rewrites `scripts/api-forgotten-exports.baseline.json`.
// The TypeScript version notice prints once, not once for each entry, and a pass prints one line.
// Public API snapshots, as in docx-editor/scripts/api-extractor.mjs: one report per published entry
// point in docs/api/<slug>/<entry>.api.md. `--local` rewrites the snapshots; without it, any drift
// fails. Packages must be built first (`bun run build:packages`). DOCX Editor commit adf0e178.
import { Extractor, ExtractorConfig } from '@microsoft/api-extractor';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  orphanReports,
  packageArgumentProblem,
  reportsBelow,
  unlistedEntries,
} from './lib/api-entries.mjs';
import {
  evaluateForgottenExportPolicy,
  nextForgottenExportBaseline,
} from './lib/api-extractor-forgotten-exports.mjs';
import { apiReportConfigObject } from './lib/api-report-config.mjs';
import { PACKAGES } from './lib-packages.mjs';

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const local = process.argv.includes('--local');
const record = process.argv.includes('--record-forgotten-exports');
const argumentProblem = packageArgumentProblem(process.argv.slice(2), PACKAGES);
if (argumentProblem) {
  console.error(argumentProblem);
  process.exit(1);
}
if (record && (!local || process.argv.includes('--package'))) {
  console.error('--record-forgotten-exports needs --local, and it records every package at once.');
  process.exit(1);
}
const baselinePath = resolve(repoRoot, 'scripts/api-forgotten-exports.baseline.json');
const baselines = existsSync(baselinePath) ? JSON.parse(readFileSync(baselinePath, 'utf8')) : {};
const recorded = {};
const only = process.argv.includes('--package')
  ? process.argv[process.argv.indexOf('--package') + 1]
  : null;
let failed = false;
let preamble = false;
const unbuilt = [];
for (const pkg of PACKAGES) {
  if (only && pkg.name !== only) continue;
  const root = resolve(repoRoot, pkg.root);
  const manifest = JSON.parse(readFileSync(resolve(root, 'package.json'), 'utf8'));
  for (const subpath of unlistedEntries(pkg, manifest.exports)) {
    console.error(`${pkg.name}: add the export ${subpath} to scripts/lib-packages.mjs.`);
    failed = true;
  }
  // With no build, the forgotten-export messages are missing too, so only the build is the fix.
  if (pkg.entries.some((entry) => !existsSync(resolve(root, 'dist', `${entry}.d.ts`)))) {
    unbuilt.push(pkg.name);
    failed = true;
    continue;
  }
  const forgotten = [];
  for (const entry of pkg.entries) {
    const dts = resolve(root, 'dist', `${entry}.d.ts`);
    const reportFolder = resolve(repoRoot, 'docs/api', pkg.slug);
    mkdirSync(reportFolder, { recursive: true });
    const config = ExtractorConfig.prepare({
      configObject: apiReportConfigObject({
        root,
        dts,
        reportFileName: `${entry}.api.md`,
        reportFolder,
        reportTempFolder: resolve(repoRoot, '.cache/api-temp', pkg.slug),
      }),
      configObjectFullPath: resolve(root, 'api-extractor.json'),
      packageJsonFullPath: resolve(root, 'package.json'),
    });
    const result = Extractor.invoke(config, {
      localBuild: local,
      showVerboseMessages: false,
      messageCallback: (message) => {
        if (message.messageId === 'console-preamble') {
          message.handled = preamble;
          preamble = true;
          return;
        }
        if (message.messageId !== 'ae-forgotten-export') return;
        forgotten.push(`${entry}: ${message.text}`);
        message.handled = true;
      },
    });
    if (!result.succeeded && result.apiReportChanged) {
      console.error(
        `${pkg.name}/${entry}: API report drift. Run bun run build:packages, then bun run api:extract, and commit.`
      );
      failed = true;
    } else if (!result.succeeded) {
      console.error(
        `${pkg.name}/${entry}: API Extractor failed with ${result.errorCount} errors and ${result.warningCount} warnings. The messages of this entry name each cause.`
      );
      failed = true;
    }
  }
  const baseline = baselines[pkg.name] ?? {};
  const policy = evaluateForgottenExportPolicy({
    packageName: pkg.name,
    isLocal: local,
    allowlist: {},
    baseline,
    messages: forgotten,
  });
  if (record) {
    const next = nextForgottenExportBaseline({ baseline, policy, record: true });
    if (Object.keys(next).length > 0) recorded[pkg.name] = next;
    continue;
  }
  for (const { entry, symbol } of policy.unallowlisted) {
    console.error(
      `${pkg.name}/${entry}: a public signature reaches ${symbol}, and this entry does not export it. Export it from this entry, or review it and run node scripts/api-extractor.mjs --local --record-forgotten-exports.`
    );
    failed = true;
  }
  for (const { entry, symbol } of policy.staleBaseline) {
    console.error(
      `${pkg.name}/${entry}: the forgotten-export baseline lists ${symbol}, which no longer warns. Run node scripts/api-extractor.mjs --local --record-forgotten-exports.`
    );
    failed = true;
  }
}
if (unbuilt.length > 0)
  console.error(
    `${unbuilt.join(', ')}: no declarations in dist/. Run bun run build:packages first.`
  );
if (record && !failed) writeFileSync(baselinePath, `${JSON.stringify(recorded, null, 2)}\n`);
if (!only) {
  const apiRoot = resolve(repoRoot, 'docs/api');
  const reports = reportsBelow(apiRoot);
  for (const report of orphanReports(PACKAGES, reports)) {
    console.error(`docs/api/${report}: no published entry produces this report. Delete it.`);
    failed = true;
  }
}
if (!failed)
  console.log(local ? 'API reports written from the build.' : 'API reports match the build.');
process.exit(failed ? 1 : 0);
