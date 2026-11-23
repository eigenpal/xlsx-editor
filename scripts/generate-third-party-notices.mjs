// Copied from pptx-editor/scripts/generate-third-party-notices.mjs at d4502552 (EigenPal).
// Differences: the HarfBuzz notice check runs for `@xlsx-editor.dev/core`. As the DOCX Editor
// script does (docx-editor/scripts/generate-third-party-notices.mjs at 35a29d6d): a missing dist
// and a `--check` failure name the command that fixes them, a `--check` failure names the state
// of the file, a passing `--check` prints one line, and a write prints one line for each package.
// Also as the DOCX Editor script does: a failure is all or nothing, so the run writes no file
// until every package passes and deletes the notices on a failure; the license text is each
// `LICENSE`, `LICENCE`, `COPYING`, or `UNLICENSE` file (with a suffix such as `-MIT` and a text
// extension) and each `NOTICE` file; and every path resolves from the repository root. Like
// PPTX Editor, it leaves out the DOCX Editor check that each shipped bundle is in a metafile
// (each package has one tsup config, so no metafile can replace another), the `files` check
// (`check:package-artifacts` reads the tarball), the DOCX Editor preamble and package table, its
// `LicenseRef-` paragraph for Pro packages (none exists yet), its `--package=` option, its write
// only on a change, and its `### filename` heading for a single license file. Neither sister has
// these, which this repository adds so that every failure names its fix and reaches the delete:
// an error that a package or a notice module throws (such as a HarfBuzz notice failure) is
// reported with its fix and fails the whole run; a metafile with no `inputs` object is refused;
// a failed write deletes the notices; and the success lines print only after every write. The
// texts of the module load, thrown error, write, and delete failures are in
// `scripts/lib/notices-errors.mjs`, with unit tests.
// `scripts/__tests__/generate-third-party-notices.test.ts` runs a module load failure, a thrown
// package error, an invalid metafile, a failed package after one that renders, a failed write, a
// failed delete, a `--check` with no build and with a stale or missing notice, a dist with no
// metafile, an inlined package with no license text, the license and NOTICE texts of an inlined
// package (also from the Bun store path), one license text with no heading, and a passing run
// with the exact text of each notice.
//
// The PPTX Editor header, unchanged (it names the DOCX Editor commit that PPTX Editor copied):
// THIRD_PARTY_NOTICES.md per published package, from the bundler metafile: every package that
// tsup inlined into dist must ship its license text (docx-editor/scripts/generate-third-party-notices.mjs).
// `--check` compares against the current dist. DOCX Editor commit adf0e178.
import { readFileSync, writeFileSync, existsSync, readdirSync, rmSync, statSync } from 'node:fs';
import { dirname, join, relative, resolve } from 'node:path';
import { createRequire } from 'node:module';
import { PACKAGES } from './lib-packages.mjs';
import { deleteError, moduleLoadError, packageError, writeError } from './lib/notices-errors.mjs';

const check = process.argv.includes('--check');
// Paths resolve from the repository root, as in DOCX Editor, so a run from another directory
// reads and deletes only this checkout's files.
const repoRoot = resolve(import.meta.dirname, '..');
const noticesOf = (pkg) => resolve(repoRoot, pkg.root, 'THIRD_PARTY_NOTICES.md');
/** Delete every notice, so a failed run leaves no stale or partial file behind. */
function deleteNotices() {
  if (check) return;
  for (const pkg of PACKAGES) {
    try {
      rmSync(noticesOf(pkg), { force: true });
    } catch (error) {
      console.error(deleteError(relative(repoRoot, noticesOf(pkg)), error));
    }
  }
}
// These modules read tracked files when they load; a failure there also fails the whole run.
/** Import `specifier`, or report the module and the files that it reads, delete, and exit. */
async function load(specifier, file, reads) {
  try {
    return await import(specifier);
  } catch (error) {
    console.error(moduleLoadError(file, reads, error));
    deleteNotices();
    process.exit(1);
  }
}
// `vendored-notices.mjs` imports `harfbuzz-notices.mjs`, so that module loads first: a failure
// in it, or in the manifest that it reads, names those files.
const { validateHarfBuzzNotices } = await load(
  './harfbuzz-notices.mjs',
  'scripts/harfbuzz-notices.mjs',
  ['packages/core/licenses/harfbuzz-manifest.json']
);
const { VENDORED } = await load('./vendored-notices.mjs', 'scripts/vendored-notices.mjs', [
  'reference/ecma-376/NOTICE',
]);
// From docx-editor/scripts/generate-third-party-notices.mjs at 35a29d6d.
const LICENSE_FILE =
  /^(?:licen[cs]e|copying|unlicen[cs]e)(?:[-_][a-z0-9._-]+)?(?:\.(?:md|txt|rst))?$/i;
const NOTICE_FILE = /^notice(?:\.(?:md|txt|rst))?$/i;
/** The non-empty texts of the files in `dir` whose names match `pattern`, by file name. */
function licenseTextsIn(dir, pattern) {
  return readdirSync(dir)
    .filter((entry) => pattern.test(entry) && statSync(join(dir, entry)).isFile())
    .sort()
    .map((filename) => ({ filename, text: readFileSync(join(dir, filename), 'utf8').trim() }))
    .filter(({ text }) => text.length > 0);
}
/** The installed directory of `dep` as `root` resolves it (Bun installs it per workspace). */
function packageDir(dep, root) {
  for (const candidate of [
    resolve(root, 'node_modules', dep),
    resolve(repoRoot, 'node_modules', dep),
  ])
    if (existsSync(join(candidate, 'package.json'))) return candidate;
  let dir = dirname(createRequire(resolve(root, 'package.json')).resolve(dep));
  while (dir !== dirname(dir)) {
    const manifest = join(dir, 'package.json');
    if (existsSync(manifest) && JSON.parse(readFileSync(manifest, 'utf8')).name === dep) return dir;
    dir = dirname(dir);
  }
  return resolve(repoRoot, 'node_modules', dep);
}

/**
 * The notices of `pkg`, or null after it reports a problem. A thrown error is caught by the
 * caller, so every failure reaches the all-or-nothing delete.
 */
function renderNotices(pkg) {
  const root = resolve(repoRoot, pkg.root);
  const dist = resolve(root, 'dist');
  if (!existsSync(dist)) {
    console.error(`${pkg.name}: no dist. Run \`bun run build:packages\` first.`);
    return null;
  }
  if (pkg.name === '@xlsx-editor.dev/core') validateHarfBuzzNotices();
  const inlined = new Set();
  const metafiles = readdirSync(dist).filter(
    (n) => n.startsWith('metafile-') && n.endsWith('.json')
  );
  // Without a metafile the inlined packages are unknown: fail, never claim "bundles no third-party code".
  if (metafiles.length === 0) {
    console.error(`${pkg.name}: no metafile in dist. Set \`metafile: true\` in tsup.config.ts.`);
    return null;
  }
  for (const f of metafiles) {
    let meta;
    let reason = 'it has no inputs object';
    try {
      meta = JSON.parse(readFileSync(join(dist, f), 'utf8'));
    } catch (error) {
      reason = error.message;
    }
    if (typeof meta?.inputs !== 'object' || meta.inputs === null || Array.isArray(meta.inputs)) {
      console.error(
        `${pkg.name}: dist/${f} is not a valid metafile (${reason}). ` +
          'Run `bun run build:packages`.'
      );
      return null;
    }
    for (const input of Object.keys(meta.inputs)) {
      const m = /node_modules\/(?:\.bun\/[^/]+\/node_modules\/)?((?:@[^/]+\/)?[^/]+)\//.exec(input);
      if (m) inlined.add(m[1]);
    }
  }
  let ok = true;
  let text = `# Third-party notices for ${pkg.name}\n\n`;
  const vendored = VENDORED[pkg.name] ?? [];
  if (inlined.size === 0 && vendored.length === 0)
    text += 'This package bundles no third-party code.\n';
  for (const entry of vendored) text += `${entry}\n`;
  for (const dep of [...inlined].sort()) {
    const base = packageDir(dep, root);
    const licenses = licenseTextsIn(base, LICENSE_FILE);
    if (licenses.length === 0) {
      console.error(
        `${pkg.name}: inlined ${dep} ships no license text (looked for LICENSE, LICENCE, ` +
          `COPYING, or UNLICENSE in ${relative(repoRoot, base)}). ` +
          `Keep it external in the tsup config of ${pkg.name}, or use a package that ships one.`
      );
      ok = false;
      continue;
    }
    const texts = [...licenses, ...licenseTextsIn(base, NOTICE_FILE)];
    text += `## ${dep}\n\n`;
    for (const { filename, text: body } of texts)
      text += `${texts.length > 1 ? `### ${filename}\n\n` : ''}\`\`\`\n${body}\n\`\`\`\n\n`;
  }
  return ok ? { inlined, text } : null;
}

let failed = false;
const renders = [];
for (const pkg of PACKAGES) {
  try {
    const render = renderNotices(pkg);
    if (render) renders.push({ pkg, ...render });
    else failed = true;
  } catch (error) {
    console.error(packageError(pkg.name, error));
    failed = true;
  }
}
// A partial notice claims licenses that it leaves out, so a failed run writes nothing and
// deletes the notices that it would replace.
if (failed) deleteNotices();
const written = [];
for (const { pkg, inlined, text } of failed && !check ? [] : renders) {
  const out = noticesOf(pkg);
  if (check) {
    if (!existsSync(out) || readFileSync(out, 'utf8') !== text) {
      const state = existsSync(out) ? 'out of date' : 'missing';
      console.error(
        `${pkg.name}: THIRD_PARTY_NOTICES.md is ${state}. Run \`bun run notices:generate\`.`
      );
      failed = true;
    }
  } else {
    try {
      writeFileSync(out, text);
    } catch (error) {
      console.error(writeError(relative(repoRoot, out), error));
      failed = true;
      deleteNotices();
      break;
    }
    written.push(
      `${pkg.name}: ${inlined.size} bundled third-party package${inlined.size === 1 ? '' : 's'} -> ${relative(repoRoot, out)}`
    );
  }
}
// A failed write deletes the notices, so a line prints only when every write succeeded.
if (!failed) for (const line of written) console.log(line);
if (check && !failed) console.log('Third-party notices are up to date.');
process.exit(failed ? 1 : 0);
