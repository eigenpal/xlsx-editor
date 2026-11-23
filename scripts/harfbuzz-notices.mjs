// Copied from pptx-editor/scripts/harfbuzz-notices.mjs at d4502552 (EigenPal). Differences: each
// error names its fix (the upgrade steps in `packages/core/licenses/README.md`, a restore from
// git, or a rebuild of a stale `dist`); a missing binary or notice has its own error; a notice
// or binary error names the file from the repository root; and the version error names the
// three versions and `bun install` for a stale install.
//
// The JavaScript metafile cannot identify code inside the copied WebAssembly binary.
import { createHash } from 'node:crypto';
import { existsSync, readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, relative, resolve } from 'node:path';

const repoRoot = resolve(import.meta.dirname, '..');
const coreRoot = resolve(repoRoot, 'packages/core');
export const HARFBUZZ_MANIFEST = JSON.parse(
  readFileSync(resolve(coreRoot, 'licenses/harfbuzz-manifest.json'), 'utf8')
);

const STEPS = 'the upgrade steps in packages/core/licenses/README.md';
const sha256 = (bytes) => createHash('sha256').update(bytes).digest('hex');

/** Fail before publishing if a dependency upgrade bypasses the notice inventory. */
export function validateHarfBuzzNotices(options = {}) {
  const root = options.coreRoot ?? coreRoot;
  const manifest = JSON.parse(
    readFileSync(resolve(root, 'licenses/harfbuzz-manifest.json'), 'utf8')
  );
  const dependencyRoot =
    options.dependencyRoot ??
    resolve(dirname(createRequire(resolve(root, 'package.json')).resolve('harfbuzzjs')), '..');
  const dependency = JSON.parse(readFileSync(resolve(dependencyRoot, 'package.json'), 'utf8'));
  const declared = JSON.parse(readFileSync(resolve(root, 'package.json'), 'utf8'));
  if (
    dependency.name !== manifest.package.name ||
    dependency.version !== manifest.package.version ||
    declared.devDependencies.harfbuzzjs !== manifest.package.version
  ) {
    throw new Error(
      `HarfBuzz version changed: installed ${dependency.version}, declared ` +
        `${declared.devDependencies.harfbuzzjs}, manifest ${manifest.package.version}. If the ` +
        'installed version differs from the declared one, run `bun install`. Otherwise, update ' +
        `the component licenses and manifest before publishing, as ${STEPS} say.`
    );
  }
  for (const file of [
    resolve(dependencyRoot, 'dist/harfbuzz.wasm'),
    resolve(root, 'dist/harfbuzz.wasm'),
  ]) {
    if (!existsSync(file)) {
      throw new Error(
        `HarfBuzz binary is missing: ${relative(repoRoot, file)}. Run \`bun install\` and ` +
          '`bun run build:packages`.'
      );
    }
    if (sha256(readFileSync(file)) !== manifest.package.wasmSha256) {
      throw new Error(
        `HarfBuzz binary differs from the notice manifest: ${relative(repoRoot, file)}. If ` +
          `the file is in a package dist, run \`bun run build:packages\`. Otherwise, follow ${STEPS}.`
      );
    }
  }
  for (const notice of manifest.notices) {
    const file = resolve(root, 'licenses', notice.file);
    if (!existsSync(file)) {
      throw new Error(
        `HarfBuzz notice is missing: ${relative(repoRoot, file)}. Restore the file from git. If ` +
          `the notice changed upstream, follow ${STEPS}.`
      );
    }
    if (sha256(readFileSync(file)) !== notice.sha256) {
      throw new Error(
        `HarfBuzz notice differs from the source manifest: ${relative(repoRoot, file)}. Restore ` +
          `the file from git. If the notice changed upstream, follow ${STEPS}.`
      );
    }
  }
}

export const HARFBUZZ_NOTICE =
  '## HarfBuzz\n\n' +
  '`dist/harfbuzz.wasm` contains HarfBuzz, distributed by `harfbuzzjs` (its license is below). ' +
  'The component versions, source URLs, and notice checksums are in `licenses/harfbuzz-manifest.json`. ' +
  'The following files contain the component copyright and permission notices. ' +
  'The source and runtime notices also cover optional upstream components.\n\n' +
  HARFBUZZ_MANIFEST.notices
    .map((notice) => `- ${notice.component}: \`licenses/${notice.file}\``)
    .join('\n') +
  '\n';
