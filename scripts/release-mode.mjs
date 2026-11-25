// Adapted from pptx-editor/scripts/release-mode.mjs at 8e843068 (EigenPal). Differences: no
// canary channel, because changesets refuses a snapshot release in prerelease mode; and
// `prereleaseTag` reports the dist-tag of `.changeset/pre.json` for the run log.
//
// Decides which path the Release workflow takes:
//   pending — changesets wait in .changeset/, so the run only updates the release PR;
//   publish — no changesets wait, so the run builds and publishes.
// In prerelease mode, `changeset version` moves each changeset that it used into
// `.changeset/pre/`, so a used changeset never counts as pending.
import { appendFileSync, existsSync, readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';

const json = (path) => JSON.parse(readFileSync(path, 'utf8'));

export function pendingChangesets(root) {
  const directory = join(root, '.changeset');
  return readdirSync(directory, { withFileTypes: true })
    .filter((entry) => entry.isFile() && entry.name.endsWith('.md') && entry.name !== 'README.md')
    .map((entry) => entry.name);
}

// The fixed group is the published package set; scripts/release-mode.test.ts proves it
// equals every public workspace. Manifests resolve by package name, not directory name.
export function publicPackages(root) {
  const manifests = new Map(
    readdirSync(join(root, 'packages'), { withFileTypes: true })
      .filter((entry) => entry.isDirectory())
      .map((entry) => join(root, 'packages', entry.name, 'package.json'))
      .filter((path) => existsSync(path))
      .map((path) => json(path))
      .map((manifest) => [manifest.name, manifest])
  );
  return json(join(root, '.changeset/config.json'))
    .fixed.flat()
    .map((name) => {
      const manifest = manifests.get(name);
      if (!manifest) throw new Error(`No packages/*/package.json is named ${name}`);
      return manifest;
    })
    .filter((manifest) => !manifest.private)
    .map(({ name, version }) => ({ name, version }));
}

// The prerelease tag of `.changeset/pre.json`, or null outside prerelease mode.
// `scripts/release-publish.mjs` picks the dist-tag of each release.
export function prereleaseTag(root) {
  const path = join(root, '.changeset/pre.json');
  if (!existsSync(path)) return null;
  const state = json(path);
  return state.mode === 'pre' ? state.tag : null;
}

export function releaseMode(root) {
  return { mode: pendingChangesets(root).length > 0 ? 'pending' : 'publish' };
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  try {
    const { mode } = releaseMode(process.cwd());
    const tag = prereleaseTag(process.cwd());
    console.log(`Release mode: ${mode} (${tag ? `prerelease: ${tag}` : 'stable'})`);
    if (process.env.GITHUB_OUTPUT) appendFileSync(process.env.GITHUB_OUTPUT, `mode=${mode}\n`);
  } catch (error) {
    console.error(error.message);
    process.exitCode = 1;
  }
}
