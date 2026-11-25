// Publishes the packages for the Release workflow (`.github/workflows/release.yml`).
//
// In prerelease mode, `changeset publish` sends each version to the dist-tag of
// `.changeset/pre.json` (`beta`) and refuses `--tag`. So `latest`, which is the version
// that the npm page and a plain `npm install` show, stays on an old beta. npm trusted
// publishing covers `npm publish` only, so no later step can move `latest` without a token.
//
// This script keeps the changesets publish path and changes only the dist-tag:
//   1. `changeset publish-plan` writes the plan, with one dist-tag for each package.
//   2. A package that has no stable version on the registry gets `latest` in place of the
//      prerelease dist-tag. The newest beta is then the version that users see.
//   3. `changeset pack` packs the plan, and `changeset publish --from-pack-dir` publishes
//      the tarballs and makes the git tags.
// After the first stable release of a package, step 2 changes nothing for it: a later beta
// goes to the prerelease dist-tag, and `latest` stays on the stable version.
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { prereleaseTag } from './release-mode.mjs';

// A stable version has no prerelease part and no build part.
export const isStable = (version) => /^\d+\.\d+\.\d+$/.test(version);

// Returns the plan with `latest` on each publish entry that carries the prerelease
// dist-tag, when `hasStableVersion(name)` is false. Other entries stay as they are.
export function retargetPlan(plan, preTag, hasStableVersion) {
  if (preTag === null) return plan;
  return plan.map((group) =>
    group.map((release) =>
      release.kind === 'publish' && release.tag === preTag && !hasStableVersion(release.name)
        ? { ...release, tag: 'latest' }
        : release
    )
  );
}

// Reads the published versions of a package. A package that the registry does not have
// has no versions. Any other failed read stops the release: a guess here could put a beta
// on `latest` over a stable version.
function publishedVersions(name) {
  let output;
  try {
    output = execFileSync('npm', ['view', name, 'versions', '--json'], {
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
    });
  } catch (error) {
    if (`${error.stdout}${error.stderr}`.includes('E404')) return [];
    throw new Error(`Cannot read the published versions of ${name}: ${error.stderr}`);
  }
  const versions = JSON.parse(output);
  return Array.isArray(versions) ? versions : [versions];
}

function main() {
  const root = process.cwd();
  const directory = mkdtempSync(join(process.env.RUNNER_TEMP ?? tmpdir(), 'release-publish-'));
  const planPath = join(directory, 'publish-plan.json');
  const changeset = (...args) => execFileSync('bunx', ['changeset', ...args], { stdio: 'inherit' });

  changeset('publish-plan', '--output', planPath);
  const file = JSON.parse(readFileSync(planPath, 'utf8'));
  const plan = retargetPlan(file.plan, prereleaseTag(root), (name) =>
    publishedVersions(name).some(isStable)
  );
  for (const release of plan.flat())
    if (release.kind === 'publish')
      console.log(`Publish ${release.name}@${release.version} to the dist-tag ${release.tag}`);
  writeFileSync(planPath, JSON.stringify({ ...file, plan }, null, 2));

  const packDirectory = join(directory, 'pack');
  changeset('pack', '--from-publish-plan', planPath, '--out-dir', packDirectory);
  changeset('publish', '--from-pack-dir', packDirectory);
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  try {
    main();
  } catch (error) {
    console.error(error.message);
    process.exitCode = 1;
  }
}
