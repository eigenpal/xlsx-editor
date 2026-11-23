// Copied from pptx-editor/scripts/check-consumer-install.mjs at e22f3f3b (EigenPal), lines 37-70.
// Differences: `run` and the pack loop are exported functions, so `check:consumer-install` and
// `check:consumer-floors` share them; the missing-build error names `bun run build:packages`; and
// each manifest is packed with its `devDependencies`, so the staged `package.json` is not written
// again.
//
// Packs each published package the way `changeset publish` does, for the checks that install the
// tarballs into an empty consumer app.
import { cpSync, existsSync, mkdirSync, readFileSync } from 'node:fs';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { PUBLISHED_PACKAGES as PACKAGES } from '../lib-packages.mjs';

/**
 * @param {string} command The program.
 * @param {string[]} args Its arguments.
 * @param {string} cwd The working directory.
 * @returns {string} The standard output.
 */
export function run(command, args, cwd) {
  const result = spawnSync(command, args, { cwd, encoding: 'utf8', stdio: 'pipe' });
  if (result.status !== 0) {
    process.stderr.write(result.stdout ?? '');
    process.stderr.write(result.stderr ?? '');
    throw new Error(`${command} ${args.join(' ')} failed with exit code ${result.status}`);
  }
  return result.stdout;
}

/**
 * Stages each package as it would be published. `changeset publish` runs `npm publish`, which
 * copies a `workspace:` range into the tarball unchanged, so such a range fails here instead of
 * being rewritten.
 *
 * @param {string} root The repository root.
 * @param {string} tempRoot The directory for the staged copies and the tarballs.
 * @returns {Map<string, string>} The tarball path of each package name.
 */
export function packPublishedPackages(root, tempRoot) {
  const stageDir = path.join(tempRoot, 'staged');
  const packDir = path.join(tempRoot, 'packs');
  const tarballs = new Map();
  mkdirSync(packDir, { recursive: true });
  for (const pkg of PACKAGES) {
    const source = path.join(root, pkg.root);
    if (!existsSync(path.join(source, 'dist')))
      throw new Error(`${pkg.name} has no build. Run \`bun run build:packages\` first.`);
    const target = path.join(stageDir, pkg.slug);
    cpSync(source, target, {
      recursive: true,
      filter: (file) => !file.includes(`${path.sep}node_modules`),
    });
    const manifest = JSON.parse(readFileSync(path.join(target, 'package.json'), 'utf8'));
    for (const field of ['dependencies', 'peerDependencies', 'optionalDependencies']) {
      for (const [name, range] of Object.entries(manifest[field] ?? {}))
        if (String(range).startsWith('workspace:'))
          throw new Error(
            `${pkg.name}: ${field} lists ${name}@${range}; npm publishes it as it is`
          );
    }
    const file = run('npm', ['pack', '--silent', '--pack-destination', packDir], target).trim();
    tarballs.set(pkg.name, path.join(packDir, file.split('\n').pop()));
  }
  return tarballs;
}
