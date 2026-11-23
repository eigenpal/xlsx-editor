#!/usr/bin/env node
// Copied from pptx-editor/scripts/check-package-artifacts.mjs at e22f3f3b (EigenPal). Differences:
// the scope is `@xlsx-editor.dev`. The Ecma notice check also covers the bundled predefined
// function names (`PREDEFINED_FUNCTIONS`), a second ECMA-376 table that PPTX Editor does not have.
// The tarball checks of THIRD_PARTY_NOTICES.md, `licenses/`, the license header, and the Ecma
// notice are from PPTX Editor at d4502552. A missing THIRD_PARTY_NOTICES.md names the
// command that writes it, and a required file that is on disk but not in the tarball names the
// `files` entry. A missing Ecma notice names the build and the generator to run. The `licenses/`
// list skips dotfiles: no notice is a dotfile, and npm never publishes system files such as
// `.DS_Store`, so a local system file does not fail the check. A missing table names the license
// texts that list the bundled tables, a missing tracked file names a restore from git, and the
// license header error names the banner to add. The remedy choice, the `licenses/` list, and the
// Ecma match are in `scripts/lib/package-artifact-rules.mjs`, with unit tests.
// It adds checks that neither sister script has. The adapter and i18n manifests need `keywords` and
// an https `homepage`, as the DOCX Editor adapter and i18n manifests have. Each relative import of
// a built `.js` file must reach a file, so a missing chunk fails before a publish. Each export
// target must be in the tarball, not only on disk. Each bare import in `dist` must name the package
// itself, a declared dependency or peer, or a Node.js built-in, so no built file imports a private
// workspace package such as `@eigenpal/ooxml-core`. This check reads built files only: the packed
// core manifest still lists that package in `devDependencies`, which a consumer's install does not
// read. Each request for a published package must name a subpath of its export map. It also fixes
// the `sideEffects` value of each package, reads a side-effect import that has no space, and names
// the build command when a `dist` is missing. Published exports and declarations must be
// self-contained (docx-editor/scripts/check-package-artifacts.mjs). Adapted from DOCX Editor at
// commit adf0e178.
import { execFileSync } from 'node:child_process';
import { builtinModules } from 'node:module';
import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs';
import path from 'node:path';
import {
  bareImportRequests,
  exposesWorkspacePath,
  hasEcmaLegalComment,
  listingFieldProblems,
  missingFileRemedy,
  packageNameOf,
  relativeImportRequests,
  requiredLicenseFiles,
  sideEffectsProblem,
  unexportedRequestProblem,
} from './lib/package-artifact-rules.mjs';

const root = path.resolve(import.meta.dirname, '..');
// Discover publishable packages so new packages cannot bypass artifact checks.
const packageDirs = readdirSync(path.join(root, 'packages')).filter((directory) => {
  const manifest = path.join(root, 'packages', directory, 'package.json');
  return existsSync(manifest) && JSON.parse(readFileSync(manifest, 'utf8')).private !== true;
});
const errors = [];
const publishedManifests = new Map(
  packageDirs.map((directory) => {
    const manifest = JSON.parse(
      readFileSync(path.join(root, 'packages', directory, 'package.json'), 'utf8')
    );
    return [manifest.name, manifest];
  })
);

const unbuilt = packageDirs.filter(
  (directory) => !existsSync(path.join(root, 'packages', directory, 'dist'))
);
if (unbuilt.length > 0) {
  console.error(`No build of ${unbuilt.join(', ')}. Run \`bun run build:packages\` first.`);
  process.exit(1);
}

// What a published artifact may name when it reaches for the engine: exactly the
// subpaths core's own export map declares. Deriving the set instead of listing it
// keeps this honest when core adds or retires a lane — a hand-written denylist ages
// into false positives on subpaths that have since become public.
const corePackageJson = JSON.parse(
  readFileSync(path.join(root, 'packages/core/package.json'), 'utf8')
);
const publicCoreSubpaths = new Set(
  Object.keys(corePackageJson.exports ?? {})
    .filter((subpath) => subpath.startsWith('./'))
    .map((subpath) => subpath.slice(2))
);

function filesBelow(directory) {
  if (!existsSync(directory)) return [];
  return readdirSync(directory).flatMap((entry) => {
    const absolute = path.join(directory, entry);
    return statSync(absolute).isDirectory() ? filesBelow(absolute) : [absolute];
  });
}

const builtins = new Set(builtinModules);

function exportTargets(value) {
  if (typeof value === 'string') return [value];
  if (!value || typeof value !== 'object') return [];
  return Object.values(value).flatMap(exportTargets);
}

for (const packageDir of packageDirs) {
  const packageRoot = path.join(root, 'packages', packageDir);
  const packageJson = JSON.parse(readFileSync(path.join(packageRoot, 'package.json'), 'utf8'));

  const sideEffects = sideEffectsProblem(packageJson);
  if (sideEffects) errors.push(sideEffects);
  errors.push(...listingFieldProblems(packageJson));

  for (const [subpath, declaration] of Object.entries(packageJson.exports ?? {})) {
    for (const target of exportTargets(declaration)) {
      if (!target.startsWith('./')) continue;
      if (!existsSync(path.resolve(packageRoot, target))) {
        errors.push(`${packageJson.name} ${subpath} points to missing ${target}`);
      }
    }
  }

  const declared = new Set([
    packageJson.name,
    ...Object.keys(packageJson.dependencies ?? {}),
    ...Object.keys(packageJson.peerDependencies ?? {}),
  ]);
  const artifacts = ['dist', 'types'].flatMap((directory) =>
    filesBelow(path.join(packageRoot, directory))
  );
  for (const artifact of artifacts) {
    if (!/\.(?:[cm]?js|d\.[cm]?ts)$/.test(artifact)) continue;
    const content = readFileSync(artifact, 'utf8');
    const isDeclaration = /\.d\.[cm]?ts$/.test(artifact);
    if (isDeclaration && exposesWorkspacePath(content)) {
      errors.push(`${path.relative(root, artifact)} exposes a workspace-only source path`);
    }
    if (isDeclaration) {
      for (const [, request] of content.matchAll(
        /\b(?:from\s*|import\s*\(\s*)['"](\.{1,2}\/[^'"]+)['"]/g
      )) {
        const target = path.resolve(path.dirname(artifact), request);
        const candidates = [
          target,
          target.replace(/\.([cm]?)js$/, '.d.$1ts'),
          `${target}.d.ts`,
          path.join(target, 'index.d.ts'),
        ];
        if (!candidates.some(existsSync)) {
          errors.push(`${path.relative(root, artifact)} references missing declaration ${request}`);
        }
      }
    }

    // A consumer's install has only the declared packages, so any other name fails there.
    for (const request of bareImportRequests(content)) {
      if (request.startsWith('node:') || builtins.has(packageNameOf(request))) continue;
      if (!declared.has(packageNameOf(request))) {
        errors.push(
          `${path.relative(root, artifact)} imports ${request}, which ${packageJson.name} does not declare as a dependency or peer`
        );
      }
      const unexported = unexportedRequestProblem(request, publishedManifests);
      if (unexported) errors.push(`${path.relative(root, artifact)} imports ${unexported}`);
    }

    // A subpath core does not export resolves to nothing on a consumer's disk, so the
    // failure lands at their `npm install`, not ours.
    if (packageDir !== 'core') {
      for (const [, subpath] of content.matchAll(/@xlsx-editor\.dev\/core\/([\w./-]+)/g)) {
        if (!publicCoreSubpaths.has(subpath)) {
          errors.push(
            `${path.relative(root, artifact)} imports @xlsx-editor.dev/core/${subpath}, which core does not export`
          );
        }
      }
    }

    if (!isDeclaration && artifact.endsWith('.js')) {
      for (const request of relativeImportRequests(content)) {
        const target = path.resolve(path.dirname(artifact), request);
        if (!existsSync(target)) {
          errors.push(`${path.relative(root, artifact)} imports missing file ${request}`);
        }
      }
      for (const match of content.matchAll(/\brequire\(\s*['"](\.{1,2}\/[^'"]+)['"]\s*\)/g)) {
        const request = match[1];
        const target = path.resolve(path.dirname(artifact), request);
        const candidates = [target, `${target}.js`, path.join(target, 'index.js')];
        if (!candidates.some(existsSync)) {
          errors.push(
            `${path.relative(root, artifact)} requires missing private artifact ${request}`
          );
        }
      }
    }
  }
}

// What npm would publish. npm skips a missing `files` entry without an error, so a lost license
// or notice file would publish silently: read the tarball list and require them.
for (const packageDir of packageDirs) {
  const packageRoot = path.join(root, 'packages', packageDir);
  const packageJson = JSON.parse(readFileSync(path.join(packageRoot, 'package.json'), 'utf8'));
  const name = packageJson.name;
  const [packed] = JSON.parse(
    execFileSync('npm', ['pack', '--dry-run', '--json', '--ignore-scripts'], {
      cwd: packageRoot,
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
      // npm is `npm.cmd` on Windows, which only a shell resolves.
      shell: process.platform === 'win32',
    })
  );
  const files = new Set(packed.files.map((file) => file.path));
  const proprietary = String(packageJson.license).startsWith('LicenseRef-');
  const required = ['README.md', 'THIRD_PARTY_NOTICES.md', proprietary ? 'LICENSE.md' : 'LICENSE'];
  // The component notices of the bundled HarfBuzz binary.
  if (existsSync(path.join(packageRoot, 'licenses')))
    required.push(...requiredLicenseFiles(readdirSync(path.join(packageRoot, 'licenses'))));
  for (const file of required) {
    if (files.has(file)) continue;
    const remedy = missingFileRemedy(file, existsSync(path.join(packageRoot, file)), packageDir);
    errors.push(`${name} would publish without ${file}.${remedy}`);
  }
  // An export target that exists on disk but that `files` leaves out is missing on npm.
  for (const [subpath, declaration] of Object.entries(packageJson.exports ?? {})) {
    for (const target of exportTargets(declaration)) {
      if (target.startsWith('./') && !files.has(target.slice(2)))
        errors.push(`${name} ${subpath} points to ${target}, which the tarball leaves out`);
    }
  }
  for (const file of files) {
    if (/\.map$|(?:^|\/)metafile-[^/]*\.json$/.test(file))
      errors.push(`${name} would publish ${file}, which holds source or build paths`);
  }

  const scripts = [...files].filter((file) => /^dist\/.*\.[cm]?js$/.test(file));
  const declarations = [...files].filter((file) => /^dist\/.*\.d\.[cm]?ts$/.test(file));
  // A home directory or CI workspace path is a build machine path, wherever it was checked out.
  const machinePath = /(?:^|[\s'"`(])\/(?:Users|home|runner|private\/var|tmp)\/[\w.-]+\//;
  for (const file of [...scripts, ...declarations]) {
    const content = readFileSync(path.join(packageRoot, file), 'utf8');
    if (content.includes(root) || machinePath.test(content))
      errors.push(`${name} ${file} contains a build machine path`);
  }
  for (const file of [...scripts, ...declarations]) {
    const content = readFileSync(path.join(packageRoot, file), 'utf8');
    // Every file of a proprietary package opens with its license header (the package's
    // `license-header.txt`, added with the first EigenPal Pro License package).
    if (proprietary && !/^\/\*!\s*Copyright \(c\) \d{4} Eigenpal, Inc\./.test(content))
      errors.push(
        `${name} ${file} does not start with the license header. Add the package's ` +
          '`license-header.txt` as the `/*!` banner in its tsup config, then run ' +
          '`bun run build:packages`.'
      );
  }
  // The Ecma notice travels with each bundled ECMA-376 table, inside the code itself.
  if (name === '@xlsx-editor.dev/core') {
    const TABLE_LISTS =
      '`LICENSE`, `AGENTS.md`, `packages/core/LICENSE`, `packages/core/README.md`, ' +
      '`packages/ooxml-core/README.md`, ' +
      '`reference/ecma-376/NOTICE`, `reference/README.md`, `scripts/vendored-notices.mjs`, and ' +
      '`scripts/notices-sources.test.ts`';
    for (const [table, label, build] of [
      ['PRESET_SHAPES', 'the preset geometry', 'geometry:build'],
      ['PREDEFINED_FUNCTIONS', 'the predefined function names', 'functions:build'],
    ]) {
      const bundling = scripts.filter((file) =>
        readFileSync(path.join(packageRoot, file), 'utf8').includes(table)
      );
      for (const file of bundling) {
        if (!hasEcmaLegalComment(readFileSync(path.join(packageRoot, file), 'utf8')))
          errors.push(
            `${name} ${file} bundles ${label} without the Ecma notice. Run \`bun run ` +
              `build:packages\`. If the generated table has no notice, run \`bun run ${build}\` ` +
              'first. Keep `/*!` legal comments in the core build.'
          );
      }
      if (bundling.length === 0)
        errors.push(
          `${name} dist has no file that names ${table}. If the table was renamed, update its ` +
            'name in `scripts/check-package-artifacts.mjs`. If core no longer bundles it, remove ' +
            `it from this check and from the texts that list the bundled tables, such as ` +
            `${TABLE_LISTS}.`
        );
    }
  }
}

if (errors.length > 0) {
  console.error(errors.join('\n'));
  process.exit(1);
}

console.log(
  'Published package exports and declarations are self-contained, and every tarball carries its README, license, and notices.'
);
