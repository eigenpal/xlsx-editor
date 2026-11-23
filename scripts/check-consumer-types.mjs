#!/usr/bin/env node
// Copied from pptx-editor/scripts/check-consumer-types.mjs at e22f3f3b (EigenPal). Differences:
// the scope is `@xlsx-editor.dev`, the linked packages are core, react, vue, and i18n, the
// compiler is `bun check`, the repository's type checker (AGENTS.md, "TypeScript"),
// and the failure message names the TypeScript 7 declaration build in place of tsup's `dts`.
// The message asks for a changeset only after the first publish, because AGENTS.md
// ("Releasing") says that no pull request has a changeset until then. The comment on what the
// source build cannot see drops the declaration bundle that names `@eigenpal/ooxml-core`, and
// adds that the linked workspace folders resolve that package, so `check:package-artifacts`
// refuses such an import. The temp directory prefix is `xlsx-editor-consumer-types-`.
//
// Compiles `packages/core/src/__tests__/consumer.test-d.ts` against the BUILT declarations,
// the ones a host gets from npm.
//
// `bun run typecheck` compiles the same file with `packages/core/tsconfig.consumer.json`,
// which maps each package name to its sources. That is fast and needs no build, but it cannot
// see a break that exists only in `dist`: an `exports` entry whose `types` file is missing, a
// subpath that `exports` drops, or a type that two bundles declare twice so that core's
// `XlsxEditorInstance` and the one the React adapter hands out no longer match. The linked
// workspace folders still resolve a private package such as `@eigenpal/ooxml-core`, so
// `check:package-artifacts` refuses an undeclared bare import in `dist` instead.
//
// This script links the packages into an empty project's `node_modules`, so TypeScript
// resolves `@xlsx-editor.dev/core/editor` through `package.json` `exports`, as it does for a
// host. Needs built packages (`bun run build:packages`).
// Usage: node scripts/check-consumer-types.mjs [--keep]
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  rmSync,
  symlinkSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';

const ROOT = path.resolve(import.meta.dirname, '..');
const TEST_FILE = 'packages/core/src/__tests__/consumer.test-d.ts';
const LINKED = {
  '@xlsx-editor.dev/core': 'packages/core',
  '@xlsx-editor.dev/react': 'packages/react',
  '@xlsx-editor.dev/vue': 'packages/vue',
  '@xlsx-editor.dev/i18n': 'packages/i18n',
};

for (const [name, root] of Object.entries(LINKED)) {
  if (!existsSync(path.join(ROOT, root, 'dist'))) {
    console.error(`${name}: no dist. Run \`bun run build:packages\` first.`);
    process.exit(1);
  }
}

const keep = process.argv.includes('--keep');
const app = mkdtempSync(path.join(tmpdir(), 'xlsx-editor-consumer-types-'));
try {
  for (const [name, root] of Object.entries(LINKED)) {
    const link = path.join(app, 'node_modules', name);
    mkdirSync(path.dirname(link), { recursive: true });
    symlinkSync(path.join(ROOT, root), link, 'dir');
  }
  copyFileSync(path.join(ROOT, TEST_FILE), path.join(app, 'consumer.ts'));
  writeFileSync(
    path.join(app, 'tsconfig.json'),
    JSON.stringify(
      {
        extends: path.join(ROOT, 'tsconfig.base.json'),
        // A host's settings: no Bun types. `skipLibCheck` stays on, as in the repo and in the
        // common app templates, so the check judges the public types, not errors inside them.
        compilerOptions: { types: [], noEmit: true },
        files: ['consumer.ts'],
      },
      null,
      2
    )
  );

  const result = spawnSync('bun', ['check', '-p', app, '--no-pretty'], {
    cwd: ROOT,
    encoding: 'utf8',
  });
  if (result.status !== 0) {
    // Paths print relative to the repository; the copied test file prints as its source.
    const copied = path.relative(ROOT, path.join(app, 'consumer.ts'));
    const output = `${result.stdout}${result.stderr}`.replaceAll(copied, TEST_FILE);
    console.error(output.trim());
    console.error(
      [
        '',
        `${TEST_FILE} does not compile against the built declarations (dist/*.d.ts).`,
        'If `bun run typecheck` passes, the break is in the build, not in the sources: check the',
        '`exports` map of the package, its declaration build (scripts/build-declarations.mjs, or',
        'scripts/build-core-declarations.mjs for core), and any type that',
        'two bundles now declare twice. If the change to the public types is intended, update the',
        'test file. After the first publish, also add a changeset (minor for an addition, major',
        'for a break).',
      ].join('\n')
    );
    process.exit(1);
  }
  console.log(
    'Consumer types compile against the built declarations of core, react, vue, and i18n.'
  );
} finally {
  if (keep) console.log(`Kept ${app}`);
  else rmSync(app, { recursive: true, force: true });
}
