// Runs the failure paths of `scripts/generate-third-party-notices.mjs` in a temporary copy of the
// repository files that it reads, so a failing run never deletes the notices of this checkout.
// The error texts themselves are tested in `notices-errors.test.ts`.
import { afterEach, describe, expect, test } from 'bun:test';
import {
  cpSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';

import { VENDORED } from '../vendored-notices.mjs';

const repoRoot = resolve(import.meta.dirname, '../..');
const COPIED = [
  'scripts/generate-third-party-notices.mjs',
  'scripts/lib-packages.mjs',
  'scripts/lib/notices-errors.mjs',
  'scripts/vendored-notices.mjs',
  'scripts/harfbuzz-notices.mjs',
  'reference/ecma-376/NOTICE',
  'packages/core/licenses/harfbuzz-manifest.json',
];
const PACKAGE_DIRS = ['packages/i18n', 'packages/core', 'packages/react'];

let root: string | undefined;
afterEach(() => {
  if (root) rmSync(root, { recursive: true, force: true });
  root = undefined;
});

/**
 * A temporary copy with a notice in each package and no build output. Its empty `node_modules`
 * turns off Bun auto-install, so no package resolves from the global cache or the registry.
 */
function copy(): string {
  root = mkdtempSync(join(tmpdir(), 'xlsx-notices-'));
  for (const file of COPIED) cpSync(join(repoRoot, file), join(root, file), { recursive: true });
  mkdirSync(join(root, 'node_modules'));
  for (const dir of PACKAGE_DIRS) {
    mkdirSync(join(root, dir), { recursive: true });
    writeFileSync(join(root, dir, 'THIRD_PARTY_NOTICES.md'), 'old notice\n');
  }
  return root;
}

/** List only i18n and react in the copy, each with an empty metafile, so both render. */
function renderable(dir: string) {
  writeFileSync(
    join(dir, 'scripts/lib-packages.mjs'),
    'export const PACKAGES = [\n' +
      "  { name: '@xlsx-editor.dev/i18n', root: 'packages/i18n' },\n" +
      "  { name: '@xlsx-editor.dev/react', root: 'packages/react' },\n" +
      '];\n'
  );
  for (const pkg of ['packages/i18n', 'packages/react']) {
    mkdirSync(join(dir, pkg, 'dist'));
    writeFileSync(join(dir, pkg, 'dist/metafile-esm.json'), '{"inputs":{}}');
  }
}

/**
 * Install `fake-dep` in the copy with `files`, and inline it into the i18n build, from both the
 * plain path and the Bun store path, so both name the same package.
 */
function inlineFakeDep(dir: string, files: Record<string, string>) {
  const dep = join(dir, 'node_modules/fake-dep');
  mkdirSync(dep);
  writeFileSync(join(dep, 'package.json'), '{"name":"fake-dep","version":"1.0.0"}');
  for (const [name, text] of Object.entries(files)) writeFileSync(join(dep, name), text);
  writeFileSync(
    join(dir, 'packages/i18n/dist/metafile-esm.json'),
    JSON.stringify({
      inputs: {
        '../../node_modules/fake-dep/index.js': {},
        '../../node_modules/.bun/fake-dep@1.0.0/node_modules/fake-dep/util.js': {},
      },
    })
  );
}

function run(dir: string, ...args: string[]) {
  const result = Bun.spawnSync({
    cmd: [process.execPath, join(dir, 'scripts/generate-third-party-notices.mjs'), ...args],
    cwd: tmpdir(),
    stdout: 'pipe',
    stderr: 'pipe',
  });
  return {
    code: result.exitCode,
    stdout: result.stdout.toString(),
    stderr: result.stderr.toString(),
  };
}

const notices = (dir: string) =>
  PACKAGE_DIRS.filter((pkg) => existsSync(join(dir, pkg, 'THIRD_PARTY_NOTICES.md')));

describe('generate-third-party-notices failure paths', () => {
  test('a broken HarfBuzz manifest names the HarfBuzz module and the manifest', () => {
    const dir = copy();
    writeFileSync(join(dir, 'packages/core/licenses/harfbuzz-manifest.json'), '{ broken');
    const { code, stderr } = run(dir);
    expect(code).toBe(1);
    expect(stderr).toContain('scripts/harfbuzz-notices.mjs failed to load');
    expect(stderr).toContain('`packages/core/licenses/harfbuzz-manifest.json`');
    expect(notices(dir)).toEqual([]);
  });

  test('a broken Ecma notice names the vendored notices module and the notice', () => {
    const dir = copy();
    writeFileSync(join(dir, 'reference/ecma-376/NOTICE'), 'no heading\n');
    const { code, stderr } = run(dir, '--check');
    expect(code).toBe(1);
    expect(stderr).toContain('scripts/vendored-notices.mjs failed to load');
    expect(stderr).toContain('`reference/ecma-376/NOTICE`');
    expect(notices(dir)).toEqual(PACKAGE_DIRS);
  });

  test('a broken metafile fails the run and deletes every notice', () => {
    const dir = copy();
    mkdirSync(join(dir, 'packages/i18n/dist'));
    writeFileSync(join(dir, 'packages/i18n/dist/metafile-esm.json'), '{"inputs":[]}');
    const { code, stderr } = run(dir);
    expect(code).toBe(1);
    expect(stderr).toContain(
      '@xlsx-editor.dev/i18n: dist/metafile-esm.json is not a valid metafile (it has no inputs ' +
        'object). Run `bun run build:packages`.'
    );
    expect(notices(dir)).toEqual([]);
  });

  test('an invalid metafile names its parse error', () => {
    const dir = copy();
    mkdirSync(join(dir, 'packages/i18n/dist'));
    writeFileSync(join(dir, 'packages/i18n/dist/metafile-esm.json'), '{ broken');
    const { code, stderr } = run(dir);
    expect(code).toBe(1);
    expect(stderr).toMatch(
      /@xlsx-editor\.dev\/i18n: dist\/metafile-esm\.json is not a valid metafile \(JSON Parse error: .+\)\. Run `bun run build:packages`\./
    );
    expect(notices(dir)).toEqual([]);
  });

  test('a package that fails after one that renders leaves no notice', () => {
    const dir = copy();
    renderable(dir);
    rmSync(join(dir, 'packages/react/dist'), { recursive: true });
    const { code, stdout, stderr } = run(dir);
    expect(code).toBe(1);
    expect(stderr).toContain('@xlsx-editor.dev/react: no dist.');
    expect(stdout).not.toContain('->');
    expect(existsSync(join(dir, 'packages/i18n/THIRD_PARTY_NOTICES.md'))).toBe(false);
  });

  test('a failed --check run keeps every notice', () => {
    const dir = copy();
    const { code, stderr } = run(dir, '--check');
    expect(code).toBe(1);
    expect(stderr).toContain('@xlsx-editor.dev/i18n: no dist. Run `bun run build:packages` first.');
    expect(notices(dir)).toEqual(PACKAGE_DIRS);
  });

  test('a notice that cannot be deleted names its fix', () => {
    const dir = copy();
    rmSync(join(dir, 'packages/react/THIRD_PARTY_NOTICES.md'));
    mkdirSync(join(dir, 'packages/react/THIRD_PARTY_NOTICES.md/held'), { recursive: true });
    const { code, stderr } = run(dir);
    expect(code).toBe(1);
    expect(stderr).toContain('Could not delete packages/react/THIRD_PARTY_NOTICES.md');
    expect(notices(dir)).toEqual(['packages/react']);
  });

  test('an error that a package throws fails the run and deletes every notice', () => {
    const dir = copy();
    // The copy has no core package.json and installs no harfbuzzjs, so the HarfBuzz check
    // throws for core. The cause depends on what resolves above the temporary directory.
    mkdirSync(join(dir, 'packages/core/dist'));
    const { code, stderr } = run(dir);
    expect(code).toBe(1);
    expect(stderr).toMatch(
      /@xlsx-editor\.dev\/core: [^]*\. Run `bun install` and `bun run build:packages`, then run this script again\./
    );
    expect(notices(dir)).toEqual([]);
  });

  test('a failed write deletes the notices written before it and prints no success line', () => {
    const dir = copy();
    renderable(dir);
    rmSync(join(dir, 'packages/react/THIRD_PARTY_NOTICES.md'));
    mkdirSync(join(dir, 'packages/react/THIRD_PARTY_NOTICES.md/held'), { recursive: true });
    const { code, stdout, stderr } = run(dir);
    expect(code).toBe(1);
    expect(stderr).toContain('Could not write packages/react/THIRD_PARTY_NOTICES.md');
    expect(stdout).not.toContain('->');
    expect(existsSync(join(dir, 'packages/i18n/THIRD_PARTY_NOTICES.md'))).toBe(false);
  });

  test('a dist with no metafile fails the run and deletes every notice', () => {
    const dir = copy();
    mkdirSync(join(dir, 'packages/i18n/dist'));
    const { code, stderr } = run(dir);
    expect(code).toBe(1);
    expect(stderr).toContain(
      '@xlsx-editor.dev/i18n: no metafile in dist. Set `metafile: true` in tsup.config.ts.'
    );
    expect(notices(dir)).toEqual([]);
  });

  test('a passing run writes each notice and prints one line for each', () => {
    const dir = copy();
    renderable(dir);
    const { code, stdout } = run(dir);
    expect(code).toBe(0);
    expect(readFileSync(join(dir, 'packages/i18n/THIRD_PARTY_NOTICES.md'), 'utf8')).toBe(
      '# Third-party notices for @xlsx-editor.dev/i18n\n\nThis package bundles no third-party code.\n'
    );
    expect(readFileSync(join(dir, 'packages/react/THIRD_PARTY_NOTICES.md'), 'utf8')).toBe(
      `# Third-party notices for @xlsx-editor.dev/react\n\n${VENDORED['@xlsx-editor.dev/react'].join('\n')}\n`
    );
    expect(stdout).toContain(
      '@xlsx-editor.dev/i18n: 0 bundled third-party packages -> packages/i18n/THIRD_PARTY_NOTICES.md'
    );
    expect(stdout).toContain('-> packages/react/THIRD_PARTY_NOTICES.md');
    const checked = run(dir, '--check');
    expect(checked.code).toBe(0);
    expect(checked.stdout).toBe('Third-party notices are up to date.\n');
  });

  test('one license text ships with no file name heading', () => {
    const dir = copy();
    renderable(dir);
    inlineFakeDep(dir, { LICENSE: 'MIT text' });
    expect(run(dir).code).toBe(0);
    expect(readFileSync(join(dir, 'packages/i18n/THIRD_PARTY_NOTICES.md'), 'utf8')).toBe(
      '# Third-party notices for @xlsx-editor.dev/i18n\n\n## fake-dep\n\n```\nMIT text\n```\n\n'
    );
  });

  test('an inlined package with no license text fails the run and deletes the notice of each listed package', () => {
    const dir = copy();
    renderable(dir);
    inlineFakeDep(dir, {});
    const { code, stderr } = run(dir);
    expect(code).toBe(1);
    expect(stderr).toContain(
      '@xlsx-editor.dev/i18n: inlined fake-dep ships no license text (looked for LICENSE, ' +
        'LICENCE, COPYING, or UNLICENSE in node_modules/fake-dep). Keep it external in the tsup ' +
        'config of @xlsx-editor.dev/i18n, or use a package that ships one.'
    );
    expect(notices(dir)).toEqual(['packages/core']);
  });

  test('an inlined package ships its license and NOTICE texts, each under its file name', () => {
    const dir = copy();
    renderable(dir);
    inlineFakeDep(dir, {
      'LICENSE-MIT': 'MIT text',
      NOTICE: 'Notice text',
      'README.md': 'Not a license',
    });
    expect(run(dir).code).toBe(0);
    const text = readFileSync(join(dir, 'packages/i18n/THIRD_PARTY_NOTICES.md'), 'utf8');
    expect(text).toBe(
      '# Third-party notices for @xlsx-editor.dev/i18n\n\n## fake-dep\n\n' +
        '### LICENSE-MIT\n\n```\nMIT text\n```\n\n### NOTICE\n\n```\nNotice text\n```\n\n'
    );
  });

  test('--check names a stale or missing notice and keeps the files', () => {
    const dir = copy();
    renderable(dir);
    expect(run(dir).code).toBe(0);
    writeFileSync(join(dir, 'packages/i18n/THIRD_PARTY_NOTICES.md'), 'old notice\n');
    rmSync(join(dir, 'packages/react/THIRD_PARTY_NOTICES.md'));
    const { code, stderr } = run(dir, '--check');
    expect(code).toBe(1);
    expect(stderr).toContain(
      '@xlsx-editor.dev/i18n: THIRD_PARTY_NOTICES.md is out of date. Run `bun run notices:generate`.'
    );
    expect(stderr).toContain(
      '@xlsx-editor.dev/react: THIRD_PARTY_NOTICES.md is missing. Run `bun run notices:generate`.'
    );
    expect(notices(dir)).toEqual(['packages/i18n', 'packages/core']);
  });
});
