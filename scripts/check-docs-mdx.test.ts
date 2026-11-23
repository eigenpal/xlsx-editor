// Tests for scripts/check-docs-mdx.mjs. Each case runs a copy of the script on pages in a new folder.
import { expect, test } from 'bun:test';
import { copyFileSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { spawnSync } from 'node:child_process';

function check(pages: Record<string, string>) {
  // The script reads `docs/site/content` next to its own `scripts` folder.
  const root = mkdtempSync(join(tmpdir(), 'docs-mdx-'));
  try {
    mkdirSync(join(root, 'scripts'));
    copyFileSync(
      join(import.meta.dir, 'check-docs-mdx.mjs'),
      join(root, 'scripts/check-docs-mdx.mjs')
    );
    for (const [path, text] of Object.entries(pages)) {
      const file = join(root, 'docs/site/content', path);
      mkdirSync(dirname(file), { recursive: true });
      writeFileSync(file, text);
    }
    const result = spawnSync('node', [join(root, 'scripts/check-docs-mdx.mjs')], {
      cwd: root,
      encoding: 'utf8',
    });
    return { status: result.status, output: result.stdout + result.stderr };
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
}

const PAGE = '---\ntitle: Rows\n---\n\n';

test('refuses a bare {name} in prose and names the file, line, and fix', () => {
  const result = check({ 'guides/rows.mdx': `${PAGE}Select the Edit {label} row.\n` });
  expect(result.status).toBe(1);
  expect(result.output).toContain('docs/site/content/guides/rows.mdx:5  {label}');
  expect(result.output).toContain('Backtick the braces, or escape them');
});

test('passes the same name in a code span, a fence, an escape, and frontmatter', () => {
  const result = check({
    'span.mdx': `${PAGE}Select the \`Edit {label}\` row.\n`,
    'fence.mdx': `${PAGE}\`\`\`ts\nconst row = {label};\n\`\`\`\n`,
    'escape.mdx': `${PAGE}Select the Edit \\{label\\} row.\n`,
    'frontmatter.mdx': '---\ntitle: Edit {label}\n---\n\nSelect a row.\n',
  });
  expect(result).toEqual({ status: 0, output: expect.stringContaining('docs mdx: OK — 4 pages') });
});

test('passes a name that the page binds itself', () => {
  const result = check({
    'bound.mdx': `${PAGE}export const label = 'Rows';\n\nSelect the Edit {label} row.\n`,
  });
  expect(result.status).toBe(0);
});

test('refuses an empty content folder', () => {
  const result = check({ 'meta.json': '{"pages":[]}' });
  expect(result.status).toBe(1);
  expect(result.output).toContain('no MDX under docs/site/content/');
});
