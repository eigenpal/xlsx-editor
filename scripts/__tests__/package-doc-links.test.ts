import { afterEach, describe, expect, test } from 'bun:test';
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { headingAnchors, installedDocProblems, markdownLinks } from '../lib/package-doc-links.mjs';

const REPOSITORY = 'https://github.com/eigenpal/xlsx-editor/blob/main/packages';
let directory: string | undefined;

function install(files: Record<string, string>): Map<string, string> {
  directory = mkdtempSync(join(tmpdir(), 'package-doc-links-'));
  for (const [file, text] of Object.entries(files)) {
    mkdirSync(dirname(join(directory, file)), { recursive: true });
    writeFileSync(join(directory, file), text);
  }
  return new Map([
    ['react', join(directory, 'react')],
    ['core', join(directory, 'core')],
  ]);
}

afterEach(() => {
  if (directory) rmSync(directory, { recursive: true, force: true });
  directory = undefined;
});

describe('headingAnchors', () => {
  test('reads code, punctuation, and repeated headings', () => {
    const anchors = headingAnchors(
      '# Use `save()` now\n## Sheets, clipboard\n## Sheets, clipboard'
    );
    expect([...anchors]).toEqual(['use-save-now', 'sheets-clipboard', 'sheets-clipboard-1']);
  });

  test('skips a heading in fenced code', () => {
    expect([...headingAnchors('```\n# Not a heading\n```\n')]).toEqual([]);
  });
});

describe('markdownLinks', () => {
  test('skips code spans and fenced code', () => {
    const text = '[a](docs/a.md) `[b](b.md)`\n```\n[c](c.md)\n```\n[`d` e](f.md#g)';
    expect(markdownLinks(text)).toEqual(['docs/a.md', 'f.md#g']);
  });
});

describe('installedDocProblems', () => {
  const core = { 'core/README.md': '# Core\n## Open a workbook\n' };

  test('accepts links to files, headings, and another installed package', () => {
    const roots = install({
      ...core,
      'react/README.md': `# React\n[a](docs/a.md#b) [c](${REPOSITORY}/core/README.md#open-a-workbook)\n`,
      'react/docs/a.md':
        '# A\n## B\n[up](../README.md#react) [self](#b) [web](https://example.com)\n',
    });
    expect(installedDocProblems(roots, directory!)).toEqual([]);
  });

  test('drops a query on a link into an installed package', () => {
    const roots = install({
      ...core,
      'react/README.md': `# React\n[a](${REPOSITORY}/core/README.md?plain=1)\n`,
    });
    expect(installedDocProblems(roots, directory!)).toEqual([]);
  });

  test('checks a link into an installed package that has a query', () => {
    const link = `${REPOSITORY}/core/gone.md?plain=1`;
    const roots = install({ ...core, 'react/README.md': `# React\n[a](${link})\n` });
    expect(installedDocProblems(roots, directory!)).toEqual([
      `react/README.md: the link ${link} names no installed file`,
    ]);
  });

  test('refuses a guide that the package does not hold', () => {
    const roots = install({ ...core, 'react/README.md': '# React\n[a](docs/a.md)\n' });
    expect(installedDocProblems(roots, directory!)).toEqual([
      'react/README.md: the link docs/a.md names no installed file',
    ]);
  });

  test('refuses a heading that the target does not hold', () => {
    const roots = install({
      ...core,
      'react/README.md': `# React\n[a](${REPOSITORY}/core/README.md#missing)\n`,
    });
    expect(installedDocProblems(roots, directory!)[0]).toContain('names no heading');
  });

  test('refuses a link that leaves the package', () => {
    const roots = install({ ...core, 'react/README.md': '# React\n[a](../core/README.md)\n' });
    expect(installedDocProblems(roots, directory!)[0]).toContain('leaves the package');
  });

  test('refuses a guide that the package README does not link', () => {
    const roots = install({ ...core, 'react/README.md': '# React\n', 'react/docs/a.md': '# A\n' });
    expect(installedDocProblems(roots, directory!)).toEqual([
      'react/docs/a.md: the package README has no link to this guide',
    ]);
  });

  test('refuses a guide that only another guide links', () => {
    const roots = install({
      ...core,
      'react/README.md': '# React\n[a](docs/a.md)\n',
      'react/docs/a.md': '# A\n[b](b.md)\n',
      'react/docs/b.md': '# B\n',
    });
    expect(installedDocProblems(roots, directory!)).toEqual([
      'react/docs/b.md: the package README has no link to this guide',
    ]);
  });

  test('reads an installed guide once, also when another package links it', () => {
    const roots = install({
      'core/README.md': '# Core\n[g](docs/guide.md)\n',
      'core/docs/guide.md': `# Guide\n[x](${REPOSITORY}/core/gone.md)\n`,
      'react/README.md': `# React\n[g](${REPOSITORY}/core/docs/guide.md)\n`,
    });
    expect(installedDocProblems(roots, directory!)).toEqual([
      `core/docs/guide.md: the link ${REPOSITORY}/core/gone.md names no installed file`,
    ]);
  });

  describe('a link to a repository file outside packages/', () => {
    const SITE = 'https://github.com/eigenpal/xlsx-editor/blob/main/docs/site/content';
    const GUIDE = `${REPOSITORY}/core/README.md`;

    function withCheckout(link: string, page = "---\ntitle: 'Page'\n---\n\n## Heading\n") {
      const roots = install({
        ...core,
        'react/README.md': `# React\n[a](${link})\n`,
        'checkout/docs/site/content/page.mdx': page,
      });
      return { roots, checkout: join(directory!, 'checkout') };
    }

    test('resolves to a file and a heading of the checkout', () => {
      const { roots, checkout } = withCheckout(`${SITE}/page.mdx#heading`);
      expect(installedDocProblems(roots, checkout)).toEqual([]);
    });

    test('drops a query', () => {
      const { roots, checkout } = withCheckout(`${SITE}/page.mdx?plain=1`);
      expect(installedDocProblems(roots, checkout)).toEqual([]);
    });

    test('checks a link that has a query', () => {
      const { roots, checkout } = withCheckout(`${SITE}/gone.mdx?plain=1`);
      expect(installedDocProblems(roots, checkout)).toEqual([
        `react/README.md: the link ${SITE}/gone.mdx?plain=1 names no repository file`,
      ]);
    });

    test('refuses a file that the checkout does not hold', () => {
      const { roots, checkout } = withCheckout(`${SITE}/gone.mdx`);
      expect(installedDocProblems(roots, checkout)).toEqual([
        `react/README.md: the link ${SITE}/gone.mdx names no repository file`,
      ]);
    });

    test('refuses a heading that the file does not hold', () => {
      const { roots, checkout } = withCheckout(`${SITE}/page.mdx#missing`);
      expect(installedDocProblems(roots, checkout)[0]).toContain('names no heading');
    });

    test('refuses a link that leaves the checkout', () => {
      const { roots, checkout } = withCheckout(
        'https://github.com/eigenpal/xlsx-editor/blob/main/../outside.md'
      );
      expect(installedDocProblems(roots, checkout)[0]).toContain('leaves the repository');
    });

    test('reads the links of a linked page to this repository, one level deep', () => {
      const page = `# Page\n[site](/docs/other) [ok](${GUIDE}#open-a-workbook) [bad](${GUIDE}#gone)\n`;
      const { roots, checkout } = withCheckout(`${SITE}/page.mdx`, page);
      expect(installedDocProblems(roots, checkout)).toEqual([
        `docs/site/content/page.mdx: the link ${GUIDE}#gone names no heading`,
      ]);
    });

    test('reads the links of a linked Markdown page too', () => {
      const roots = install({
        ...core,
        'react/README.md': `# React\n[a](https://github.com/eigenpal/xlsx-editor/blob/main/docs/page.md)\n`,
        'checkout/docs/page.md': `# Page\n[b](${GUIDE}#gone)\n`,
      });
      expect(installedDocProblems(roots, join(directory!, 'checkout'))).toEqual([
        `docs/page.md: the link ${GUIDE}#gone names no heading`,
      ]);
    });

    test('reads a page once, also when two READMEs link it', () => {
      const roots = install({
        'core/README.md': `# Core\n## Open a workbook\n[a](${SITE}/page.mdx)\n`,
        'react/README.md': `# React\n[a](${SITE}/page.mdx)\n`,
        'checkout/docs/site/content/page.mdx': `# Page\n[b](${GUIDE}#gone)\n`,
      });
      expect(installedDocProblems(roots, join(directory!, 'checkout'))).toEqual([
        `docs/site/content/page.mdx: the link ${GUIDE}#gone names no heading`,
      ]);
    });

    test('reads the links of a page that only a guide links', () => {
      const roots = install({
        ...core,
        'react/README.md': '# React\n[a](docs/a.md)\n',
        'react/docs/a.md': `# A\n[b](${SITE}/page.mdx)\n`,
        'checkout/docs/site/content/page.mdx': `# Page\n[c](${GUIDE}#gone)\n`,
      });
      expect(installedDocProblems(roots, join(directory!, 'checkout'))).toEqual([
        `docs/site/content/page.mdx: the link ${GUIDE}#gone names no heading`,
      ]);
    });

    test('refuses a link of a page to a file that the package does not hold', () => {
      const roots = install({
        ...core,
        'react/README.md': `# React\n[a](${SITE}/page.mdx)\n`,
        'checkout/docs/site/content/page.mdx': `# Page\n[b](${REPOSITORY}/core/gone.md)\n`,
      });
      expect(installedDocProblems(roots, join(directory!, 'checkout'))).toEqual([
        `docs/site/content/page.mdx: the link ${REPOSITORY}/core/gone.md names no installed file`,
      ]);
    });

    test('does not read the links of a linked file that is not Markdown', () => {
      const roots = install({
        ...core,
        'react/README.md': `# React\n[a](https://github.com/eigenpal/xlsx-editor/blob/main/scripts/a.ts)\n`,
        'checkout/scripts/a.ts': `// [b](${GUIDE}#gone)\n`,
      });
      expect(installedDocProblems(roots, join(directory!, 'checkout'))).toEqual([]);
    });

    test('does not read the links of a page that a linked page links', () => {
      const second = `${SITE}/second.mdx`;
      const roots = install({
        ...core,
        'react/README.md': `# React\n[a](${SITE}/page.mdx)\n`,
        'checkout/docs/site/content/page.mdx': `# Page\n[b](${second})\n`,
        'checkout/docs/site/content/second.mdx': `# Second\n[c](${GUIDE}#gone)\n`,
      });
      expect(installedDocProblems(roots, join(directory!, 'checkout'))).toEqual([]);
    });

    test('resolves dot segments before it picks the installed package or the checkout', () => {
      const roots = install({
        ...core,
        'react/README.md': `# React\n[a](${SITE}/../../../packages/core/gone.md)\n`,
        'checkout/packages/core/gone.md': '# Gone\n',
      });
      expect(installedDocProblems(roots, join(directory!, 'checkout'))).toEqual([
        `react/README.md: the link ${SITE}/../../../packages/core/gone.md names no installed file`,
      ]);
    });

    test('resolves a link into a package in the installed package, not the checkout', () => {
      const { roots, checkout } = withCheckout(`${GUIDE}#open-a-workbook`);
      expect(installedDocProblems(roots, checkout)).toEqual([]);
    });
  });
});
