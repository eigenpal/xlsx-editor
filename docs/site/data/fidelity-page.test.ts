// Copied from pptx-editor/docs/site/data/fidelity-page.test.ts at 412d750f (EigenPal).
// XLSX Editor changes: it reads `support-matrix.ts` in place of `pptx-features.ts`, with
// `SUPPORT_CATEGORY_LABELS`, `supportFeatures`, and `supportFeatureById` in place of
// `FEATURE_CATEGORY_LABELS`, `pptxFeatures`, and `pptxFeatureById`, and `xlsx-fidelity.mdx` in
// place of `pptx-fidelity.mdx`.
import { describe, expect, test } from 'bun:test';
import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { SUPPORT_CATEGORY_LABELS, supportFeatureById, supportFeatures } from './support-matrix.ts';

// The docs site renders the matrix through `<FeatureMatrix>`, `<FeatureSummary>`, and
// `<FeatureBadge>` (docs/site/README.md). These tests keep the tags in the pages valid.

const CONTENT = join(import.meta.dir, '../content');

/** Every `.mdx` page under `docs/site/content`, by its path relative to that directory. */
function pages(dir = CONTENT, prefix = ''): string[] {
  return readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
    const name = `${prefix}${entry.name}`;
    if (entry.isDirectory()) return pages(join(dir, entry.name), `${name}/`);
    return entry.name.endsWith('.mdx') ? [name] : [];
  });
}

const read = (page: string): string => readFileSync(join(CONTENT, page), 'utf8');

describe('the feature matrix on the docs site', () => {
  test('the fidelity page shows each category that has rows once, in label order', () => {
    const page = read('xlsx-fidelity.mdx');
    const shown = [...page.matchAll(/<FeatureMatrix category="([^"]+)" \/>/g)].map((m) => m[1]);
    const expected = Object.keys(SUPPORT_CATEGORY_LABELS).filter((category) =>
      supportFeatures.some((row) => row.category === category)
    );
    expect(shown).toEqual(expected);
    expect(page).toContain('<FeatureSummary />');
  });

  test('every FeatureMatrix tag names a category, and every FeatureBadge tag names a row', () => {
    const wrong = pages().flatMap((page) => {
      const text = read(page);
      const categories = [...text.matchAll(/<FeatureMatrix category="([^"]*)"/g)]
        .map((m) => m[1])
        .filter((category) => !Object.hasOwn(SUPPORT_CATEGORY_LABELS, category));
      const ids = [...text.matchAll(/<FeatureBadge id="([^"]*)"/g)]
        .map((m) => m[1])
        .filter((id) => !Object.hasOwn(supportFeatureById, id));
      return [...categories, ...ids].map((value) => `${page}: ${value}`);
    });
    expect(wrong).toEqual([]);
  });

  test('every docsLink of a row names a page of the docs', () => {
    // `/docs/<slug>` is `content/<slug>.mdx` or `content/<slug>/index.mdx` (docs/site/README.md).
    const slugs = new Set(
      pages().map((page) => page.replace(/(\/index)?\.mdx$/, '').replace(/^index$/, ''))
    );
    const broken = supportFeatures
      .filter((row) => row.docsLink !== undefined)
      .filter((row) => !slugs.has(row.docsLink!.replace(/^\/docs\/?/, '').replace(/#.*$/, '')))
      .map((row) => `${row.id}: ${row.docsLink}`);
    expect(broken).toEqual([]);
  });
});
