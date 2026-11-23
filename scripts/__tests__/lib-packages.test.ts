// `scripts/lib-packages.mjs` is the only list that `api:check` and the consumer checks read, while
// `check:package-artifacts` and `check:package-size` find the non-private packages on disk. This
// test keeps the publishable packages of the list equal to the ones on disk, so a new publishable
// package cannot ship without an API report. The list may also name a private package, whose API
// is reviewed the same way, while it names a package that exists and marks it `private`.
import { describe, expect, test } from 'bun:test';
import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { PACKAGES, PUBLISHED_PACKAGES } from '../lib-packages.mjs';

const ROOT = resolve(import.meta.dirname, '..', '..');

interface Manifest {
  name: string;
  private?: boolean;
}

function manifestAt(root: string): Manifest | undefined {
  const file = join(ROOT, root, 'package.json');
  return existsSync(file) ? (JSON.parse(readFileSync(file, 'utf8')) as Manifest) : undefined;
}

function publishablePackages(): { name: string; root: string }[] {
  const found: { name: string; root: string }[] = [];
  for (const dir of readdirSync(join(ROOT, 'packages')).sort()) {
    const manifest = manifestAt(`packages/${dir}`);
    if (manifest && manifest.private !== true)
      found.push({ name: manifest.name, root: `packages/${dir}` });
  }
  return found;
}

describe('PACKAGES', () => {
  test('lists each non-private package under packages/', () => {
    const listed = PACKAGES.filter(({ root }) => manifestAt(root)?.private !== true)
      .map(({ name, root }) => ({ name, root }))
      .sort((a, b) => a.root.localeCompare(b.root));
    expect(listed).toEqual(publishablePackages());
  });

  test('names only packages that exist under packages/, by their manifest name', () => {
    for (const { name, root } of PACKAGES) {
      expect(root.startsWith('packages/')).toBe(true);
      expect(manifestAt(root)?.name).toBe(name);
    }
  });

  test('marks a package private exactly when its manifest is private', () => {
    for (const pkg of PACKAGES)
      expect(pkg.private === true).toBe(manifestAt(pkg.root)?.private === true);
    expect(PUBLISHED_PACKAGES.map(({ root }) => root).sort()).toEqual(
      publishablePackages().map(({ root }) => root)
    );
  });
});
