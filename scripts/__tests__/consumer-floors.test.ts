import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { FLOORS, floorProblems } from '../lib/consumer-floors.mjs';

const repoRoot = resolve(import.meta.dir, '../..');
const read = (file: string) => readFileSync(resolve(repoRoot, file), 'utf8');
const stated = () => ({
  coreReadme: read('packages/core/README.md'),
  reactReadme: read('packages/react/README.md'),
  vueReadme: read('packages/vue/README.md'),
  engines: JSON.parse(read('packages/core/package.json')).engines?.node as string | undefined,
  reactPeer: JSON.parse(read('packages/react/package.json')).peerDependencies.react as string,
  ci: read('.github/workflows/ci.yml'),
});

describe('floorProblems', () => {
  test('the repository states the tested floors', () => {
    expect(floorProblems(stated())).toEqual([]);
  });

  test('fails when a README states another Node.js floor', () => {
    const text = stated();
    const coreReadme = text.coreReadme.replace('Node.js 20.16.0 or', 'Node.js 20.10.0 or');
    expect(floorProblems({ ...text, coreReadme })[0]).toContain('states 20.10.0 or 22.3.0');
    expect(floorProblems({ ...text, coreReadme })[0]).toContain('scripts/lib/consumer-floors.mjs');
  });

  test('fails when a README states another TypeScript or Vite floor', () => {
    const text = stated();
    const reactReadme = text.reactReadme
      .replace('TypeScript 5.7 or later', 'TypeScript 5.5 or later')
      .replace('use version 7.0.0 or later', 'use version 6.4.6 or later');
    expect(floorProblems({ ...text, reactReadme })).toHaveLength(2);
    const vueReadme = text.vueReadme.replace('TypeScript 5.7 or later', 'TypeScript 5.5 or later');
    expect(floorProblems({ ...text, vueReadme })[0]).toContain('packages/vue/README.md');
  });

  test('fails when a webpack or React 18 floor differs', () => {
    const text = stated();
    const coreReadme = text.coreReadme.replace('webpack 5.83.0 or later', 'webpack 5.0.0 or later');
    expect(floorProblems({ ...text, coreReadme })[0]).toContain('webpack');
    const reactReadme = text.reactReadme.replace('18.2.6 or later', '18.2.0 or later');
    expect(floorProblems({ ...text, reactReadme })[0]).toContain('@types/react');
    expect(floorProblems({ ...text, reactPeer: '^19.0.0' })[0]).toContain('react peer');
  });

  test('fails when the tested floor moves and the READMEs stay', () => {
    expect(floorProblems(stated(), { ...FLOORS, vite: '7.1.0' }).length).toBeGreaterThan(0);
  });

  test('fails when the engines field or the CI job leaves out a floor', () => {
    const text = stated();
    expect(floorProblems({ ...text, engines: '>=20' })[0]).toContain('engines');
    const ci = text.ci.replace('floor: [tools, 20.16.0, 22.3.0]', 'floor: [tools, 22.3.0]');
    expect(floorProblems({ ...text, ci })[0]).toContain('floors job');
  });
});
