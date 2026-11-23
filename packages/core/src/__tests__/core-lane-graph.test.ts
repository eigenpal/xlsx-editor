// Taken through pptx-editor/packages/core/src/__tests__/core-lane-graph.test.ts at 1943d2c9 (EigenPal), adapted for XLSX Editor.
// Differences: the export test also checks that each subpath points at its built `types` and
// `import` files, and a test checks that the tsconfig `paths` name the source of every subpath,
// because `scripts/build-core-declarations.mjs` reads each subpath's source from those paths.

import { describe, expect, test } from 'bun:test';
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join, resolve, relative } from 'node:path';
import {
  CORE_LANES,
  CORE_PACKAGE_ROOT as CORE,
  existingLanePath,
  laneTopologicalOrder,
  type LaneName,
} from './core-lane-graph.ts';

function sources(dir: string): string[] {
  const out: string[] = [];
  for (const name of readdirSync(dir)) {
    const p = join(dir, name);
    if (statSync(p).isDirectory()) {
      if (name !== '__tests__') out.push(...sources(p));
    } else if (/\.tsx?$/.test(name)) out.push(p);
  }
  return out;
}

function laneOf(file: string): LaneName | undefined {
  const rel = relative(join(CORE, 'src'), file).split('/')[0];
  return (Object.keys(CORE_LANES) as LaneName[]).find(
    (l) => CORE_LANES[l].directory === `src/${rel}`
  );
}

describe('core lane graph', () => {
  test('the DAG is acyclic', () => {
    expect(laneTopologicalOrder()).toHaveLength(Object.keys(CORE_LANES).length);
  });

  test('every lane resolves to real source', () => {
    for (const lane of Object.keys(CORE_LANES) as LaneName[])
      expect(sources(existingLanePath(lane)).length).toBeGreaterThan(0);
  });

  test('relative imports only cross declared edges', () => {
    const violations: string[] = [];
    for (const lane of Object.keys(CORE_LANES) as LaneName[]) {
      for (const file of sources(existingLanePath(lane))) {
        const text = readFileSync(file, 'utf8');
        for (const m of text.matchAll(/from\s+'(\.[^']+)'/g)) {
          const target = laneOf(resolve(file, '..', m[1]));
          if (target && target !== lane && !CORE_LANES[lane].mayImport.includes(target))
            violations.push(`${relative(CORE, file)} imports ${target}`);
        }
      }
    }
    expect(violations).toEqual([]);
  });

  test('every lane is a package subpath and a build entry, and nothing else is', () => {
    const manifest = JSON.parse(readFileSync(join(CORE, 'package.json'), 'utf8')) as {
      exports: Record<string, unknown>;
    };
    const tsup = readFileSync(join(CORE, 'tsup.config.ts'), 'utf8');
    const entryBlock = /entry:\s*\{([\s\S]*?)\}/.exec(tsup)?.[1] ?? '';
    const entries = new Map(
      [...entryBlock.matchAll(/(\w+):\s*'([^']+)'/g)].map((m) => [m[1], m[2]] as const)
    );
    const codeSubpaths = Object.keys(manifest.exports).filter(
      (k) => k !== '.' && !/\.(css|json|wasm)$/.test(k)
    );
    const laneSubpaths = (Object.keys(CORE_LANES) as LaneName[]).map((l) => CORE_LANES[l].subpath);
    expect(codeSubpaths.sort()).toEqual(laneSubpaths.sort());
    for (const lane of Object.keys(CORE_LANES) as LaneName[]) {
      const name = CORE_LANES[lane].subpath.slice(2);
      expect(entries.get(name)).toBe(`${CORE_LANES[lane].directory}/index.ts`);
      expect(manifest.exports[CORE_LANES[lane].subpath]).toEqual({
        types: `./dist/${name}.d.ts`,
        import: `./dist/${name}.js`,
      });
    }
  });

  test('tsconfig paths name the source of every subpath, for the declaration build', () => {
    // scripts/build-core-declarations.mjs reads each subpath's source from these paths.
    const tsconfig = JSON.parse(readFileSync(join(CORE, 'tsconfig.json'), 'utf8')) as {
      compilerOptions: { paths: Record<string, string[]> };
    };
    const own = Object.entries(tsconfig.compilerOptions.paths).filter(([specifier]) =>
      specifier.startsWith('@xlsx-editor.dev/core')
    );
    const expected = [
      ['@xlsx-editor.dev/core', ['./src/index.ts']],
      ...(Object.keys(CORE_LANES) as LaneName[]).map((lane) => [
        `@xlsx-editor.dev/core${CORE_LANES[lane].subpath.slice(1)}`,
        [`./${CORE_LANES[lane].directory}/index.ts`],
      ]),
    ];
    expect(Object.fromEntries(own)).toEqual(Object.fromEntries(expected));
  });

  test('neutral lanes do not name DOM types or globals', () => {
    const violations: string[] = [];
    for (const lane of Object.keys(CORE_LANES) as LaneName[]) {
      if (CORE_LANES[lane].environment !== 'neutral') continue;
      for (const file of sources(existingLanePath(lane))) {
        // Comments are prose: "…per document." names no DOM global.
        const text = readFileSync(file, 'utf8')
          .replace(/\/\*[\s\S]*?\*\//g, '')
          .replace(/\/\/.*$/gm, '');
        if (/\b(document|window|navigator)\.|\b(HTML\w*Element|SVG\w*Element)\b/.test(text))
          violations.push(relative(CORE, file));
      }
    }
    expect(violations).toEqual([]);
  });
});
