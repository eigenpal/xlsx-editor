import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import {
  bareImportRequests,
  exposesWorkspacePath,
  hasEcmaLegalComment,
  listingFieldProblems,
  missingFileRemedy,
  relativeImportRequests,
  requiredLicenseFiles,
  sideEffectsProblem,
  unexportedRequestProblem,
} from '../lib/package-artifact-rules.mjs';
import { PACKAGES } from '../lib-packages.mjs';

const repoRoot = resolve(import.meta.dir, '../..');

describe('bareImportRequests', () => {
  test('finds a side-effect import with and without a space', () => {
    expect(bareImportRequests('import "a";import"@b/c";')).toEqual(['a', '@b/c']);
  });

  test('finds static, dynamic, and require forms', () => {
    const text = 'import x from "a";await import("b");require("c");export * from "d";';
    expect(bareImportRequests(text)).toEqual(['a', 'b', 'c', 'd']);
  });

  test('skips relative requests', () => {
    expect(bareImportRequests('import "./a.js";import x from "../b.js";')).toEqual([]);
  });

  test('finds a triple-slash type reference', () => {
    expect(bareImportRequests('/// <reference types="left-pad" />\nexport {};')).toEqual([
      'left-pad',
    ]);
    expect(bareImportRequests('/// <reference path="./a.d.ts" />')).toEqual([]);
  });

  test('finds a template literal request with no substitution', () => {
    expect(bareImportRequests('await import(`left-pad`);require(`b`);')).toEqual(['left-pad', 'b']);
    expect(bareImportRequests('await import(`./${name}.js`);')).toEqual([]);
  });
});

describe('relativeImportRequests', () => {
  test('finds static, re-export, side-effect, and dynamic forms', () => {
    const text = 'import a from "./a.js";export*from"./b.js";import"./c.js";import("../d.js");';
    expect(relativeImportRequests(text)).toEqual(['./a.js', './b.js', './c.js', '../d.js']);
  });

  test('skips bare requests', () => {
    expect(relativeImportRequests('import a from "react";')).toEqual([]);
  });
});

describe('listingFieldProblems', () => {
  test('refuses a missing field on an adapter', () => {
    expect(listingFieldProblems({ name: '@xlsx-editor.dev/react' })).toEqual([
      '@xlsx-editor.dev/react has no keywords',
      '@xlsx-editor.dev/react has no https homepage',
    ]);
  });

  test('leaves core out, as the DOCX Editor core manifest has neither field', () => {
    expect(listingFieldProblems({ name: '@xlsx-editor.dev/core' })).toEqual([]);
  });

  test('accepts each published manifest', () => {
    for (const pkg of PACKAGES) {
      const file = resolve(repoRoot, pkg.root, 'package.json');
      expect(listingFieldProblems(JSON.parse(readFileSync(file, 'utf8')))).toEqual([]);
    }
  });
});

describe('unexportedRequestProblem', () => {
  const manifests = new Map([
    ['@xlsx-editor.dev/i18n', { exports: { '.': './a.js', './en': './b.js' } }],
    ['@xlsx-editor.dev/core', { exports: { './store': './c.js' } }],
  ]);

  test('accepts the subpaths of the export map', () => {
    expect(unexportedRequestProblem('@xlsx-editor.dev/i18n', manifests)).toBeNull();
    expect(unexportedRequestProblem('@xlsx-editor.dev/i18n/en', manifests)).toBeNull();
  });

  test('refuses a subpath of any published package that its export map does not hold', () => {
    expect(unexportedRequestProblem('@xlsx-editor.dev/i18n/missing', manifests)).toContain(
      '@xlsx-editor.dev/i18n does not export'
    );
    expect(unexportedRequestProblem('@xlsx-editor.dev/core', manifests)).toContain('not export');
  });

  test('leaves other packages to the dependency rule', () => {
    expect(unexportedRequestProblem('react/jsx-runtime', manifests)).toBeNull();
  });
});

describe('exposesWorkspacePath', () => {
  test('finds a source path with and without a leading slash', () => {
    expect(exposesWorkspacePath('see `packages/core/src/store/a.ts` at 1')).toBe(true);
    expect(exposesWorkspacePath('from "/repo/packages/react/src/a"')).toBe(true);
    expect(exposesWorkspacePath('from "../../core/src/a"')).toBe(true);
  });

  test('accepts text that names no workspace source', () => {
    expect(exposesWorkspacePath('`core/src/automation/a.ts` of that repository')).toBe(false);
    expect(exposesWorkspacePath('node_modules/my-packages/core/src/a')).toBe(false);
  });
});

describe('sideEffectsProblem', () => {
  const css = { './styles.css': './a.css' };

  test('accepts the shipped value of a package with a stylesheet', () => {
    expect(sideEffectsProblem({ name: 'a', sideEffects: ['*.css'], exports: css })).toBeNull();
  });

  test('refuses false, a missing field, and a wider list on a package with a stylesheet', () => {
    for (const sideEffects of [false, undefined, ['*.css', '*.js']])
      expect(sideEffectsProblem({ name: 'a', sideEffects, exports: css })).toContain('a sets');
  });

  test('requires false on a package with no stylesheet', () => {
    expect(sideEffectsProblem({ name: 'a', sideEffects: false, exports: {} })).toBeNull();
    expect(sideEffectsProblem({ name: 'a', exports: {} })).toContain('expected false');
  });

  test('accepts each published manifest', () => {
    for (const pkg of PACKAGES) {
      const file = resolve(repoRoot, pkg.root, 'package.json');
      expect(sideEffectsProblem(JSON.parse(readFileSync(file, 'utf8')))).toBeNull();
    }
  });
});

describe('missingFileRemedy', () => {
  test('names the files entry for a file on disk that the tarball leaves out', () => {
    expect(missingFileRemedy('licenses/a.txt', true, 'core')).toBe(
      ' Add it to `files` in packages/core/package.json.'
    );
    expect(missingFileRemedy('THIRD_PARTY_NOTICES.md', true, 'react')).toContain(
      'packages/react/package.json'
    );
  });

  test('names the notices command for notices that are not on disk', () => {
    expect(missingFileRemedy('THIRD_PARTY_NOTICES.md', false, 'core')).toBe(
      ' Run `bun run notices:generate` after `bun run build:packages`.'
    );
  });

  test('names a restore for a tracked file that is not on disk', () => {
    expect(missingFileRemedy('README.md', false, 'core')).toBe(
      ' Restore packages/core/README.md from git.'
    );
  });
});

describe('requiredLicenseFiles', () => {
  test('requires each notice and skips dotfiles such as the system files that npm never publishes', () => {
    expect(requiredLicenseFiles(['.DS_Store', 'HarfBuzz-MIT.txt', '._HarfBuzz-MIT.txt'])).toEqual([
      'licenses/HarfBuzz-MIT.txt',
    ]);
  });
});

describe('hasEcmaLegalComment', () => {
  const line = '© Ecma International';

  test('accepts the copyright line inside one legal comment', () => {
    expect(
      hasEcmaLegalComment(`/*!\n * Table.\n *\n * ${line} 2016\n */\nexport const T = 1;`)
    ).toBe(true);
    expect(hasEcmaLegalComment(`/*! other */\n/*! ${line} */`)).toBe(true);
  });

  test('refuses the line outside a legal comment', () => {
    expect(hasEcmaLegalComment(`/* ${line} */`)).toBe(false);
    expect(hasEcmaLegalComment(`// ${line}`)).toBe(false);
    expect(hasEcmaLegalComment(`/*! Table. */ const a = '${line}'; /* x */`)).toBe(false);
    expect(hasEcmaLegalComment(`/*! ${line}`)).toBe(false);
  });
});
