import { describe, expect, test } from 'bun:test';
import {
  deleteError,
  moduleLoadError,
  packageError,
  withPeriod,
  writeError,
} from '../lib/notices-errors.mjs';

describe('withPeriod', () => {
  test('ends a message with one period', () => {
    expect(withPeriod('Unexpected EOF')).toBe('Unexpected EOF.');
    expect(withPeriod('Restore the file.')).toBe('Restore the file.');
  });
});

describe('moduleLoadError', () => {
  test('names the module and the files that it reads, and the fix', () => {
    expect(
      moduleLoadError(
        'scripts/vendored-notices.mjs',
        ['reference/ecma-376/NOTICE'],
        new Error('reference/ecma-376/NOTICE has no COPYRIGHT NOTICE heading.')
      )
    ).toBe(
      'scripts/vendored-notices.mjs failed to load: reference/ecma-376/NOTICE has no COPYRIGHT ' +
        'NOTICE heading. Fix `scripts/vendored-notices.mjs` or `reference/ecma-376/NOTICE`, or ' +
        'restore them from git, then run this script again.'
    );
  });
});

describe('packageError', () => {
  test('keeps the fix that a HarfBuzz notice error names', () => {
    expect(
      packageError(
        '@xlsx-editor.dev/core',
        new Error('HarfBuzz notice is missing: a.txt. Restore the file from git.')
      )
    ).toBe('@xlsx-editor.dev/core: HarfBuzz notice is missing: a.txt. Restore the file from git.');
  });

  test('names a rebuild, then a restore, for any other error', () => {
    expect(
      packageError('@xlsx-editor.dev/core', new Error("Cannot find package 'harfbuzzjs'"))
    ).toBe(
      "@xlsx-editor.dev/core: Cannot find package 'harfbuzzjs'. Run `bun install` and " +
        '`bun run build:packages`, then run this script again. If the error names a tracked ' +
        'file, fix it or restore it from git.'
    );
  });
});

describe('writeError', () => {
  test('names the file and the fix', () => {
    expect(
      writeError('packages/core/THIRD_PARTY_NOTICES.md', new Error('EACCES: permission denied'))
    ).toBe(
      'Could not write packages/core/THIRD_PARTY_NOTICES.md: EACCES: permission denied. Make the ' +
        'file and its directory writable, then run this script again.'
    );
  });
});

describe('deleteError', () => {
  test('names the file, the fix, and the risk', () => {
    expect(
      deleteError('packages/react/THIRD_PARTY_NOTICES.md', new Error('EISDIR: is a directory'))
    ).toBe(
      'Could not delete packages/react/THIRD_PARTY_NOTICES.md: EISDIR: is a directory. Delete ' +
        'it, then run this script again. Until then, the file can name licenses that the build ' +
        'does not ship.'
    );
  });
});

describe('a thrown value that is not an Error', () => {
  test('gives its text, so the run still reaches the delete', () => {
    expect(packageError('@xlsx-editor.dev/i18n', 'broken')).toBe(
      '@xlsx-editor.dev/i18n: broken. Run `bun install` and `bun run build:packages`, then run ' +
        'this script again. If the error names a tracked file, fix it or restore it from git.'
    );
    expect(writeError('a.md', undefined)).toBe(
      'Could not write a.md: undefined. Make the file and its directory writable, then run this ' +
        'script again.'
    );
  });
});
