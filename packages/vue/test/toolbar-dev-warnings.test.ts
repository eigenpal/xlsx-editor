// The toolbar's problems print in development only, once each, with the `[xlsx-editor]` prefix.
// A copy of docx-editor/packages/react/test/toolbar-dev-warnings.test.ts at 10eeaf67 (EigenPal),
// pointed at `warnToolbarProblems`, which prints a list of problems where DOCX Editor's
// `toolbarDevWarning` prints one. The twin is packages/react/src/__tests__/toolbar-dev-warnings.test.ts.

import { afterEach, expect, mock, test } from 'bun:test';
import { developmentFrom, isDevelopment } from '../src/is-development.ts';
import { warnToolbarProblems } from '../src/toolbar/host-parts.tsx';

const savedProcess = globalThis.process;
const savedEnv = process.env.NODE_ENV;
const savedWarn = console.warn;

function restoreProcess(): void {
  Object.defineProperty(globalThis, 'process', {
    value: savedProcess,
    configurable: true,
    writable: true,
  });
}

afterEach(() => {
  restoreProcess();
  process.env.NODE_ENV = savedEnv;
  console.warn = savedWarn;
});

function withoutProcess<T>(run: () => T): T {
  delete (globalThis as { process?: unknown }).process;
  try {
    return run();
  } finally {
    restoreProcess();
  }
}

test('a page without a process global counts as production and prints nothing', () => {
  const warn = mock(() => {});
  console.warn = warn;
  expect(withoutProcess(() => isDevelopment())).toBe(false);
  withoutProcess(() => warnToolbarProblems(['no process global']));
  expect(warn).not.toHaveBeenCalled();
});

test('a production build prints nothing and a development build prints once', () => {
  const warn = mock((_message: string) => {});
  console.warn = warn;
  process.env.NODE_ENV = 'production';
  expect(isDevelopment()).toBe(false);
  warnToolbarProblems(['production build']);
  expect(warn).not.toHaveBeenCalled();
  process.env.NODE_ENV = 'development';
  expect(isDevelopment()).toBe(true);
  warnToolbarProblems(['development build']);
  warnToolbarProblems(['development build']);
  expect(warn.mock.calls).toEqual([['[xlsx-editor] development build']]);
});

test('a Vite build answers through import.meta.env.DEV, in development and in production', () => {
  // Vite leaves no `process` in the browser; its flag decides either way.
  expect(developmentFrom({ viteEnv: { DEV: true }, nodeEnv: undefined, hasProcess: false })).toBe(
    true
  );
  expect(developmentFrom({ viteEnv: { DEV: false }, nodeEnv: undefined, hasProcess: false })).toBe(
    false
  );
  // A production Vite build stays silent even where a `process` shim says development.
  expect(
    developmentFrom({ viteEnv: { DEV: false }, nodeEnv: 'development', hasProcess: true })
  ).toBe(false);
});

test('without a Vite flag the Node convention decides', () => {
  expect(developmentFrom({ viteEnv: undefined, nodeEnv: 'development', hasProcess: true })).toBe(
    true
  );
  expect(developmentFrom({ viteEnv: {}, nodeEnv: 'production', hasProcess: true })).toBe(false);
  expect(developmentFrom({ viteEnv: undefined, nodeEnv: undefined, hasProcess: false })).toBe(
    false
  );
});
