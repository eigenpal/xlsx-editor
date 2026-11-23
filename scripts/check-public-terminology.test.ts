// Tests for scripts/check-public-terminology.mjs, after PPTX Editor's test of its check at
// 8dc88c4d. Each case runs the script in a new git repository.
import { expect, test } from 'bun:test';
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { spawnSync } from 'node:child_process';

function check(files: Record<string, string>) {
  const root = mkdtempSync(join(tmpdir(), 'xlsx-terminology-'));
  try {
    const all: Record<string, string> = {
      'README.md': 'Spreadsheet editor',
      'scripts/check-public-terminology.mjs': readFileSync(
        join(import.meta.dir, 'check-public-terminology.mjs'),
        'utf8'
      ),
      ...files,
    };
    for (const [path, text] of Object.entries(all)) {
      const target = join(root, path);
      mkdirSync(dirname(target), { recursive: true });
      writeFileSync(target, text);
    }
    spawnSync('git', ['init', '-q'], { cwd: root });
    return spawnSync('node', [join(root, 'scripts/check-public-terminology.mjs')], {
      cwd: root,
      encoding: 'utf8',
    });
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
}

const app = ['Ex', 'cel'].join('');
const vendor = ['Micro', 'soft'].join('');
const control = ['Active', 'X'].join('');

test('passes plain product text', () => {
  const result = check({ 'docs/guide.md': 'Open a workbook and edit a cell.' });
  expect(result.status).toBe(0);
  expect(result.stdout).toContain('Public terminology check passed');
});

test('refuses the spreadsheet app name in any tracked text file', () => {
  const result = check({
    'docs/guide.md': `Open the file in ${app}.`,
    'packages/core/src/a.ts': `// Matches ${app} rounding.`,
  });
  expect(result.status).toBe(1);
  expect(result.stderr).toContain('docs/guide.md:1');
  expect(result.stderr).toContain('packages/core/src/a.ts:1');
});

test('does not match a longer word that starts with the name', () => {
  const result = check({ 'docs/guide.md': `An ${app.toLowerCase()}lent result.` });
  expect(result.status).toBe(0);
});

test('refuses the names inside identifiers and with separators', () => {
  const forms = [
    `is${app}Compat`,
    `${app.toUpperCase()}_EPOCH`,
    `${vendor}${app}`,
    ['Google', 'Sheets'].join('\u00a0'),
    ['Libre', 'Office'].join('-'),
    ['Power', 'Point'].join(' '),
    ['O', '365'].join(''),
  ];
  for (const form of forms) {
    const result = check({ 'packages/core/src/a.ts': `const x = '${form}';` });
    expect(result.status, form).toBe(1);
  }
});

test('passes format-defined and ordinary text that looks close', () => {
  const lines = [
    `urn:schemas-${vendor.toLowerCase()}-com:vml`,
    'const activeXfId = 0;',
    'the active xf of the cell',
    'an interactive x axis',
    `An ${app.toLowerCase()}lent result.`,
    `xmlns:x="urn:schemas-${vendor.toLowerCase()}-com:office:${app.toLowerCase()}"`,
    `http://schemas.${vendor.toLowerCase()}.com/office/2006/${control}`,
    'const hasOfficeTheme = true;',
    'This parser excels at sparse sheets.',
    'const activeX = cell.x + 1;',
    'an active x-coordinate',
    'the lotus position',
    'forms office',
    'const apiWorkbook = load();',
    'const multiWorker = 2;',
    '@media (-ms-high-contrast: active) {}',
    'locale ms-MY',
    'It reads only Office Open XML packages.',
    'const isStringNumeric = true;',
    'const isBreakingSoftHyphen = false;',
    `${app}sior`,
  ];
  for (const line of lines) {
    const result = check({ 'docs/guide.md': line });
    expect(result.status, line).toBe(0);
  }
});

test('refuses more product names and a file name', () => {
  const forms = [
    ['Only', 'Office'].join(''),
    ['Lo', 'tus123'].join(''),
    ['Power', ' Query'].join(''),
    ["Apple's", ' Numbers'].join(''),
    ['s', 'office --headless'].join(''),
    `${control}Control`,
    ['Google', ' Docs'].join(''),
    ['Smart', 'sheet'].join(''),
    ['MS', '-OI29500'].join(''),
    ['one', 'drive'].join(''),
    ['Google', '%20Sheets'].join(''),
    ['Star', 'Office'].join(''),
    ['a Google', ' Sheet'].join(''),
  ];
  for (const form of forms) {
    const result = check({ 'docs/guide.md': form });
    expect(result.status, form).toBe(1);
  }
  const named = check({ [`docs/${app}-compat.png`]: '' });
  expect(named.status).toBe(1);
  expect(named.stderr).toContain('(file name)');
});

test('reads XML and relationship parts', () => {
  const result = check({ 'fixtures/app.xml': `<Application>${app}</Application>` });
  expect(result.status).toBe(1);
  expect(result.stderr).toContain('fixtures/app.xml:1');
});

test('refuses a vendor URL that is not a namespace URI', () => {
  const result = check({ 'docs/guide.md': `See https://learn.${vendor.toLowerCase()}.com/x.` });
  expect(result.status).toBe(1);
});

test('refuses the vendor name, but not a namespace URI that the format defines', () => {
  const host = `http://schemas.${vendor.toLowerCase()}.com/office/spreadsheetml/2009/9/main`;
  expect(check({ 'a.ts': `const NS = '${host}';` }).status).toBe(0);
  const result = check({ 'a.md': `A ${vendor} format.` });
  expect(result.status).toBe(1);
  expect(result.stderr).toContain('a.md:1');
});

test('refuses the trademarked control name in authored text', () => {
  const result = check({
    'packages/i18n/en.json': JSON.stringify({ label: `${control} controls` }),
  });
  expect(result.status).toBe(1);
  expect(result.stderr).toContain('packages/i18n/en.json:1');
});

test('skips the verbatim Ecma files and the legal texts', () => {
  const result = check({
    LICENSE: vendor,
    'CLA.md': vendor,
    'reference/ecma-376/NOTICE.txt': vendor,
    'packages/pro/LICENSE.md': vendor,
  });
  expect(result.status).toBe(0);
});

test('takes exact, format-defined exceptions from the allowlist', () => {
  const allow = (line: string) => JSON.stringify({ allowed: [{ file: 'src/sinks.ts', line }] });
  const files = { 'src/sinks.ts': `  '${control}',\n` };
  expect(
    check({
      ...files,
      'scripts/public-terminology-allow.json': allow(`'${control}',`),
    }).status
  ).toBe(0);
  const changed = check({
    ...files,
    'scripts/public-terminology-allow.json': allow(`'${control}'`),
  });
  expect(changed.status).toBe(1);
  expect(changed.stderr).toContain('src/sinks.ts:1');
  expect(changed.stderr).toContain('Stale entries');
});

test('allows a format-defined part path as a file name entry', () => {
  const path = `xl/${control}/${control}1.xml`;
  const result = check({
    [path]: '<x/>',
    'scripts/public-terminology-allow.json': JSON.stringify({
      allowed: [{ file: path, line: '(file name)' }],
    }),
  });
  expect(result.status).toBe(0);
});

test('reads files of any type except compressed and image formats', () => {
  const result = check({ 'fixtures/ref.pdf': `/Producer (${vendor} ${app})` });
  expect(result.status).toBe(1);
  expect(result.stderr).toContain('fixtures/ref.pdf:1');
});

test('passes the exact npm package names of the build tools, and nothing more', () => {
  const scope = `@${vendor.toLowerCase()}`;
  const pass = check({
    'package.json': `"${scope}/api-extractor": "^7.60.1",`,
    'bun.lock': `"${scope}/tsdoc-config": ["${scope}/tsdoc-config@0.18.3", { "${scope}/tsdoc": "0.17.1" }],`,
  });
  expect(pass.status).toBe(0);
  const fail = check({
    'docs/guide.md': `Install ${scope}/other-tool.`,
    'scripts/a.mjs': `// ${scope}/api-extractor-cli`,
  });
  expect(fail.status).toBe(1);
  expect(fail.stderr).toContain('docs/guide.md:1');
  expect(fail.stderr).toContain('scripts/a.mjs:1');
});
