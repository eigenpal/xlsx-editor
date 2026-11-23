// Tests for scripts/check-security-grep.mjs. Each case runs the script in a new temporary folder.
import { describe, expect, test } from 'bun:test';
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { spawnSync } from 'node:child_process';

const SCRIPT = join(import.meta.dir, 'check-security-grep.mjs');

function grep(files: Record<string, string>) {
  const root = mkdtempSync(join(tmpdir(), 'xlsx-security-grep-'));
  try {
    mkdirSync(join(root, 'packages/core/src'), { recursive: true });
    for (const [path, text] of Object.entries(files)) {
      mkdirSync(dirname(join(root, path)), { recursive: true });
      writeFileSync(join(root, path), text);
    }
    return spawnSync('node', [SCRIPT], { cwd: root, encoding: 'utf8' });
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
}

test('passes code with no sink', () => {
  const result = grep({ 'packages/core/src/a.ts': 'node.textContent = value;' });
  expect(result.status).toBe(0);
});

test('refuses each sink form', () => {
  const sinks = [
    'el.innerHTML = s;',
    "el['innerHTML'] = s;",
    'imgEl.src = target;',
    "el.setAttribute('href', t);",
    'const a = <a href={target}>x</a>;',
    '<img :src="u" />',
    "el.setAttributeNS(XLINK, 'xlink:href', t);",
    'frame.srcdoc = html;',
    'const style = { fontFamily: `${name}` };',
    'range.createContextualFragment(html);',
    "new DOMParser().parseFromString(s, 'text/html');",
    'defaultView.open(url);',
    'ownerDocument.write(s);',
    'defaultView.location.assign(url);',
    'el.setHTMLUnsafe(s);',
    'el.setAttribute(`href`, t);',
    'defaultView.location = url;',
  ];
  for (const sink of sinks) {
    const result = grep({ 'packages/core/src/a.ts': sink });
    expect(result.status, sink).toBe(1);
  }
});

test('passes ordinary code that looks close to a sink', () => {
  const lines = [
    'let allocation = n;',
    'const location = { row, col };',
    'const relocation = 0;',
    'reader.src = bytes;',
    "parser.parseFromString(s, 'application/xml');",
    'if (link.href === other) return;',
    "el.setAttribute('href', sanitizeHref(target));",
    '// Plain text only: no `innerHTML`',
    'this.location = next;',
    'view.open();',
  ];
  for (const line of lines) {
    const result = grep({ 'packages/core/src/a.ts': line });
    expect(result.status, line).toBe(0);
  }
});

describe('the generated blob URL exception', () => {
  const FILE = 'examples/vite/src/download.ts';
  const CODE = [
    'export function downloadBlob(blob: Blob, name: string): void {',
    '  const objectUrl = URL.createObjectURL(blob);',
    "  const link = document.createElement('a');",
    "  link.setAttribute('href', objectUrl);",
    '  link.download = name;',
    '  link.click();',
    '  setTimeout(() => URL.revokeObjectURL(objectUrl), 1000);',
    '}',
  ];
  const status = (lines: string[], file = FILE) => grep({ [file]: lines.join('\n') }).status;

  test('passes the exact code, with whole-line comments and blank lines', () => {
    expect(status(CODE)).toBe(0);
    expect(status(['// The download.', '', ...CODE, ''])).toBe(0);
  });

  test('passes the file of the demo', () => {
    const shipped = readFileSync(join(import.meta.dir, '..', FILE), 'utf8');
    expect(grep({ [FILE]: shipped }).status).toBe(0);
  });

  test('passes CRLF line ends, and refuses a lone CR with the reason and the remedy', () => {
    expect(grep({ [FILE]: CODE.join('\r\n') }).status).toBe(0);
    expect(grep({ [FILE]: `// The download.\r\n${CODE.join('\r\n')}\r\n` }).status).toBe(0);
    const hidden = grep({ [FILE]: `// note\rconst other = 1;\n${CODE.join('\n')}` });
    expect(hidden.status).toBe(1);
    expect(hidden.stderr).toContain('a lone CR, U+2028, or U+2029');
    expect(hidden.stderr).toContain('Restore the code of the file to `BLOB_DOWNLOAD_CODE`');
  });

  test('names the changed line and the remedy when the code differs', () => {
    const changed = CODE.map((line, i) => (i === 4 ? '  link.download = `${name}.xlsx`;' : line));
    const result = grep({ [FILE]: ['// The download.', ...changed].join('\n') });
    expect(result.status).toBe(1);
    expect(result.stderr).toContain(`${FILE}:5: link.setAttribute('href', objectUrl);`);
    expect(result.stderr).toContain(
      'line 6 is "link.download = `${name}.xlsx`;", but the permitted line is "link.download = name;"'
    );
    expect(result.stderr).toContain('Restore the code of the file to `BLOB_DOWNLOAD_CODE`');
  });

  test('names code that ends early or continues after the permitted code', () => {
    const early = grep({ [FILE]: CODE.slice(0, 5).join('\n') });
    expect(early.stderr).toContain('the code ends before the permitted line "link.click();"');
    const extra = grep({ [FILE]: [...CODE, 'const other = 1;'].join('\n') });
    expect(extra.stderr).toContain('line 9 adds code after the permitted code');
  });

  test('refuses the same code in another file', () => {
    expect(status(CODE, 'examples/vite/src/other.ts')).toBe(1);
  });

  test('refuses a second function with the same sink line', () => {
    const second = [
      'export function openFromFile(objectUrl: string, link: HTMLAnchorElement): void {',
      "  link.setAttribute('href', objectUrl);",
      '}',
    ];
    expect(status([...CODE, ...second])).toBe(1);
  });

  test('refuses code that takes a string, or a URL that is not the object URL', () => {
    const replace = (from: number, line: string) => CODE.map((l, i) => (i === from ? line : l));
    expect(
      status(replace(0, 'export function downloadBlob(blob: string, name: string): void {'))
    ).toBe(1);
    expect(status(replace(1, '  const objectUrl = blob;'))).toBe(1);
    expect(status(replace(3, "  link.setAttribute('href', name);"))).toBe(1);
  });

  test('refuses the needed lines inside a block comment over other code', () => {
    const hidden = [
      '/*',
      ...CODE.slice(0, 2),
      '*/',
      'export function downloadBlob(objectUrl: string, link: HTMLAnchorElement): void {',
      "  link.setAttribute('href', objectUrl);",
      '}',
    ];
    expect(status(hidden)).toBe(1);
  });

  test('refuses code that a one-line block comment or a line separator hides', () => {
    expect(status(['/* */ const other = 1;', ...CODE])).toBe(1);
    expect(status([`// note const other = 1;`, ...CODE])).toBe(1);
  });
});

test('reads JavaScript sources and skips tests', () => {
  expect(grep({ 'packages/core/src/z.js': 'el.innerHTML = s;' }).status).toBe(1);
  expect(grep({ 'packages/core/src/a.test.ts': 'el.innerHTML = s;' }).status).toBe(0);
});
