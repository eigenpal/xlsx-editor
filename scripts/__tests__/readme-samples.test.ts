import { describe, expect, test } from 'bun:test';
import { readdirSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { extractSample, GUIDE_SAMPLES } from '../lib/readme-samples.mjs';

const repoRoot = resolve(import.meta.dir, '../..');
const readme = (name: string) =>
  readFileSync(resolve(repoRoot, 'packages', name, 'README.md'), 'utf8');

describe('extractSample', () => {
  test('returns the fenced block after the marker', () => {
    const text = 'Intro\n\n<!-- sample: one -->\n\n```ts\nconst a = 1;\n```\n\nEnd\n';
    expect(extractSample(text, 'one')).toBe('const a = 1;\n');
  });

  test('refuses a missing marker', () => {
    expect(() => extractSample('```ts\nx\n```\n', 'one')).toThrow('expected one marker, found 0');
  });

  test('refuses a repeated marker', () => {
    const text = '<!-- sample: one -->\n```ts\nx\n```\n<!-- sample: one -->\n```ts\ny\n```\n';
    expect(() => extractSample(text, 'one')).toThrow('expected one marker, found 2');
  });

  test('refuses a marker with no fenced block', () => {
    expect(() => extractSample('<!-- sample: one -->\n\nText\n', 'one')).toThrow(
      'no fenced block follows'
    );
  });

  test('skips a prettier-ignore comment between the marker and the block', () => {
    const text = '<!-- sample: one -->\n\n<!-- prettier-ignore -->\n```ts\nconst a = 1;\n```\n';
    expect(extractSample(text, 'one')).toBe('const a = 1;\n');
  });

  test('refuses a fenced block that does not end', () => {
    expect(() => extractSample('<!-- sample: one -->\n```ts\nx\n', 'one')).toThrow('does not end');
  });
});

describe('the README samples that check:consumer-install builds', () => {
  test('the core README marks the host element and the opening script', () => {
    const core = readme('core');
    expect(extractSample(core, 'core-host')).toContain('class="xlsx-editor xlsx-editor-viewport"');
    expect(extractSample(core, 'core-open')).toContain('createXlsxEditor');
  });

  test('the React README marks the editor and the custom layout', () => {
    const react = readme('react');
    expect(extractSample(react, 'react-editor')).toContain('<XlsxEditor document={bytes}');
    expect(extractSample(react, 'react-layout')).toContain('<XlsxEditor.Root document={bytes}>');
  });
});

describe('the React guide samples that check:consumer-install type-checks', () => {
  const docs = resolve(repoRoot, 'packages/react/docs');

  test('each listed sample has its marker', () => {
    for (const { guide, name } of GUIDE_SAMPLES) {
      expect(
        extractSample(readFileSync(resolve(docs, guide), 'utf8'), name).length
      ).toBeGreaterThan(0);
    }
  });

  test('each code block of a guide has a marker that the check reads', () => {
    const read = new Set([...GUIDE_SAMPLES.map(({ name }) => name), 'react-unsaved']);
    for (const guide of readdirSync(docs).filter((file) => file.endsWith('.md'))) {
      const lines = readFileSync(resolve(docs, guide), 'utf8').split('\n');
      let inside = false;
      for (const [index, line] of lines.entries()) {
        if (!line.startsWith('```')) continue;
        inside = !inside;
        if (!inside) continue;
        const marker = lines
          .slice(0, index)
          .reverse()
          .find((text) => text.trim() !== '' && text.trim() !== '<!-- prettier-ignore -->');
        const name = /^<!-- sample: (.+) -->$/.exec(marker?.trim() ?? '')?.[1];
        expect({ guide, line: index + 1, read: name !== undefined && read.has(name) }).toEqual({
          guide,
          line: index + 1,
          read: true,
        });
      }
    }
  });
});
