import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { APPROXIMATIONS } from '../approximations.ts';

const PUBLIC_LIST = join(
  import.meta.dir,
  '../../../../../docs/site/content/known-approximations.mdx'
);

/** The bold titles of the items under "Known approximations" in the public list. */
function documentedTitles(): string[] {
  const doc = readFileSync(PUBLIC_LIST, 'utf8');
  const start = doc.indexOf('## Known approximations');
  expect(start).toBeGreaterThanOrEqual(0);
  const end = doc.indexOf('\n## ', start + 3);
  const section = doc.slice(start, end < 0 ? undefined : end);
  return [...section.matchAll(/^- \*\*(.+?)\.\*\*/gm)].map((m) => m[1]!);
}

describe('the approximation list', () => {
  test('lists exactly the approximations of the public list, in order', () => {
    expect(APPROXIMATIONS.map((a) => a.title)).toEqual(documentedTitles());
  });

  test('each unmeasured case appears word for word in the text of its public item', () => {
    const doc = readFileSync(PUBLIC_LIST, 'utf8');
    for (const a of APPROXIMATIONS) {
      const start = doc.indexOf(`- **${a.title}.**`);
      const end = doc.indexOf('\n- **', start + 1);
      const item = doc.slice(start, end < 0 ? undefined : end);
      for (const text of a.unmeasured ?? []) expect(item, a.id).toContain(text);
    }
  });

  test('ids are unique', () => {
    expect(new Set(APPROXIMATIONS.map((a) => a.id)).size).toBe(APPROXIMATIONS.length);
  });
});
