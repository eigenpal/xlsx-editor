// Each known gap names its cause: an approximation whose public item names the fixture and the
// recorded value (AGENTS.md, "Fidelity"). A gap that moves, or a new one, fails here until the
// public list says so.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { APPROXIMATIONS } from '../../packages/core/src/layout/approximations.ts';
import { GAP_CAUSES, KNOWN_GAPS, type FailureKind } from './thresholds.ts';

const PUBLIC_LIST = readFileSync(
  join(import.meta.dir, '..', '..', 'docs', 'site', 'content', 'known-approximations.mdx'),
  'utf8'
);

/** The text of the public item with this bold title. */
function itemText(title: string): string {
  const start = PUBLIC_LIST.indexOf(`- **${title}.**`);
  expect(start).toBeGreaterThanOrEqual(0);
  const end = PUBLIC_LIST.indexOf('\n- **', start + 1);
  return PUBLIC_LIST.slice(start, end < 0 ? undefined : end);
}

/** The value of a check as the public list writes it. */
function written(check: FailureKind, value: number): string {
  if (check === 'pixels') return `${(value * 100).toFixed(3)}% of the pixels`;
  if (check === 'baseline' || check === 'line-position')
    return `${value} point${value === 1 ? '' : 's'}`;
  return String(value);
}

describe('known gap causes', () => {
  // Every list: the list of each raster toolchain of each font environment.
  const gaps = Object.values(KNOWN_GAPS).flatMap((lists) =>
    (lists ?? []).flatMap((list) =>
      Object.entries(list.fixtures).flatMap(([fixture, gap]) =>
        Object.entries(gap).map(([check, value]) => ({
          fixture,
          check: check as FailureKind,
          value: value as number,
        }))
      )
    )
  );

  test('every known gap has an approximation whose item names the fixture and the value', () => {
    expect(gaps.length).toBeGreaterThan(0);
    for (const { fixture, check, value } of gaps) {
      const id = GAP_CAUSES[fixture]?.[check];
      const approximation = APPROXIMATIONS.find((a) => a.id === id);
      expect(approximation, `${fixture} ${check}`).toBeDefined();
      const text = itemText(approximation!.title);
      expect(text).toContain(`\`${fixture}\``);
      expect(text).toContain(written(check, value));
    }
  });

  test('every cause names a recorded gap', () => {
    for (const [fixture, checks] of Object.entries(GAP_CAUSES))
      for (const check of Object.keys(checks))
        expect(gaps.some((g) => g.fixture === fixture && g.check === check)).toBe(true);
  });
});
