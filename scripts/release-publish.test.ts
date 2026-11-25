import { expect, test } from 'bun:test';
import { isStable, retargetPlan, type PlannedRelease } from './release-publish.mjs';

const release = (name: string, tag: string): PlannedRelease => ({
  kind: 'publish',
  name,
  version: '0.1.0-beta.2',
  access: 'public',
  tag,
});

test('a beta of a package with no stable version goes to latest', () => {
  const plan = [[release('@scope/core', 'beta')], [release('@scope/react', 'beta')]];
  const result = retargetPlan(plan, 'beta', () => false);
  expect(result.flat().map((entry) => entry.tag)).toEqual(['latest', 'latest']);
  expect(plan[0][0].tag).toBe('beta');
});

test('a beta of a package with a stable version keeps the prerelease dist-tag', () => {
  const plan = [[release('@scope/core', 'beta'), release('@scope/react', 'beta')]];
  const result = retargetPlan(plan, 'beta', (name) => name === '@scope/core');
  expect(result.flat().map((entry) => entry.tag)).toEqual(['beta', 'latest']);
});

test('outside prerelease mode the plan does not change', () => {
  const plan = [[release('@scope/core', 'latest')]];
  expect(retargetPlan(plan, null, () => false)).toBe(plan);
});

test('a tag-only entry and another dist-tag do not change', () => {
  const tagOnly: PlannedRelease = { kind: 'tag-only', name: '@scope/internal', version: '1.0.0' };
  const plan = [[tagOnly, release('@scope/core', 'next')]];
  expect(retargetPlan(plan, 'beta', () => false)).toEqual(plan);
});

test('only a version with no prerelease part is stable', () => {
  expect(isStable('0.1.0')).toBe(true);
  expect(isStable('0.1.0-beta.1')).toBe(false);
  expect(isStable('0.0.0-placeholder.0')).toBe(false);
  expect(isStable('0.0.0-stage')).toBe(false);
});
