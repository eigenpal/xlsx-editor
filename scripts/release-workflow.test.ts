// Adapted from pptx-editor/scripts/release-workflow.test.ts at 8e843068 (EigenPal). Differences:
// no canary, post-release, or label cases; cases for trusted publishing and the artifact
// gate. Evaluates the workflow's own conditions, so a changed `if:` that lets a skipped
// gate reach a publish fails here.
import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';

interface Step {
  id?: string;
  name?: string;
  run?: string;
  uses?: string;
  if?: string;
  with?: Record<string, string>;
  env?: Record<string, string>;
  'continue-on-error'?: boolean;
}
interface Job {
  needs?: string | string[];
  if?: string;
  'runs-on'?: string;
  permissions?: Record<string, string>;
  outputs?: Record<string, string>;
  'continue-on-error'?: boolean;
  steps: Step[];
}
const source = readFileSync(new URL('../.github/workflows/release.yml', import.meta.url), 'utf8');
const workflow = Bun.YAML.parse(source) as {
  on: Record<string, unknown>;
  concurrency: { group: string; 'cancel-in-progress': boolean };
  jobs: Record<string, Job>;
};

// Evaluate a workflow expression with the given contexts: `||`, `&&`, `!`, `==`, `!=`,
// parentheses, string literals, context paths, and status calls such as `failure()`. As in
// GitHub Actions, `||` and `&&` answer an operand, and a missing value compares unequal to a
// string. A parser, not `new Function`, so the repository's eval ban holds.
function evaluate(expression: string, contexts: Record<string, unknown>): unknown {
  const source = expression.replace(/^\$\{\{\s*/, '').replace(/\s*\}\}$/, '');
  const tokens = source.match(/'[^']*'|\|\||&&|[!=]=|[!()]|[A-Za-z_][\w.-]*/g) ?? [];
  if (tokens.join('') !== source.replace(/\s+/g, '')) throw new Error(`Cannot parse ${source}`);
  let index = 0;
  const peek = (): string | undefined => tokens[index];
  const take = (expected?: string): string => {
    const token = tokens[index++];
    if (token === undefined || (expected !== undefined && token !== expected))
      throw new Error(`Expected ${expected ?? 'a token'} in ${source}`);
    return token;
  };
  const literals: Record<string, unknown> = { true: true, false: false, null: null };
  const lookup = (path: string): unknown => {
    if (path in literals) return literals[path];
    const [root, ...keys] = path.split('.');
    // A context the test did not supply is a test error, not a missing value.
    if (!(root in contexts)) throw new Error(`No context ${root} for ${source}`);
    return keys.reduce<unknown>(
      (value, key) => (value as Record<string, unknown> | undefined)?.[key],
      contexts[root]
    );
  };
  const primary = (): unknown => {
    const token = take();
    if (token === '!') return !primary();
    if (token === '(') {
      const value = or();
      take(')');
      return value;
    }
    if (token.startsWith("'")) return token.slice(1, -1);
    if (peek() === '(') {
      take('(');
      take(')');
      return (lookup(token) as () => unknown)();
    }
    return lookup(token);
  };
  const equality = (): unknown => {
    let value = primary();
    while (peek() === '==' || peek() === '!=') {
      const operator = take();
      const right = primary();
      value = operator === '==' ? value === right : value !== right;
    }
    return value;
  };
  const and = (): unknown => {
    let value = equality();
    while (peek() === '&&') {
      take();
      const right = equality();
      value = value ? right : value;
    }
    return value;
  };
  const or = (): unknown => {
    let value = and();
    while (peek() === '||') {
      take();
      const right = and();
      value = value ? value : right;
    }
    return value;
  };
  const value = or();
  if (index !== tokens.length) throw new Error(`Unexpected ${peek()} in ${source}`);
  return value;
}
const push = { github: { event_name: 'push', ref: 'refs/heads/main' } };
function needs(mode: string, results: Record<string, string> = {}) {
  return {
    mode: { result: results.mode ?? 'success', outputs: { mode } },
    checks: { result: results.checks ?? 'success' },
    build: { result: results.build ?? 'success' },
  };
}
const skipped = { checks: 'skipped', build: 'skipped' };
const stable = (job: string, contexts: Record<string, unknown>, cancelled = false): unknown =>
  evaluate(workflow.jobs[job].if!, { ...push, cancelled: () => cancelled, ...contexts });

test('publication requires every prepublish gate to succeed', () => {
  expect(workflow.jobs.release.needs).toEqual(['mode', 'checks', 'build']);
  expect(stable('release', { needs: needs('publish') })).toBe(true);
  for (const gate of ['checks', 'build'])
    for (const result of ['failure', 'cancelled', 'skipped'])
      expect(stable('release', { needs: needs('publish', { [gate]: result }) })).toBe(false);
  expect(stable('release', { needs: needs('publish') }, true)).toBe(false);
  expect(stable('release', { needs: needs('publish', { mode: 'failure' }) })).toBe(false);
});

test('pending changesets take the PR-update path without the prepublish gates', () => {
  expect(stable('release', { needs: needs('pending', skipped) })).toBe(true);
  for (const job of ['test-plan', 'checks', 'build']) {
    expect(stable(job, { needs: needs('publish') })).toBe(true);
    expect(stable(job, { needs: needs('pending') })).toBe(false);
    expect(workflow.jobs[job]['continue-on-error']).not.toBe(true);
  }
});

test('no job runs outside main, so other branches never publish', () => {
  for (const ref of ['refs/heads/changeset-release/main', 'refs/heads/feature', 'refs/tags/v1']) {
    const github = { event_name: 'workflow_dispatch', ref };
    for (const job of ['test-plan', 'checks', 'build'])
      expect(stable(job, { github, needs: needs('publish') })).toBe(false);
    expect(stable('release', { github, needs: needs('publish') })).toBe(false);
    expect(stable('release', { github, needs: needs('pending', skipped) })).toBe(false);
  }
});

test('the build job builds, writes the notices, and checks the artifacts in that order', () => {
  const steps = workflow.jobs.build.steps;
  const index = (run: string) => steps.findIndex((step) => step.run === run);
  const build = index('bun run build:packages');
  const notices = index('bun run notices:generate');
  const artifacts = index('bun run check:package-artifacts');
  const upload = steps.findIndex((step) => step.name === 'Upload publish artifacts');
  expect(build).toBeGreaterThanOrEqual(0);
  expect(notices).toBeGreaterThan(build);
  expect(artifacts).toBeGreaterThan(notices);
  expect(upload).toBeGreaterThan(artifacts);
  for (const step of steps.slice(build, upload + 1)) {
    expect(step['continue-on-error']).toBeUndefined();
    expect(step.if).toBeUndefined();
  }
});

// npm binds each package's trusted publisher to this file name and to a GitHub-hosted
// runner. A long-lived npm token would bypass it.
test('only the release job publishes, through npm trusted publishing', () => {
  for (const [name, job] of Object.entries(workflow.jobs)) {
    expect(job.permissions?.['id-token']).toBe(name === 'release' ? 'write' : undefined);
    expect(job['runs-on']).toBe('ubuntu-24.04');
  }
  expect(source).not.toMatch(/NPM_TOKEN|NODE_AUTH_TOKEN/);
  const setup = workflow.jobs.release.steps.find((step) =>
    step.uses?.startsWith('actions/setup-node')
  );
  expect(setup?.with?.['node-version']).toBe('24');
  expect(setup?.with?.['registry-url']).toBe('https://registry.npmjs.org');
});

test('the release publishes the dist that the build job checked', () => {
  const steps = workflow.jobs.release.steps;
  const download = steps.findIndex((step) => step.uses?.startsWith('actions/download-artifact'));
  const publish = steps.findIndex((step) => step.id === 'changesets');
  expect(steps[download].with?.name).toBe('publish-dist');
  expect(publish).toBeGreaterThan(download);
  expect(steps[publish].with?.['publish-script']).toBe('bunx changeset publish');
  expect(steps.some((step) => step.run?.includes('build:packages'))).toBe(false);
});
