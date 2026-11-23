// The lowest versions that the package READMEs state, and the versions that
// `scripts/check-consumer-floors.mjs` runs. A floor changes here and in the READMEs together.
export const FLOORS = {
  node: ['20.16.0', '22.3.0'],
  typescript: '5.7.2',
  vite: '7.0.0',
  webpack: '5.83.0',
  react: '18.0.0',
  typesReact: '18.2.6',
};

const minorOf = (version) => version.split('.').slice(0, 2).join('.');

/**
 * @param {{ coreReadme: string, reactReadme: string, vueReadme: string,
 *   engines: string | undefined, reactPeer: string | undefined, ci: string }} stated The texts
 *   that state a floor.
 * @param {typeof FLOORS} floors The tested floors.
 * @returns {string[]} Each statement that differs from the tested floors.
 */
export function floorProblems(
  { coreReadme, reactReadme, vueReadme, engines, reactPeer, ci },
  floors = FLOORS
) {
  const problems = [];
  const expect = (where, actual, wanted) => {
    if (actual !== wanted)
      problems.push(
        `${where} states ${actual ?? 'no floor'}, and the tested floor in ` +
          `scripts/lib/consumer-floors.mjs is ${wanted}`
      );
  };
  const [first, second] = floors.node;
  const node = coreReadme.match(
    /Node\.js (\d+\.\d+\.\d+) or later in the \d+\.x release line, or Node\.js (\d+\.\d+\.\d+) or later/
  );
  expect(
    'packages/core/README.md (Node.js)',
    node ? `${node[1]} or ${node[2]}` : undefined,
    `${first} or ${second}`
  );
  expect('packages/core/package.json (engines)', engines, `^${first} || >=${second}`);
  for (const [name, text] of [
    ['packages/core/README.md', coreReadme],
    ['packages/react/README.md', reactReadme],
    ['packages/vue/README.md', vueReadme],
  ]) {
    expect(
      `${name} (TypeScript)`,
      text.match(/TypeScript (\d+\.\d+) or later/)?.[1],
      minorOf(floors.typescript)
    );
    expect(
      `${name} (Vite)`,
      text.match(/With Vite, use version (\d+\.\d+\.\d+) or later/)?.[1],
      floors.vite
    );
  }
  expect(
    'packages/core/README.md (webpack)',
    coreReadme.match(/webpack (\d+\.\d+\.\d+) or later/)?.[1],
    floors.webpack
  );
  expect(
    'packages/react/README.md (@types/react)',
    reactReadme.match(/`@types\/react` (\d+\.\d+\.\d+) or later/)?.[1],
    floors.typesReact
  );
  expect('packages/react/package.json (react peer)', reactPeer?.split(' ')[0], `^${floors.react}`);
  expect(
    '.github/workflows/ci.yml (floors job)',
    ci.match(/^\s*floor: \[(.+)\]$/m)?.[1],
    ['tools', ...floors.node].join(', ')
  );
  return problems;
}
