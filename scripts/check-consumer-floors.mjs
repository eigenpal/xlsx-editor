#!/usr/bin/env node
// Holds the lowest versions that the package READMEs state. The statements must equal
// `scripts/lib/consumer-floors.mjs`, and the packed packages must work at each floor.
//
// With no flag: the README sample and every published entry compile on the TypeScript floor with
// `bundler` and `node16`, the Vite floor builds an app that shows a worksheet, and the Vite floor's
// dev server shows one with the `optimizeDeps` setting of `packages/core/README.md`. The webpack
// floor emits the text shaper, the fonts, and the stylesheet. React 18.0.0 with `@types/react`
// 18.2.6 compiles and shows a worksheet. Vite 7 does not run on the Node.js floors, so the Node.js
// floors have their own mode.
// With `--node`: this Node.js must be a floor. Every published entry imports, a workbook reads,
// and the text shaper loads, which is how a server uses core. Then the server samples of
// `packages/core/README.md` that `check:consumer-install` runs (`core-server-save` and
// `core-server-measure`) run as written, with only their types removed, because the Node.js
// floors cannot run TypeScript.
//
// Needs built packages (`bun run build:packages`), the npm registry, and, with no flag,
// Playwright's Chromium. CI runs each floor in the `floors` job.
import {
  cpSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { spawn } from 'node:child_process';
import ts from 'typescript';
import { FLOORS, floorProblems } from './lib/consumer-floors.mjs';
import { packPublishedPackages, run } from './lib/consumer-packs.mjs';
import { extractSample } from './lib/readme-samples.mjs';
import { PUBLISHED_PACKAGES as PACKAGES } from './lib-packages.mjs';

const ROOT = path.resolve(import.meta.dirname, '..');
const read = (file) => readFileSync(path.join(ROOT, file), 'utf8');
const coreReadme = read('packages/core/README.md');
const reactReadme = read('packages/react/README.md');
const vueReadme = read('packages/vue/README.md');
const rootManifest = JSON.parse(read('package.json'));
const nodeMode = process.argv.includes('--node');

const problems = floorProblems({
  coreReadme,
  reactReadme,
  vueReadme,
  engines: JSON.parse(read('packages/core/package.json')).engines?.node,
  reactPeer: JSON.parse(read('packages/react/package.json')).peerDependencies.react,
  ci: read('.github/workflows/ci.yml'),
});
if (nodeMode && !FLOORS.node.includes(process.versions.node))
  problems.push(
    `Node.js ${process.versions.node} is not a floor: ${FLOORS.node.join(', ')}. ` +
      'Run --node with one of these versions, or read the result of the CI floors job.'
  );
if (problems.length > 0) {
  console.error(problems.join('\n'));
  process.exit(1);
}

const tempRoot = mkdtempSync(path.join(tmpdir(), 'xlsx-editor-floors-'));
const appDir = path.join(tempRoot, 'app');
const write = (file, text) => writeFileSync(path.join(appDir, file), text);
const failures = [];

function installApp(devDependencies) {
  const tarballs = packPublishedPackages(ROOT, tempRoot);
  mkdirSync(path.join(appDir, 'src'), { recursive: true });
  mkdirSync(path.join(appDir, 'public'), { recursive: true });
  cpSync(
    path.join(ROOT, 'examples/vite/public/sample.xlsx'),
    path.join(appDir, 'public/workbook.xlsx')
  );
  write(
    'package.json',
    JSON.stringify({
      name: 'xlsx-editor-floors',
      private: true,
      type: 'module',
      dependencies: {
        ...Object.fromEntries([...tarballs].map(([name, file]) => [name, `file:${file}`])),
        react: rootManifest.devDependencies.react,
        'react-dom': rootManifest.devDependencies['react-dom'],
      },
      devDependencies,
    })
  );
  run('npm', ['install', '--no-audit', '--no-fund', '--loglevel=error'], appDir);
}

function checkNodeFloor() {
  installApp({});
  const entries = PACKAGES.flatMap(({ name, entries: names }) =>
    names.map((entry) => (entry === 'index' ? name : `${name}/${entry}`))
  );
  run(
    process.execPath,
    [
      '--input-type=module',
      '--eval',
      `import { readFileSync } from 'node:fs';
      for (const entry of ${JSON.stringify(entries)}) await import(entry);
      const { readXlsx } = await import('@xlsx-editor.dev/core/store');
      const read = readXlsx(new Uint8Array(readFileSync('public/workbook.xlsx')));
      if (!read.ok) throw new Error('readXlsx refused the sample: ' + JSON.stringify(read.error));
      const { initializeHarfBuzz } = await import('@xlsx-editor.dev/core/layout');
      await initializeHarfBuzz();`,
    ],
    appDir
  );
  // The server samples of the core README, as `check:consumer-install` runs them: a read and a
  // save, then the measurer and the grid metrics.
  cpSync(path.join(appDir, 'public/workbook.xlsx'), path.join(appDir, 'in.xlsx'));
  for (const name of ['core-server-save', 'core-server-measure']) {
    const { outputText } = ts.transpileModule(extractSample(coreReadme, name), {
      compilerOptions: { module: ts.ModuleKind.ESNext, target: ts.ScriptTarget.ES2022 },
    });
    write(`${name}.mjs`, outputText);
    run(process.execPath, [`${name}.mjs`], appDir);
  }
  if (!existsSync(path.join(appDir, 'out.xlsx')))
    failures.push('core README server sample: no out.xlsx');
}

// Starts a Vite command, and ends its process group after `visit`.
async function withServer(args, port, visit) {
  const vite = path.join(appDir, 'node_modules/vite/bin/vite.js');
  const server = spawn(process.execPath, [vite, ...args, '--port', port, '--strictPort'], {
    cwd: appDir,
    stdio: 'pipe',
    detached: true,
  });
  try {
    await new Promise((resolve, reject) => {
      const timer = setTimeout(() => reject(new Error(`vite ${args[0]} did not start`)), 30_000);
      server.stdout.on('data', (chunk) => {
        if (String(chunk).includes(port)) {
          clearTimeout(timer);
          resolve();
        }
      });
    });
    await visit(`http://localhost:${port}/`);
  } finally {
    try {
      process.kill(-server.pid, 'SIGTERM');
    } catch {
      // The group has already exited.
    }
  }
}

// The webpack floor emits the text shaper, the fonts, and the stylesheet with no setup.
function checkWebpackFloor() {
  write(
    'src/webpack-entry.js',
    `import { XlsxEditor } from '@xlsx-editor.dev/react';
import '@xlsx-editor.dev/react/styles.css';

console.log(XlsxEditor);
`
  );
  write(
    'webpack-build.mjs',
    `import path from 'node:path';
import MiniCssExtractPlugin from 'mini-css-extract-plugin';
import webpack from 'webpack';

const config = {
  mode: 'production',
  context: import.meta.dirname,
  entry: './src/webpack-entry.js',
  output: { path: path.join(import.meta.dirname, 'dist-webpack') },
  module: { rules: [{ test: /\\.css$/, use: [MiniCssExtractPlugin.loader, 'css-loader'] }] },
  plugins: [new MiniCssExtractPlugin()],
  performance: { hints: false },
};
webpack(config, (error, stats) => {
  if (error || stats.hasErrors()) {
    console.error(error ?? stats.toString({ all: false, errors: true }));
    process.exit(1);
  }
});
`
  );
  run(process.execPath, ['webpack-build.mjs'], appDir);
  const emitted = readdirSync(path.join(appDir, 'dist-webpack'));
  for (const [label, pattern] of [
    ['text shaper', /\.wasm$/],
    ['font', /\.(?:ttf|otf)$/],
    ['stylesheet', /\.css$/],
  ]) {
    if (!emitted.some((name) => pattern.test(name)))
      failures.push(`webpack ${FLOORS.webpack}: the build emitted no ${label}`);
  }
}

async function checkToolFloors() {
  installApp({
    typescript: FLOORS.typescript,
    vite: FLOORS.vite,
    webpack: FLOORS.webpack,
    'css-loader': '7.1.2',
    'mini-css-extract-plugin': '2.9.4',
    '@types/react': rootManifest.devDependencies['@types/react'],
    '@types/react-dom': rootManifest.devDependencies['@types/react-dom'],
  });
  write('src/sample.tsx', extractSample(reactReadme, 'react-editor'));
  // Every published entry, so that the TypeScript floor reads each declaration file.
  write(
    'src/entries.ts',
    PACKAGES.flatMap(({ name, entries }) =>
      entries.map((entry) => (entry === 'index' ? name : `${name}/${entry}`))
    )
      .map((entry, index) => `import * as entry${index} from '${entry}';\n`)
      .join('')
  );
  write(
    'src/main.tsx',
    `import { createRoot } from 'react-dom/client';
import { Editor } from './sample.tsx';

createRoot(document.getElementById('app')!).render(<Editor />);
`
  );
  write(
    'index.html',
    '<!doctype html><html><head><meta charset="UTF-8" /></head><body><div id="app"></div><script type="module" src="/src/main.tsx"></script></body></html>\n'
  );
  const compilerOptions = {
    strict: true,
    noEmit: true,
    target: 'ES2022',
    allowImportingTsExtensions: true,
    jsx: 'react-jsx',
    lib: ['ES2022', 'DOM', 'DOM.Iterable'],
    types: [],
    // The published declarations are what the TypeScript floor holds.
    skipLibCheck: false,
  };
  const projects = ['tsconfig.json', 'tsconfig.node16.json'];
  for (const [file, resolution] of [
    [projects[0], { module: 'ESNext', moduleResolution: 'bundler' }],
    [projects[1], { module: 'Node16', moduleResolution: 'node16' }],
  ]) {
    write(
      file,
      JSON.stringify({ compilerOptions: { ...compilerOptions, ...resolution }, include: ['src'] })
    );
    run(process.execPath, ['node_modules/typescript/bin/tsc', '--project', file], appDir);
  }

  run(process.execPath, ['node_modules/vite/bin/vite.js', 'build', '--logLevel', 'warn'], appDir);
  const { chromium } = await import('@playwright/test');
  const browser = await chromium.launch();
  const showsWorksheet = (label) => async (url) => {
    const page = await browser.newPage();
    try {
      await page.goto(url);
      await page.waitForSelector('.xlsx-editor [role=grid][aria-label^="Worksheet "]', {
        timeout: 60_000,
      });
    } catch {
      failures.push(`${label}: no worksheet showed within 60 s`);
    } finally {
      await page.close();
    }
  };
  try {
    await withServer(['preview'], '4331', showsWorksheet(`vite ${FLOORS.vite} build`));
    write('vite.config.ts', extractSample(coreReadme, 'vite-7-dev'));
    await withServer(['--force'], '4332', showsWorksheet(`vite ${FLOORS.vite} dev server`));
    checkWebpackFloor();
    // The React 18 row replaces React and its types in the same app. `@types/react` 18.2.6 reads
    // `scheduler/tracing`, which later releases of `@types/scheduler` do not have.
    run(
      'npm',
      [
        'install',
        '--no-audit',
        '--no-fund',
        '--loglevel=error',
        '--save-exact',
        `react@${FLOORS.react}`,
        `react-dom@${FLOORS.react}`,
        `@types/react@${FLOORS.typesReact}`,
        '@types/react-dom@18.2.4',
        '@types/scheduler@0.16.8',
      ],
      appDir
    );
    for (const file of projects)
      run(process.execPath, ['node_modules/typescript/bin/tsc', '--project', file], appDir);
    run(process.execPath, ['node_modules/vite/bin/vite.js', 'build', '--logLevel', 'warn'], appDir);
    await withServer(['preview'], '4333', showsWorksheet(`React ${FLOORS.react} build`));
  } finally {
    await browser.close();
  }
}

try {
  if (nodeMode) checkNodeFloor();
  else await checkToolFloors();
} catch (error) {
  failures.push(error.message);
} finally {
  rmSync(tempRoot, { recursive: true, force: true });
}

if (failures.length > 0) {
  console.error(failures.join('\n'));
  process.exit(1);
}
console.log(
  nodeMode
    ? `Consumer floors: the packages run on Node.js ${process.versions.node}.`
    : `Consumer floors: TypeScript ${FLOORS.typescript}, Vite ${FLOORS.vite}, webpack ${FLOORS.webpack}, and React ${FLOORS.react} work with the packed packages.`
);
