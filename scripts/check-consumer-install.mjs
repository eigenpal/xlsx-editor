#!/usr/bin/env node
// Copied from pptx-editor/scripts/check-consumer-install.mjs at e22f3f3b (EigenPal). Differences:
// the packages are `@xlsx-editor.dev/*`; the app shows `examples/vite/public/sample.xlsx` and
// waits for the worksheet grid; and the Node PDF export is left out, because XLSX Editor has no
// PDF package yet (AGENTS.md, "Packages"). In its place, Node.js imports every published entry,
// reads the sample workbook, loads the text shaper, and runs the two server samples of the core
// README and the sample of the i18n README, as DOCX Editor's check imports its packages under
// Node.js. CI runs it on one Node.js version only. The PDF step also checked the substitute
// fonts; here the Vite build must emit the font assets, and the browser's font report
// must show that every packaged face loaded. The preview port is 4329, the editor takes a
// fixed height of 600 instead of a `100vh` container, its `onReady` hands the editor to the
// font report check, and the README comment names `packages/react/README.md`. XLSX Editor also
// builds the README samples that a `<!-- sample: name -->` comment marks, as written, type-checks
// them, and checks each grid in Chromium, so a README sample that does not run fails here. It
// type-checks the marked React guide samples too, each fragment after the declarations that its
// prelude in `scripts/lib/readme-samples.mjs` holds. It
// resolves each documented asset subpath through the installed `exports` maps, where PPTX Editor
// checks only that the WASM file exists, and it packs each manifest with its `devDependencies`.
// A webpack 5 production build must keep both published stylesheets, because webpack honors the
// `sideEffects` field where the Vite build does not show a wrong value. The unsaved-edits sample
// must report no unsaved edits after each save. Each link of the installed READMEs and guides
// must reach a file and a heading of the installed packages. The samples also compile with no
// bundler types on TypeScript 6.0.3 and 7.0.2, the first React sample compiles with
// `exactOptionalPropertyTypes`, the Vite dev server shows a worksheet with every packaged face, and
// webpack 5 must emit the text shaper and the fonts.
//
// Installs the packages the way a consumer does: from tarballs, into an empty project, with npm
// resolving what it finds inside them. Then it builds a React app with Vite and with Bun's
// bundler, and opens the Vite build in Chromium to check that a worksheet shows
// (docx-editor/scripts/check-consumer-install.mjs, sized for XLSX Editor). DOCX Editor commit
// adf0e178.
//
// This is the only check that reads a published manifest instead of a workspace one, so it
// is the only place where a `workspace:` range, a missing `exports` subpath, a `files` list
// that drops a needed file, or a WASM asset that a bundler cannot find (for example,
// when `harfbuzzjs` stays external) fails before a user sees it.
//
// Needs built packages (`bun run build:packages`), Playwright's Chromium, and network access to
// the npm registry.
// Usage: node scripts/check-consumer-install.mjs [--keep]
import {
  cpSync,
  existsSync,
  mkdtempSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { spawn, spawnSync } from 'node:child_process';
import { packPublishedPackages, run } from './lib/consumer-packs.mjs';
import { installedDocProblems } from './lib/package-doc-links.mjs';
import { extractSample, GUIDE_SAMPLES } from './lib/readme-samples.mjs';
import { PUBLISHED_PACKAGES as PACKAGES } from './lib-packages.mjs';

const ROOT = path.resolve(import.meta.dirname, '..');
const keep = process.argv.includes('--keep');
const tempRoot = mkdtempSync(path.join(tmpdir(), 'xlsx-editor-consumer-'));
const appDir = path.join(tempRoot, 'app');

// The non-code subpaths that the package READMEs tell a consumer to import or serve.
const DOCUMENTED_ASSETS = [
  '@xlsx-editor.dev/core/harfbuzz.wasm',
  '@xlsx-editor.dev/core/styles/editor.css',
  '@xlsx-editor.dev/react/styles.css',
];

// The README samples that the app builds as written. Each page shows one sample.
const README_PAGES = [
  { label: 'core README sample', page: 'core-sample.html' },
  { label: 'React README sample', page: 'react-sample.html' },
  { label: 'React README layout sample', page: 'react-layout-sample.html' },
  { label: 'React guide unsaved-edits sample', page: 'react-unsaved-sample.html' },
];

// Exact versions, so that the check does not change when a new release of a tool appears.
const WEBPACK_TOOLS = {
  webpack: '5.112.2',
  'css-loader': '7.1.2',
  'mini-css-extract-plugin': '2.9.4',
};

// Exact versions: each compiles the samples with no bundler types.
const SAMPLE_COMPILERS = {
  typescript6: 'npm:typescript@6.0.3',
  typescript7: 'npm:typescript@7.0.2',
};

const rootManifest = JSON.parse(readFileSync(path.join(ROOT, 'package.json'), 'utf8'));
let tarballs;
try {
  tarballs = packPublishedPackages(ROOT, tempRoot);
} catch (error) {
  console.error(error.message);
  rmSync(tempRoot, { recursive: true, force: true });
  process.exit(1);
}

// A minimal consumer app, written the way packages/react/README.md tells users to write it.
const dep = (name) => rootManifest.devDependencies[name];
mkdirSync(path.join(appDir, 'src'), { recursive: true });
mkdirSync(path.join(appDir, 'public'), { recursive: true });
writeFileSync(
  path.join(appDir, 'package.json'),
  JSON.stringify(
    {
      name: 'xlsx-editor-consumer',
      private: true,
      type: 'module',
      dependencies: {
        ...Object.fromEntries([...tarballs].map(([name, file]) => [name, `file:${file}`])),
        react: dep('react'),
        'react-dom': dep('react-dom'),
      },
      devDependencies: {
        vite: dep('vite'),
        '@vitejs/plugin-react': dep('@vitejs/plugin-react'),
        typescript: dep('typescript'),
        '@types/react': dep('@types/react'),
        '@types/react-dom': dep('@types/react-dom'),
        ...WEBPACK_TOOLS,
        ...SAMPLE_COMPILERS,
      },
    },
    null,
    2
  )
);
writeFileSync(
  path.join(appDir, 'vite.config.js'),
  `import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  build: {
    rollupOptions: {
      input: ${JSON.stringify(['index.html', ...README_PAGES.map(({ page }) => page)])},
    },
  },
});
`
);
writeFileSync(
  path.join(appDir, 'index.html'),
  '<!doctype html><html><head><meta charset="UTF-8" /></head><body><div id="app"></div><script type="module" src="/src/main.jsx"></script></body></html>\n'
);
writeFileSync(
  path.join(appDir, 'src', 'main.jsx'),
  `import { createRoot } from 'react-dom/client';
import { XlsxEditor } from '@xlsx-editor.dev/react';
import '@xlsx-editor.dev/react/styles.css';

const bytes = new Uint8Array(await (await fetch('/sample.xlsx')).arrayBuffer());
// The check reads the font report through the editor that onReady hands out.
createRoot(document.getElementById('app')).render(
  <XlsxEditor
    document={bytes}
    style={{ height: 600 }}
    onReady={(editor) => {
      window.__xlsxEditor = editor;
    }}
  />
);
`
);
cpSync(
  path.join(ROOT, 'examples/vite/public/sample.xlsx'),
  path.join(appDir, 'public/sample.xlsx')
);

writeReadmeSamples();

const failures = [];
try {
  run('npm', ['install', '--no-audit', '--no-fund', '--loglevel=error'], appDir);

  // The links resolve in what npm installed, so a guide that `files` leaves out fails here. A
  // link to a docs site page resolves in this checkout.
  failures.push(
    ...installedDocProblems(
      new Map(
        PACKAGES.map((pkg) => [
          path.basename(pkg.root),
          path.join(appDir, 'node_modules', pkg.name),
        ])
      ),
      ROOT
    )
  );

  // The samples compile as written, against the packed declarations.
  for (const compiler of Object.keys(SAMPLE_COMPILERS))
    run('node', [`node_modules/${compiler}/bin/tsc`, '--project', 'tsconfig.samples.json'], appDir);
  run('node', ['node_modules/typescript6/bin/tsc', '--project', 'tsconfig.exact.json'], appDir);

  // Vite: the build must emit HarfBuzz's WASM, and the app must show a worksheet.
  run('npx', ['vite', 'build', '--logLevel', 'warn'], appDir);
  const assets = readdirSync(path.join(appDir, 'dist', 'assets'));
  if (!assets.some((name) => name.endsWith('.wasm')))
    failures.push('vite: the build emitted no .wasm asset (harfbuzzjs)');
  // The default substitute faces of @docx-editor.dev/fonts load through the same
  // `new URL(..., import.meta.url)` pattern, so the build must emit them too.
  if (!assets.some((name) => /\.(?:ttf|otf)$/.test(name)))
    failures.push('vite: the build emitted no font asset (@docx-editor.dev/fonts)');
  await checkRender(failures);

  // Node.js: every published entry imports, a workbook reads, and HarfBuzz loads from the
  // installed WASM, which is how a server uses core.
  const entries = PACKAGES.flatMap(({ name, entries: names }) =>
    names.map((entry) => (entry === 'index' ? name : `${name}/${entry}`))
  );
  run(
    'node',
    [
      '--input-type=module',
      '--eval',
      `import { readFileSync } from 'node:fs';
      for (const entry of ${JSON.stringify(entries)}) await import(entry);
      const { readXlsx } = await import('@xlsx-editor.dev/core/store');
      const read = readXlsx(new Uint8Array(readFileSync('public/sample.xlsx')));
      if (!read.ok) throw new Error('readXlsx refused the sample: ' + JSON.stringify(read.error));
      const { initializeHarfBuzz } = await import('@xlsx-editor.dev/core/layout');
      await initializeHarfBuzz();`,
    ],
    appDir
  );
  // The server samples of the core README: a read and a save, then the measurer.
  cpSync(path.join(appDir, 'public/sample.xlsx'), path.join(appDir, 'in.xlsx'));
  run('node', ['server-save.ts'], appDir);
  if (!existsSync(path.join(appDir, 'out.xlsx')))
    failures.push('core README server sample: no out.xlsx');
  run('node', ['server-measure.ts'], appDir);
  run('node', ['i18n-sample.ts'], appDir);

  // Bun: the browser bundle must build. No check opens it in a browser, so no README states
  // that a Bun app works.
  const bun = spawnSync(
    'bun',
    ['build', 'src/main.jsx', '--outdir', 'dist-bun', '--target', 'browser', '--splitting'],
    { cwd: appDir, encoding: 'utf8' }
  );
  if (bun.status !== 0) failures.push(`bun build failed: ${bun.stderr.trim().split('\n')[0]}`);

  checkWebpackStylesheets(failures);

  // The asset subpaths that the READMEs document resolve through the installed `exports` maps.
  const resolveFromApp = createRequire(path.join(appDir, 'package.json')).resolve;
  for (const subpath of DOCUMENTED_ASSETS) {
    try {
      resolveFromApp(subpath);
    } catch (error) {
      failures.push(`${subpath} does not resolve from the installed package: ${error.code}`);
    }
  }
} finally {
  if (keep) console.log(`Kept ${tempRoot}`);
  else rmSync(tempRoot, { recursive: true, force: true });
}

function writeReadmeSamples() {
  const readme = (name) => readFileSync(path.join(ROOT, 'packages', name, 'README.md'), 'utf8');
  const page = (body, script) =>
    `<!doctype html><html><head><meta charset="UTF-8" /></head><body>${body}<script type="module" src="/src/${script}"></script></body></html>\n`;
  const write = (file, text) => writeFileSync(path.join(appDir, file), text);
  const mount = (name, source) =>
    `import { createRoot } from 'react-dom/client';
import { ${name} } from './${source}';

const response = await fetch('/workbook.xlsx');
const bytes = new Uint8Array(await response.arrayBuffer());
createRoot(document.getElementById('app')!).render(<${name} bytes={bytes} />);
`;
  const core = readme('core');
  const react = readme('react');
  write('core-sample.html', page(extractSample(core, 'core-host'), 'core-sample.ts'));
  // The server samples run under Node.js as written, outside `src/` and the browser build.
  write('server-save.ts', extractSample(core, 'core-server-save'));
  write('server-measure.ts', extractSample(core, 'core-server-measure'));
  // The i18n README sample needs no DOM, so it type-checks with the others and runs under Node.js.
  write('i18n-sample.ts', extractSample(readme('i18n'), 'i18n-use'));
  write('src/core-sample.ts', extractSample(core, 'core-open'));
  write('react-sample.html', page('<div id="app"></div>', 'react-sample-main.tsx'));
  write('src/react-sample.tsx', extractSample(react, 'react-editor'));
  // The first sample fetches its own bytes, so the page mounts it with no props.
  write(
    'src/react-sample-main.tsx',
    `import { createRoot } from 'react-dom/client';
import { Editor } from './react-sample.tsx';

createRoot(document.getElementById('app')!).render(<Editor />);
`
  );
  write('react-layout-sample.html', page('<div id="app"></div>', 'react-layout-main.tsx'));
  // The layout sample is one JSX expression, so it needs a component around it.
  write(
    'src/react-layout-sample.tsx',
    `import { XlsxEditor } from '@xlsx-editor.dev/react';
import '@xlsx-editor.dev/react/styles.css';

export function Layout({ bytes }: { bytes: Uint8Array }) {
  return (
${extractSample(react, 'react-layout')}  );
}
`
  );
  write('src/react-layout-main.tsx', mount('Layout', 'react-layout-sample.tsx'));
  write('react-unsaved-sample.html', page('<div id="app"></div>', 'react-unsaved-main.tsx'));
  const propsGuide = readFileSync(
    path.join(ROOT, 'packages/react/docs/layout-and-props.md'),
    'utf8'
  );
  write('src/react-unsaved-sample.tsx', extractSample(propsGuide, 'react-unsaved'));
  write('src/react-unsaved-main.tsx', mount('TrackedEditor', 'react-unsaved-sample.tsx'));
  // The other guide samples only type-check. A fragment gets the declarations of its prelude.
  for (const { guide, name, prelude } of GUIDE_SAMPLES) {
    const text = readFileSync(path.join(ROOT, 'packages/react/docs', guide), 'utf8');
    write(`src/react-${name}.tsx`, `${prelude}${prelude ? '\n' : ''}${extractSample(text, name)}`);
  }
  write(
    'src/webpack-react.js',
    `import { XlsxEditor } from '@xlsx-editor.dev/react';
import '@xlsx-editor.dev/react/styles.css';

console.log(XlsxEditor);
`
  );
  write(
    'src/webpack-core.js',
    `import { createXlsxEditor } from '@xlsx-editor.dev/core/editor';
import '@xlsx-editor.dev/core/styles/editor.css';

console.log(createXlsxEditor);
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
  entry: { react: './src/webpack-react.js', core: './src/webpack-core.js' },
  output: { path: path.join(import.meta.dirname, 'dist-webpack') },
  module: { rules: [{ test: /\\.css$/, use: [MiniCssExtractPlugin.loader, 'css-loader'] }] },
  plugins: [new MiniCssExtractPlugin({ filename: '[name].css' })],
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
  const compilerOptions = {
    strict: true,
    noEmit: true,
    target: 'ES2022',
    module: 'ESNext',
    moduleResolution: 'bundler',
    allowImportingTsExtensions: true,
    jsx: 'react-jsx',
    lib: ['ES2022', 'DOM', 'DOM.Iterable'],
    types: [],
    skipLibCheck: true,
  };
  write(
    'tsconfig.samples.json',
    JSON.stringify({
      compilerOptions,
      include: ['src/core-sample.ts', 'src/react-*.tsx', 'i18n-sample.ts'],
    })
  );
  write(
    'tsconfig.exact.json',
    JSON.stringify({
      compilerOptions: { ...compilerOptions, exactOptionalPropertyTypes: true },
      include: ['src/react-sample.tsx'],
    })
  );
  cpSync(
    path.join(ROOT, 'examples/vite/public/sample.xlsx'),
    path.join(appDir, 'public/workbook.xlsx')
  );
}

// webpack removes an import that the `sideEffects` field marks as free of side effects, with no
// error, so only the emitted file shows that a stylesheet survived.
function checkWebpackStylesheets(out) {
  run('node', ['webpack-build.mjs'], appDir);
  for (const name of ['react', 'core']) {
    const file = path.join(appDir, 'dist-webpack', `${name}.css`);
    if (!existsSync(file) || !readFileSync(file, 'utf8').includes('.xlsx-active-cell'))
      out.push(`webpack: the ${name} entry emitted no editor stylesheet`);
  }
  const emitted = readdirSync(path.join(appDir, 'dist-webpack'));
  if (!emitted.some((name) => name.endsWith('.wasm')))
    out.push('webpack: the build emitted no .wasm asset (harfbuzzjs)');
  if (!emitted.some((name) => /\.(?:ttf|otf)$/.test(name)))
    out.push('webpack: the build emitted no font asset (@docx-editor.dev/fonts)');
}

// A save with nothing to commit does not call `onChange`, so the sample must clear its flag itself,
// after a committed edit and after an edit that is still open.
async function checkUnsavedSample(browser, out) {
  const label = 'React guide unsaved-edits sample';
  const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
  const status = page.locator('#app > span');
  const expectStatus = async (text, step) => {
    try {
      await page.waitForFunction(
        (wanted) => document.querySelector('#app > span')?.textContent === wanted,
        text,
        { timeout: 5_000 }
      );
    } catch {
      out.push(`${label}: ${step}: the status is "${await status.textContent()}"`);
    }
  };
  try {
    await page.goto('http://localhost:4329/react-unsaved-sample.html');
    await page.waitForSelector('.xlsx-scroller .xlsx-active-cell', {
      state: 'attached',
      timeout: 30_000,
    });
    const save = page.locator('#app > button');
    await expectStatus('No unsaved edits', 'after the open');
    await page.locator('.xlsx-scroller').focus();
    await page.keyboard.type('9');
    await page.keyboard.press('Enter');
    await expectStatus('Unsaved edits', 'after a committed edit');
    await save.click();
    await expectStatus('No unsaved edits', 'after the save of a committed edit');
    await page.locator('.xlsx-scroller').focus();
    await page.keyboard.type('7');
    await expectStatus('Unsaved edits', 'during an open cell edit');
    await save.click();
    await expectStatus('No unsaved edits', 'after the save of an open cell edit');
    await page.locator('.xlsx-scroller').focus();
    await page.keyboard.press('ArrowDown');
    await expectStatus('No unsaved edits', 'after a selection change');
  } catch (error) {
    out.push(`${label}: ${error.message.split('\n')[0]}`);
  } finally {
    await page.close();
  }
}

// A sample passes when its grid has a height, scrolls under the wheel, and shows the active cell
// border. Without the stylesheet scope or a height, the workbook still opens, so a worksheet alone
// proves nothing here.
async function checkReadmeSample(browser, { label, page: file }, out) {
  const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
  const errors = [];
  page.on('pageerror', (error) => errors.push(error.message));
  page.on('console', (message) => {
    if (message.type() === 'error') errors.push(message.text());
  });
  try {
    await page.goto(`http://localhost:4329/${file}`);
    try {
      await page.waitForSelector('.xlsx-scroller .xlsx-active-cell', {
        state: 'attached',
        timeout: 30_000,
      });
    } catch {
      out.push(`${label}: no active cell showed within 30 s`);
      return;
    }
    const scroller = page.locator('.xlsx-scroller');
    const box = await scroller.boundingBox();
    if (!box || box.height < 100 || box.height > 600) {
      out.push(`${label}: the grid height is ${box?.height ?? 'unknown'} px, expected 100 to 600`);
      return;
    }
    // Before the scroll: the active cell leaves the page when it scrolls out of view.
    const border = await page.locator('.xlsx-active-cell').evaluate((element) => {
      const style = getComputedStyle(element);
      return {
        width: Number.parseFloat(style.borderTopWidth),
        style: style.borderTopStyle,
        token: style.getPropertyValue('--doc-primary').trim(),
      };
    });
    if (border.width <= 0 || border.style === 'none' || border.token === '')
      out.push(`${label}: the active cell has no border (${JSON.stringify(border)})`);
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.mouse.wheel(0, 400);
    try {
      await page.waitForFunction(
        () => (document.querySelector('.xlsx-scroller')?.scrollTop ?? 0) > 0,
        null,
        { timeout: 5_000 }
      );
    } catch {
      out.push(`${label}: the grid did not scroll under the mouse wheel`);
    }
    for (const error of errors) out.push(`${label}: console error: ${error}`);
  } finally {
    await page.close();
  }
}

// The dev server of the Vite version that the app installs serves the text shaper and the fonts
// with no setup, as the READMEs state.
async function checkDevServer(browser, out) {
  const vite = path.join(appDir, 'node_modules/vite/bin/vite.js');
  const server = spawn(process.execPath, [vite, '--port', '4330', '--strictPort'], {
    cwd: appDir,
    stdio: 'pipe',
    detached: true,
  });
  try {
    await new Promise((resolve, reject) => {
      const timer = setTimeout(
        () => reject(new Error('the vite dev server did not start')),
        30_000
      );
      server.stdout.on('data', (chunk) => {
        if (String(chunk).includes('4330')) {
          clearTimeout(timer);
          resolve();
        }
      });
    });
    const page = await browser.newPage();
    await page.goto('http://localhost:4330/');
    try {
      await page.waitForSelector('.xlsx-editor [role=grid][aria-label^="Worksheet "]', {
        timeout: 60_000,
      });
    } catch {
      out.push('vite dev server: no worksheet showed within 60 s');
    }
    const fonts = await page.evaluate(() => {
      const report = window.__xlsxEditor?.fontReport();
      return report
        ? { faces: report.faces.length, failures: report.substituteFailures.length }
        : null;
    });
    if (!fonts || fonts.faces === 0 || fonts.failures > 0)
      out.push(`vite dev server: the packaged faces did not load (${JSON.stringify(fonts)})`);
    await page.close();
  } finally {
    try {
      process.kill(-server.pid, 'SIGTERM');
    } catch {
      // The group has already exited.
    }
  }
}

async function checkRender(out) {
  // Vite's own bin in its own process group. Through `npx`, killing the child on Linux left
  // the Vite server running, and its open stdout pipe kept this script alive until CI's timeout.
  const vite = path.join(appDir, 'node_modules/vite/bin/vite.js');
  const preview = spawn(process.execPath, [vite, 'preview', '--port', '4329', '--strictPort'], {
    cwd: appDir,
    stdio: 'pipe',
    detached: true,
  });
  try {
    await new Promise((resolve, reject) => {
      const timer = setTimeout(() => reject(new Error('vite preview did not start')), 30_000);
      preview.stdout.on('data', (chunk) => {
        if (String(chunk).includes('4329')) {
          clearTimeout(timer);
          resolve();
        }
      });
    });
    const { chromium } = await import('@playwright/test');
    const browser = await chromium.launch();
    try {
      const page = await browser.newPage();
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      page.on('console', (message) => {
        if (message.type() === 'error') errors.push(message.text());
      });
      await page.goto('http://localhost:4329/');
      try {
        // The grid names its worksheet once the workbook is open and its first sheet shows.
        await page.waitForSelector('.xlsx-editor [role=grid][aria-label^="Worksheet "]', {
          timeout: 30_000,
        });
      } catch {
        out.push('vite: no worksheet showed within 30 s');
      }
      // A workbook opens after its fonts settle, and a face that did not load only warns. So
      // read the report: every packaged substitute face must have loaded.
      const fonts = await page.evaluate(() => {
        const report = window.__xlsxEditor?.fontReport();
        return report
          ? { faces: report.faces.length, failures: report.substituteFailures.length }
          : null;
      });
      if (!fonts) out.push('vite: onReady did not hand out the editor');
      else if (fonts.faces === 0 || fonts.failures > 0)
        out.push(
          `vite: ${fonts.failures} substitute faces did not load (${fonts.faces} faces in force)`
        );
      for (const error of errors) out.push(`vite: console error: ${error}`);
      await page.close();
      for (const sample of README_PAGES) await checkReadmeSample(browser, sample, out);
      await checkUnsavedSample(browser, out);
      await checkDevServer(browser, out);
    } finally {
      await browser.close();
    }
  } finally {
    // The negative pid signals the whole group: the server and anything it started.
    try {
      process.kill(-preview.pid, 'SIGTERM');
    } catch {
      // The group has already exited.
    }
  }
}

if (failures.length > 0) {
  console.error(failures.join('\n'));
  process.exit(1);
}
console.log(
  'Consumer install: every entry imports under Node.js, the Vite app builds and shows a worksheet with every packaged face, the links of the installed docs resolve, the marked samples compile and show a working grid, the unsaved-edits sample clears after a save, the Vite dev server shows a worksheet, webpack keeps both stylesheets and emits the shaper and fonts, and the Bun bundle builds.'
);
