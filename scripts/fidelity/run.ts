// `bun run fidelity`: render each fixture with the editor's own layout and paint in headless
// Chromium, take a screenshot of the canvas for the pixel score, print it to PDF for the line
// scores, and compare both with the fixture's reference PDF (design D8).
//
// Same role as pptx-editor/scripts/fidelity/run.ts at 1943d2c9 (EigenPal). Each fixture:
// 1. The harness page (`page/`) lays out and paints the first worksheet with the print profile,
//    the fonts of the run, and the number locale of the fixture's reference record.
// 2. Chromium takes a screenshot of the page at a device pixel ratio of 3 (288 dpi), and Poppler
//    draws the reference PDF at 288 dpi; both reduce to 144 dpi (`poppler.ts`).
// 3. Chromium also prints the page to PDF, between `harnessPrint` and `harnessEndPrint` of the
//    page. The page paints its canvas again while it prints, and the PDF keeps the glyphs that
//    the paint drew as text, at their painted places; the runner refuses a PDF unless that
//    print paint ran once. pdf.js reads the lines of both PDFs, and one `groupLines` groups
//    both (`pdf-lines.ts`), as in the source. The PDF leaves out a glyph that a cell clip hides
//    entirely: with the clip rectangles that both PDFs carry (`pdf-clips.ts`), `compareLines`
//    tells such a glyph from a missing one.
// 4. `compare.ts` scores pixels and lines against `thresholds.ts`; `known-gaps.ts` checks the
//    scores against the known gaps of the font environment and the raster toolchain
//    (`gateRun`). In CI, a run needs the list of exactly its tools.
//
//   bun run fidelity [fixture names...] [--update-gaps [--accept-worse]] [--help]
//
// `--update-gaps` writes the list of the run's tools (`planUpdate` of `update-gaps.ts`), unless
// it is worse and `--accept-worse` is not given. It then runs `bun run fidelity:check-gaps`
// against the merge base and fails when the new list is looser. The command runs through
// `cli.ts`, which loads this module: a module that does not load prints one line and exits 2.
//
// Needs Poppler (`pdftoppm`) and Playwright's Chromium
// (`bunx playwright install chromium`). Reference fonts: `fonts.ts`.
//
// Differences from PPTX Editor:
// - Exit code 2 also marks a run that measured nothing because a tool is missing, the harness
//   server or the browser does not start, the reference fonts cannot be read, or the lock is
//   held or cannot be taken. Such a run gives no score, as a usage error gives none, so CI tells
//   it from a failed gate (exit code 1). As in the source, `main` reads the arguments and the
//   reference fonts, so each of these errors prints one line. An error that stops the harness
//   after it starts prints one line and exits with 1, as in the source (`harnessStopped`).
// - A run clears only the entries of the report folder that it writes (`clearsEntry`), and only
//   after the server and the browser start. The source clears the whole folder of its set on a
//   run of every deck.
// - `XLSX_FIDELITY_REPORT_DIR` names another report folder, and `XLSX_FIDELITY_MANIFEST` another
//   reference manifest. Tests and probes set them; the user docs do not name them. The helper
//   tools read the same report folder (`reportDir`), and a relative value resolves against the
//   repository root. A run in a shared folder keeps every entry that it does not write. The run
//   prints the report path relative to the current folder, as the source does, or absolute when
//   the folder is outside it.
// - `main` reads the reference manifest before it takes the lock. A manifest that cannot be read
//   prints one line and exits with code 2, and the report of the earlier run stays.
// - After `--update-gaps` writes a list, the run starts the merge-base check
//   (`bun run fidelity:check-gaps --require-thresholds`) and exits by its result (`updateExit`).
//   The source writes the list and starts no check.
//
// As in the source, a run that stops after it cleared the report folder, such as one that a
// signal stops, loses the report of the earlier run.

import { spawnSync } from 'node:child_process';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import type { IncomingMessage, ServerResponse } from 'node:http';
import { join } from 'node:path';
import { chromium, type Browser, type Page } from '@playwright/test';
import { PNG } from 'pngjs';
import { createServer, type Plugin } from 'vite';
import { FIXTURES } from '../fixtures/catalog.ts';
import {
  checkPoppler,
  chromiumPid,
  eachFixture,
  firstLine,
  harnessStopped,
  TOOL_EXIT,
  withHarness,
} from './chromium.ts';
import { checkValues, compareLines, comparePixels, failureKinds } from './compare.ts';
import { FOREIGN_REFERENCE_REMEDY, familyId, foreignReferenceFonts } from './font-names.ts';
import { printPainted } from './print.ts';
import {
  facesIdentity,
  FONTS_ENV,
  referenceFaces,
  postScriptFamilies,
  referenceFontDirs,
  workbookFamilies,
  type LocalFace,
} from './fonts.ts';
import { fixCauses, gateRun, writeGaps, type FixtureResult, type GapTools } from './known-gaps.ts';
import { groupLines, pdfItems, PDFJS_VERSION, printedPageProblem } from './pdf-lines.ts';
import { downsample, popplerVersion, rasterPdf, SUPERSAMPLE, type Raster } from './poppler.ts';
import {
  FIXTURE_DIR,
  MANIFEST,
  readManifest,
  referenceProblems,
  type Manifest,
} from './references.ts';
import {
  clipLogLines,
  COVERAGE_GAPS,
  coverageGapLine,
  oneCharacterGap,
  oneCharacterLines,
  reportDir,
  reportPathText,
  rotatedItems,
  rotatedTextGap,
  writeReport,
} from './report.ts';
import { FIDELITY_USAGE, fidelityArgs } from './run-args.ts';
import { clearReport, clearsEntry, releaseOnSignals, takeRunLock } from './run-lock.ts';
import { newListRefusal, planUpdate, updateExit, type CheckRun } from './update-gaps.ts';
import { KNOWN_GAPS, type FontEnvironment, type KnownGapList } from './thresholds.ts';
import { DEVICE_SCALE } from './page/scale.ts';

const REPO = join(import.meta.dir, '..', '..');
const REPORT_DIR = reportDir();
const PAGE_DIR = join(import.meta.dir, 'page');

/** What `main` reads before the harness starts: the arguments and the reference fonts. */
interface RunSetup {
  readonly names: readonly string[];
  readonly updateGaps: boolean;
  /** `--accept-worse`: `--update-gaps` also writes a worse list (`planUpdate`). */
  readonly acceptWorse: boolean;
  readonly selected: typeof FIXTURES;
  /** The run measures every fixture. */
  readonly complete: boolean;
  readonly fontDirs: readonly string[];
  readonly env: FontEnvironment;
  /** The families of the reference font files by PostScript name (`foreignReferenceFonts`). */
  readonly knownFamilies: ReadonlyMap<string, string>;
  /** The reference manifest (`fixtures/references.json`). */
  readonly manifest: Manifest;
}

/** The environment variable that names another reference manifest, for tests only. */
export const MANIFEST_ENV = 'XLSX_FIDELITY_MANIFEST';

/**
 * Read the arguments, the reference manifest, and the reference fonts, or give the exit code: 0
 * after `--help`, and `TOOL_EXIT` after a usage error, a manifest that cannot be read, or a
 * reference font folder that cannot be read. Each error prints one line. Nothing is cleared yet.
 */
function readSetup(): RunSetup | number {
  const parsed = fidelityArgs(
    process.argv.slice(2),
    FIXTURES.map((f) => f.name)
  );
  if ('help' in parsed) {
    console.log(FIDELITY_USAGE);
    return 0;
  }
  if ('error' in parsed) {
    console.error(`fidelity: ${parsed.error}`);
    return TOOL_EXIT;
  }
  const { names, updateGaps, acceptWorse } = parsed;
  const selected = names.length ? FIXTURES.filter((f) => names.includes(f.name)) : FIXTURES;
  const manifestPath = process.env[MANIFEST_ENV] || MANIFEST;
  let manifest: Manifest;
  try {
    manifest = readManifest(manifestPath);
  } catch (error) {
    console.error(`fidelity: the reference manifest cannot be read: ${firstLine(error)}`);
    return TOOL_EXIT;
  }
  try {
    const fontDirs = referenceFontDirs();
    return {
      names,
      updateGaps,
      acceptWorse,
      selected,
      complete: selected.length === FIXTURES.length,
      fontDirs,
      env: fontDirs.length ? 'reference' : 'substitute',
      knownFamilies: postScriptFamilies(fontDirs),
      manifest,
    };
  } catch (error) {
    console.error(
      `fidelity: the reference fonts in ${FONTS_ENV} cannot be read: ${firstLine(error)}`
    );
    return TOOL_EXIT;
  }
}

/** The faces that the page loads for the fixture that renders now. */
let served: LocalFace[] = [];

/** Serves the fixtures and the reference faces to the harness page. */
const harnessFiles: Plugin = {
  name: 'xlsx-fidelity-files',
  configureServer(server) {
    server.middlewares.use((req: IncomingMessage, res: ServerResponse, next: () => void) => {
      const url = req.url ?? '';
      const fixture = /^\/fixtures\/([a-z0-9-]+)\.xlsx$/.exec(url);
      if (fixture && FIXTURES.some((f) => f.name === fixture[1])) {
        res.setHeader('content-type', 'application/octet-stream');
        res.end(readFileSync(join(FIXTURE_DIR, `${fixture[1]}.xlsx`)));
        return;
      }
      if (url === '/__harness/fonts.json') {
        res.setHeader('content-type', 'application/json');
        res.end(
          JSON.stringify(
            served.map((f, i) => ({
              family: f.family,
              bold: f.bold,
              italic: f.italic,
              url: `/__harness/font/${i}`,
            }))
          )
        );
        return;
      }
      const font = /^\/__harness\/font\/(\d+)$/.exec(url);
      if (font && Number(font[1]) < served.length) {
        res.setHeader('content-type', 'font/ttf');
        res.end(served[Number(font[1])].bytes);
        return;
      }
      next();
    });
  },
};

const toPng = (r: Raster): Buffer =>
  PNG.sync.write(
    Object.assign(new PNG({ width: r.width, height: r.height }), { data: Buffer.from(r.data) })
  );

async function main(): Promise<number> {
  const setup = readSetup();
  if (typeof setup === 'number') return setup;
  // The lock comes before the tools, so that a held lock gives its own line. The run clears the
  // report folder only after the server and the browser start (`measureWith`): a missing tool
  // keeps the report of the earlier run.
  let lock: ReturnType<typeof takeRunLock>;
  try {
    lock = takeRunLock(REPORT_DIR);
  } catch (error) {
    // A report folder that cannot be written: the run measures nothing.
    console.error(`fidelity: the run lock cannot be taken: ${firstLine(error)}`);
    return TOOL_EXIT;
  }
  if ('held' in lock) {
    console.error(`fidelity: ${lock.held}`);
    return TOOL_EXIT;
  }
  // A run that a signal stops releases the lock; a run that is killed leaves a stale lock, which
  // the next run takes over (`takeRunLock`).
  const unregister = releaseOnSignals(lock);
  try {
    const tools = checkPoppler(popplerVersion);
    if ('error' in tools) {
      console.error(`fidelity: ${tools.error}`);
      return tools.code;
    }
    return await measure(setup, tools.poppler);
  } finally {
    unregister();
    lock.release();
  }
}

/**
 * `--update-gaps`: plan the list of the run (`planUpdate`), print its changes, and write it
 * unless it is worse without `--accept-worse`. Gives the written list, or why it wrote nothing.
 */
function recordGaps(
  lists: readonly KnownGapList[],
  results: readonly FixtureResult[],
  tools: GapTools,
  setup: RunSetup
): { readonly written?: KnownGapList; readonly refused?: string; readonly failed?: string } {
  const plan = planUpdate(lists, results, tools, setup.complete, setup.acceptWorse);
  if ('refused' in plan) return { refused: plan.refused };
  for (const e of plan.changes.looser) console.log(`LOOSER: ${e}`);
  for (const e of plan.changes.added) console.log(`NEW: ${e}`);
  for (const e of plan.changes.removed) console.log(`REMOVED: ${e}`);
  if (!plan.write) return {};
  try {
    writeGaps(join(import.meta.dir, 'thresholds.ts'), setup.env, plan.next);
  } catch (error) {
    // The run still writes its report; `updateExit` gives the line and the exit code.
    return { failed: firstLine(error) };
  }
  console.log(
    `Recorded the known gaps of the ${setup.env} font environment for ${tools.poppler} and ${tools.chromium}.`
  );
  return { written: plan.next };
}

/** The server of the harness page: its address, for example `http://127.0.0.1:53712/`. */
interface HarnessServer {
  readonly url: string;
  close(): Promise<void>;
}

/**
 * Start the server of the harness page. A server with no local address does not start, as in
 * the source. Here `withHarness` then prints one line and exits with code 2.
 */
async function startServer(): Promise<HarnessServer> {
  const server = await createServer({
    root: PAGE_DIR,
    configFile: false,
    logLevel: 'error',
    plugins: [harnessFiles],
    optimizeDeps: { exclude: ['harfbuzzjs', '@docx-editor.dev/fonts'] },
    // As in the source: the page opens the exact address of this server, and edits during a run
    // do not reload the page.
    server: {
      host: '127.0.0.1',
      port: 0,
      strictPort: true,
      hmr: false,
      watch: null,
      fs: { allow: [REPO] },
    },
  });
  try {
    await server.listen();
  } catch (error) {
    await server.close();
    throw error;
  }
  const url = server.resolvedUrls?.local[0];
  if (!url) {
    await server.close();
    throw new Error('the harness server has no local address');
  }
  return { url, close: () => server.close() };
}

async function measure(setup: RunSetup, poppler: string): Promise<number> {
  // `withHarness` closes the browser and the server, also when the run throws.
  const done = await withHarness(
    startServer,
    () => chromium.launch(),
    chromiumPid,
    (server, browser) => measureWith(setup, server, browser, poppler)
  );
  if ('error' in done) {
    console.error(`fidelity: ${done.error}`);
    return done.code;
  }
  return done.result;
}

async function measureWith(
  setup: RunSetup,
  server: HarnessServer,
  browser: Browser,
  poppler: string
): Promise<number> {
  const { selected, complete, updateGaps, fontDirs, env, knownFamilies } = setup;
  const tools = {
    poppler: `${poppler}, pdf.js ${PDFJS_VERSION}`,
    chromium: `chromium ${browser.version()}`,
  };
  const lists = KNOWN_GAPS[env] ?? [];
  // A partial update for tools that have no list is refused before the run measures anything.
  const refusal = updateGaps ? newListRefusal(lists, tools, complete) : undefined;
  if (refusal) {
    console.error(`fidelity: ${refusal}`);
    return TOOL_EXIT;
  }
  // The run writes only the folders of its fixtures and the report files (`clearsEntry`): the
  // folders of other fixtures stay from an earlier run, and so does any other entry. The source
  // clears its whole report folder on a run of every deck.
  const names = selected.map((f) => f.name);
  try {
    clearReport(REPORT_DIR, (entry) => clearsEntry(names, entry));
  } catch (error) {
    console.error(`fidelity: the report folder cannot be cleared: ${firstLine(error)}`);
    return TOOL_EXIT;
  }
  const { manifest } = setup;
  const allFaces: LocalFace[] = [];
  // An error of one fixture, from its reference to its page, fails that fixture only
  // (`eachFixture`), and the next fixture still runs.
  const results = await eachFixture<FixtureResult>(
    names,
    async (name) => {
      let page: Page | undefined;
      try {
        // The gate names the fixture (`checkGate`), so its problems do not repeat the name.
        const problems = referenceProblems(name, manifest).map((p) =>
          p.startsWith(`${name}: `) ? p.slice(name.length + 2) : p
        );
        if (problems.length) {
          return { name, status: 'error', error: problems.join('; ') };
        }
        const xlsx = new Uint8Array(readFileSync(join(FIXTURE_DIR, `${name}.xlsx`)));
        served = env === 'reference' ? referenceFaces(fontDirs, workbookFamilies(xlsx)) : [];
        allFaces.push(...served);
        const dir = join(REPORT_DIR, name);
        mkdirSync(dir, { recursive: true });
        page = await browser.newPage({
          viewport: { width: 816, height: 1056 },
          deviceScaleFactor: DEVICE_SCALE,
        });
        await page.goto(new URL('index.html', server.url).href);
        await page.waitForFunction(() => typeof window.harnessRender === 'function');
        const locale = manifest.references[name]!.locale;
        const report = await page.evaluate(([n, l, f]) => window.harnessRender(n, l, f), [
          name,
          locale,
          env,
        ] as const);
        if (!report.ok) {
          return {
            name,
            status: 'error',
            error: `the page did not render: ${report.problem}`,
          };
        }
        const shot = PNG.sync.read(
          await page.screenshot({ type: 'png', clip: { x: 0, y: 0, width: 816, height: 1056 } })
        );
        const a = downsample(
          { width: shot.width, height: shot.height, data: new Uint8Array(shot.data) },
          SUPERSAMPLE
        );
        // The painted lines come from what the page painted: Chromium prints the page, the page
        // paints its canvas again while it prints (`print.ts`), and the PDF keeps the drawn glyphs
        // as text. pdf.js reads it as it reads the reference PDF.
        const printed = await printPainted(page);
        writeFileSync(join(dir, 'ours.pdf'), printed);
        const ours = await pdfItems(printed);
        const reference = join(FIXTURE_DIR, `${name}.pdf`);
        const ref = await pdfItems(reference);
        const named = new Set([...workbookFamilies(xlsx)].map(familyId));
        const foreignRef = foreignReferenceFonts(ref.fonts, named, knownFamilies);
        if (foreignRef.length) {
          return {
            name,
            status: 'error',
            error: `the reference PDF draws with fonts that it does not embed or that the workbook does not name: ${foreignRef.join(', ')}; ${FOREIGN_REFERENCE_REMEDY}`,
          };
        }
        const { faces, ...pageReport } = report;
        const loaded = new Map((faces ?? []).map((f) => [f.postScript, f]));
        const printProblem = printedPageProblem(
          ours,
          ref.size,
          report.texts ?? 0,
          new Set(loaded.keys())
        );
        if (printProblem) {
          return { name, status: 'error', error: printProblem };
        }
        const b = rasterPdf(reference);
        const pixels = comparePixels(a, b, true);
        const lines = compareLines(
          groupLines(ours.items, ref.size[0], ref.size[1]),
          groupLines(ref.items, ref.size[0], ref.size[1]),
          { ours: ours.runs, reference: ref.runs }
        );
        writeFileSync(join(dir, 'ours.png'), toPng(a));
        writeFileSync(join(dir, 'reference.png'), toPng(b));
        if (pixels.diff) writeFileSync(join(dir, 'diff.png'), pixels.diff);
        return {
          name,
          status: 'compared',
          failures: failureKinds(pixels, lines, name),
          values: checkValues(pixels, lines),
          linesTotal: lines.total,
          linesMatched: lines.matched,
          clippedLines: lines.clipped,
          ...(lines.uncut.length ? { uncutLines: lines.uncut } : {}),
          ...(lines.lostUncut.length ? { lostUncutLines: lines.lostUncut } : {}),
          ...(lines.extraCut.length ? { extraCutLines: lines.extraCut } : {}),
          ...(lines.unknownReference.length
            ? { unknownReferenceEnds: lines.unknownReference }
            : {}),
          ...(lines.bothCut.length ? { bothCutLines: lines.bothCut } : {}),
          missingLines: lines.missing,
          extraLines: lines.extra,
          ...(ref.rotated ? { rotatedItems: ref.rotated } : {}),
          ...(lines.oneCharacter ? { oneCharacterLines: lines.oneCharacter } : {}),
          ...(pixels.diffBox ? { diffBox: pixels.diffBox } : {}),
          page: pageReport,
        };
      } finally {
        await page?.close();
      }
    },
    (name, error) => ({ name, status: 'error', error })
  );
  const fonts =
    env === 'reference' ? facesIdentity(allFaces) : '@docx-editor.dev/fonts substitutes';
  const update = updateGaps ? recordGaps(lists, results, tools, setup) : undefined;
  const { list, problems } = gateRun({
    results,
    env,
    tools,
    lists,
    complete,
    ci: Boolean(process.env.CI),
    ...(update?.written ? { recorded: update.written } : {}),
  });
  writeReport(REPORT_DIR, results, {
    env,
    fonts,
    tools,
    ...(list ? { gaps: list } : {}),
    problems,
  });
  for (const r of results) {
    const v = r.values;
    const text = v
      ? `pixels ${(v.pixels * 100).toFixed(3)}%, lines ${r.linesMatched}/${r.linesTotal} (clipped ${r.clippedLines?.length ?? 0}), baseline ${v.baseline.toFixed(2)} pt, position ${v['line-position'].toFixed(2)} pt, extra ${v['extra-text']}`
      : r.error;
    // The status follows the gate: a `fix` problem fails the fixture also when its checks pass.
    const causes = fixCauses(r.name, problems);
    const status = [
      ...(r.failures?.length
        ? [
            `fails ${r.failures.join(', ')}${list && Object.hasOwn(list.fixtures, r.name) && !causes.includes('gap') ? ' (known gap)' : ''}`,
          ]
        : []),
      ...(causes.length ? [`fails the gate (${causes.join(', ')})`] : []),
    ];
    console.log(
      `${r.name}: ${status.length ? status.join('; ') : r.status === 'error' ? 'error' : 'pass'} (${text})`
    );
    for (const line of clipLogLines(r)) console.log(line);
  }
  console.log(
    `Font environment: ${env} (${fonts}). Report: ${reportPathText(process.cwd(), REPORT_DIR)}`
  );
  for (const gap of COVERAGE_GAPS) console.log(coverageGapLine(gap));
  const single = oneCharacterLines(results);
  if (single) console.log(oneCharacterGap(single));
  const rotated = rotatedItems(results);
  if (rotated) console.log(rotatedTextGap(rotated));
  for (const p of problems) console.error(`${p.kind}: ${p.text}`);
  const status = problems.some((p) => p.kind !== 'note') ? 1 : 0;
  if (!update) return status;
  let check: CheckRun | undefined;
  if (update.written) {
    // A new record must not loosen the gate: run the merge-base ratchet in a new process, so that
    // it reads the rewritten thresholds.ts.
    console.log('Known gaps against the merge base (bun run fidelity:check-gaps):');
    check = spawnSync(
      process.execPath,
      [join(import.meta.dir, 'cli.ts'), 'check-gaps', '--require-thresholds'],
      {
        cwd: REPO,
        stdio: 'inherit',
      }
    );
  }
  const exit = updateExit(status, update, check);
  if (exit.line) console.error(exit.line);
  return exit.code;
}

/**
 * `bun run fidelity` (`cli.ts run`). Any other error stops the harness: one line, never a stack
 * trace, as in the source.
 */
export function command(): Promise<number> {
  return main().catch((error: unknown) => {
    console.error(harnessStopped(error));
    return 1;
  });
}

if (import.meta.main) process.exit(await command());
