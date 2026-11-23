// A control for the glyph pixels of one fixture: `bun scripts/fidelity/rasterizer-control.ts
// <fixture>`, after `bun run fidelity`. It draws the reference PDF of the fixture with pdf.js in
// the harness Chromium, as canvas text at 288 dpi, and scores three pairs with the harness rules
// (the 2× box downsample and `comparePixels`):
//
// - the editor page against the pdf.js raster: both are canvas text in the same Chromium, so this pair
//   shows what paint and layout add;
// - the pdf.js raster against the Poppler raster of the same PDF: this pair shows what the two
//   rasterizers add to the same page;
// - the editor page against the Poppler raster: the score of `bun run fidelity`.
//
// It reads `<fixture>/ours.png` and `reference.png` in the report folder of the run (`reportDir`:
// `fidelity-report/`, or `XLSX_FIDELITY_REPORT_DIR`), and writes nothing. It prints
// the Chromium, pdf.js, and Poppler versions, and warns when `ours.png` is older than the PDF. A
// wrong fixture name or a missing input prints its line and exits with code 2, as the run does
// (`helperInputs`).

import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { chromium } from '@playwright/test';
import { PNG } from 'pngjs';
import { comparePixels } from './compare.ts';
import { COMPARE_DPI, downsample, SUPERSAMPLE, type Raster } from './poppler.ts';
import { FIXTURES } from '../fixtures/catalog.ts';
import { helperInputs } from './helper-inputs.ts';
import { FIXTURE_DIR } from './references.ts';

const inputs = helperInputs(
  'rasterizer-control.ts',
  process.argv[2],
  ['ours.png', 'reference.png'],
  FIXTURES.map((f) => f.name),
  FIXTURE_DIR
);
if (inputs.errors.length) {
  for (const line of inputs.errors) console.error(line);
  process.exit(2);
}
if (inputs.warning) console.warn(inputs.warning);
const name = inputs.name;
const REPORT = inputs.report;
const PDFJS_PACKAGE = createRequire(import.meta.url).resolve('pdfjs-dist/package.json');
const PDFJS = join(dirname(PDFJS_PACKAGE), 'build');
const PDFJS_VERSION = (JSON.parse(readFileSync(PDFJS_PACKAGE, 'utf8')) as { version: string })
  .version;

const load = (file: string): Raster => {
  const png = PNG.sync.read(readFileSync(file));
  return { width: png.width, height: png.height, data: new Uint8Array(png.data) };
};

const html = (body: string) => new Response(body, { headers: { 'content-type': 'text/html' } });
const script = (file: string) =>
  new Response(readFileSync(join(PDFJS, file)), { headers: { 'content-type': 'text/javascript' } });

/** The reference PDF drawn by pdf.js as canvas text in Chromium, at the harness resolution. */
async function pdfjsRaster(): Promise<Raster> {
  const server = Bun.serve({
    hostname: '127.0.0.1',
    port: 0,
    fetch(req) {
      const path = new URL(req.url).pathname;
      if (path === '/') return html('<!doctype html><body></body>');
      if (path === '/doc.pdf') return new Response(readFileSync(join(FIXTURE_DIR, `${name}.pdf`)));
      if (path === '/pdf.mjs' || path === '/pdf.worker.mjs') return script(path.slice(1));
      return new Response('not found', { status: 404 });
    },
  });
  let browser: Awaited<ReturnType<typeof chromium.launch>> | undefined;
  try {
    browser = await chromium.launch();
    console.log(`${name}: chromium ${browser.version()}, pdf.js ${PDFJS_VERSION}`);
    const page = await browser.newPage({ viewport: { width: 816, height: 1056 } });
    await page.goto(`http://127.0.0.1:${server.port}/`);
    const url = await page.evaluate(
      async (scale) => {
        const src = '/pdf.mjs';
        const lib = await import(src);
        lib.GlobalWorkerOptions.workerSrc = '/pdf.worker.mjs';
        const doc = await lib.getDocument({ url: '/doc.pdf' }).promise;
        const first = await doc.getPage(1);
        const viewport = first.getViewport({ scale });
        const canvas = document.createElement('canvas');
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        const g = canvas.getContext('2d')!;
        g.fillStyle = '#fff';
        g.fillRect(0, 0, canvas.width, canvas.height);
        await first.render({ canvasContext: g, canvas, viewport }).promise;
        return canvas.toDataURL('image/png');
      },
      (COMPARE_DPI * SUPERSAMPLE) / 72
    );
    const png = PNG.sync.read(Buffer.from(url.slice(url.indexOf(',') + 1), 'base64'));
    return downsample(
      { width: png.width, height: png.height, data: new Uint8Array(png.data) },
      SUPERSAMPLE
    );
  } finally {
    await browser?.close();
    await server.stop();
  }
}

const popplerVersion = Bun.spawnSync(['pdftoppm', '-v']).stderr.toString().split('\n')[0];
console.log(`${name}: ${popplerVersion}`);
const ours = load(join(REPORT, 'ours.png'));
const poppler = load(join(REPORT, 'reference.png'));
const pdfjs = await pdfjsRaster();
const pairs: [string, Raster, Raster][] = [
  ['editor page vs pdf.js canvas text', ours, pdfjs],
  ['pdf.js canvas text vs Poppler', pdfjs, poppler],
  ['editor page vs Poppler', ours, poppler],
];
for (const [label, a, b] of pairs) {
  const share = comparePixels(a, b).mismatch;
  const px = Math.round(share * a.width * a.height);
  console.log(`${name}: ${label}: ${px} px (${(share * 100).toFixed(3)}%)`);
}
