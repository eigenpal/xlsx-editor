// The harness page: paints the printed page of one fixture with the editor's own layout and
// paint, so the harness measures what the editor paints (design D8). `run.ts` serves it, opens
// it in headless Chromium at a fixed device pixel ratio, takes a screenshot of the page for the
// pixel score, and prints the page to PDF for the line scores.
//
// The page is Letter (612 x 792 pt). Every fixture prints its first worksheet from A1 at 100 %
// inside 0.5 inch margins, with print gridlines and headings off (scripts/fixtures/builder.ts).
// Layout uses the print profile and the number locale of the fixture's reference record. One
// point is 4/3 CSS pixels, so the 816 x 1056 CSS pixel page is the Letter page.

import {
  createHarfBuzzMeasurer,
  initializeHarfBuzz,
  layoutSheet,
  PRINT_PROFILE,
  type FontSource,
} from '../../../packages/core/src/layout/index.ts';
import { readXlsx, workbookIndex } from '../../../packages/core/src/store/index.ts';
import {
  createPictureStore,
  paintSheet,
  registerFontFaces,
} from '../../../packages/core/src/output/index.ts';
import { loadSubstituteFonts } from '../../../packages/core/src/editor/substitute-fonts.ts';
import type { NumberLocale } from '../../../packages/core/src/cascade/index.ts';
import { postScriptName, type PageFace } from '../font-names.ts';
import { PX_PER_PT } from './scale.ts';

/** What one render reports to the runner. */
export interface PageResult {
  readonly ok: boolean;
  readonly problem?: string;
  readonly sheet?: string;
  readonly cells?: number;
  readonly texts?: number;
  readonly drawings?: number;
  /** Layout problems (`SheetLayout.problems`), by code. */
  readonly layoutProblems?: readonly string[];
  /** The fonts that measured and painted text, by origin. */
  readonly fonts?: string;
  /**
   * The faces that the page loaded, with the PostScript name that a PDF writes for each. The
   * runner refuses a printed PDF that draws with another font.
   */
  readonly faces?: readonly PageFace[];
}

/** A face that the runner serves from the reference font folder. */
interface ServedFace {
  readonly family: string;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly url: string;
}

const PAGE_PT = { width: 612, height: 792 };
const MARGIN_PT = 36;

/** The paint of the rendered fixture: for print, and back to the screenshot's paint. */
let printer: { print(): unknown; endPrint(): unknown } | undefined;
/** Whether `harnessPrint` armed the print paint, and how often the print paint ran since. */
let armed = false;
let printPaints = 0;

async function fontSources(mode: string): Promise<FontSource[]> {
  if (mode === 'reference') {
    const list = (await (await fetch('/__harness/fonts.json')).json()) as ServedFace[];
    return Promise.all(
      list.map(async (f) => ({
        family: f.family,
        bold: f.bold,
        italic: f.italic,
        bytes: new Uint8Array(await (await fetch(f.url)).arrayBuffer()),
        origin: 'reference',
      }))
    );
  }
  const { sources } = await loadSubstituteFonts();
  return [...sources];
}

/** Render the fixture whose file name (`fixtureFile`, `scripts/fixtures/catalog.ts`) is `file`. */
async function render(file: string, locale: NumberLocale, fonts: string): Promise<PageResult> {
  await initializeHarfBuzz();
  const sources = await fontSources(fonts);
  if (sources.length === 0) return { ok: false, problem: 'no font loaded' };
  const measurer = createHarfBuzzMeasurer(sources);
  const bytes = new Uint8Array(await (await fetch(`/fixtures/${file}`)).arrayBuffer());
  const read = readXlsx(bytes);
  if (!read.ok) return { ok: false, problem: read.error.reason };
  const index = workbookIndex(read.pkg);
  const sheet = index?.sheets.find((s) => s.kind === 'worksheet' && s.state === 'visible');
  if (!index || !sheet) return { ok: false, problem: 'no visible worksheet' };
  const width = PAGE_PT.width - 2 * MARGIN_PT;
  const height = PAGE_PT.height - 2 * MARGIN_PT;
  const layout = layoutSheet(
    read.pkg,
    sheet.name,
    { width, height, scrollX: 0, scrollY: 0 },
    { measurer, profile: PRINT_PROFILE, locale }
  );
  if (!layout) return { ok: false, problem: 'no layout' };
  const faces = registerFontFaces(document, measurer.faces);
  await faces.ready;
  const pictures = createPictureStore(read.pkg, window);
  const host = document.getElementById('page')!;
  host.style.left = `${MARGIN_PT * PX_PER_PT}px`;
  host.style.top = `${MARGIN_PT * PX_PER_PT}px`;
  const canvas = document.createElement('canvas');
  host.replaceChildren(canvas);
  const paint = (devicePixelRatio = window.devicePixelRatio) =>
    paintSheet(canvas, layout, {
      scale: PX_PER_PT,
      width,
      height,
      gridlines: layout.printGridLines,
      pictures,
      devicePixelRatio,
      // The reference PDFs start the pattern tile grid at the bottom-left corner of the page,
      // the origin of PDF pattern space.
      patternOrigin: { x: -MARGIN_PT, y: PAGE_PT.height - MARGIN_PT },
    });
  let stats = paint();
  for (let tries = 0; stats.pendingPictures > 0 && tries < 50; tries += 1) {
    await new Promise<void>((resolve) => {
      const off = pictures.onReady(() => {
        off();
        resolve();
      });
      setTimeout(resolve, 200);
    });
    stats = paint();
  }
  printer = {
    print: () => paint(1),
    endPrint: () => paint(),
  };
  const origins = [...new Set(measurer.faces.map((f) => f.origin))].sort().join(', ');
  return {
    ok: true,
    sheet: sheet.name,
    cells: stats.cells,
    texts: stats.texts,
    drawings: stats.drawings,
    layoutProblems: [...new Set(layout.problems.map((p) => p.code))].sort(),
    fonts: origins,
    faces: measurer.faces.map((f) => ({
      postScript: postScriptName(f.bytes) ?? '',
    })),
  };
}

// `run.ts` reads the painted text lines from a PDF that Chromium prints of this page. Chromium
// keeps the drawing commands of a canvas only when the canvas is painted during `beforeprint`: a
// canvas painted before the print, even with print media emulated, prints as an image with no text
// (measured). So `harnessPrint` arms the print paint, the `beforeprint` handler paints, and
// `harnessEndPrint` reports how often it painted, so that the runner can refuse a PDF that holds an
// earlier paint. The print paint uses one backing pixel for each CSS pixel: at the device pixel
// ratio of the screenshot (3), the printed canvas scale is rounded, and the glyphs move by up to
// 1.6% of their distance from the canvas origin (measured).
window.addEventListener('beforeprint', () => {
  if (!armed || !printer) return;
  printer.print();
  printPaints += 1;
});

/** Arm the print paint of the rendered fixture. False when no fixture is rendered. */
function harnessPrint(): boolean {
  armed = printer !== undefined;
  printPaints = 0;
  return armed;
}

/** Disarm the print paint, paint for the screen again, and return how often the print paint ran. */
function harnessEndPrint(): number {
  armed = false;
  printer?.endPrint();
  return printPaints;
}

declare global {
  interface Window {
    harnessRender: (file: string, locale: NumberLocale, fonts: string) => Promise<PageResult>;
    harnessPrint: () => boolean;
    harnessEndPrint: () => number;
  }
}

window.harnessRender = render;
window.harnessPrint = harnessPrint;
window.harnessEndPrint = harnessEndPrint;
