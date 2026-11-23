// Raster helpers of the fidelity harness, and the line type of `compare.ts`.
//
// Same role as the raster part of pptx-editor/scripts/fidelity/pdf.ts at 1943d2c9 (EigenPal).
// Difference: Poppler's `pdftoppm` draws the reference PDF instead of pdf.js, and every raster is
// supersampled: drawn at twice the compared density and reduced with a 2 x 2 box filter, so
// sub-pixel edges come out alike. The painted page is a screenshot of the canvas at the same
// density (`run.ts`), reduced by the same filter. Test tools only; nothing here is bundled.

import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { PNG } from 'pngjs';

/** Pixels, RGBA, row by row. */
export interface Raster {
  readonly width: number;
  readonly height: number;
  readonly data: Uint8Array;
}

/** One text line of a page, in points from the page's top-left. */
/** One item of a `PdfLine`. */
export interface LineItem {
  readonly at: number;
  readonly length: number;
  readonly xMin: number;
  readonly xMax: number;
}

export interface PdfLine {
  /** The words of the line, joined by one space. */
  readonly text: string;
  readonly xMin: number;
  readonly xMax: number;
  /** The baseline, from the page's top. */
  readonly baseline: number;
  /** The line starts with a one-character item in another font: a list marker. */
  readonly marker?: boolean;
  /** The largest font size of the line's items, in points. */
  readonly size?: number;
  /**
   * The items of the line: `at` counts the characters of the line's text without whitespace
   * before the item, `length` counts the item's own, and `xMin` and `xMax` are its edges.
   */
  readonly items?: readonly LineItem[];
}

/** The density at which the harness compares rasters, in pixels per inch. */
export const COMPARE_DPI = 144;
/** The supersampling factor: each raster is drawn at `COMPARE_DPI * SUPERSAMPLE`. */
export const SUPERSAMPLE = 2;

function run(cmd: string, args: readonly string[]): string {
  const out = spawnSync(cmd, args, { encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });
  if (out.error) throw new Error(`${cmd} did not start: ${out.error.message}`);
  if (out.status !== 0) throw new Error(`${cmd} failed: ${out.stderr.trim()}`);
  return out.stdout;
}

/** The Poppler version, for the known-gap list. */
export function popplerVersion(): string {
  const out = spawnSync('pdftoppm', ['-v'], { encoding: 'utf8' });
  if (out.error) return 'missing';
  const m = /version\s+([\d.]+)/.exec(`${out.stdout}${out.stderr}`);
  return m ? `poppler ${m[1]}` : 'poppler (unknown version)';
}

/** Reduce a raster by `factor` on each side with a box filter. */
export function downsample(src: Raster, factor: number): Raster {
  const width = Math.floor(src.width / factor);
  const height = Math.floor(src.height / factor);
  const data = new Uint8Array(width * height * 4);
  const n = factor * factor;
  for (let y = 0; y < height; y += 1)
    for (let x = 0; x < width; x += 1) {
      const sum = [0, 0, 0, 0];
      for (let dy = 0; dy < factor; dy += 1)
        for (let dx = 0; dx < factor; dx += 1) {
          const i = ((y * factor + dy) * src.width + (x * factor + dx)) * 4;
          for (let c = 0; c < 4; c += 1) sum[c] += src.data[i + c];
        }
      const o = (y * width + x) * 4;
      for (let c = 0; c < 4; c += 1) data[o + c] = Math.round(sum[c] / n);
    }
  return { width, height, data };
}

/** Page 1 of a PDF as a supersampled raster at `COMPARE_DPI`. */
export function rasterPdf(pdfPath: string): Raster {
  const dir = mkdtempSync(join(tmpdir(), 'xlsx-fidelity-'));
  try {
    const prefix = join(dir, 'page');
    run('pdftoppm', [
      '-png',
      '-singlefile',
      '-f',
      '1',
      '-l',
      '1',
      '-r',
      String(COMPARE_DPI * SUPERSAMPLE),
      pdfPath,
      prefix,
    ]);
    const png = PNG.sync.read(readFileSync(`${prefix}.png`));
    return downsample(
      { width: png.width, height: png.height, data: new Uint8Array(png.data) },
      SUPERSAMPLE
    );
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

/**
 * Whitespace runs as one space, trimmed: `normalize` of pptx-editor/scripts/fidelity/pdf.ts at
 * 1943d2c9, without its CJK radical mapping, which no fixture needs yet.
 */
export const normalize = (text: string): string => text.replace(/\s+/g, ' ').trim();
