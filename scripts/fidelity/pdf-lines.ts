// The text lines of both sides of a comparison. Copied from
// pptx-editor/scripts/fidelity/pdf.ts at 1943d2c9 (EigenPal): `TextItem`, `groupLines`, and the
// item reading of `openPdf().lines`. As in the source, one pdf.js reading takes the items of
// both PDFs: the reference PDF, and the PDF that Chromium prints of the painted page
// (`page/main.ts`, `run.ts`). Both go through the one `groupLines`. Differences: a line also
// records its font size and the edges of its items, for the join of `compareLines`,
// `pdfItems` counts the rotated items that it leaves out, and `groupLines` starts a new line
// where the clip of the next item differs (`TextItem.clip`).
//
// Compared with 1943d2c9; later commits 67ae7309, 06534363, and b887a55d checked: 67ae7309 and
// 06534363 do not change this file, and b887a55d is ported: pdf.js 6 closes a document through
// its loading task (`loadingTask.destroy()`), and has no `isEvalSupported` option, because it
// has no code path that compiles PDF functions with `new Function`. It already sets
// `fontExtraProperties`.

import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { getDocument, OPS, version } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { baseFontName } from './font-names.ts';
import {
  clipOfItem,
  otherClip,
  splitItemByRuns,
  textRuns,
  type TextClipBox,
  type TextRun,
} from './pdf-clips.ts';
import { THRESHOLDS } from './thresholds.ts';
import type { LineItem, PdfLine } from './poppler.ts';

/** The pdf.js version that reads the reference PDFs. */
export const PDFJS_VERSION: string = version;

const PDFJS_DIR = dirname(Bun.resolveSync('pdfjs-dist/package.json', import.meta.dir));

/** One text item of a page: its text, origin, advance width, and font size, in points. */
export interface TextItem {
  readonly text: string;
  readonly x: number;
  /** The baseline, from the page's top. */
  readonly y: number;
  readonly width: number;
  readonly size: number;
  /** The PDF font of the item (pdf.js `fontName`). */
  readonly font?: string;
  /** The PostScript name of that font, without a subset prefix (`pdfItems`). */
  readonly postScript?: string;
  /** The clip of the text run that draws the item (`clipOfItem`), when one draws it. */
  readonly clip?: TextClipBox;
}

/** Items whose baselines are this close (points) share a row (`THRESHOLDS.sameRowPt`). */
const SAME_ROW_PT = THRESHOLDS.sameRowPt;
/** A gap wider than this share of the font size starts a new line on the same row. */
export const LINE_GAP_EM = 1;
/**
 * A gap wider than this share of the font size is a space between two items. The gate does not
 * guard this constant, nor `LINE_GAP_EM`.
 */
const SPACE_GAP_EM = 0.2;

/**
 * Group text items into lines. Items with an origin outside the page are left out. Items whose
 * baselines are within `SAME_ROW_PT` form a row; a row splits where the gap between two items is
 * wider than `LINE_GAP_EM` of the font size (columns, tabs, two shapes side by side), and where
 * the clip of the next item differs from the clip of the item before it: the text of a cell
 * that runs past its clip and the text of the next cell stay two lines, also where they
 * overlap. Both PDFs of a comparison go through this one function, so the same layout gives the
 * same lines.
 */
export function groupLines(
  items: readonly TextItem[],
  widthPt: number,
  heightPt: number
): PdfLine[] {
  const inside = items
    .filter((i) => i.x >= 0 && i.x <= widthPt && i.y >= 0 && i.y <= heightPt)
    .sort((p, q) => p.y - q.y || p.x - q.x);
  const rows: TextItem[][] = [];
  for (const item of inside) {
    const row = rows.at(-1);
    if (row && Math.abs(item.y - row[0].y) <= SAME_ROW_PT) row.push(item);
    else rows.push([item]);
  }
  const lines: PdfLine[] = [];
  for (const row of rows) {
    row.sort((p, q) => p.x - q.x);
    let run: TextItem[] = [];
    const flush = (): void => {
      if (run.length) lines.push(lineOf(run));
      run = [];
    };
    for (const item of row) {
      const last = run.at(-1);
      if (
        last &&
        (item.x - (last.x + last.width) > LINE_GAP_EM * Math.max(item.size, last.size) ||
          otherClip(last.clip, item.clip))
      )
        flush();
      run.push(item);
    }
    flush();
  }
  return lines.filter((l) => l.text.trim() !== '');
}

function lineOf(run: readonly TextItem[]): PdfLine {
  let text = '';
  let main = run[0];
  let size = 0;
  const items: LineItem[] = [];
  run.forEach((item, i) => {
    items.push({
      at: text.replace(/\s+/g, '').length,
      length: item.text.replace(/\s+/g, '').length,
      xMin: item.x,
      xMax: item.x + item.width,
    });
    size = Math.max(size, item.size);
    const last = run[i - 1];
    if (
      last &&
      item.x - (last.x + last.width) > SPACE_GAP_EM * Math.max(item.size, last.size) &&
      !/\s$/.test(text) &&
      !/^\s/.test(item.text)
    )
      text += ' ';
    text += item.text;
    if (item.text.trim().length > main.text.trim().length) main = item;
  });
  const end = run.at(-1)!;
  const marker = run.length > 1 && run[0].text.trim().length === 1 && run[0].font !== run[1].font;
  return {
    text: text.replace(/\s+/g, ' ').trim(),
    xMin: run[0].x,
    xMax: end.x + end.width,
    baseline: main.y,
    ...(marker ? { marker } : {}),
    size,
    items,
  };
}

/** The text of page 1 of a PDF (`pdfItems`). */
export interface PdfItems {
  /** The unrotated text items, in points from the page's top-left. */
  readonly items: TextItem[];
  /** How many text items with glyphs were left out as rotated, mirrored, or vertical. */
  readonly rotated: number;
  /** The page size of page 1 in points. */
  readonly size: [number, number];
  /** The number of pages of the PDF. */
  readonly pages: number;
  /** The fonts that page 1 draws with, by PostScript name, and whether the PDF embeds each. */
  readonly fonts: readonly { readonly name: string; readonly embedded: boolean }[];
  /** The text runs of page 1 with their clips (`textRuns`). */
  readonly runs: readonly TextRun[];
}

/** Page sizes within this distance (points) are the same size. */
const SAME_SIZE_PT = 0.5;

/**
 * Why the PDF that Chromium printed of the painted page cannot give the painted lines, or
 * undefined. The page must be one page of the size of the reference page. Every font it draws with
 * must be an embedded face that the page loaded (`loaded`, by PostScript name), as PPTX Editor's
 * `foreignOurFonts` requires: a system font means that the paint asked for a face that it does not
 * have. It must hold text when the paint drew text: a canvas that Chromium prints as an image holds
 * none, and every line would then count as missing for a reason that is not the render. This
 * catches only the loss of all text; the line scores catch the loss of some lines.
 */
export function printedPageProblem(
  ours: PdfItems,
  referenceSize: readonly [number, number],
  paintedTexts: number,
  loaded: ReadonlySet<string> = new Set()
): string | undefined {
  if (ours.pages !== 1)
    return `the printed PDF has ${ours.pages} pages, not 1; check the \`@page\` size and the canvas size of the harness page (\`page/index.html\`, \`page/main.ts\`)`;
  const foreign = ours.fonts
    .filter((f) => !f.embedded || !loaded.has(f.name))
    .map((f) => (f.embedded ? f.name : `${f.name} (not embedded)`));
  if (foreign.length)
    return `the printed PDF draws with fonts that the page did not load: ${foreign.join(', ')}; make the paint use the faces that the measurer loaded, or load the missing face`;
  const [w, h] = ours.size;
  if (
    Math.abs(w - referenceSize[0]) > SAME_SIZE_PT ||
    Math.abs(h - referenceSize[1]) > SAME_SIZE_PT
  )
    return `the printed page is ${w} x ${h} pt, the reference page ${referenceSize[0]} x ${referenceSize[1]} pt; set the \`@page\` size in \`page/index.html\` to the reference page size`;
  if (paintedTexts > 0 && ours.items.length + ours.rotated === 0)
    return `the paint drew ${paintedTexts} texts, but the printed page holds no text; paint the canvas during \`beforeprint\` (\`page/main.ts\`), or Chromium prints it as an image`;
  return undefined;
}

/** The text of page 1 of a PDF, given as a path or as bytes. */
export async function pdfItems(source: string | Uint8Array): Promise<PdfItems> {
  const task = getDocument({
    // pdf.js takes the buffer over: give it a copy.
    data: typeof source === 'string' ? new Uint8Array(readFileSync(source)) : source.slice(),
    useSystemFonts: false,
    // pdf.js gives a font's `name` only with its extra properties.
    fontExtraProperties: true,
    cMapUrl: `${join(PDFJS_DIR, 'cmaps')}/`,
    cMapPacked: true,
    verbosity: 0,
  });
  const doc = await task.promise;
  try {
    const page = await doc.getPage(1);
    const viewport = page.getViewport({ scale: 1 });
    // The font of each `setFont`: its PostScript name, and whether the PDF embeds it.
    const ops = await page.getOperatorList();
    const named = new Map<string, { name: string; embedded: boolean }>();
    ops.fnArray.forEach((fn, i) => {
      if (fn !== OPS.setFont) return;
      const id = String((ops.argsArray[i] as unknown[])[0]);
      const font = page.commonObjs.get(id) as { name?: unknown; missingFile?: boolean } | undefined;
      named.set(id, {
        name: baseFontName(String(font?.name ?? 'unnamed')),
        embedded: font?.missingFile !== true,
      });
    });
    const runs = textRuns(ops.fnArray, ops.argsArray, (x, y) => {
      const [px, py] = viewport.convertToViewportPoint(x, y);
      return [px, py];
    });
    const items: TextItem[] = [];
    let rotated = 0;
    for (const item of (await page.getTextContent()).items) {
      // Whitespace items carry no glyph; spaces come from the gaps instead (`groupLines`).
      if (!('str' in item) || !item.str.trim()) continue;
      const [a, b, c, d, e, f] = item.transform as number[];
      // Rotated, mirrored, and vertical text is checked by pixels only.
      if (Math.abs(b) > 1e-6 || Math.abs(c) > 1e-6 || a <= 0 || d <= 0) {
        rotated += 1;
        continue;
      }
      const [x, y] = viewport.convertToViewportPoint(e, f);
      const postScript = named.get(item.fontName)?.name;
      const whole = { text: item.str, x, width: item.width };
      for (const part of splitItemByRuns(runs, whole, y, SAME_ROW_PT)) {
        const clip = clipOfItem(runs, part.x, y, SAME_ROW_PT);
        items.push({
          text: part.text,
          x: part.x,
          y,
          width: part.width,
          size: d,
          font: item.fontName,
          ...(postScript ? { postScript } : {}),
          ...(clip ? { clip } : {}),
        });
      }
    }
    const fonts = [
      ...new Map([...named.values()].map((f) => [`${f.name}\0${f.embedded}`, f])).values(),
    ].sort((a, b) => a.name.localeCompare(b.name));
    return {
      items,
      rotated,
      size: [viewport.width, viewport.height],
      pages: doc.numPages,
      fonts,
      runs,
    };
  } finally {
    await task.destroy();
  }
}
