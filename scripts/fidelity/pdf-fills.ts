// Read the fill geometry of a reference PDF: the tiling patterns of pattern fills, the axial
// shadings of linear gradients, and the images of path gradients. Tests compare the paint tables
// with what the reference PDF of fixture `fills-patterns` draws, so the tables cannot drift from
// it.
//
// The reader handles the PDF that the reference export writes: uncompressed object syntax with
// Flate streams. It reads the page content in order, and pairs each pattern fill and each shading
// with the rectangle that it paints, in page points with the origin at the top left.

import { inflateSync } from 'node:zlib';

const decoder = new TextDecoder('latin1');

interface PdfObjects {
  readonly dict: (id: number) => string;
  readonly stream: (id: number) => Uint8Array;
}

function objects(bytes: Uint8Array): PdfObjects {
  const text = decoder.decode(bytes);
  const at = new Map<number, number>();
  // The first match of each object wins: a later one can be bytes inside a binary stream.
  for (const m of text.matchAll(/(\d+) 0 obj/g))
    if (!at.has(Number(m[1]))) at.set(Number(m[1]), m.index!);
  const body = (id: number): string => {
    const start = at.get(id);
    if (start === undefined) throw new Error(`no object ${id}`);
    return text.slice(start, text.indexOf('endobj', start));
  };
  return {
    dict: (id) => body(id).split('stream')[0]!,
    stream: (id) => {
      const start = at.get(id)!;
      const s = text.indexOf('stream', start) + 'stream'.length;
      const from = text[s] === '\r' ? s + 2 : s + 1;
      const end = text.indexOf('endstream', from);
      const raw = bytes.subarray(from, end);
      return body(id).includes('/FlateDecode') ? new Uint8Array(inflateSync(raw)) : raw;
    },
  };
}

const ref = (dict: string, key: string): number => {
  const m = new RegExp(`/${key}\\s+(\\d+)\\s+0\\s+R`).exec(dict);
  if (!m) throw new Error(`no /${key}`);
  return Number(m[1]);
};

/** A pattern fill on the page: its rectangle, tile step in points, and tile colors by row. */
export interface PdfPatternFill {
  /** Top-left origin, points. */
  readonly rect: { x: number; y: number; width: number; height: number };
  readonly step: number;
  /** The tile pattern space has no matrix: its grid starts at the page origin. */
  readonly pageAnchored: boolean;
  /** `RRGGBB` of each tile pixel, row 0 at the top. */
  readonly pixels: readonly (readonly string[])[];
}

/** A path gradient on the page: its cell rectangle, and its pixels, row 0 at the top. */
export interface PdfCellImage {
  readonly rect: { x: number; y: number; width: number; height: number };
  readonly pixels: readonly (readonly string[])[];
}

/** A linear gradient on the page: its cell rectangle, and its end points in the unit square. */
export interface PdfAxialShading {
  readonly rect: { x: number; y: number; width: number; height: number };
  /** `[x0, y0, x1, y1]` in the unit square of the cell, y down. */
  readonly coords: readonly [number, number, number, number];
  /** The color at the start and the end, `RRGGBB`. */
  readonly from: string;
  readonly to: string;
  /** The color at a position from 0 to 1, from the sampled function. */
  readonly colorAt: (position: number) => string;
}

/**
 * Fail unless the matrix `a 0 0 d e f` maps the unit square onto the clip `x y w h`: a cell
 * reading of the shading or image would otherwise be wrong with no error.
 */
function onClip(matrix: readonly string[], clip: readonly string[], what: string): void {
  const [a, d, e, f] = matrix.map(Number) as [number, number, number, number];
  const [x, y, w, h] = clip.map(Number) as [number, number, number, number];
  const bottom = d < 0 ? f + d : f;
  const near = (p: number, q: number) => Math.abs(p - q) < 0.01;
  if (!(near(a, w) && near(Math.abs(d), h) && near(e, x) && near(bottom, y)))
    throw new Error(`${what}: the unit square is not the clip rectangle`);
}

const hex = (b: Uint8Array, i: number): string =>
  Array.from(b.subarray(i, i + 3), (v) => v.toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();

/** The pattern fills, axial shadings, and path gradient images of page 1. */
export function readPdfFills(bytes: Uint8Array): {
  patterns: PdfPatternFill[];
  shadings: PdfAxialShading[];
  images: PdfCellImage[];
} {
  const pdf = objects(bytes);
  const text = decoder.decode(bytes);
  const pageId = Number(/(\d+) 0 obj\s*<<\s*\/Type \/Page\b(?!s)/.exec(text)?.[1]);
  const page = pdf.dict(pageId);
  const height = Number(/\/MediaBox \[\s*\S+ \S+ \S+ (\S+)\s*\]/.exec(page)?.[1]);
  const resId = /\/Resources (\d+) 0 R/.exec(page);
  const resources = resId ? pdf.dict(Number(resId[1])) : page;
  const content = decoder.decode(pdf.stream(ref(page, 'Contents'))).replace(/\s+/g, ' ');
  const patterns: PdfPatternFill[] = [];
  const fillRe = /\/(P\d+) scn (\S+) (\S+) m (\S+) \S+ l \S+ (\S+) l/g;
  for (const m of content.matchAll(fillRe)) {
    const [, name, x0, y0, x1, y1] = m;
    const patternId = ref(resources, name!);
    const dict = pdf.dict(patternId);
    // The tile is one image of step x step points with row 0 at the top. Another placement would
    // read the tile flipped or at another phase, so it fails here.
    const step = Number(/\/XStep (\S+)/.exec(dict)?.[1]);
    const placed =
      Number(/\/YStep (\S+)/.exec(dict)?.[1]) === step &&
      new RegExp(`/BBox \\[\\s*0 0 ${step} ${step}\\s*\\]`).test(dict) &&
      decoder
        .decode(pdf.stream(patternId))
        .replace(/\s+/g, ' ')
        .includes(`q ${step} 0 0 ${step} 0 0 cm /Im`);
    if (!placed) throw new Error(`pattern ${name}: the tile is not one image of the step`);
    const image = pdf.dict(ref(pdf.dict(ref(dict, 'Resources')), 'Im\\d+'));
    const width = Number(/\/Width (\d+)/.exec(image)?.[1]);
    const data = pdf.stream(ref(pdf.dict(ref(dict, 'Resources')), 'Im\\d+'));
    const pixels = Array.from({ length: Number(/\/Height (\d+)/.exec(image)?.[1]) }, (_, r) =>
      Array.from({ length: width }, (_, c) => hex(data, (r * width + c) * 3))
    );
    const top = height - Math.max(Number(y0), Number(y1));
    patterns.push({
      rect: {
        x: Number(x0),
        y: top,
        width: Number(x1) - Number(x0),
        height: Math.abs(Number(y1) - Number(y0)),
      },
      step,
      pageAnchored: !dict.includes('/Matrix'),
      pixels,
    });
  }
  const shadings: PdfAxialShading[] = [];
  // A clip to the cell, then one or more shadings in the unit square of the cell (y down). The
  // shading whose axis lies in the unit square is the gradient; the others paint outside the clip.
  const groupRe = /q (\S+) (\S+) (\S+) (\S+) re W n \/Perceptual ri ((?:q [^Q]*? sh Q ?)+)/g;
  for (const m of content.matchAll(groupRe)) {
    const [, x, y, w, h, body] = m;
    for (const sh of body!.matchAll(/q (\S+) 0 0 (\S+) (\S+) (\S+) cm \/(Sh\d+) sh/g)) {
      const dict = pdf.dict(ref(resources, sh[5]!));
      const coords = /\/Coords \[\s*(\S+) (\S+) (\S+) (\S+)\s*\]/.exec(dict)!.slice(1).map(Number);
      if (coords.some((v) => v < 0 || v > 1)) continue;
      onClip([sh[1]!, sh[2]!, sh[3]!, sh[4]!], [x!, y!, w!, h!], sh[5]!);
      // A negative height puts y = 0 at the top of the cell; a positive one at the bottom.
      if (Number(sh[2]) > 0) [coords[1], coords[3]] = [1 - coords[1]!, 1 - coords[3]!];
      // The colors are a sampled function of 8 bit RGB over 0 to 1; another kind fails here.
      const fnId = ref(dict, 'Function');
      const fn = pdf.dict(fnId).replace(/\s+/g, ' ');
      if (
        !/\/FunctionType 0\b/.test(fn) ||
        !/\/BitsPerSample 8\b/.test(fn) ||
        !/\/Domain \[ ?0 1 ?\]/.test(fn) ||
        /\/Encode|\/Decode/.test(fn)
      )
        throw new Error(`${sh[5]}: the shading function is not 8 bit samples over 0 to 1`);
      const samples = pdf.stream(fnId);
      shadings.push({
        rect: {
          x: Number(x),
          y: height - Number(y) - Number(h),
          width: Number(w),
          height: Number(h),
        },
        coords: coords as unknown as [number, number, number, number],
        from: hex(samples, 0),
        to: hex(samples, samples.length - 3),
        colorAt: (t) => hex(samples, 3 * Math.round(t * (samples.length / 3 - 1))),
      });
    }
  }
  // Path gradients: an image clipped to its cell. A clip of the whole page holds the rows of a
  // merged range, which the export draws one image at a time.
  const images: PdfCellImage[] = [];
  const imageRe =
    /q (\S+) (\S+) (\S+) (\S+) re W n \/Perceptual ri q (\S+) 0 0 (\S+) (\S+) (\S+) cm \/(Im\d+) Do Q/g;
  for (const m of content.matchAll(imageRe)) {
    const [, x, y, w, h, a, d, e, f, name] = m;
    if (Number(w) >= Number(/\/MediaBox \[\s*\S+ \S+ (\S+)/.exec(page)?.[1])) continue;
    onClip([a!, d!, e!, f!], [x!, y!, w!, h!], name!);
    const id = ref(resources, name!);
    const dict = pdf.dict(id);
    const width = Number(/\/Width (\d+)/.exec(dict)?.[1]);
    const rows = Number(/\/Height (\d+)/.exec(dict)?.[1]);
    const data = pdf.stream(id);
    // A positive height draws image row 0 at the top of the cell.
    const flipped = Number(d) < 0;
    images.push({
      rect: {
        x: Number(x),
        y: height - Number(y) - Number(h),
        width: Number(w),
        height: Number(h),
      },
      pixels: Array.from({ length: rows }, (_, r) => {
        const source = flipped ? rows - 1 - r : r;
        return Array.from({ length: width }, (_, c) => hex(data, (source * width + c) * 3));
      }),
    });
  }
  return { patterns, shadings, images };
}
