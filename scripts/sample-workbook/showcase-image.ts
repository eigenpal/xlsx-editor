// The one picture of the sample workbook: a small landscape drawn pixel by pixel, so the sample
// holds no third-party image. The encoder is pure JavaScript (fflate), so the bytes are the
// same on every machine.
//
// Same role as pptx-editor/scripts/sample-deck/showcase-image.ts at 79a5d22d (EigenPal).
// Difference: it writes the PNG itself instead of through pngjs, whose output depends on the
// zlib build of the runtime, so the committed file can be compared byte for byte.

import { zlibSync } from 'fflate';

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(bytes: Uint8Array): number {
  let c = 0xffffffff;
  for (const b of bytes) c = CRC_TABLE[(c ^ b) & 0xff]! ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function chunk(type: string, data: Uint8Array): Uint8Array {
  const out = new Uint8Array(12 + data.length);
  const view = new DataView(out.buffer);
  view.setUint32(0, data.length);
  for (let i = 0; i < 4; i += 1) out[4 + i] = type.charCodeAt(i);
  out.set(data, 8);
  view.setUint32(8 + data.length, crc32(out.subarray(4, 8 + data.length)));
  return out;
}

/** Encode RGB pixels (`width * height * 3` bytes) as a PNG file. */
function encodePng(width: number, height: number, rgb: Uint8Array): Uint8Array {
  const raw = new Uint8Array(height * (1 + width * 3));
  for (let y = 0; y < height; y += 1)
    raw.set(rgb.subarray(y * width * 3, (y + 1) * width * 3), y * (1 + width * 3) + 1);
  const header = new Uint8Array(13);
  const view = new DataView(header.buffer);
  view.setUint32(0, width);
  view.setUint32(4, height);
  header[8] = 8; // bit depth
  header[9] = 2; // truecolor
  const parts = [
    Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', header),
    chunk('IDAT', zlibSync(raw, { level: 9 })),
    chunk('IEND', new Uint8Array(0)),
  ];
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let at = 0;
  for (const p of parts) {
    out.set(p, at);
    at += p.length;
  }
  return out;
}

const hex = (h: string) =>
  [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16)) as [number, number, number];

/** The picture's size in pixels. */
export const PICTURE_SIZE = { width: 240, height: 150 } as const;

/** Mix two RGB hex colors: `t` 0 gives `a`, 1 gives `b`. */
const mix = (a: string, b: string, t: number): [number, number, number] => {
  const [x, y] = [hex(a), hex(b)];
  return [0, 1, 2].map((i) => Math.round(x[i]! + (y[i]! - x[i]!) * t)) as [number, number, number];
};

/**
 * A landscape: a sky that fades to the horizon, a sun, and two hills in the greens of the brand
 * lockup. It is a plain illustration, so it does not suggest a chart, which the engine does not
 * draw.
 */
export function landscapePng(): Uint8Array {
  const { width, height } = PICTURE_SIZE;
  const px = new Uint8Array(width * height * 3);
  const set = (x: number, y: number, [r, g, b]: readonly number[]) => {
    const i = (y * width + x) * 3;
    px[i] = r!;
    px[i + 1] = g!;
    px[i + 2] = b!;
  };
  // A hill is a parabola: plain arithmetic, so every runtime computes the same pixels.
  const hill = (x: number, center: number, base: number, size: number) => {
    const t = (x - center) / (size * 2.2);
    return base - size * Math.max(0, 1 - t * t);
  };
  for (let y = 0; y < height; y += 1)
    for (let x = 0; x < width; x += 1) {
      let color = mix('9CC9E8', 'EAF4FA', y / height);
      if ((x - 176) * (x - 176) + (y - 38) * (y - 38) <= 18 * 18) color = hex('F6C343');
      if (y >= hill(x, 60, 122, 52)) color = mix('7FBF9A', '4FA474', (y - 70) / 80);
      if (y >= hill(x, 170, 138, 44)) color = mix('2A8C55', '107C41', (y - 94) / 56);
      set(x, y, color);
    }
  return encodePng(width, height, px);
}
