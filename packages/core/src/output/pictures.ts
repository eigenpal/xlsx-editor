// Pictures from package bytes (design D7, AGENTS.md "Security"). A picture part decodes with the
// host window's `createImageBitmap` from a `Blob` of its bytes: no URL is made, so no image sink
// and no network request exists. Only raster types whose header size `imageSize` reads are drawn
// (PNG, JPEG, GIF, BMP); any other picture paints as a placeholder (approximation "Drawing
// objects"). The header size is read before decoding. A picture whose size cannot be read, or
// that is larger than `MAX_PICTURE_PIXELS`, is refused, because the bytes come from the file.

import { imageSize } from '@eigenpal/ooxml-core';
import { binaryPartOf, type XlsxPackage } from '../store/index.ts';

/** A picture larger than this many pixels is not decoded. */
export const MAX_PICTURE_PIXELS = 50_000_000;

const RASTER_TYPES = new Set(['image/png', 'image/jpeg', 'image/gif', 'image/bmp']);

/** The window functions that decoding needs: the window that owns the paint target. */
export interface PictureWindow {
  readonly Blob: typeof Blob;
  createImageBitmap(image: Blob): Promise<ImageBitmap>;
}

/** The state of one picture part. */
export type PictureState =
  | { readonly kind: 'loading' }
  | { readonly kind: 'ready'; readonly bitmap: ImageBitmap }
  | { readonly kind: 'unsupported'; readonly reason: string };

/** Decoded pictures of one package, by part name. */
export interface PictureStore {
  /** The state of a picture part; the first call starts decoding. */
  get(partName: string, contentType: string | undefined): PictureState;
  /** Called once each time a picture finishes decoding, so the host can paint again. */
  onReady(listener: () => void): () => void;
  dispose(): void;
}

/** A picture store over the bytes of `pkg`, decoding in `win`. */
export function createPictureStore(pkg: XlsxPackage, win: PictureWindow | undefined): PictureStore {
  const states = new Map<string, PictureState>();
  const listeners = new Set<() => void>();
  let disposed = false;
  const settle = (partName: string, state: PictureState): void => {
    if (disposed) {
      if (state.kind === 'ready') state.bitmap.close();
      return;
    }
    states.set(partName, state);
    for (const l of [...listeners]) l();
  };
  return {
    get(partName, contentType) {
      const known = states.get(partName);
      if (known) return known;
      const type = (contentType ?? '').toLowerCase();
      const bytes = binaryPartOf(pkg, partName);
      let state: PictureState;
      if (!bytes) state = { kind: 'unsupported', reason: 'missing-part' };
      else if (!RASTER_TYPES.has(type)) state = { kind: 'unsupported', reason: 'content-type' };
      else if (!win || typeof win.createImageBitmap !== 'function')
        state = { kind: 'unsupported', reason: 'no-decoder' };
      else {
        const size = imageSize(bytes);
        if (!size) state = { kind: 'unsupported', reason: 'unknown-size' };
        else if (size.width * size.height > MAX_PICTURE_PIXELS)
          state = { kind: 'unsupported', reason: 'too-large' };
        else {
          state = { kind: 'loading' };
          win.createImageBitmap(new win.Blob([bytes.slice()], { type })).then(
            (bitmap) => settle(partName, { kind: 'ready', bitmap }),
            () => settle(partName, { kind: 'unsupported', reason: 'decode' })
          );
        }
      }
      states.set(partName, state);
      return state;
    },
    onReady(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    dispose() {
      disposed = true;
      for (const s of states.values()) if (s.kind === 'ready') s.bitmap.close();
      states.clear();
      listeners.clear();
    },
  };
}
