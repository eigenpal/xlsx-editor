// The fonts of one editor: the faces that measure and paint text, loaded once from the host's
// source, the measurer built over them, and their `@font-face` rules in the host document.
// `replace` swaps the source in place, as DOCX Editor's `fonts` prop does: the new faces load
// first, and only a load that succeeds replaces the faces that show the workbook.

import {
  createHarfBuzzMeasurer,
  initializeHarfBuzz,
  type FontSource,
  type RegisteredFace,
  type TextMeasurer,
} from '../layout/index.ts';
import { registerFontFaces, type FontFaceHandle } from '../output/index.ts';
import {
  loadSubstituteFonts,
  type FontReport,
  type SubstituteFontFailure,
} from './substitute-fonts.ts';

/** The faces of an editor: bytes, or a function that loads them once. */
export type EditorFontSource = readonly FontSource[] | (() => Promise<readonly FontSource[]>);

/**
 * Whether two font sources hold the same faces: the same function, or lists of equal length
 * whose faces have the same family, style, origin, and bytes (by identity). A host can pass a
 * new list with each render; only a change of faces loads them again.
 */
export function sameFontSource(
  a: EditorFontSource | undefined,
  b: EditorFontSource | undefined
): boolean {
  if (a === b) return true;
  if (!a || !b || typeof a === 'function' || typeof b === 'function') return false;
  if (a.length !== b.length) return false;
  return a.every((x, i) => {
    const y = b[i]!;
    return (
      x.family === y.family &&
      x.bold === y.bold &&
      x.italic === y.italic &&
      x.origin === y.origin &&
      x.bytes === y.bytes
    );
  });
}

export type FontMeasurer = TextMeasurer & { readonly faces: readonly RegisteredFace[] };

/** A measurer, with the packaged substitute faces that did not load for it. */
interface Loaded {
  readonly measurer: FontMeasurer;
  readonly failures: readonly SubstituteFontFailure[];
}

const NO_FAILURES: readonly SubstituteFontFailure[] = Object.freeze([]);

/** The loader of the packaged substitute faces: `loadSubstituteFonts`, or a test's own. */
type SubstituteLoader = () => ReturnType<typeof loadSubstituteFonts>;

async function measurerOf(
  source: EditorFontSource | undefined,
  loadSubstitutes: SubstituteLoader
): Promise<Loaded> {
  await initializeHarfBuzz();
  // No source: the packaged substitute faces, as DOCX Editor's and PPTX Editor's cores do. The
  // faces that did not load are kept for `fontReport`, as PPTX Editor's session keeps them.
  let failures = NO_FAILURES;
  let sources: readonly FontSource[];
  if (typeof source === 'function') sources = await source();
  else if (source) sources = source;
  else {
    const loaded = await loadSubstitutes();
    sources = loaded.sources;
    if (loaded.failures.length > 0) failures = Object.freeze([...loaded.failures]);
  }
  if (sources.length === 0) {
    const lost = failures.map((f) => f.file || f.family).join(', ');
    throw new Error(
      lost ? `no font sources; substitute faces did not load: ${lost}` : 'no font sources'
    );
  }
  return { measurer: createHarfBuzzMeasurer(sources), failures };
}

export class EditorFonts {
  #source: EditorFontSource | undefined;
  #measurer: FontMeasurer | null = null;
  #failures: readonly SubstituteFontFailure[] = NO_FAILURES;
  #load: Promise<Loaded> | null = null;
  /** Bumped by each `replace`, so a load of an older source never takes effect. */
  #generation = 0;
  #faces: { doc: Document; handle: FontFaceHandle } | null = null;

  readonly #loadSubstitutes: SubstituteLoader;

  constructor(
    source: EditorFontSource | undefined,
    loadSubstitutes: SubstituteLoader = loadSubstituteFonts
  ) {
    this.#source = source;
    this.#loadSubstitutes = loadSubstitutes;
  }

  /** Whether `source` holds the same faces as the current source (`sameFontSource`). */
  holds(source: EditorFontSource | undefined): boolean {
    return sameFontSource(source, this.#source);
  }

  /** The measurer, or null before the first load. */
  get measurer(): FontMeasurer | null {
    return this.#measurer;
  }

  /** The faces in force and the packaged substitute faces that did not load for them. */
  get report(): FontReport {
    const faces = (this.#measurer?.faces ?? []).map((f) => ({
      family: f.family,
      origin: f.origin,
    }));
    return { faces, substituteFailures: this.#failures };
  }

  /** The packaged substitute faces that did not load for the faces in force. */
  get substituteFailures(): readonly SubstituteFontFailure[] {
    return this.#failures;
  }

  /**
   * Load the faces once. A failed load can be tried again. When `replace` changes the source
   * while the load runs, the load of the old source is dropped and the new source loads.
   */
  async ensure(): Promise<void> {
    while (!this.#measurer) {
      const generation = this.#generation;
      const load = (this.#load ??= measurerOf(this.#source, this.#loadSubstitutes));
      let loaded: Loaded;
      try {
        loaded = await load;
      } catch (cause) {
        if (this.#load === load) this.#load = null;
        if (generation !== this.#generation) continue;
        throw cause;
      }
      if (generation !== this.#generation || this.#measurer) continue;
      this.#measurer = loaded.measurer;
      this.#failures = loaded.failures;
    }
  }

  /**
   * Replace the source. Before the first load, the next `ensure` reads it, and a load that
   * runs is dropped. After it, the new faces load now and replace the shown faces only when
   * they load; a failure keeps the old faces and rejects.
   */
  async replace(source: EditorFontSource | undefined): Promise<void> {
    this.#source = source;
    const generation = (this.#generation += 1);
    if (!this.#measurer) {
      this.#load = null;
      return;
    }
    const next = await measurerOf(source, this.#loadSubstitutes);
    if (generation !== this.#generation) return;
    this.#measurer = next.measurer;
    this.#failures = next.failures;
    this.#load = null;
    this.#faces?.handle.dispose();
    this.#faces = null;
  }

  /** Register the faces in `doc` once; `onReady` runs when the browser has them. */
  register(doc: Document, onReady: () => void): void {
    if (!this.#measurer || this.#faces?.doc === doc) return;
    this.#faces?.handle.dispose();
    const handle = registerFontFaces(doc, this.#measurer.faces);
    this.#faces = { doc, handle };
    void handle.ready.then(onReady);
  }

  /** Remove the `@font-face` rules from the document they were registered in. */
  unregister(): void {
    this.#faces?.handle.dispose();
    this.#faces = null;
  }
}
