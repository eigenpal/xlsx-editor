// The metric-compatible substitute faces of `@docx-editor.dev/fonts` as editor font sources (design
// D9). The editor loads them when the host passes no fonts, as DOCX Editor's and PPTX Editor's
// cores do; a host can also call this itself. A workbook never makes the engine fetch. The package
// fetches its own packaged files, checks each one by size, and reports a face that did not load.

import { type ALL_DEFAULT_FONT_FAMILIES, loadDefaultFonts } from '@docx-editor.dev/fonts';
import type { FontSource } from '../layout/index.ts';
import type { EditorFailure } from './editor-model.ts';

/**
 * A face that did not load, from `loadSubstituteFonts`.
 *
 * @public
 */
export interface SubstituteFontFailure {
  readonly family: string;
  readonly file: string;
  readonly diagnostic: string;
}

/**
 * The packaged families whose substitute faces layout draws with (`FONT_SUBSTITUTES`). The
 * package also packs Selawik for Segoe UI, but layout draws Segoe UI with Liberation Sans, as
 * PPTX Editor's measurer does. Loaded, Selawik would only join the glyph fallback faces, which
 * would change the fallback order. A test keeps this list equal to the packaged families whose
 * substitute `FONT_SUBSTITUTES` names.
 *
 * @internal
 */
export const SUBSTITUTE_FAMILIES = Object.freeze([
  'Calibri',
  'Cambria',
  'Times New Roman',
  'Arial',
  'Courier New',
  'Century Gothic',
] as const satisfies readonly (typeof ALL_DEFAULT_FONT_FAMILIES)[number][]);

/**
 * How long one substitute face download may send no byte before it fails. A slow download that
 * still sends bytes goes on, up to `SUBSTITUTE_FONT_TOTAL_MS`.
 *
 * @internal
 */
export const SUBSTITUTE_FONT_STALL_MS = 10_000;

/**
 * The longest time that the substitute face downloads of one load may take. The faces that did
 * not arrive by then fail, so a very slow network cannot hold the open of a workbook.
 *
 * @internal
 */
export const SUBSTITUTE_FONT_TOTAL_MS = 30_000;

/**
 * The substitute face downloads that run at one time. A browser queues more than this over
 * HTTP/1.1, and a download that waits in that queue sends no byte through no fault of its own.
 *
 * @internal
 */
export const SUBSTITUTE_FONT_DOWNLOADS = 6;

/**
 * The part of the stall limit that applies after the first stall of a batch. Stalled faces can
 * fill the download slots group after group, so a full limit for each group would hold the open
 * for many limits. With 24 faces, six at one time, stalled faces hold the open for at most 17.5
 * seconds.
 *
 * @internal
 */
export const SUBSTITUTE_FONT_LATER_STALL_PART = 0.25;

/** The call shape of `fetch`, with none of the extra members that a runtime adds to it. */
type FetchCall = (input: RequestInfo | URL, init?: RequestInit) => Promise<Response>;

/** `work`, or the reason of `signal` when it aborts first, also when `work` ignores the signal. */
function untilAbort<T>(work: Promise<T>, signal: AbortSignal): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const stop = () => reject(signal.reason);
    if (signal.aborted) stop();
    else signal.addEventListener('abort', stop, { once: true });
    work.then(resolve, reject).finally(() => signal.removeEventListener('abort', stop));
  });
}

/** One download that runs. */
interface RunningDownload {
  readonly control: AbortController;
  /** Whether the download started after another download of the batch stalled. */
  readonly afterStall: boolean;
  /** The bytes reads of the batch when the download started. */
  readonly readsAtStart: number;
  timer?: ReturnType<typeof setTimeout>;
  deadline: number;
  /** The limit that the timer of the download counts. */
  waitMs: number;
}

/**
 * `base` for one batch of downloads. At most `maxDownloads` run at one time. A download that
 * sends no byte for `stallMs` fails, and the next one in the queue starts. After the first stall,
 * the limit is shorter (`SUBSTITUTE_FONT_LATER_STALL_PART`). A download that started after a stall
 * and saw no byte from any download shows that the network gives nothing, so the queue fails too.
 * The batch ends at `totalMs` in all. Each read races the abort, as `fetchPptxSource` of
 * pptx-editor/packages/core/src/editor/pptx-source.ts at e22f3f3b (EigenPal) does.
 *
 * @internal
 */
export function stallLimitedFetch(
  base: FetchCall,
  stallMs: number,
  maxDownloads = SUBSTITUTE_FONT_DOWNLOADS,
  totalMs = SUBSTITUTE_FONT_TOTAL_MS
): FetchCall {
  const running = new Set<RunningDownload>();
  // The slots in use. A slot is taken at the call, before the download starts.
  let taken = 0;
  const queue: { start: () => void; fail: (reason: unknown) => void }[] = [];
  let stalled = false;
  let reads = 0;
  let limitMs = stallMs;
  let total: ReturnType<typeof setTimeout> | undefined;

  const failQueue = (reason: unknown) => {
    for (const waiting of queue.splice(0)) waiting.fail(reason);
  };
  const arm = (download: RunningDownload) => {
    clearTimeout(download.timer);
    download.deadline = Date.now() + limitMs;
    download.waitMs = limitMs;
    download.timer = setTimeout(() => onStall(download), limitMs);
  };
  const onStall = (download: RunningDownload) => {
    const reason = new Error(`no bytes for ${download.waitMs} ms`);
    if (download.afterStall && reads === download.readsAtStart) failQueue(reason);
    if (!stalled) {
      stalled = true;
      limitMs = stallMs * SUBSTITUTE_FONT_LATER_STALL_PART;
      // A download that runs gets the shorter limit too, and never a longer one.
      for (const other of running)
        if (other !== download && other.deadline > Date.now() + limitMs) arm(other);
    }
    download.control.abort(reason);
  };
  const onTotal = () => {
    const reason = new Error(`the downloads passed ${totalMs} ms`);
    failQueue(reason);
    for (const download of running) download.control.abort(reason);
  };

  /** Waits for a free download slot. It rejects when the caller aborts or the network is dead. */
  const slot = (outer: AbortSignal | null | undefined) =>
    new Promise<void>((resolve, reject) => {
      if (taken < maxDownloads) {
        taken += 1;
        return resolve();
      }
      const entry = {
        start: () => {
          outer?.removeEventListener('abort', leave);
          resolve();
        },
        fail: (reason: unknown) => {
          outer?.removeEventListener('abort', leave);
          reject(reason);
        },
      };
      const leave = () => {
        queue.splice(queue.indexOf(entry), 1);
        reject(outer?.reason);
      };
      outer?.addEventListener('abort', leave, { once: true });
      queue.push(entry);
    });

  return async (input, init) => {
    const outer = init?.signal;
    outer?.throwIfAborted();
    total ??= setTimeout(onTotal, totalMs);
    await slot(outer);
    const download: RunningDownload = {
      control: new AbortController(),
      afterStall: stalled,
      readsAtStart: reads,
      deadline: 0,
      waitMs: limitMs,
    };
    const forward = () => download.control.abort(outer?.reason);
    outer?.addEventListener('abort', forward, { once: true });
    if (outer?.aborted) forward();
    running.add(download);
    arm(download);
    const progress = () => {
      reads += 1;
      arm(download);
    };
    try {
      const signal = download.control.signal;
      const response = await untilAbort(base(input, { ...init, signal }), signal);
      progress();
      if (!response.body) return response;
      const reader = response.body.getReader();
      const chunks: Uint8Array<ArrayBuffer>[] = [];
      try {
        for (;;) {
          const { done, value } = await untilAbort(reader.read(), signal);
          if (done) break;
          chunks.push(value);
          progress();
        }
      } catch (cause) {
        void reader.cancel(cause).catch(() => undefined);
        throw cause;
      }
      return new Response(new Blob(chunks), {
        status: response.status,
        statusText: response.statusText,
        headers: response.headers,
      });
    } finally {
      clearTimeout(download.timer);
      outer?.removeEventListener('abort', forward);
      running.delete(download);
      const next = queue.shift();
      if (next) next.start();
      else taken -= 1;
      if (running.size === 0 && queue.length === 0 && taken === 0) {
        clearTimeout(total);
        total = undefined;
      }
    }
  };
}

/**
 * Fetch the packaged substitute faces (Carlito, Caladea, Liberation Sans, Serif, and Mono, and
 * TeX Gyre Adventor). Resolves with every face that loaded and the failures. A face whose
 * download stalls ({@link stallLimitedFetch}) is a failure.
 *
 * @public
 */
export async function loadSubstituteFonts(
  options: { readonly signal?: AbortSignal } = {}
): Promise<{
  readonly sources: readonly FontSource[];
  readonly failures: readonly SubstituteFontFailure[];
}> {
  const fragment = await loadDefaultFonts({
    families: SUBSTITUTE_FAMILIES,
    // Read `fetch` at call time, so a host or a test that replaces it is the one in force.
    fetcher: stallLimitedFetch(
      (input, init) => fetch(input, init),
      SUBSTITUTE_FONT_STALL_MS
    ) as typeof fetch,
    ...(options.signal ? { signal: options.signal } : {}),
  });
  return {
    sources: fragment.sources.map((s) => ({
      family: s.request.family,
      bold: s.request.weight >= 600,
      italic: s.request.style === 'italic',
      bytes: s.bytes,
      origin: 'substitute',
    })),
    failures: fragment.failures,
  };
}

/**
 * The fonts of an editor, after PPTX Editor's `FontReport`: the faces in force, and the packaged
 * substitute faces that did not load. Text in a family whose face failed measures and paints
 * with another face.
 *
 * @public
 */
export interface FontReport {
  readonly faces: readonly { readonly family: string; readonly origin: string }[];
  readonly substituteFailures: readonly SubstituteFontFailure[];
}

/**
 * The failure that reports packaged substitute faces that did not load, for `onFontError`, or
 * null when every face loaded.
 *
 * @internal
 */
export function substituteFontFailure(
  failures: readonly SubstituteFontFailure[]
): EditorFailure | null {
  if (failures.length === 0) return null;
  const files = failures.map((f) => f.file || f.family).join(', ');
  return {
    code: 'fonts',
    detail: `substitute faces did not load: ${files}`,
    substituteFailures: failures,
  };
}

/**
 * Hand a font failure to the host's `onFontError`, or print one warning when the host has none,
 * so a failure is never silent, as PPTX Editor's `reportFontError` does.
 *
 * @internal
 */
export function reportFontError(
  failure: EditorFailure,
  handler: ((failure: EditorFailure) => void) | undefined
): void {
  if (handler) handler(failure);
  else console.warn(`[xlsx-editor] fonts: ${failure.detail ?? 'the faces did not load'}`);
}
