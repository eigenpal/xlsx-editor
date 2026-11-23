// The page probe of the memory bench: the resources that an editor must release, counted in the
// page. Installed before the demo's scripts, it wraps the browser functions that make each
// resource: object URLs, observers, workers, image bitmaps, intervals, and listeners on `window`
// and `document`. `document.fonts` gives the font faces.
//
// Copied from PPTX Editor's `e2e/memory.bench.ts` at d4502552 (EigenPal), where it first landed in
// f448820c (`installProbe`). Differences: the type is `Resources`, not `Probe`, it does not count
// SVG `image` elements, which the workbook grid does not make, and the module exports the type and
// `installProbe`, so `browser-bench.ts` can use them.

export interface Resources {
  blobs: number;
  fonts: number;
  observers: number;
  workers: number;
  bitmaps: number;
  intervals: number;
  windowListeners: number;
  documentListeners: number;
}

export function installProbe(): void {
  const w = window as unknown as { __memProbe: () => Resources };
  const blobs = new Set<string>();
  const create = URL.createObjectURL.bind(URL);
  const revoke = URL.revokeObjectURL.bind(URL);
  URL.createObjectURL = (o: Blob | MediaSource) => {
    const url = create(o);
    blobs.add(url);
    return url;
  };
  URL.revokeObjectURL = (url: string) => {
    blobs.delete(url);
    revoke(url);
  };

  // Observers that observe something and are neither disconnected nor collected.
  let observers = 0;
  const gone = new FinalizationRegistry<null>(() => (observers -= 1));
  for (const name of ['ResizeObserver', 'MutationObserver', 'IntersectionObserver'] as const) {
    const proto = (window[name] as { prototype: { observe: unknown; disconnect: unknown } })
      ?.prototype;
    if (!proto) continue;
    const live = new WeakSet<object>();
    const observe = proto.observe as (...a: unknown[]) => void;
    const disconnect = proto.disconnect as () => void;
    proto.observe = function (this: object, ...a: unknown[]) {
      if (!live.has(this)) {
        live.add(this);
        observers += 1;
        gone.register(this, null, this);
      }
      return observe.apply(this, a);
    };
    proto.disconnect = function (this: object) {
      if (live.delete(this)) {
        observers -= 1;
        gone.unregister(this);
      }
      return disconnect.call(this);
    };
  }

  let workers = 0;
  const WorkerCtor = window.Worker;
  if (WorkerCtor) {
    window.Worker = class extends WorkerCtor {
      #alive = true;
      constructor(...a: ConstructorParameters<typeof Worker>) {
        super(...a);
        workers += 1;
      }
      override terminate(): void {
        if (this.#alive) workers -= 1;
        this.#alive = false;
        super.terminate();
      }
    };
  }

  let bitmaps = 0;
  const bitmapGone = new FinalizationRegistry<null>(() => (bitmaps -= 1));
  const makeBitmap = window.createImageBitmap.bind(window) as (
    ...a: unknown[]
  ) => Promise<ImageBitmap>;
  (window as unknown as { createImageBitmap: unknown }).createImageBitmap = async (
    ...a: unknown[]
  ) => {
    const bitmap = await makeBitmap(...a);
    bitmaps += 1;
    bitmapGone.register(bitmap, null, bitmap);
    return bitmap;
  };
  const close = ImageBitmap.prototype.close;
  ImageBitmap.prototype.close = function (this: ImageBitmap) {
    if (bitmapGone.unregister(this)) bitmaps -= 1;
    close.call(this);
  };

  const intervals = new Set<number>();
  const setI = window.setInterval.bind(window);
  const clearI = window.clearInterval.bind(window);
  (window as unknown as { setInterval: unknown }).setInterval = (...a: unknown[]) => {
    const id = (setI as (...b: unknown[]) => number)(...a);
    intervals.add(id);
    return id;
  };
  (window as unknown as { clearInterval: unknown }).clearInterval = (id?: number) => {
    if (id !== undefined) intervals.delete(id);
    clearI(id);
  };

  // Listeners on `window` and `document`, which outlive any editor that forgets them.
  const listeners = new Map<EventTarget, Map<string, Set<unknown>>>();
  const key = (type: string, options: unknown): string =>
    `${type}:${typeof options === 'boolean' ? options : !!(options as { capture?: boolean })?.capture}`;
  const add = EventTarget.prototype.addEventListener;
  const remove = EventTarget.prototype.removeEventListener;
  EventTarget.prototype.addEventListener = function (this: EventTarget, type, fn, options) {
    if ((this === window || this === document) && fn) {
      const byType = listeners.get(this) ?? new Map<string, Set<unknown>>();
      listeners.set(this, byType);
      const k = key(type, options);
      const set = byType.get(k) ?? new Set();
      byType.set(k, set);
      set.add(fn);
      const signal = (options as AddEventListenerOptions | undefined)?.signal;
      signal?.addEventListener('abort', () => set.delete(fn), { once: true });
    }
    return add.call(this, type, fn, options);
  };
  EventTarget.prototype.removeEventListener = function (this: EventTarget, type, fn, options) {
    listeners.get(this)?.get(key(type, options))?.delete(fn);
    return remove.call(this, type, fn, options);
  };
  const count = (target: EventTarget): number =>
    [...(listeners.get(target)?.values() ?? [])].reduce((sum, set) => sum + set.size, 0);

  w.__memProbe = () => ({
    blobs: blobs.size,
    fonts: document.fonts?.size ?? 0,
    observers,
    workers,
    bitmaps,
    intervals: intervals.size,
    windowListeners: count(window),
    documentListeners: count(document),
  });
}
