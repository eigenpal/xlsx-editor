// The zoom of Mod+wheel and a trackpad pinch (a wheel event with the control key): one zoom step
// for each 100 pixels of `deltaY`, so a trackpad that sends many small deltas zooms as a wheel
// notch does. A touch screen has no pinch zoom: the grid takes `touch-action: pan-x pan-y`, so
// two fingers scroll the grid and the zoom controls change the zoom.

/** The `deltaY` of Mod+wheel events that makes one zoom step, in pixels (one wheel notch). */
const WHEEL_ZOOM_STEP = 100;

/** Adds the deltas of zoom wheel events and tells when they make a step. */
export class WheelZoom {
  #sum = 0;

  /**
   * The zoom step of one wheel event: 1 to zoom in, -1 to zoom out, 0 for none yet. An event
   * without the control or meta key is not a zoom (`undefined`), and the caller lets it scroll.
   */
  step(e: WheelEvent): 1 | -1 | 0 | undefined {
    if (!(e.ctrlKey || e.metaKey)) return undefined;
    const scale = e.deltaMode === 1 ? 33 : e.deltaMode === 2 ? 400 : 1;
    const delta = e.deltaY * scale;
    if (!Number.isFinite(delta) || delta === 0) return 0;
    // A change of direction starts the count again.
    if (Math.sign(delta) !== Math.sign(this.#sum)) this.#sum = 0;
    this.#sum += delta;
    if (Math.abs(this.#sum) < WHEEL_ZOOM_STEP) return 0;
    const dir = this.#sum < 0 ? 1 : -1;
    this.#sum = 0;
    return dir;
  }

  /**
   * Handle a wheel event on the grid: a zoom event zooms one step through `zoomStep` when its
   * deltas make one, and never scrolls or zooms the page. Any other event scrolls.
   */
  handle(e: WheelEvent, zoomStep: (direction: 1 | -1) => void): void {
    const step = this.step(e);
    if (step === undefined) return;
    e.preventDefault();
    if (step !== 0) zoomStep(step);
  }
}
