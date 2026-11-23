// The frame of a host page around the editor: the part of the window where a chrome part can
// show, and the scale of a host transform.
//
// A host can put the editor in a container that clips its overflow (a drawer, a sheet, a card
// with `overflow: hidden`), or that scales it (`transform: scale()`). A popup that measures its
// room against the window then runs past the container's edge and is cut off, and a pointer
// position in window pixels points at the wrong cell of the grid, whose geometry is in layout
// pixels. `visibleArea` gives the room that stays visible, and `windowScale` converts window
// pixels to layout pixels.
//
// After pptx-editor/packages/core/src/editor/fixed-frame.ts at 3bf89a88 (EigenPal): `holdsFixed`,
// `fixedHolder`, `inTopLayer`, `clipsAny`, and the walk of `visibleView` are copied. Intended
// differences: the chrome popups here are `position: absolute` children of their trigger, not
// fixed panels, so the walk starts at the element's own containing block, not at a fixed holder,
// and the element's own overflow does not count; there is no probe and no `placeInFixedFrame`,
// because no chrome part here is `position: fixed` (the dialogs are modal, in the top layer); the
// area is returned as edges; and `windowScale` is new, with the rule of `formula-field.ts`.

/** A part of the window, in window pixels. @internal */
export interface VisibleArea {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

/** The scale from layout pixels to window pixels, on each axis. @internal */
export interface WindowScale {
  readonly x: number;
  readonly y: number;
}

// The values that make an element the containing block of its `position: fixed` descendants.
const NONE = /^(none|normal|auto|visible|flat)?$/;
const CONTAIN = /paint|layout|strict|content/;
const WILL_CHANGE = /transform|translate|scale|rotate|perspective|filter/;

/** Whether an element's style makes it the containing block of fixed descendants. */
function holdsFixed(style: CSSStyleDeclaration): boolean {
  // A browser without a property reads it as undefined: that sets nothing.
  const set = (value: string | undefined): boolean => !NONE.test(value ?? '');
  const webkit = style as unknown as { webkitBackdropFilter?: string };
  return (
    set(style.transform) ||
    set(style.translate) ||
    set(style.scale) ||
    set(style.rotate) ||
    set(style.perspective) ||
    set(style.filter) ||
    set(style.backdropFilter) ||
    set(webkit.webkitBackdropFilter) ||
    set(style.transformStyle) ||
    style.contentVisibility === 'auto' ||
    CONTAIN.test(style.contain ?? '') ||
    set(style.containerType) ||
    WILL_CHANGE.test(style.willChange ?? '')
  );
}

/** Whether an element is in the top layer, where the page around it neither moves nor clips it. */
function inTopLayer(el: Element): boolean {
  try {
    return el.matches(':modal, :popover-open');
  } catch {
    // A browser without these selectors has no top layer of this kind.
    return false;
  }
}

/** The nearest element at or above `el` that holds fixed descendants, or null. */
function fixedHolder(el: Element | null, win: Window): Element | null {
  for (; el; el = el.parentElement) {
    if (holdsFixed(win.getComputedStyle(el))) return el;
    // A modal dialog or an open popover is in the top layer: no ancestor holds or clips it.
    if (inTopLayer(el)) return null;
  }
  return null;
}

/** Whether an overflow value clips (`hidden`, `clip`, `auto`, `scroll`); unset does not. */
const clips = (overflow: string | undefined): boolean => !!overflow && overflow !== 'visible';
const clipsAny = (style: CSSStyleDeclaration): boolean =>
  clips(style.overflowX) ||
  clips(style.overflowY) ||
  clips(style.overflow) ||
  // Paint containment clips to the box as overflow does.
  /paint|strict|content/.test(style.contain ?? '') ||
  style.contentVisibility === 'auto';

/**
 * The next box up the containing-block chain of `el`: a fixed element escapes every ancestor up
 * to the next one that holds fixed elements, and an absolute one escapes static ones.
 */
function containingBox(el: Element, style: CSSStyleDeclaration, win: Window): Element | null {
  if (style.position === 'fixed') return fixedHolder(el.parentElement, win);
  let next = el.parentElement;
  if (style.position === 'absolute')
    while (
      next &&
      win.getComputedStyle(next).position === 'static' &&
      !holdsFixed(win.getComputedStyle(next))
    )
      next = next.parentElement;
  return next;
}

/**
 * The part of the window where `el` can show: the window, cut to the boxes up its
 * containing-block chain that clip their overflow. The element's own overflow does not count,
 * and an ancestor outside that chain does not clip it. With no window, the area is empty.
 * @internal
 */
export function visibleArea(el: Element): VisibleArea {
  const doc = el.ownerDocument;
  const win = doc.defaultView;
  if (!win) return { left: 0, top: 0, right: 0, bottom: 0 };
  let left = 0;
  let top = 0;
  let right = win.innerWidth;
  let bottom = win.innerHeight;
  let at: Element | null = inTopLayer(el) ? null : containingBox(el, win.getComputedStyle(el), win);
  while (at && at !== doc.documentElement) {
    const style = win.getComputedStyle(at);
    // The body's overflow moves to the window when the root's is visible: it clips nothing.
    const toWindow = at === doc.body && !clipsAny(win.getComputedStyle(doc.documentElement));
    if (clipsAny(style) && !toWindow) {
      const box = at.getBoundingClientRect();
      left = Math.max(left, box.left);
      top = Math.max(top, box.top);
      right = Math.min(right, box.right);
      bottom = Math.min(bottom, box.bottom);
    }
    if (inTopLayer(at)) break;
    at = containingBox(at, style, win);
  }
  return { left, top, right: Math.max(left, right), bottom: Math.max(top, bottom) };
}

/**
 * The scale from the layout pixels of `el` to window pixels: more than 1 inside a host that
 * scales the editor up, less than 1 inside one that scales it down. A distance in window pixels
 * divided by it is a distance in the element's layout pixels. 1 when the element has no layout
 * box (hidden, or a DOM with no layout), and 1 when the shown size is less than one pixel from
 * the layout size: the layout size is rounded to whole pixels, and a fractional width with no
 * transform is not a scale. A transform that rotates or skews the editor is not a scale, and
 * the ratio of its bounding box does not convert it. @internal
 */
export function windowScale(el: HTMLElement): WindowScale {
  const box = el.getBoundingClientRect();
  const ratio = (shown: number, laid: number): number => {
    if (Math.abs(shown - laid) < 1) return 1;
    const value = laid > 0 ? shown / laid : 0;
    return Number.isFinite(value) && value > 0 ? value : 1;
  };
  return { x: ratio(box.width, el.offsetWidth), y: ratio(box.height, el.offsetHeight) };
}
