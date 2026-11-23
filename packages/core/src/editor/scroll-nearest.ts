// Scroll an element into view in its nearest scrolling ancestor only, never the host page.
//
// Shared code: a copy of PPTX Editor (eigenpal/pptx-editor) `origin/main` 79a5d22d,
// `packages/core/src/editor/rail-orientation.ts:202-271` (`SCROLLS`, `scrollParent`, `px`,
// `nearest`, and `scrollIntoNearestScroller`), equal to the source. Difference: only these
// functions, without the rail orientation helpers of that file. Here `within` is the tab list
// of the sheet tabs, or the overflow panel of the toolbar.

const SCROLLS = /(auto|scroll)/;

type Axes = { readonly el: HTMLElement; readonly x: boolean; readonly y: boolean };

/**
 * The nearest ancestor of `el` that scrolls, and the axes on which it does, or null. The search
 * never goes past `within`: the editor's own pane, so a host container is never scrolled.
 */
function scrollParent(el: HTMLElement, within: Element | null): Axes | null {
  const view = el.ownerDocument.defaultView;
  for (let p = el.parentElement; p; p = p.parentElement) {
    const style = view?.getComputedStyle(p);
    if (!style) return null;
    const x = SCROLLS.test(style.overflowX) && p.scrollWidth > p.clientWidth;
    const y = SCROLLS.test(style.overflowY) && p.scrollHeight > p.clientHeight;
    if (x || y) return { el: p, x, y };
    if (p === within) return null;
  }
  return null;
}

/** A `scroll-padding-*` length in pixels; `auto` and anything else read as 0. */
const px = (value: string | undefined): number => parseFloat(value ?? '') || 0;

/** The least distance that brings `[start, start + size]` into `[pad0, length - pad1]`. */
function nearest(start: number, size: number, length: number, pad0: number, pad1: number): number {
  if (start < pad0) return start - pad0;
  const end = start + size - (length - pad1);
  return end > 0 ? Math.min(end, start - pad0) : 0;
}

/**
 * Scroll `target` into view in its nearest scrolling ancestor only, by the least distance on
 * each axis that scrolls, and inside its `scroll-padding`. Unlike `scrollIntoView`, it leaves
 * the outer page alone: a host that slides the editor in from off screen, or keeps it in a
 * clipped panel, does not see its layout move. `within` bounds the search (the rail or the
 * sorter): an ancestor outside it is never scrolled. @internal
 */
export function scrollIntoNearestScroller(
  target: HTMLElement,
  within: Element | null = null
): void {
  const found = scrollParent(target, within);
  if (!found) return;
  const { el: scroller } = found;
  const style = scroller.ownerDocument.defaultView?.getComputedStyle(scroller);
  const box = target.getBoundingClientRect();
  const view = scroller.getBoundingClientRect();
  const top = box.top - view.top - scroller.clientTop;
  const left = box.left - view.left - scroller.clientLeft;
  const dy = found.y
    ? nearest(
        top,
        box.height,
        scroller.clientHeight,
        px(style?.scrollPaddingTop),
        px(style?.scrollPaddingBottom)
      )
    : 0;
  const dx = found.x
    ? nearest(
        left,
        box.width,
        scroller.clientWidth,
        px(style?.scrollPaddingLeft),
        px(style?.scrollPaddingRight)
      )
    : 0;
  if (dx || dy) scroller.scrollBy?.({ left: dx, top: dy });
}
