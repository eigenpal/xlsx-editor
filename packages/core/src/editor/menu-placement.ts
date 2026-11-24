// Where an open menu panel of the menu bar stands, so that it stays inside the viewport. After
// `toolbarPanelPlacement` of docx-editor/packages/core/src/editor/toolbar-fit.ts and its hook
// docx-editor/packages/react/src/editor/menu/menu-panel-placement.ts at ad686574 (EigenPal).
// Differences from DOCX Editor:
// - It places every panel of the bar, not only the "More" panel: a bar menu's panel lines up with
//   its trigger's start edge first (the "More" panel with its end edge), and a submenu opens at
//   the end of its row first, and toward the start of the bar when the end has no room.
// - It reads the direction: in a right-to-left bar, start and end swap sides.
// - A submenu also moves up when its bottom would pass the bottom of the viewport. A bar menu's
//   panel keeps its top under the trigger, and its height is limited to the viewport space
//   below the trigger: a tall menu on a short viewport scrolls (`data-scrolls`). A submenu of a
//   panel that scrolls is placed against the viewport (`position: fixed`), so the scroll box of
//   its panel does not clip it, and it follows its row when that panel scrolls.
// - The DOM step (`placeMenuPanel`) is core's, so the React and Vue panels place the same way.

import { hostWindowOf } from './host-realm.ts';

/** The margin that a panel keeps from each viewport edge, in px. @internal */
export const MENU_PANEL_MARGIN = 8;

/** The gap between a submenu panel and its row, in px (the stylesheet's `100% + 4px`). */
const SUBMENU_GAP = 4;
/** How far a submenu panel stands above its row, in px (the stylesheet's `top: -5px`). */
const SUBMENU_RAISE = 5;
/** The gap between a bar menu's panel and its trigger, in px (the stylesheet's `100% + 2px`). */
const PANEL_GAP = 2;

/** A box in viewport px. @internal */
export interface MenuBox {
  readonly left: number;
  readonly right: number;
  readonly top: number;
  readonly bottom: number;
}

/** Where a panel opens, in viewport px. @internal */
export interface MenuPanelPlacement {
  readonly left: number;
  readonly top: number;
  /** The side of the anchor where the panel opens, or `clamped` when no side has room. */
  readonly side: 'start' | 'end' | 'clamped';
  /**
   * The tallest the panel may be: the viewport space below the trigger for a bar menu's panel,
   * and the viewport less both margins for a submenu.
   */
  readonly maxHeight: number;
}

/**
 * Keep a panel inside the viewport. `anchor` is the trigger of a bar menu's panel, or the row of
 * a submenu panel. A bar menu's panel lines up with the trigger's start edge first, or its end
 * edge with `prefer: 'end'` (the "More" menu at the end of the bar); a submenu opens at the end
 * of its row first. When the preferred side has no room the panel takes the other side, and when
 * neither has room it is clamped to the margins. A submenu also moves up as far as it must to
 * keep its bottom inside the viewport. `maxHeight` limits the panel to the viewport. @internal
 */
export function menuPanelPlacement(input: {
  readonly submenu: boolean;
  readonly rtl: boolean;
  readonly prefer?: 'start' | 'end';
  readonly anchor: MenuBox;
  readonly width: number;
  readonly height: number;
  readonly viewportWidth: number;
  readonly viewportHeight: number;
  readonly margin?: number;
}): MenuPanelPlacement {
  const margin = input.margin ?? MENU_PANEL_MARGIN;
  const { anchor, rtl } = input;
  const width = Math.min(input.width, Math.max(0, input.viewportWidth - margin * 2));
  const min = margin;
  const max = input.viewportWidth - margin - width;
  // The left edge of the panel on each side of the anchor.
  const starts = input.submenu
    ? rtl
      ? anchor.right + SUBMENU_GAP
      : anchor.left - SUBMENU_GAP - width
    : rtl
      ? anchor.right - width
      : anchor.left;
  const ends = input.submenu
    ? rtl
      ? anchor.left - SUBMENU_GAP - width
      : anchor.right + SUBMENU_GAP
    : rtl
      ? anchor.left
      : anchor.right - width;
  const order: ['start' | 'end', number][] =
    input.submenu || input.prefer === 'end'
      ? [
          ['end', ends],
          ['start', starts],
        ]
      : [
          ['start', starts],
          ['end', ends],
        ];
  const fits = order.find(([, left]) => left >= min && left <= max);
  const [side, left] = fits ?? ['clamped' as const, Math.max(min, Math.min(order[0]![1], max))];
  const top = input.submenu
    ? Math.max(
        margin,
        Math.min(anchor.top - SUBMENU_RAISE, input.viewportHeight - margin - input.height)
      )
    : anchor.bottom + PANEL_GAP;
  const bottom = input.viewportHeight - margin;
  const maxHeight = Math.max(0, input.submenu ? bottom - margin : bottom - top);
  return { left, top, side, maxHeight };
}

/**
 * Place an open panel of the menu bar inside the viewport of its own window: it measures the
 * panel and its anchor (the parent element: the bar menu or the submenu row), and sets its
 * `left`, `top`, and `max-height` against that parent. A panel taller than its limit gets
 * `data-scrolls`, which the stylesheet scrolls. Returns a function that stops the placement on a
 * resize of the window. @internal
 */
export function placeMenuPanel(
  panel: HTMLElement,
  options: { readonly submenu: boolean; readonly prefer?: 'start' | 'end' }
): () => void {
  const view = hostWindowOf(panel);
  const place = (): void => {
    const parent = panel.parentElement;
    if (!parent || panel.hidden) return;
    const box = parent.getBoundingClientRect();
    const anchor = options.submenu
      ? (parent.querySelector<HTMLElement>(':scope > [aria-haspopup="menu"]') ?? parent)
      : parent;
    const at = menuPanelPlacement({
      submenu: options.submenu,
      rtl: view.getComputedStyle(panel).direction === 'rtl',
      ...(options.prefer ? { prefer: options.prefer } : {}),
      anchor: anchor.getBoundingClientRect(),
      width: panel.offsetWidth,
      height: panel.offsetHeight,
      viewportWidth: view.innerWidth,
      viewportHeight: view.innerHeight,
    });
    panel.style.maxHeight = `${at.maxHeight}px`;
    panel.style.right = 'auto';
    // In a panel that scrolls, a submenu stands against the viewport, out of the scroll box.
    const fixed = options.submenu && !!parent.closest('.xlsx-menu-panel[data-scrolls]');
    panel.style.position = fixed ? 'fixed' : '';
    const origin = fixed ? { left: 0, top: 0 } : box;
    panel.style.left = `${at.left - origin.left}px`;
    panel.style.top = `${at.top - origin.top}px`;
    if (fixed) {
      // A transformed host moves the box of `fixed`: the difference corrects a translation.
      const moved = panel.getBoundingClientRect();
      panel.style.left = `${2 * at.left - moved.left}px`;
      panel.style.top = `${2 * at.top - moved.top}px`;
    }
    panel.toggleAttribute('data-scrolls', panel.scrollHeight > panel.clientHeight + 1);
  };
  place();
  view.addEventListener('resize', place);
  // A submenu follows its row when the panel that holds the row scrolls.
  const outer = options.submenu
    ? panel.parentElement?.closest<HTMLElement>('.xlsx-menu-panel')
    : null;
  outer?.addEventListener('scroll', place, { passive: true });
  return () => {
    view.removeEventListener('resize', place);
    outer?.removeEventListener('scroll', place);
    for (const name of ['left', 'right', 'top', 'max-height', 'position'])
      panel.style.removeProperty(name);
    panel.removeAttribute('data-scrolls');
  };
}
