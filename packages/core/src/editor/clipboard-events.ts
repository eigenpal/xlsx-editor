// The browser clipboard events of the grid (design D4 of the change
// `workbook-structure-editing`): Mod+C, Mod+X, and Mod+V reach the grid as `copy`, `cut`, and
// `paste` events, because the grid is focusable and its key table lets these keys pass
// (`grid-keys.ts`).
//
// Follows docx-editor `packages/core/src/editor/surface-input.ts:466-508` at f3116761
// (`createClipboardHandlers`): copy and cut set the clipboard data and prevent the default; a
// refused copy also prevents it, so the browser copies nothing else; paste prevents the default
// and passes the text to the engine. Differences: the second type is the private
// `CLIPBOARD_TYPE` token, not `text/html`, and a paste reads only whether `text/html` exists,
// never the HTML.
//
// An event acts only while the grid has the focus, when its target is the grid or a node
// inside the grid that takes no text. A browser can fire a clipboard event at the node of the page's
// selection, not at the focused element: Firefox fires it at the stage of a focused grid, and
// at the body when the page has no selection. So when the grid takes the focus, the page
// selection collapses into the grid, as a press on text moves it. A press on the grid cannot
// move it, because the grid cancels the press, so a press on the focused grid collapses it too.
// The focus comes from the grid's own root. In WebKit, the clipboard keys still do not reach a
// grid in a shadow root (`docs/site/content/known-approximations.mdx`, "Clipboard keys in a
// shadow root"). The cell editor keeps its own text clipboard: its events target its input, which
// has the focus, and pass through.
//
// The listeners sit on the grid. Difference from pptx-editor
// `packages/core/src/editor/editor-clipboard.ts` (`bindEditorClipboard`), which listens on the
// document for events outside its slide surface, inside the editor's own element: the grid has
// no such second region, and a document listener would also take a copy that the host page
// starts elsewhere while the grid keeps the focus. A listener on the grid runs before any
// ancestor, so it needs no `defaultPrevented` check, and an ancestor that stops the event
// cannot block it.

import { CLIPBOARD_TYPE, type CopyResult, type PasteData } from './clipboard.ts';
import type { EditRefusal, EditResult } from './cell-input.ts';

/** What the clipboard events call on the editor. */
export interface ClipboardEventHost {
  copy(): CopyResult;
  cut(): CopyResult;
  paste(data: PasteData): EditResult;
}

/** The count of `copy` and `cut` events that each grid handled, for the menu bar's commands. */
const gridHandled = new WeakMap<Element, number>();
/** The refusal of the last `copy` or `cut` event that each grid handled, or null when it wrote. */
const gridRefusals = new WeakMap<Element, EditRefusal | null>();

/**
 * How many `copy` and `cut` events the grid element has handled: written, or refused with the
 * refusal of the copy. A command that fires the event (the menu bar's Copy and Cut) reads it
 * before and after, so it knows that the grid took the event, whatever a host listener did to
 * its propagation, and runs no second copy that would show the refusal again. @internal
 */
export function gridClipboardHandled(grid: Element): number {
  return gridHandled.get(grid) ?? 0;
}

/**
 * The refusal of the last `copy` or `cut` event that the grid element handled, or null when it
 * wrote the clipboard, so the menu bar's command returns the same result as the copy. @internal
 */
export function gridClipboardRefusal(grid: Element): EditRefusal | null {
  return gridRefusals.get(grid) ?? null;
}

/** The `copy`, `cut`, and `paste` listeners of a grid element. */
export function clipboardHandlers(
  host: ClipboardEventHost,
  grid: Element
): {
  readonly onCopy: (event: ClipboardEvent) => void;
  readonly onCut: (event: ClipboardEvent) => void;
  readonly onPaste: (event: ClipboardEvent) => void;
  readonly onFocus: (event: FocusEvent) => void;
  readonly onPress: () => void;
} {
  // The grid's event: at the grid, or at a node inside it that takes no text, while focused.
  // The page selection stays in the grid after the focus leaves it, so an engine that fires at
  // the selection node can target the grid itself while another control has the focus.
  const focused = (): boolean =>
    (grid.getRootNode() as Document | ShadowRoot).activeElement === grid;
  const ours = (event: ClipboardEvent): boolean => focused() && !takesText(event.target, grid);
  const write = (event: ClipboardEvent, run: () => CopyResult) => {
    const data = event.clipboardData;
    if (!ours(event) || !data) return;
    event.preventDefault();
    const r = run();
    gridHandled.set(grid, gridClipboardHandled(grid) + 1);
    gridRefusals.set(grid, r.ok ? null : r.refusal);
    if (!r.ok) return;
    data.setData('text/plain', r.content.text);
    data.setData(CLIPBOARD_TYPE, r.content.token);
  };
  // The page selection moves into the grid, so the next clipboard event targets the grid. The
  // surface is not editable (`grid-surface.ts`). If an engine still moves the focus to an
  // editing host that holds the selection, the grid takes the focus back.
  const claimSelection = () => {
    const selection = grid.ownerDocument.getSelection();
    const at = selection?.anchorNode;
    if (!selection || (at && grid.contains(at))) return;
    selection.collapse(grid, 0);
    if (!focused()) (grid as HTMLElement).focus({ preventScroll: true });
  };
  return {
    onCopy: (event) => write(event, () => host.copy()),
    onCut: (event) => write(event, () => host.cut()),
    onPaste: (event) => {
      if (!ours(event)) return;
      event.preventDefault();
      const data = event.clipboardData;
      if (!data) return;
      host.paste({
        text: data.getData('text/plain'),
        token: data.getData(CLIPBOARD_TYPE),
        // Only the list of types is read: the HTML itself never is.
        html: Array.from(data.types).includes('text/html'),
      });
    },
    onFocus: (event) => {
      if (event.target === grid) claimSelection();
    },
    // A press on the grid that already has the focus fires no focus event.
    onPress: () => {
      if (focused()) claimSelection();
    },
  };
}

/**
 * Whether `target` is, or is inside, an element of the grid that takes text: a form field or
 * editable content. Its clipboard events belong to it. The walk stops at the grid, and it reads
 * the `contenteditable` attribute and not `isContentEditable`, because a grid inside editable
 * content of the host page inherits that state and still owns its events. The test reads
 * `nodeType` and the tag name, not `instanceof`, so it holds for a node of another window
 * (`host-realm.ts`).
 */
function takesText(target: EventTarget | null, grid: Element): boolean {
  const node = target as Node | null;
  let el = (node?.nodeType === 1 ? node : node?.parentNode) as Element | null | undefined;
  for (; el && el.nodeType === 1 && el !== grid; el = el.parentElement) {
    if (/^(INPUT|TEXTAREA|SELECT)$/i.test(el.tagName)) return true;
    const editable = el.getAttribute('contenteditable');
    if (editable !== null && editable.toLowerCase() !== 'false') return true;
  }
  return false;
}
