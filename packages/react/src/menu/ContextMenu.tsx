/** @jsxImportSource react */
// The grid context menu. After docx-editor/packages/react/src/editor/contextmenu/
// DocxEditorContextMenu.tsx and its `start.tsx` at f5cc27dfb (EigenPal): it opens at the pointer
// on a secondary press, and at the active cell on Shift+F10 or the Menu key; its rows are
// chrome slots with the registry's labels, icons, commands, and enabled states; a row child that
// names a default row replaces it in place, `hidden` removes it, other children follow the
// default rows, `Start` children go first, and `preset={false}` shows only the children.
// `disabled` gives the browser its own menu back. The rows are the menu bar's parts
// (`XlsxEditor.Menu.Item`, `.Row`, `.Cut`, and so on), so one row part serves both menus.
//
// Differences from DOCX Editor:
// - The engine makes the request (`EditorEvents.contextMenu`) and selects the pressed cells
//   first, so the React and Vue menus read one request.
// - The default rows follow what the press hit: cells, whole rows, whole columns, or the sheet
//   (`gridContextMenuEntries`).
// - `useContextMenuTarget` returns what the menu acts on (the sheet, the range, and its kind),
//   not an element: the grid is painted, so a press lands on no element of a cell.
// - The component takes no `t`: the menu bar takes none either. Use the root's `translations`.

import {
  Children,
  createContext,
  isValidElement,
  useCallback,
  useContext,
  useEffect,
  useId,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type ReactElement,
  type ReactNode,
} from 'react';
import {
  gridContextMenuEntries,
  guardToolbarMousedown,
  hostWindowOf,
  isApplePlatform,
  listenSheetMenuDismiss,
  placeContextMenu,
  systemClipboardReader,
  type EditorSnapshot,
  type GridContextMenuRequest,
  type GridContextMenuTargetKind,
  type GridRange,
  type ToolbarHost,
} from '@xlsx-editor.dev/core/editor';
import { useEditorState, useTranslation, useXlsxEditor } from '../context.ts';
import { ToolbarHostContext } from '../toolbar/slot.tsx';
import { NumberFormatDialogHost } from '../toolbar/NumberFormatDialog.tsx';
import {
  MENU_NAMED_ROWS,
  MenuBarContext,
  MenuPanel,
  XlsxEditorMenuGroup,
  XlsxEditorMenuItem,
  XlsxEditorMenuRow,
  XlsxEditorMenuSeparator,
  XlsxEditorMenuSubmenu,
  panelRows,
  type MenuBarState,
  type XlsxEditorMenuNamedRow,
} from './parts.tsx';

/**
 * What the open grid context menu acts on.
 *
 * @public
 */
export interface ContextMenuTarget {
  /** Cells, whole rows, whole columns, or the whole sheet. */
  readonly kind: GridContextMenuTargetKind;
  /** The sheet, 0-based. */
  readonly sheetIndex: number;
  /** The selected range, 0-based and inclusive. */
  readonly range: GridRange;
}

/**
 * Where the open grid context menu stands and how it opened.
 *
 * @public
 */
export interface ContextMenuAnchor {
  /** The window x of the point that the menu opened at, in CSS pixels. */
  readonly x: number;
  /** The window y of the point that the menu opened at, in CSS pixels. */
  readonly y: number;
  /** `'pointer'` for a secondary press, `'keyboard'` for Shift+F10 or the Menu key. */
  readonly trigger: 'pointer' | 'keyboard';
}

const RequestContext = createContext<GridContextMenuRequest | null>(null);

/**
 * What the open grid context menu acts on, or null while it is closed. A row of your own reads
 * it to act on the pressed cells.
 *
 * @public
 */
export function useContextMenuTarget(): ContextMenuTarget | null {
  const request = useContext(RequestContext);
  return useMemo(
    () =>
      request ? { kind: request.kind, sheetIndex: request.sheetIndex, range: request.range } : null,
    [request]
  );
}

/**
 * Where the open grid context menu stands and how it opened, or null while it is closed.
 *
 * @public
 */
export function useContextMenuAnchor(): ContextMenuAnchor | null {
  const request = useContext(RequestContext);
  return useMemo(
    () => (request ? { x: request.x, y: request.y, trigger: request.trigger } : null),
    [request]
  );
}

/** Props for `XlsxEditor.ContextMenu.Start`. @public */
export interface XlsxEditorContextMenuStartProps {
  readonly children?: ReactNode;
}

/**
 * Rows that go before the default rows, with a separator after them.
 *
 * @public
 */
export function XlsxEditorContextMenuStart({ children }: XlsxEditorContextMenuStartProps) {
  return <>{children}</>;
}

/** Props for `XlsxEditor.ContextMenu`. @public */
export interface XlsxEditorContextMenuProps {
  /**
   * `false` renders only the children as the rows. The default is `true`: the default rows for
   * what the press hit, with a row child that names a default row replacing it in place, and the
   * other children after the default rows.
   */
  readonly preset?: boolean;
  /** `true` removes the menu: a secondary press then shows the browser's own menu. */
  readonly disabled?: boolean;
  /** Called when the menu opens or closes. */
  readonly onOpenChange?: (open: boolean) => void;
  /** Appended after the `xlsx-context-menu` class. */
  readonly className?: string;
  readonly children?: ReactNode;
}

const isLoading = (s: EditorSnapshot): boolean => s.status === 'loading';

/** The children of `Start` parts, and the other children. */
function splitStart(children: ReactNode): { start: ReactNode[]; rest: ReactNode[] } {
  const start: ReactNode[] = [];
  const rest: ReactNode[] = [];
  for (const child of Children.toArray(children)) {
    if (isValidElement(child) && child.type === XlsxEditorContextMenuStart)
      start.push((child.props as XlsxEditorContextMenuStartProps).children);
    else rest.push(child);
  }
  return { start, rest };
}

/** The grid context menu. Place it inside `XlsxEditor.Viewport`. */
function XlsxEditorContextMenuRoot({
  preset = true,
  disabled = false,
  onOpenChange,
  className,
  children,
}: XlsxEditorContextMenuProps) {
  const { t } = useTranslation();
  const editor = useXlsxEditor();
  const panelId = useId();
  const [request, setRequest] = useState<GridContextMenuRequest | null>(null);
  const anchor = useRef<HTMLDivElement | null>(null);
  const [anchorNode, setAnchorNode] = useState<HTMLDivElement | null>(null);
  const anchorRef = useCallback((node: HTMLDivElement | null) => {
    anchor.current = node;
    setAnchorNode(node);
  }, []);
  const focusEdge = useRef<'first' | 'last' | null>(null);
  const openChange = useRef(onOpenChange);
  openChange.current = onOpenChange;

  useEffect(() => {
    if (!editor || disabled) return undefined;
    return editor.on('contextMenu', (next) => {
      focusEdge.current = 'first';
      setRequest(next);
    });
  }, [editor, disabled]);
  useEffect(() => {
    if (disabled) setRequest(null);
  }, [disabled]);
  const open = request !== null;
  const wasOpen = useRef(false);
  useEffect(() => {
    if (wasOpen.current === open) return;
    wasOpen.current = open;
    openChange.current?.(open);
  }, [open]);
  // An open of another workbook closes the menu: its rows showed the old workbook's state.
  const loading = useEditorState(isLoading);
  useEffect(() => {
    if (loading) setRequest(null);
  }, [loading]);

  const close = useCallback(
    (focus: 'trigger' | 'grid' | 'none' | 'origin') => {
      setRequest(null);
      if (focus !== 'none') {
        editor?.focus();
        return;
      }
      // A row that moves the keyboard (a dialog) keeps it; a close that leaves it nowhere
      // gives it back to the grid.
      queueMicrotask(() => {
        const doc = anchor.current?.ownerDocument;
        const active = doc?.activeElement;
        if (!active || active === doc.body || !active.isConnected) editor?.focus();
      });
    },
    [editor]
  );

  // Placed at its point, then closed on a press outside it, a wheel, a scroll of the page, a
  // resize, and a window blur, as the tab menu (`listenSheetMenuDismiss`).
  useLayoutEffect(() => {
    const node = anchor.current;
    const panel = node?.firstElementChild as HTMLElement | null | undefined;
    if (!request || !node || !panel) return undefined;
    const stopPlace = placeContextMenu(node, panel, request.x, request.y);
    const stopDismiss = listenSheetMenuDismiss(
      node,
      (target) => node.contains(target),
      () => setRequest(null)
    );
    return () => {
      stopPlace();
      stopDismiss();
    };
  }, [request]);

  const [rtl, setRtl] = useState(false);
  const [apple, setApple] = useState(false);
  useLayoutEffect(() => {
    if (!anchorNode) return;
    setRtl(hostWindowOf(anchorNode).getComputedStyle(anchorNode).direction === 'rtl');
    setApple(isApplePlatform(hostWindowOf(anchorNode).navigator));
  }, [anchorNode]);

  const host = useMemo<ToolbarHost>(() => {
    const read = systemClipboardReader(anchorNode);
    return read ? { readClipboard: read } : {};
  }, [anchorNode]);
  const state = useMemo<MenuBarState>(
    () => ({
      openMenu: null,
      open: () => {},
      focusRequest: { current: null },
      activeMenu: null,
      setActiveMenu: () => {},
      order: [],
      rtl,
      apple,
      closeAll: close,
      triggerOf: () => null,
    }),
    [rtl, apple, close]
  );

  const label = t('contextMenu.ariaLabel');
  const { start, rest } = splitStart(children);
  const rows = request
    ? panelRows(gridContextMenuEntries(request.kind), rest, preset, label)
    : null;
  return (
    <ToolbarHostContext.Provider value={host}>
      <NumberFormatDialogHost>
        <MenuBarContext.Provider value={state}>
          <RequestContext.Provider value={request}>
            <div
              ref={anchorRef}
              className={`xlsx-context-menu${className ? ` ${className}` : ''}`}
              onMouseDown={guardToolbarMousedown}
              onContextMenu={(e) => e.preventDefault()}
            >
              {request ? (
                <MenuPanel
                  id={panelId}
                  label={label}
                  placed
                  focusEdge={focusEdge}
                  onClose={() => close('grid')}
                >
                  {start.length > 0 ? (
                    <>
                      {start}
                      <XlsxEditorMenuSeparator />
                    </>
                  ) : null}
                  {rows}
                </MenuPanel>
              ) : null}
            </div>
          </RequestContext.Provider>
        </MenuBarContext.Provider>
      </NumberFormatDialogHost>
    </ToolbarHostContext.Provider>
  );
}

/** The grid context menu with its parts as statics. @public */
export interface XlsxEditorContextMenuNamespace {
  /** The grid context menu. */
  (props: XlsxEditorContextMenuProps): ReactElement;
  /** Rows that go before the default rows. */
  readonly Start: typeof XlsxEditorContextMenuStart;
  /** One chrome slot as a live row. */
  readonly Item: typeof XlsxEditorMenuItem;
  /** A row of your own, for an action that is not a chrome slot. */
  readonly Row: typeof XlsxEditorMenuRow;
  /** A named section of rows. */
  readonly Group: typeof XlsxEditorMenuGroup;
  readonly Separator: typeof XlsxEditorMenuSeparator;
  readonly Submenu: typeof XlsxEditorMenuSubmenu;
  readonly Cut: XlsxEditorMenuNamedRow;
  readonly Copy: XlsxEditorMenuNamedRow;
  readonly Paste: XlsxEditorMenuNamedRow;
  readonly InsertRows: XlsxEditorMenuNamedRow;
  readonly InsertColumns: XlsxEditorMenuNamedRow;
  readonly DeleteRows: XlsxEditorMenuNamedRow;
  readonly DeleteColumns: XlsxEditorMenuNamedRow;
  readonly ClearContents: XlsxEditorMenuNamedRow;
  readonly ClearFormatting: XlsxEditorMenuNamedRow;
  readonly ClearAll: XlsxEditorMenuNamedRow;
  readonly MergeCells: XlsxEditorMenuNamedRow;
  readonly MergeAndCenter: XlsxEditorMenuNamedRow;
  readonly UnmergeCells: XlsxEditorMenuNamedRow;
  readonly NumberFormat: XlsxEditorMenuNamedRow;
}

/**
 * The grid context menu: it opens on a secondary press on the grid, and on Shift+F10 or the
 * Menu key, at the active cell. Its default rows are chrome slots for what the press hit, so a
 * row and its toolbar control share one label, one command, and one enabled state. Place it
 * inside `XlsxEditor.Viewport`. The packaged `XlsxEditor` mounts it; its `contextMenu` prop
 * takes these props, or `false`.
 *
 * @example
 * ```tsx
 * <XlsxEditor.ContextMenu>
 *   <XlsxEditor.ContextMenu.Paste hidden />
 *   <XlsxEditor.ContextMenu.Row label="Add comment" onSelect={addComment} />
 * </XlsxEditor.ContextMenu>
 * ```
 *
 * @public
 */
export const XlsxEditorContextMenu: XlsxEditorContextMenuNamespace = Object.assign(
  XlsxEditorContextMenuRoot,
  {
    Start: XlsxEditorContextMenuStart,
    Item: XlsxEditorMenuItem,
    Row: XlsxEditorMenuRow,
    Group: XlsxEditorMenuGroup,
    Separator: XlsxEditorMenuSeparator,
    Submenu: XlsxEditorMenuSubmenu,
    Cut: MENU_NAMED_ROWS.Cut,
    Copy: MENU_NAMED_ROWS.Copy,
    Paste: MENU_NAMED_ROWS.Paste,
    InsertRows: MENU_NAMED_ROWS.InsertRows,
    InsertColumns: MENU_NAMED_ROWS.InsertColumns,
    DeleteRows: MENU_NAMED_ROWS.DeleteRows,
    DeleteColumns: MENU_NAMED_ROWS.DeleteColumns,
    ClearContents: MENU_NAMED_ROWS.ClearContents,
    ClearFormatting: MENU_NAMED_ROWS.ClearFormatting,
    ClearAll: MENU_NAMED_ROWS.ClearAll,
    MergeCells: MENU_NAMED_ROWS.MergeCells,
    MergeAndCenter: MENU_NAMED_ROWS.MergeAndCenter,
    UnmergeCells: MENU_NAMED_ROWS.UnmergeCells,
    NumberFormat: MENU_NAMED_ROWS.NumberFormat,
  }
);
