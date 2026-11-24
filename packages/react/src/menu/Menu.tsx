/** @jsxImportSource react */
// The menu bar: File, Edit, View, Insert, and Format, derived from the chrome registry
// (`CHROME_MENUS`). After docx-editor/packages/react/src/editor/menu/DocxEditorMenu.tsx at
// ad686574 (EigenPal): with no children the bar shows every menu of the registry in order; a
// child that is a menu part replaces its menu in place, `hidden` removes it, other children show
// after the menus, and `preset={false}` shows only the children. The bar owns which menu is open
// (one at a time), the host actions that are not engine commands (Open, and Save to the host's
// `onSave`), and the shortcut that it shows on Save (Control or Command and S). When the menus do
// not fit, the bar moves them, from the end, into one "More" menu, where each is a submenu. The
// bar is one tab stop, as the WAI-ARIA menubar pattern asks. Its order and its menus come from
// core's `arrangeMenuBar` (`menu-arrangement.ts`), as in PPTX Editor. A press keeps the keyboard
// where it was; while a menu that a press opened is open, a navigation key from outside the bar
// moves the keyboard into it (core's `listenForPointerPopupKeys`), and Escape gives it back.
//
// Differences from DOCX Editor:
// - The menus of a workbook editor. There is no print, export, page setup, or help menu.
// - `onSave` receives the saved bytes (`onSave(bytes, saved)`), as the toolbar's does, where
//   DOCX Editor's `onSave()` takes none: the bar saves through the engine (`file.save`) and hands
//   over the result, so a host never calls `save()` a second time. With no `onSave`, Save is
//   disabled: the bar never downloads a file by itself.
// - The packaged Open loads the picked file, as DOCX Editor's does, through the engine's
//   picked-file load (`openPickedWorkbook`): one read of the file, and a later load wins. A file
//   that the browser cannot read or the editor cannot open shows the refusal `open-failed`, and
//   the open workbook, its edits, and its open cell edit stay. `onOpen` replaces it, and
//   `onOpenFile` tells the host the picked file.
// - The bar takes `opening`, as the toolbar does, and no `t`: the toolbar takes none either.
// - Paste reads the system clipboard text where the host window allows a page to read it.

import {
  cloneElement,
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type ReactElement,
  type ReactNode,
} from 'react';
import {
  MENU_FIXED_ATTRIBUTE,
  arrangeMenuBar,
  enterMenuPanel,
  focusMenuPanelEdge,
  guardToolbarMousedown,
  hostWindowOf,
  isApplePlatform,
  isElementNode,
  listenForOutsidePress,
  listenForPointerPopupKeys,
  listenForPopupEscape,
  listenForPopupFocusLeave,
  listenSaveKey,
  measureMenuBar,
  menuPanelOf,
  menuRowInBarProblem,
  menuStep,
  observeMenuBar,
  openPickedWorkbook,
  REACT_MENU_PART_NAMES,
  runToolbarCommand,
  systemClipboardReader,
  type EditorSnapshot,
  type SavedWorkbook,
  type ToolbarHost,
} from '@xlsx-editor.dev/core/editor';
import { useEditorState, useTranslation, useXlsxEditor } from '../context.ts';
import { ToolbarHostContext, translateKey } from '../toolbar/slot.tsx';
import { NumberFormatDialogHost } from '../toolbar/NumberFormatDialog.tsx';
import { warnToolbarProblems } from '../toolbar/host-parts.tsx';
import {
  MENU_NAMED_ROWS,
  MORE_MENU_ID,
  slotRowOf,
  MenuBarContext,
  MenuEdit,
  MenuFile,
  MenuFormat,
  MenuInsert,
  MenuPlacementContext,
  MenuView,
  MoreMenu,
  XlsxEditorMenuEntry,
  XlsxEditorMenuGroup,
  XlsxEditorMenuItem,
  XlsxEditorMenuMenu,
  XlsxEditorMenuRow,
  XlsxEditorMenuSeparator,
  XlsxEditorMenuSubmenu,
  flatMenuChildren,
  menuSpecOf,
  type MenuBarState,
  type MenuPlacement,
  type XlsxEditorMenuNamedRow,
  type XlsxEditorMenuPart,
  type XlsxEditorMenuPartProps,
} from './parts.tsx';

/** Props for `XlsxEditor.Menu`. @public */
export interface XlsxEditorMenuProps {
  /**
   * Receives the workbook bytes when File > Save is chosen, or when the user presses Control or
   * Command and S in the editor, as the toolbar's `onSave` does. Without it, File > Save is
   * disabled with the engine's reason.
   */
  readonly onSave?: (bytes: Uint8Array<ArrayBuffer>, saved: SavedWorkbook) => void;
  /** Called once when File > Save starts a delivery to `onSave`, as the toolbar's `onSaveStart`. */
  readonly onSaveStart?: (saving: Promise<SavedWorkbook | null>) => void;
  /**
   * Replaces File > Open. The default opens a file picker and passes the bytes of the picked
   * workbook to the editor's `load`: a file that the user picks, never a fetch. The default does
   * not ask first: the picked file replaces the open workbook, also with unsaved edits. Pass
   * `onOpen` to ask the user first. A picked file that the browser cannot read or the editor
   * cannot open replaces nothing: the open workbook, its edits, and an open cell edit stay, and
   * the refusal `open-failed` shows.
   */
  readonly onOpen?: () => void;
  /**
   * Called when the packaged File > Open reads a picked file, before its bytes load, so you can
   * show its name. Not called when `onOpen` replaces the packaged picker.
   */
  readonly onOpenFile?: (file: File) => void;
  /**
   * Whether your page opens another workbook now, as the toolbar's `opening`: File > Open and
   * File > Save are disabled while it is true.
   */
  readonly opening?: boolean;
  /**
   * `false` lets the bar wrap onto more lines instead of moving the menus that do not fit into
   * the "More" menu. The default is `true`.
   */
  readonly overflow?: boolean;
  /**
   * `false` renders only the children, with no registry menus. The default is `true`: a menu
   * child changes its menu in place, and other children show after the menus.
   */
  readonly preset?: boolean;
  /** Appended after the `xlsx-menubar` class. */
  readonly className?: string;
  readonly children?: ReactNode;
}

const ACCEPT =
  '.xlsx,.xltx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,' +
  'application/vnd.openxmlformats-officedocument.spreadsheetml.template';

const isLoading = (s: EditorSnapshot): boolean => s.status === 'loading';

/** The menu bar. Place it inside `XlsxEditor.Viewport`, which carries the styling scope. */
function XlsxEditorMenuRoot({
  onSave,
  onSaveStart,
  onOpen,
  onOpenFile,
  opening,
  overflow = true,
  preset = true,
  className,
  children,
}: XlsxEditorMenuProps) {
  const { t } = useTranslation();
  const editor = useXlsxEditor();
  const bar = useRef<HTMLDivElement | null>(null);
  // The mounted bar, so the host reads the clipboard of the bar's own window, not the window
  // that the module loaded in (`host-realm.ts`).
  const [barNode, setBarNode] = useState<HTMLDivElement | null>(null);
  const barRef = useCallback((node: HTMLDivElement | null) => {
    bar.current = node;
    setBarNode(node);
  }, []);
  // The row that holds the menu bar and the host's loose children: the overflow measures it.
  const row = useRef<HTMLDivElement>(null);
  const fileInput = useRef<HTMLInputElement>(null);
  const [openMenu, setOpenMenu] = useState<string | null>(null);
  const [activeMenu, setActiveMenu] = useState<string | null>(null);
  const [rtl, setRtl] = useState(false);
  const [apple, setApple] = useState(() =>
    isApplePlatform(typeof navigator === 'undefined' ? undefined : navigator)
  );
  const focusRequest = useRef<{ id: string; edge: 'first' | 'last' } | null>(null);
  const openingNow = useRef(opening ?? false);
  openingNow.current = opening ?? false;

  const packagedOpen = useCallback(() => fileInput.current?.click(), []);
  const host = useMemo<ToolbarHost>(() => {
    const read = systemClipboardReader(barNode);
    const open = onOpen ?? (editor ? packagedOpen : undefined);
    return {
      ...(onSave ? { save: onSave } : {}),
      ...(onSaveStart ? { saveStart: onSaveStart } : {}),
      ...(opening === undefined ? {} : { opening: () => openingNow.current }),
      ...(open ? { open } : {}),
      ...(read ? { readClipboard: read } : {}),
    };
  }, [onSave, onSaveStart, opening, onOpen, editor, packagedOpen, barNode]);

  // The menus of the bar, from the children and the registry.
  const flat = flatMenuChildren(children);
  const specs = new Map<string, ReturnType<typeof menuSpecOf> & object>();
  const appended: ReactNode[] = [];
  const problems: string[] = [];
  for (const child of flat) {
    const spec = menuSpecOf(child);
    const row = spec ? null : slotRowOf(child);
    if (spec) {
      specs.delete(spec.id);
      specs.set(spec.id, spec);
    } else if (row) problems.push(menuRowInBarProblem(row.slotId, REACT_MENU_PART_NAMES));
    else appended.push(child);
  }
  const arranged = arrangeMenuBar(
    new Map(
      [...specs.values()].map((s) => [
        s.id,
        {
          id: s.id,
          label: s.props.label,
          labelKey: s.props.labelKey,
          hidden: s.props.hidden,
          after: s.props.after,
        },
      ])
    ),
    preset,
    (key) => translateKey(t, key),
    (message) => problems.push(message)
  );
  const problemKey = problems.join('\n');
  useEffect(() => {
    if (problemKey) warnToolbarProblems(problemKey.split('\n'));
  }, [problemKey]);
  const ids = arranged.map((m) => m.id);
  const idKey = ids.join('\n');
  const elements: ReactElement[] = arranged.map((m) => {
    const element = specs.get(m.id)?.element;
    return element ? (
      cloneElement(element, { key: m.id })
    ) : (
      <XlsxEditorMenuMenu key={m.id} id={m.id} />
    );
  });

  // The overflow: the menus that fit stay in the bar, the rest go into the "More" menu.
  const measuring = overflow;
  const widths = useRef(new Map<string, number>());
  const [shown, setShown] = useState(Number.MAX_SAFE_INTEGER);
  useLayoutEffect(() => {
    widths.current.clear();
  }, [t, idKey]);
  useLayoutEffect(() => {
    const el = row.current;
    if (!el) return undefined;
    setRtl(el.ownerDocument.defaultView?.getComputedStyle(el).direction === 'rtl');
    setApple(isApplePlatform(hostWindowOf(el).navigator));
    if (!measuring) {
      setShown(Number.MAX_SAFE_INTEGER);
      return undefined;
    }
    const measure = () => {
      const next = measureMenuBar(el, idKey ? idKey.split('\n') : [], widths.current);
      setShown((current) => (current === next ? current : next));
    };
    measure();
    return observeMenuBar(el, measure);
  }, [measuring, idKey, shown, t]);
  const count = Math.min(shown, ids.length);
  const inBar = ids.slice(0, count);
  const inMore = ids.slice(count);
  const order = inMore.length > 0 ? [...inBar, MORE_MENU_ID] : inBar;
  const orderKey = order.join('\n');

  // An open of another workbook closes the menus: their rows showed the old workbook's state.
  const loading = useEditorState(isLoading);
  useEffect(() => {
    if (loading) setOpenMenu(null);
  }, [loading]);

  const triggerOf = useCallback(
    (id: string): HTMLElement | null =>
      bar.current?.querySelector<HTMLElement>(
        id === MORE_MENU_ID
          ? ':scope > [data-menu-more] > [role="menuitem"]'
          : `:scope > [data-menu="${id.replace(/["\\]/g, '\\$&')}"] > [role="menuitem"]`
      ) ?? null,
    []
  );
  // Whether the bar held the keyboard when the open menu opened: a keyboard open. A pointer
  // press keeps the keyboard in the grid, so a pointer open leaves this false.
  const openedInBar = useRef(false);
  const openNow = useRef<string | null>(null);
  openNow.current = openMenu;
  // The control that held the keyboard when a key moved it from outside the bar into a menu
  // opened by a pointer press: Escape gives the keyboard back to it, such as an open cell edit.
  const enteredFrom = useRef<HTMLElement | null>(null);
  const closeAll = useCallback(
    (focus: 'trigger' | 'grid' | 'none' | 'origin') => {
      const was = openMenu;
      const from = enteredFrom.current;
      enteredFrom.current = null;
      setOpenMenu(null);
      if (focus === 'trigger' && from?.isConnected) from.focus();
      else if ((focus === 'trigger' || (focus === 'origin' && openedInBar.current)) && was)
        triggerOf(was)?.focus();
      else if (focus === 'grid') editor?.focus();
    },
    [openMenu, triggerOf, editor]
  );
  const open = useCallback(
    (id: string | null, edge?: 'first' | 'last') => {
      const el = bar.current;
      // The menu shows already, so no render follows: its row takes the focus now.
      if (id && edge && id === openNow.current)
        return focusMenuPanelEdge(menuPanelOf(triggerOf(id)), edge);
      if (id && openNow.current === null)
        openedInBar.current = !!el && el.contains(el.ownerDocument.activeElement);
      focusRequest.current = id && edge ? { id, edge } : null;
      setOpenMenu(id);
      if (id) setActiveMenu(id);
    },
    [triggerOf]
  );

  // The bar's order and direction, for the keys that a menu opened by a pointer press takes.
  const layout = useRef({ order: [] as readonly string[], rtl: false });
  // While a menu is open: a press outside the bar, an Escape from outside it (the grid keeps the
  // keyboard after a pointer open), and focus that leaves it close the menu. A navigation key
  // from outside the bar moves the keyboard into the open menu (`listenForPointerPopupKeys`).
  useEffect(() => {
    const el = bar.current;
    if (openMenu === null || !el) {
      enteredFrom.current = null;
      return undefined;
    }
    const contains = (node: Node) => el.contains(node);
    const close = () => setOpenMenu(null);
    const enter = (entry: Parameters<typeof enterMenuPanel>[0]) => {
      const panel = menuPanelOf(triggerOf(openMenu));
      if (!panel) return null;
      const active = el.ownerDocument.activeElement;
      if (!enteredFrom.current && isElementNode(active))
        enteredFrom.current = active as HTMLElement;
      const { order, rtl } = layout.current;
      return enterMenuPanel(entry, panel, {
        rtl,
        step: (by) => {
          const next = order[menuStep(order.length, order.indexOf(openMenu), by)];
          if (next) open(next, 'first');
        },
      });
    };
    const stops = [
      listenForPointerPopupKeys({ popup: el, contains, enter }),
      listenForOutsidePress({ popup: el, contains, close }),
      listenForPopupEscape({
        popup: el,
        contains,
        // A key inside the bar belongs to its panels, which close one level at a time.
        skip: (event) => {
          const target = event.composedPath()[0];
          return isElementNode(target) && el.contains(target);
        },
        close,
      }),
      listenForPopupFocusLeave({ popup: el, contains, close }),
    ];
    return () => stops.forEach((stop) => stop());
  }, [openMenu, triggerOf, open]);

  // Control or Command and S runs File > Save, as the toolbar binds it: the shortcut text on
  // the row names a key that works. The first listener of the editor takes the key.
  const saves = !!onSave;
  useEffect(() => {
    const el = bar.current;
    if (!el || !editor || !saves) return undefined;
    return listenSaveKey(el, () => runToolbarCommand(editor, 'file.save', undefined, host));
  }, [editor, saves, host]);

  layout.current = { order: orderKey ? orderKey.split('\n') : [], rtl };
  const state = useMemo<MenuBarState>(() => {
    const list = orderKey ? orderKey.split('\n') : [];
    return {
      openMenu,
      open,
      focusRequest,
      activeMenu: activeMenu !== null && list.includes(activeMenu) ? activeMenu : (list[0] ?? null),
      setActiveMenu,
      order: list,
      rtl,
      apple,
      closeAll,
      triggerOf,
    };
  }, [openMenu, open, activeMenu, orderKey, rtl, apple, closeAll, triggerOf]);
  const barKey = inBar.join('\n');
  const moreKey = inMore.join('\n');
  const barPlacement = useMemo<MenuPlacement>(
    () => ({ kind: 'bar', ids: new Set(barKey ? barKey.split('\n') : []) }),
    [barKey]
  );
  const morePlacement = useMemo<MenuPlacement>(
    () => ({ kind: 'more', ids: new Set(moreKey ? moreKey.split('\n') : []) }),
    [moreKey]
  );

  return (
    <ToolbarHostContext.Provider value={host}>
      <NumberFormatDialogHost>
        <MenuBarContext.Provider value={state}>
          {/* The host's loose children sit beside the menu bar element, not in it: they are no
              rows of the menubar and keep their own tab stops. */}
          <div
            ref={row}
            className={`xlsx-menubar${className ? ` ${className}` : ''}`}
            {...(measuring ? { 'data-overflow': '' } : {})}
            onMouseDown={guardToolbarMousedown}
          >
            <div
              ref={barRef}
              role="menubar"
              aria-label={t('menu.label')}
              className="xlsx-menubar-menus"
            >
              <MenuPlacementContext.Provider value={barPlacement}>
                {elements}
              </MenuPlacementContext.Provider>
              {inMore.length > 0 ? (
                <MoreMenu>
                  <MenuPlacementContext.Provider value={morePlacement}>
                    {elements}
                  </MenuPlacementContext.Provider>
                </MoreMenu>
              ) : null}
            </div>
            {appended.length > 0 ? (
              <div className="xlsx-menubar-fixed" {...{ [MENU_FIXED_ATTRIBUTE]: '' }}>
                {appended}
              </div>
            ) : null}
          </div>
          {/* Opening a workbook is a file that the user picks, never a fetch. A host `onOpen`
              replaces the picker, so the bar adds no file input of its own then. */}
          {onOpen ? null : (
            <input
              ref={fileInput}
              type="file"
              accept={ACCEPT}
              hidden
              tabIndex={-1}
              aria-hidden="true"
              onChange={(event) => {
                const file = event.target.files?.[0];
                // Cleared, so a pick of the same file again fires a change.
                event.target.value = '';
                if (!file || !editor) return;
                onOpenFile?.(file);
                // The load starts now, so a later load wins over this pick during its read.
                void openPickedWorkbook(editor, file);
              }}
            />
          )}
        </MenuBarContext.Provider>
      </NumberFormatDialogHost>
    </ToolbarHostContext.Provider>
  );
}

/** The menu bar with its parts as statics. @public */
export interface XlsxEditorMenuNamespace {
  /** The menu bar. */
  (props: XlsxEditorMenuProps): ReactElement;
  /** A menu of the bar: a built-in id changes that menu, another id adds a host menu. */
  readonly Menu: typeof XlsxEditorMenuMenu;
  readonly File: XlsxEditorMenuPart;
  readonly Edit: XlsxEditorMenuPart;
  readonly View: XlsxEditorMenuPart;
  readonly Insert: XlsxEditorMenuPart;
  readonly Format: XlsxEditorMenuPart;
  /** One chrome slot as a live row. */
  readonly Item: typeof XlsxEditorMenuItem;
  /** A row of your own, for an action that is not a chrome slot. */
  readonly Row: typeof XlsxEditorMenuRow;
  /** A named section of rows. */
  readonly Group: typeof XlsxEditorMenuGroup;
  readonly Separator: typeof XlsxEditorMenuSeparator;
  readonly Submenu: typeof XlsxEditorMenuSubmenu;
  /** One registry entry as its row. */
  readonly Entry: typeof XlsxEditorMenuEntry;
  readonly Open: XlsxEditorMenuNamedRow;
  readonly Save: XlsxEditorMenuNamedRow;
  readonly Undo: XlsxEditorMenuNamedRow;
  readonly Redo: XlsxEditorMenuNamedRow;
  readonly Cut: XlsxEditorMenuNamedRow;
  readonly Copy: XlsxEditorMenuNamedRow;
  readonly Paste: XlsxEditorMenuNamedRow;
  readonly SelectAll: XlsxEditorMenuNamedRow;
  readonly ClearContents: XlsxEditorMenuNamedRow;
  readonly ClearFormatting: XlsxEditorMenuNamedRow;
  readonly ClearAll: XlsxEditorMenuNamedRow;
  readonly DeleteRows: XlsxEditorMenuNamedRow;
  readonly DeleteColumns: XlsxEditorMenuNamedRow;
  readonly ZoomIn: XlsxEditorMenuNamedRow;
  readonly ZoomOut: XlsxEditorMenuNamedRow;
  readonly ZoomReset: XlsxEditorMenuNamedRow;
  readonly InsertRows: XlsxEditorMenuNamedRow;
  readonly InsertColumns: XlsxEditorMenuNamedRow;
  readonly InsertSheet: XlsxEditorMenuNamedRow;
  readonly Bold: XlsxEditorMenuNamedRow;
  readonly Italic: XlsxEditorMenuNamedRow;
  readonly Underline: XlsxEditorMenuNamedRow;
  readonly AlignLeft: XlsxEditorMenuNamedRow;
  readonly AlignCenter: XlsxEditorMenuNamedRow;
  readonly AlignRight: XlsxEditorMenuNamedRow;
  readonly AlignTop: XlsxEditorMenuNamedRow;
  readonly AlignMiddle: XlsxEditorMenuNamedRow;
  readonly AlignBottom: XlsxEditorMenuNamedRow;
  readonly WrapText: XlsxEditorMenuNamedRow;
  readonly MergeCells: XlsxEditorMenuNamedRow;
  readonly MergeAndCenter: XlsxEditorMenuNamedRow;
  readonly UnmergeCells: XlsxEditorMenuNamedRow;
  readonly BordersOutline: XlsxEditorMenuNamedRow;
  readonly BordersAll: XlsxEditorMenuNamedRow;
  readonly BordersNone: XlsxEditorMenuNamedRow;
  readonly NumberFormat: XlsxEditorMenuNamedRow;
  readonly IncreaseDecimals: XlsxEditorMenuNamedRow;
  readonly DecreaseDecimals: XlsxEditorMenuNamedRow;
}

/**
 * The menu bar: File, Edit, View, Insert, and Format, from the chrome registry. Every row of a
 * built-in menu is a chrome slot, so a row and its toolbar control share one label, one command,
 * and one enabled state. Place it inside `XlsxEditor.Viewport`.
 *
 * @example
 * ```tsx
 * <XlsxEditor.Menu onSave={upload}>
 *   <XlsxEditor.Menu.Edit>
 *     <XlsxEditor.Menu.Paste hidden />
 *   </XlsxEditor.Menu.Edit>
 *   <XlsxEditor.Menu.Menu id="share" label="Share" after="format">
 *     <XlsxEditor.Menu.Row label="Copy link" onSelect={copyLink} />
 *   </XlsxEditor.Menu.Menu>
 * </XlsxEditor.Menu>
 * ```
 *
 * @public
 */
export const XlsxEditorMenu: XlsxEditorMenuNamespace = Object.assign(XlsxEditorMenuRoot, {
  Menu: XlsxEditorMenuMenu,
  File: MenuFile,
  Edit: MenuEdit,
  View: MenuView,
  Insert: MenuInsert,
  Format: MenuFormat,
  Item: XlsxEditorMenuItem,
  Row: XlsxEditorMenuRow,
  Group: XlsxEditorMenuGroup,
  Separator: XlsxEditorMenuSeparator,
  Submenu: XlsxEditorMenuSubmenu,
  Entry: XlsxEditorMenuEntry,
  ...MENU_NAMED_ROWS,
});

export type { XlsxEditorMenuPartProps };
