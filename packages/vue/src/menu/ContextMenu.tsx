/** @jsxImportSource vue */
// The grid context menu, the Vue twin of packages/react/src/menu/ContextMenu.tsx. After
// docx-editor/packages/vue/src/editor/contextmenu/ at f5cc27dfb (EigenPal): it opens at the
// pointer on a secondary press, and at the active cell on Shift+F10 or the Menu key; its rows are
// chrome slots; a row that names a default row replaces it in place, `hidden` removes it, other
// rows follow the default rows, `Start` rows go first, and `preset: false` shows only the default
// slot. `disabled` gives the browser its own menu back. The rows are the menu bar's parts.
//
// Differences from the React twin:
// - The hooks return computed refs (`UseContextMenuTargetReturn`, `UseContextMenuAnchorReturn`).

import {
  computed,
  defineComponent,
  inject,
  provide,
  shallowRef,
  useId,
  watch,
  type ComputedRef,
  type ExtractPublicPropTypes,
  type AllowedComponentProps,
  type InjectionKey,
  type PropType,
  type Ref,
  type VNode,
} from 'vue';
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
import { toolbarHostKey } from '../toolbar/slot.tsx';
import { NumberFormatDialogHost } from '../toolbar/NumberFormatDialog.tsx';
import { visibleNodes } from '../toolbar/host-parts.tsx';
import {
  MENU_NAMED_ROWS,
  MenuPanel,
  XlsxEditorMenuGroup,
  XlsxEditorMenuItem,
  XlsxEditorMenuRow,
  XlsxEditorMenuSeparator,
  XlsxEditorMenuSubmenu,
  menuBarKey,
  panelRows,
  type MenuBarState,
  type XlsxEditorMenuNamedRow,
} from './parts.tsx';

/** What the open grid context menu acts on, the twin of React's `ContextMenuTarget`. @public */
export interface ContextMenuTarget {
  /** Cells, whole rows, whole columns, or the whole sheet. */
  readonly kind: GridContextMenuTargetKind;
  /** The sheet, 0-based. */
  readonly sheetIndex: number;
  /** The selected range, 0-based and inclusive. */
  readonly range: GridRange;
}

/** Where the open grid context menu stands, the twin of React's `ContextMenuAnchor`. @public */
export interface ContextMenuAnchor {
  /** The window x of the point that the menu opened at, in CSS pixels. */
  readonly x: number;
  /** The window y of the point that the menu opened at, in CSS pixels. */
  readonly y: number;
  /** `'pointer'` for a secondary press, `'keyboard'` for Shift+F10 or the Menu key. */
  readonly trigger: 'pointer' | 'keyboard';
}

const requestKey: InjectionKey<Readonly<Ref<GridContextMenuRequest | null>>> =
  Symbol('xlsxContextMenuRequest');

/** What `useContextMenuTarget` returns. @public */
export type UseContextMenuTargetReturn = ComputedRef<ContextMenuTarget | null>;

/** What the open grid context menu acts on, or null while it is closed. @public */
export function useContextMenuTarget(): UseContextMenuTargetReturn {
  const request = inject(requestKey, shallowRef(null));
  return computed(() => {
    const r = request.value;
    return r ? { kind: r.kind, sheetIndex: r.sheetIndex, range: r.range } : null;
  });
}

/** What `useContextMenuAnchor` returns. @public */
export type UseContextMenuAnchorReturn = ComputedRef<ContextMenuAnchor | null>;

/** Where the open grid context menu stands and how it opened, or null while it is closed. @public */
export function useContextMenuAnchor(): UseContextMenuAnchorReturn {
  const request = inject(requestKey, shallowRef(null));
  return computed(() => {
    const r = request.value;
    return r ? { x: r.x, y: r.y, trigger: r.trigger } : null;
  });
}

/**
 * Rows that go before the default rows, with a separator after them, the twin of React's
 * `XlsxEditorContextMenuStart`.
 *
 * @public
 */
export const XlsxEditorContextMenuStart = defineComponent({
  name: 'XlsxEditorContextMenuStart',
  setup(_props, { slots }) {
    return () => slots.default?.();
  },
});

const xlsxEditorContextMenuProps = {
  /**
   * `false` renders only the default slot as the rows. The default is `true`: the default rows
   * for what the press hit, with a row that names a default row replacing it in place, and the
   * other rows after the default rows.
   */
  preset: { type: Boolean, default: true },
  /** `true` removes the menu: a secondary press then shows the browser's own menu. */
  disabled: { type: Boolean, default: false },
  /** Called when the menu opens (`true`) or closes (`false`). */
  onOpenChange: { type: Function as PropType<(open: boolean) => void>, default: undefined },
  /** Appended after the `xlsx-context-menu` class. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/** Props of `XlsxEditorContextMenu`, the twin of React's `XlsxEditorContextMenuProps`. @public */
export type XlsxEditorContextMenuProps = ExtractPublicPropTypes<typeof xlsxEditorContextMenuProps>;

const isLoading = (s: EditorSnapshot): boolean => s.status === 'loading';

/** The rows of `Start` parts, and the other rows. */
function splitStart(nodes: readonly VNode[]): { start: VNode[]; rest: VNode[] } {
  const start: VNode[] = [];
  const rest: VNode[] = [];
  for (const node of nodes) {
    if (node.type === XlsxEditorContextMenuStart) {
      const children = node.children as { default?: () => unknown } | null;
      start.push(...visibleNodes(children?.default?.()));
    } else rest.push(node);
  }
  return { start, rest };
}

const XlsxEditorContextMenuRoot = defineComponent({
  name: 'XlsxEditorContextMenu',
  // The host's attributes go onto the anchor element, not the dialog host around it.
  inheritAttrs: false,
  props: xlsxEditorContextMenuProps,
  setup(props, { attrs, slots }) {
    const { t } = useTranslation();
    const editor = useXlsxEditor();
    const panelId = useId();
    const request = shallowRef<GridContextMenuRequest | null>(null);
    const anchor = shallowRef<HTMLDivElement | null>(null);
    const focusEdge: { current: 'first' | 'last' | null } = { current: null };
    provide(requestKey, request);

    watch(
      [editor, () => props.disabled],
      ([ed, disabled], _old, onCleanup) => {
        if (disabled) request.value = null;
        if (!ed || disabled) return;
        onCleanup(
          ed.on('contextMenu', (next) => {
            focusEdge.current = 'first';
            request.value = next;
          })
        );
      },
      { immediate: true }
    );
    watch(
      () => request.value !== null,
      (open) => props.onOpenChange?.(open)
    );
    // An open of another workbook closes the menu: its rows showed the old workbook's state.
    const loading = useEditorState(isLoading);
    watch(loading, (now) => {
      if (now) request.value = null;
    });

    const close = (focus: 'trigger' | 'grid' | 'none' | 'origin'): void => {
      request.value = null;
      if (focus !== 'none') {
        editor.value?.focus();
        return;
      }
      // A row that moves the keyboard (a dialog) keeps it; a close that leaves it nowhere
      // gives it back to the grid.
      queueMicrotask(() => {
        const doc = anchor.value?.ownerDocument;
        const active = doc?.activeElement;
        if (!active || active === doc.body || !active.isConnected) editor.value?.focus();
      });
    };

    // Placed at its point, then closed on a press outside it, a wheel, a scroll of the page, a
    // resize, and a window blur, as the tab menu (`listenSheetMenuDismiss`).
    watch(
      [request, anchor],
      ([r, node], _old, onCleanup) => {
        const panel = node?.firstElementChild as HTMLElement | null | undefined;
        if (!r || !node || !panel) return;
        const stopPlace = placeContextMenu(node, panel, r.x, r.y);
        const stopDismiss = listenSheetMenuDismiss(
          node,
          (target) => node.contains(target),
          () => (request.value = null)
        );
        onCleanup(() => {
          stopPlace();
          stopDismiss();
        });
      },
      { flush: 'post' }
    );

    const rtl = shallowRef(false);
    const apple = shallowRef(false);
    watch(
      anchor,
      (node) => {
        if (!node) return;
        rtl.value = hostWindowOf(node).getComputedStyle(node).direction === 'rtl';
        apple.value = isApplePlatform(hostWindowOf(node).navigator);
      },
      { flush: 'post' }
    );
    const host = computed<ToolbarHost>(() => {
      const read = systemClipboardReader(anchor.value);
      return read ? { readClipboard: read } : {};
    });
    provide(toolbarHostKey, host);
    const state: MenuBarState = {
      openMenu: shallowRef(null),
      open: () => {},
      focusRequest: { current: null },
      activeMenu: computed(() => null),
      setActiveMenu: () => {},
      order: computed(() => []),
      rtl,
      apple,
      closeAll: close,
      triggerOf: () => null,
    };
    provide(menuBarKey, state);

    return () => {
      const r = request.value;
      const label = t('contextMenu.ariaLabel');
      const { start, rest } = splitStart(visibleNodes(slots.default?.()));
      return (
        <NumberFormatDialogHost>
          {{
            default: () => (
              <div
                {...attrs}
                ref={anchor}
                class={['xlsx-context-menu', props.className, attrs.class]}
                onMousedown={guardToolbarMousedown}
                onContextmenu={(e: MouseEvent) => e.preventDefault()}
              >
                {r ? (
                  <MenuPanel
                    id={panelId}
                    label={label}
                    placed
                    focusEdge={focusEdge}
                    onClose={() => close('grid')}
                  >
                    {{
                      default: () => [
                        ...(start.length > 0 ? [...start, <XlsxEditorMenuSeparator />] : []),
                        panelRows(gridContextMenuEntries(r.kind), rest, props.preset, label),
                      ],
                    }}
                  </MenuPanel>
                ) : null}
              </div>
            ),
          }}
        </NumberFormatDialogHost>
      );
    };
  },
});

/** The grid context menu with its parts as statics, the twin of React's namespace. @public */
export interface XlsxEditorContextMenuNamespace {
  /** The grid context menu. A host `class` lands on its anchor, as on a component. */
  (props: XlsxEditorContextMenuProps & AllowedComponentProps): VNode;
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
 * Menu key, at the active cell. Its default rows are chrome slots for what the press hit. Place
 * it inside `XlsxEditor.Viewport`. The packaged `XlsxEditor` mounts it; its `contextMenu` prop
 * takes these props, or `false`.
 *
 * @public
 */
export const XlsxEditorContextMenu = Object.assign(XlsxEditorContextMenuRoot, {
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
}) as unknown as XlsxEditorContextMenuNamespace;
