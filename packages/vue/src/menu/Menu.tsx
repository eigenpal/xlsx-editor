/** @jsxImportSource vue */
// The menu bar: the Vue twin of the React adapter's `menu/Menu.tsx`, which is after
// docx-editor/packages/react/src/editor/menu/DocxEditorMenu.tsx at ad686574 (EigenPal); the
// header there explains the rules and the differences from DOCX Editor. The order of the menus
// comes from core's `arrangeMenuBar`, the measure from `measureMenuBar` and `observeMenuBar`,
// and Control or Command and S from `listenSaveKey`.
//
// Differences from React: the bar state is a set of refs that the parts inject (`menuBarKey`);
// the toolbar host is a computed ref (`toolbarHostKey`), as the toolbar provides it; the measure
// runs in a watcher after the render; the bar takes the host's attributes
// (`inheritAttrs: false`), because the component also renders the file input and the number
// format dialog beside the bar; and the callbacks are function props with React's names, as on
// `XlsxEditorToolbar`.

import {
  cloneVNode,
  computed,
  defineComponent,
  h,
  provide,
  shallowRef,
  watch,
  type AllowedComponentProps,
  type ExtractPublicPropTypes,
  type PropType,
  type VNode,
} from 'vue';
import {
  MENU_FIXED_ATTRIBUTE,
  arrangeMenuBar,
  guardToolbarMousedown,
  hostWindowOf,
  isApplePlatform,
  isElementNode,
  listenForOutsidePress,
  listenForPopupEscape,
  listenForPopupFocusLeave,
  listenSaveKey,
  measureMenuBar,
  menuRowInBarProblem,
  observeMenuBar,
  openPickedWorkbook,
  VUE_MENU_PART_NAMES,
  runToolbarCommand,
  systemClipboardReader,
  type EditorSnapshot,
  type SavedWorkbook,
  type ToolbarHost,
} from '@xlsx-editor.dev/core/editor';
import { useEditorState, useTranslation, useTranslator, useXlsxEditor } from '../context.ts';
import { toolbarHostKey, translateKey } from '../toolbar/slot.tsx';
import { NumberFormatDialogHost } from '../toolbar/NumberFormatDialog.tsx';
import { vnodeFlag, visibleNodes, warnToolbarProblems } from '../toolbar/host-parts.tsx';
import {
  MENU_NAMED_ROWS,
  MORE_MENU_ID,
  slotRowOf,
  MenuEdit,
  MenuFile,
  MenuFormat,
  MenuInsert,
  MenuView,
  MoreMenu,
  PlacementScope,
  XlsxEditorMenuEntry,
  XlsxEditorMenuGroup,
  XlsxEditorMenuItem,
  XlsxEditorMenuMenu,
  XlsxEditorMenuRow,
  XlsxEditorMenuSeparator,
  XlsxEditorMenuSubmenu,
  menuBarKey,
  menuSpecOf,
  menuText,
  type MenuBarState,
  type XlsxEditorMenuNamedRow,
  type XlsxEditorMenuPart,
} from './parts.tsx';

const xlsxEditorMenuProps = {
  /**
   * Receives the workbook bytes when File > Save is chosen, or when the user presses Control or
   * Command and S in the editor, as the toolbar's `onSave` does. Without it, File > Save is
   * disabled with the engine's reason.
   */
  onSave: {
    type: Function as PropType<(bytes: Uint8Array<ArrayBuffer>, saved: SavedWorkbook) => void>,
    default: undefined,
  },
  /** Called once when File > Save starts a delivery to `onSave`, as the toolbar's `onSaveStart`. */
  onSaveStart: {
    type: Function as PropType<(saving: Promise<SavedWorkbook | null>) => void>,
    default: undefined,
  },
  /**
   * Replaces File > Open. The default opens a file picker and passes the bytes of the picked
   * workbook to the editor's `load`: a file that the user picks, never a fetch. The default does
   * not ask first: the picked file replaces the open workbook, also with unsaved edits. Pass
   * `onOpen` to ask the user first. A picked file that the browser cannot read or the editor
   * cannot open replaces nothing: the open workbook, its edits, and an open cell edit stay, and
   * the refusal `open-failed` shows.
   */
  onOpen: { type: Function as PropType<() => void>, default: undefined },
  /**
   * Called when the packaged File > Open reads a picked file, before its bytes load, so you can
   * show its name. Not called when `onOpen` replaces the packaged picker.
   */
  onOpenFile: { type: Function as PropType<(file: File) => void>, default: undefined },
  /**
   * Whether your page opens another workbook now, as the toolbar's `opening`: File > Open and
   * File > Save are disabled while it is true.
   */
  opening: { type: Boolean, default: undefined },
  /**
   * `false` lets the bar wrap onto more lines instead of moving the menus that do not fit into
   * the "More" menu. The default is `true`.
   */
  overflow: { type: Boolean, default: true },
  /**
   * `false` renders only the default slot, with no registry menus. The default is `true`: a menu
   * in the default slot changes its menu in place, and other content shows after the menus.
   */
  preset: { type: Boolean, default: true },
  /** Appended after the `xlsx-menubar` class. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/** Props of `XlsxEditorMenu`, the twin of React's `XlsxEditorMenuProps`. @public */
export type XlsxEditorMenuProps = ExtractPublicPropTypes<typeof xlsxEditorMenuProps>;

const ACCEPT =
  '.xlsx,.xltx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,' +
  'application/vnd.openxmlformats-officedocument.spreadsheetml.template';

const isLoading = (s: EditorSnapshot): boolean => s.status === 'loading';

/** The menu bar, with no statics. */
const XlsxEditorMenuRoot = defineComponent({
  name: 'XlsxEditorMenu',
  inheritAttrs: false,
  props: xlsxEditorMenuProps,
  setup(props, { attrs, slots }) {
    const { t } = useTranslation();
    const translator = useTranslator();
    const editor = useXlsxEditor();
    const bar = shallowRef<HTMLDivElement | null>(null);
    // The row that holds the menu bar and the host's loose children: the overflow measures it.
    const row = shallowRef<HTMLDivElement | null>(null);
    const fileInput = shallowRef<HTMLInputElement | null>(null);
    const openMenu = shallowRef<string | null>(null);
    const activeId = shallowRef<string | null>(null);
    const rtl = shallowRef(false);
    const apple = shallowRef(
      isApplePlatform(typeof navigator === 'undefined' ? undefined : navigator)
    );
    const focusRequest: MenuBarState['focusRequest'] = { current: null };
    const ids = shallowRef<readonly string[]>([]);
    const shown = shallowRef(Number.MAX_SAFE_INTEGER);

    const packagedOpen = (): void => fileInput.value?.click();
    const host = computed<ToolbarHost>(() => {
      const read = systemClipboardReader(bar.value);
      const open = props.onOpen ?? (editor.value ? packagedOpen : undefined);
      return {
        ...(props.onSave ? { save: props.onSave } : {}),
        ...(props.onSaveStart ? { saveStart: props.onSaveStart } : {}),
        ...(props.opening === undefined ? {} : { opening: () => props.opening ?? false }),
        ...(open ? { open } : {}),
        ...(read ? { readClipboard: read } : {}),
      };
    });
    provide(toolbarHostKey, host);

    const inBar = computed(() => ids.value.slice(0, Math.min(shown.value, ids.value.length)));
    const inMore = computed(() => ids.value.slice(Math.min(shown.value, ids.value.length)));
    const order = computed<readonly string[]>(() =>
      inMore.value.length > 0 ? [...inBar.value, MORE_MENU_ID] : inBar.value
    );
    const triggerOf = (id: string): HTMLElement | null =>
      bar.value?.querySelector<HTMLElement>(
        id === MORE_MENU_ID
          ? ':scope > [data-menu-more] > [role="menuitem"]'
          : `:scope > [data-menu="${id.replace(/["\\]/g, '\\$&')}"] > [role="menuitem"]`
      ) ?? null;
    let openedInBar = false;
    const state: MenuBarState = {
      openMenu,
      open: (id, edge) => {
        // Whether the bar held the keyboard when the menu opened: a keyboard open. A pointer
        // press keeps the keyboard in the grid, so a pointer open leaves this false.
        const el = bar.value;
        if (id && openMenu.value === null)
          openedInBar = !!el && el.contains(el.ownerDocument.activeElement);
        focusRequest.current = id && edge ? { id, edge } : null;
        openMenu.value = id;
        if (id) activeId.value = id;
      },
      focusRequest,
      activeMenu: computed(() =>
        activeId.value !== null && order.value.includes(activeId.value)
          ? activeId.value
          : (order.value[0] ?? null)
      ),
      setActiveMenu: (id) => {
        activeId.value = id;
      },
      order,
      rtl,
      apple,
      closeAll: (focus) => {
        const was = openMenu.value;
        openMenu.value = null;
        const toTrigger = focus === 'trigger' || (focus === 'origin' && openedInBar);
        if (toTrigger && was) triggerOf(was)?.focus();
        else if (focus === 'grid') editor.value?.focus();
      },
      triggerOf,
    };
    provide(menuBarKey, state);

    // An open of another workbook closes the menus: their rows showed the old workbook's state.
    const loading = useEditorState(isLoading);
    watch(loading, (now) => {
      if (now) openMenu.value = null;
    });

    // While a menu is open: a press outside the bar, an Escape from outside it, and focus that
    // leaves it close the menu.
    watch(
      [openMenu, bar],
      ([id, el], _old, onCleanup) => {
        if (id === null || !el) return;
        const contains = (node: Node) => el.contains(node);
        const close = (): void => {
          openMenu.value = null;
        };
        const stops = [
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
        onCleanup(() => stops.forEach((stop) => stop()));
      },
      { flush: 'post' }
    );

    // Control or Command and S runs File > Save, as the toolbar binds it.
    watch(
      [bar, editor, () => !!props.onSave, host],
      ([el, ed, saves, hst], _old, onCleanup) => {
        if (!el || !ed || !saves) return;
        onCleanup(listenSaveKey(el, () => runToolbarCommand(ed, 'file.save', undefined, hst)));
      },
      { immediate: true, flush: 'post' }
    );

    // The overflow: the menus that fit stay in the bar, the rest go into the "More" menu.
    const widths = new Map<string, number>();
    let measuredFor: { t: unknown; ids: string } | null = null;
    watch(
      [() => props.overflow, ids, shown, translator, row],
      ([measuring, all, , tr, el], _old, onCleanup) => {
        const key = all.join('\n');
        if (measuredFor?.t !== tr || measuredFor.ids !== key) widths.clear();
        measuredFor = { t: tr, ids: key };
        if (!el) return;
        rtl.value = el.ownerDocument.defaultView?.getComputedStyle(el).direction === 'rtl';
        apple.value = isApplePlatform(hostWindowOf(el).navigator);
        if (!measuring) {
          shown.value = Number.MAX_SAFE_INTEGER;
          return;
        }
        const measure = (): void => {
          const next = measureMenuBar(el, all, widths);
          if (shown.value !== next) shown.value = next;
        };
        measure();
        onCleanup(observeMenuBar(el, measure));
      },
      { immediate: true, flush: 'post' }
    );

    const onFile = (event: Event): void => {
      const input = event.target as HTMLInputElement;
      const file = input.files?.[0];
      // Cleared, so a pick of the same file again fires a change.
      input.value = '';
      const ed = editor.value;
      if (!file || !ed) return;
      props.onOpenFile?.(file);
      // The load starts now, so a later load wins over this pick during its read.
      void openPickedWorkbook(ed, file);
    };

    return () => {
      const specs = new Map<string, NonNullable<ReturnType<typeof menuSpecOf>>>();
      const appended: VNode[] = [];
      const problems: string[] = [];
      for (const node of visibleNodes(slots.default?.())) {
        const spec = menuSpecOf(node);
        const row = spec ? null : slotRowOf(node);
        if (spec) {
          specs.delete(spec.id);
          specs.set(spec.id, spec);
        } else if (row) problems.push(menuRowInBarProblem(row.slotId, VUE_MENU_PART_NAMES));
        else appended.push(node);
      }
      const arranged = arrangeMenuBar(
        new Map(
          [...specs.values()].map((s) => [
            s.id,
            {
              id: s.id,
              label: menuText(s.props.label),
              labelKey: menuText(s.props.labelKey),
              hidden: vnodeFlag(s.props.hidden),
              after: menuText(s.props.after),
            },
          ])
        ),
        props.preset,
        (key) => translateKey(t, key),
        (message) => problems.push(message)
      );
      warnToolbarProblems(problems);
      const next = arranged.map((m) => m.id);
      if (next.join('\n') !== ids.value.join('\n')) ids.value = next;
      const elements = (): VNode[] =>
        arranged.map((m) => {
          const spec = specs.get(m.id);
          return spec
            ? cloneVNode(spec.node, { key: m.id })
            : h(XlsxEditorMenuMenu, { key: m.id, id: m.id });
        });
      const more = inMore.value;
      return (
        <NumberFormatDialogHost>
          {{
            default: () => [
              // The host's loose children sit beside the menu bar element, not in it: they are no
              // rows of the menubar and keep their own tab stops.
              <div
                {...attrs}
                ref={row}
                class={['xlsx-menubar', props.className, attrs.class]}
                {...(props.overflow ? { 'data-overflow': '' } : {})}
                onMousedown={guardToolbarMousedown}
              >
                <div
                  ref={bar}
                  role="menubar"
                  aria-label={t('menu.label')}
                  class="xlsx-menubar-menus"
                >
                  <PlacementScope placement={{ kind: 'bar', ids: new Set(inBar.value) }}>
                    {{ default: elements }}
                  </PlacementScope>
                  {more.length > 0 ? (
                    <MoreMenu>
                      {{
                        default: () => (
                          <PlacementScope placement={{ kind: 'more', ids: new Set(more) }}>
                            {{ default: elements }}
                          </PlacementScope>
                        ),
                      }}
                    </MoreMenu>
                  ) : null}
                </div>
                {appended.length > 0 ? (
                  <div class="xlsx-menubar-fixed" {...{ [MENU_FIXED_ATTRIBUTE]: '' }}>
                    {appended}
                  </div>
                ) : null}
              </div>,
              // Opening a workbook is a file that the user picks, never a fetch. A host `onOpen`
              // replaces the picker, so the bar adds no file input of its own then.
              props.onOpen ? null : (
                <input
                  ref={fileInput}
                  type="file"
                  accept={ACCEPT}
                  hidden
                  tabindex={-1}
                  aria-hidden="true"
                  onChange={onFile}
                />
              ),
            ],
          }}
        </NumberFormatDialogHost>
      );
    };
  },
});

/**
 * The menu bar and its parts, the twin of React's `XlsxEditorMenuNamespace`. @public
 */
export interface XlsxEditorMenuNamespace {
  /** The menu bar. A host `class` and `style` land on the bar, as on a component. */
  (props: XlsxEditorMenuProps & AllowedComponentProps): VNode;
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
 * and one enabled state. Place it inside `XlsxEditorViewport`. Its parts go in its default slot.
 *
 * @example
 * ```ts
 * h(XlsxEditorMenu, { onSave: upload }, () => [
 *   h(XlsxEditorMenu.Edit, null, () => [h(XlsxEditorMenu.Paste, { hidden: true })]),
 *   h(XlsxEditorMenu.Menu, { id: 'share', label: 'Share', after: 'format' }, () => [
 *     h(XlsxEditorMenu.Row, { label: 'Copy link', onSelect: copyLink }),
 *   ]),
 * ]);
 * ```
 *
 * @public
 */
export const XlsxEditorMenu = Object.assign(XlsxEditorMenuRoot, {
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
}) as unknown as XlsxEditorMenuNamespace;
