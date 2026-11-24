/** @jsxImportSource vue */
// The parts of the menu bar: the Vue twin of the React adapter's `menu/parts.tsx`, which is
// after docx-editor/packages/react/src/editor/menu/parts.tsx at ad686574 (EigenPal); the header
// there explains the rules and the differences from DOCX Editor. The row rules and the keys are
// core's (`arrangeMenuRows`, `menuPanelKey`, `menuBarTriggerKey`), so the two adapters act the
// same.
//
// Differences from React: the bar state is a set of refs that the parts inject
// (`menuBarKey`); an icon and the rows are slots (`icon` and `default`), as Vue passes content;
// `Menu.Row` takes `onSelect` as a function prop, so a template binds it with `@select`; the
// props of a part vnode are read camelized, and a template's bare boolean attribute
// (`<Menu.Paste hidden />`) is `true` (`vnodeFlag`); and the markers are set after
// `defineComponent` (`markPart`), as on the toolbar parts.

import {
  Fragment,
  camelize,
  computed,
  defineComponent,
  type DefineComponent,
  h,
  inject,
  isVNode,
  provide,
  shallowRef,
  useId,
  watch,
  type ComputedRef,
  type ExtractPublicPropTypes,
  type InjectionKey,
  type PropType,
  type Ref,
  type VNode,
  type VNodeChild,
} from 'vue';
import {
  CHROME_MENUS,
  MENU_NAMED_PARTS,
  arrangeMenuRows,
  placeMenuPanel,
  MENU_MORE_ID,
  menuHiddenRowProblem,
  VUE_MENU_PART_NAMES,
  chromeMenuShortcut,
  chromeMenuRowKey,
  guardToolbarMousedown,
  menuBarTriggerKey,
  menuPanelKey,
  menuRowKeepsFocus,
  menuRowRole,
  menuKeyWords,
  menuShortcutText,
  menuStep,
  menuTypeahead,
  type ChromeMenuEntry,
  type ChromeMenuId,
  type ChromeSlotId,
  type MenuRowPartName,
} from '@xlsx-editor.dev/core/editor';
import { useTranslation, useXlsxEditor } from '../context.ts';
import { ChromeIcon, translateKey, useToolbarSlot } from '../toolbar/slot.tsx';
import { useNumberFormatDialogOpener } from '../toolbar/NumberFormatDialog.tsx';
import { vnodeFlag, visibleNodes, warnToolbarProblems } from '../toolbar/host-parts.tsx';
import { ICON_MORE } from '../toolbar/icons.ts';

/** The rows of a panel: the menu items, checkbox items, and submenu triggers. */
export const PANEL_ROWS =
  ':scope > [role^="menuitem"], :scope > [role="none"] > [role^="menuitem"], ' +
  ':scope > [role="group"] > [role^="menuitem"], ' +
  ':scope > [role="group"] > [role="none"] > [role^="menuitem"]';

/** The id of the "More" menu, which holds the menus that do not fit. Never a host menu id. */
export const MORE_MENU_ID = MENU_MORE_ID;

/** What the menu bar gives its menus. */
export interface MenuBarState {
  readonly openMenu: Readonly<Ref<string | null>>;
  /** Open a menu (null closes every menu). `edge` moves the keyboard to its first or last row. */
  readonly open: (id: string | null, edge?: 'first' | 'last') => void;
  /** The pending keyboard move into the panel that opens next. */
  readonly focusRequest: { current: { id: string; edge: 'first' | 'last' } | null };
  /** The menu whose trigger is the bar's one tab stop. */
  readonly activeMenu: ComputedRef<string | null>;
  readonly setActiveMenu: (id: string) => void;
  /** The ids of the triggers in the bar, in order: for the arrow keys. */
  readonly order: ComputedRef<readonly string[]>;
  readonly rtl: Readonly<Ref<boolean>>;
  readonly apple: Readonly<Ref<boolean>>;
  /**
   * Close every menu, then move the keyboard to the trigger, to the grid, or nowhere. `origin`
   * moves it to the trigger only when the bar held the keyboard as the menu opened (a keyboard
   * open); after a pointer open, the keyboard stays where it was, such as in the grid.
   */
  readonly closeAll: (focus: 'trigger' | 'grid' | 'none' | 'origin') => void;
  /** The trigger of a menu in the bar. */
  readonly triggerOf: (id: string) => HTMLElement | null;
}

export const menuBarKey: InjectionKey<MenuBarState> = Symbol('xlsxMenuBar');

/** Where a menu renders: in the bar, or as a submenu of the "More" menu. */
export type MenuPlacement =
  | { readonly kind: 'bar'; readonly ids: ReadonlySet<string> }
  | { readonly kind: 'more'; readonly ids: ReadonlySet<string> };

export const menuPlacementKey: InjectionKey<ComputedRef<MenuPlacement | null>> =
  Symbol('xlsxMenuPlacement');

/** Gives a placement to the menus in its default slot. */
export const PlacementScope = defineComponent({
  name: 'XlsxMenuPlacementScope',
  props: { placement: { type: Object as PropType<MenuPlacement>, required: true } },
  setup(props, { slots }) {
    provide(
      menuPlacementKey,
      computed(() => props.placement)
    );
    return () => slots.default?.();
  },
});

const panelKey: InjectionKey<boolean> = Symbol('xlsxMenuPanel');

/** Set a static marker that the bar and the panels read. Not a public member. */
function markPart(part: object, marker: string, value: unknown = true): void {
  (part as Record<string, unknown>)[marker] = value;
}

/** The marker of a part vnode, or undefined. */
function markerOf(node: VNode, marker: string): unknown {
  const type = node.type as unknown;
  if (typeof type !== 'object' && typeof type !== 'function') return undefined;
  return type === null ? undefined : (type as Record<string, unknown>)[marker];
}

/** The props of a vnode, camelized. */
export function vnodeProps(node: VNode): Record<string, unknown> {
  const props: Record<string, unknown> = Object.create(null);
  for (const [key, value] of Object.entries(node.props ?? {})) props[camelize(key)] = value;
  return props;
}

const text = (value: unknown): string | undefined =>
  typeof value === 'string' ? value : undefined;

// ---------------------------------------------------------------------------------------------
// Rows

/** The markup of one row, for a slot row and a host row. */
function rowButton(props: {
  readonly id: string;
  readonly label: VNodeChild;
  readonly icon: VNodeChild;
  readonly shortcut: string | undefined;
  readonly disabled: boolean;
  readonly reason: string | null;
  readonly title: string | undefined;
  readonly checked: boolean | undefined;
  /** A row of a one-of set: `menuitemradio` in place of `menuitemcheckbox`. */
  readonly radio?: boolean;
  readonly slot?: string;
  readonly className?: string | undefined;
  readonly onPress: () => void;
}) {
  const reasonId = `${props.id}-reason`;
  return (
    <button
      type="button"
      role={
        props.checked === undefined
          ? 'menuitem'
          : props.radio
            ? 'menuitemradio'
            : 'menuitemcheckbox'
      }
      tabindex={-1}
      class={['xlsx-menu-row', props.className]}
      {...(props.checked === undefined ? {} : { 'aria-checked': props.checked })}
      {...(props.disabled ? { 'aria-disabled': 'true', 'data-disabled': '' } : {})}
      {...(props.reason ? { 'aria-describedby': reasonId } : {})}
      {...(props.title ? { title: props.title } : {})}
      {...(props.slot ? { 'data-slot': props.slot } : {})}
      onClick={() => {
        if (!props.disabled) props.onPress();
      }}
    >
      <span class="xlsx-menu-icon" aria-hidden="true">
        {props.icon}
      </span>
      <span class="xlsx-menu-label">{props.label}</span>
      {props.shortcut ? (
        <span class="xlsx-menu-shortcut" aria-hidden="true">
          {props.shortcut}
        </span>
      ) : null}
      {props.reason ? (
        <span id={reasonId} hidden>
          {props.reason}
        </span>
      ) : null}
    </button>
  );
}

const xlsxEditorMenuRowProps = {
  /** The text of the row. Without it, the default slot is the text. */
  label: { type: String, default: undefined },
  /** The shortcut text at the end of the row, as you write it. */
  shortcut: { type: String, default: undefined },
  /** The row is disabled: it stays focusable, and a press does nothing. */
  disabled: { type: Boolean, default: undefined },
  /** Why the row is disabled: its title and its accessible description. */
  disabledReason: { type: String, default: undefined },
  /** The row shows a checked state (`menuitemcheckbox`, or `menuitemradio` with `selected`). */
  active: { type: Boolean, default: undefined },
  /**
   * The row is one of a one-of set, such as one alignment of several: its role is
   * `menuitemradio`, and `active` marks the one value. As DOCX Editor's prop of the same name.
   */
  selected: { type: Boolean, default: undefined },
  /** Runs on a press of an enabled row, after the menu closes. Bind it with `@select`. */
  onSelect: { type: Function as PropType<() => void>, default: undefined },
  /** Appended after the `xlsx-menu-row` class. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/** Props of `XlsxEditorMenu.Row`, the twin of React's `XlsxEditorMenuRowProps`. @public */
export type XlsxEditorMenuRowProps = ExtractPublicPropTypes<typeof xlsxEditorMenuRowProps>;

/**
 * One row of the host's own: a label, an icon (`icon` slot), a shortcut text, and an action.
 * Selecting it closes the menu, then runs `onSelect` (`@select`).
 *
 * @public
 */
export const XlsxEditorMenuRow = defineComponent({
  name: 'XlsxEditorMenuRow',
  inheritAttrs: false,
  props: xlsxEditorMenuRowProps,
  setup(props, { slots, attrs }) {
    const bar = inject(menuBarKey, null);
    const id = useId();
    return () => {
      const disabled = props.disabled === true;
      const reason = disabled ? (props.disabledReason ?? null) : null;
      const label = props.label ?? slots.default?.();
      return rowButton({
        id,
        label,
        icon: slots.icon?.() ?? null,
        shortcut: props.shortcut,
        disabled,
        reason,
        title: reason && props.label ? `${props.label}: ${reason}` : (reason ?? undefined),
        checked: props.active === undefined ? (props.selected ? false : undefined) : props.active,
        radio: props.selected === true,
        className: [props.className, text(attrs.class)].filter(Boolean).join(' ') || undefined,
        onPress: () => {
          // After a keyboard open, the keyboard goes back to the menu's trigger as the panel
          // closes; after a pointer open, it stays where it was. A command that moves it (to
          // the grid, or into a dialog) then takes it from there.
          bar?.closeAll('origin');
          props.onSelect?.();
        },
      });
    };
  },
});

const xlsxEditorMenuItemProps = {
  /** The slot of the row, such as `'edit.copy'`. */
  slotId: { type: String as PropType<ChromeSlotId>, required: true },
  /** A catalog key for the label in place of the slot's own label. */
  labelKey: { type: String, default: undefined },
  /** A label in place of the slot's own label. Wins over `labelKey`. */
  label: { type: String, default: undefined },
  /** The key chord of the row, such as `'Mod+C'`, shown for the platform. */
  shortcut: { type: String, default: undefined },
  /** Appended after the `xlsx-menu-row` class. */
  className: { type: String, default: undefined },
  /** Inside a panel, removes the registry row of the same slot. */
  hidden: { type: Boolean, default: undefined },
} as const;

/** Props of `XlsxEditorMenu.Item`, the twin of React's `XlsxEditorMenuItemProps`. @public */
export type XlsxEditorMenuItemProps = ExtractPublicPropTypes<typeof xlsxEditorMenuItemProps>;

/** One slot row: its state and words from the slot, for `XlsxEditorMenu.Item`. */
const SlotRow = defineComponent({
  name: 'XlsxMenuSlotRow',
  props: xlsxEditorMenuItemProps,
  setup(props) {
    const bar = inject(menuBarKey, null);
    const editor = useXlsxEditor();
    const { t } = useTranslation();
    const s = useToolbarSlot(props.slotId);
    const opener = useNumberFormatDialogOpener();
    const id = useId();
    return () => {
      const words = s.words.value;
      const label = props.label ?? (props.labelKey ? translateKey(t, props.labelKey) : words.label);
      const chord = props.shortcut ?? chromeMenuShortcut(props.slotId);
      return rowButton({
        id,
        label,
        icon: <ChromeIcon paths={s.control.paths} />,
        shortcut: chord
          ? menuShortcutText(
              chord,
              bar?.apple.value ?? false,
              menuKeyWords((k) => translateKey(t, k))
            )
          : undefined,
        disabled: !s.isEnabled.value,
        reason: words.description,
        title: words.title,
        checked: s.control.toggle ? s.isActive.value : undefined,
        radio: menuRowRole(props.slotId, s.control.toggle === true) === 'menuitemradio',
        slot: props.slotId,
        className: props.className,
        onPress: () => {
          // A row that keeps the keyboard (Open, Custom number format) gives it back to the
          // trigger after a keyboard open first, so a closed picker or dialog leaves it there; every other row gives it to
          // the grid after its command.
          bar?.closeAll(menuRowKeepsFocus(props.slotId) ? 'origin' : 'none');
          if (props.slotId === 'number.custom') {
            if (!opener?.open(null)) editor.value?.focus();
            return;
          }
          s.execute();
          if (!menuRowKeepsFocus(props.slotId)) editor.value?.focus();
        },
      });
    };
  },
});

/**
 * One chrome slot as a live row: its label and icon from the registry, its state from
 * `toolbarCommandState`, and its command. Selecting it closes the menu and runs the slot.
 *
 * @public
 */
export const XlsxEditorMenuItem = defineComponent({
  name: 'XlsxEditorMenuItem',
  props: xlsxEditorMenuItemProps,
  setup(props) {
    return () =>
      props.hidden
        ? null
        : h(SlotRow, {
            key: props.slotId,
            slotId: props.slotId,
            labelKey: props.labelKey,
            label: props.label,
            shortcut: props.shortcut,
            className: props.className,
          });
  },
});

markPart(XlsxEditorMenuItem, 'xlsxMenuItem');

const xlsxEditorMenuNamedRowProps = {
  labelKey: xlsxEditorMenuItemProps.labelKey,
  label: xlsxEditorMenuItemProps.label,
  shortcut: xlsxEditorMenuItemProps.shortcut,
  className: xlsxEditorMenuItemProps.className,
  hidden: xlsxEditorMenuItemProps.hidden,
} as const;

/** Props of a named row (`XlsxEditorMenu.Cut`): `Menu.Item` without its slot. @public */
export type XlsxEditorMenuNamedRowProps = ExtractPublicPropTypes<
  typeof xlsxEditorMenuNamedRowProps
>;

function defineNamedRow(name: MenuRowPartName): XlsxEditorMenuNamedRow {
  const slot = MENU_NAMED_PARTS[name].slot;
  const Part = defineComponent({
    name: `XlsxEditorMenu${name}`,
    props: xlsxEditorMenuNamedRowProps,
    setup(props) {
      return () => h(XlsxEditorMenuItem, { ...props, slotId: slot });
    },
  });
  markPart(Part, 'xlsxMenuRowSlot', slot);
  return Part as unknown as XlsxEditorMenuNamedRow;
}

/** A named row part, which stands for `Menu.Item` with its slot. @public */
export type XlsxEditorMenuNamedRow = DefineComponent<XlsxEditorMenuNamedRowProps>;

/** The named rows, one per row of core's `MENU_NAMED_PARTS`. */
export const MENU_NAMED_ROWS = Object.fromEntries(
  (Object.keys(MENU_NAMED_PARTS) as MenuRowPartName[]).map((name) => [name, defineNamedRow(name)])
) as Record<MenuRowPartName, XlsxEditorMenuNamedRow>;

/** The slot row that a vnode stands for, or null for any other vnode. */
export function slotRowOf(node: VNode): { slotId: string; hidden?: boolean } | null {
  const props = vnodeProps(node);
  const named = markerOf(node, 'xlsxMenuRowSlot');
  if (typeof named === 'string') return { slotId: named, hidden: vnodeFlag(props.hidden) };
  if (markerOf(node, 'xlsxMenuItem') === true && typeof props.slotId === 'string')
    return { slotId: props.slotId, hidden: vnodeFlag(props.hidden) };
  return null;
}

const xlsxEditorMenuGroupProps = {
  /** The heading, as you write it. Wins over `labelKey`. */
  label: { type: String, default: undefined },
  /** A catalog key for the heading. */
  labelKey: { type: String, default: undefined },
  /** Appended after the `xlsx-menu-group` class. A host `class` also works. */
  className: { type: String, default: undefined },
  /** Renders nothing. */
  hidden: { type: Boolean, default: undefined },
} as const;

/** Props of `XlsxEditorMenu.Group`, the twin of React's `XlsxEditorMenuGroupProps`. @public */
export type XlsxEditorMenuGroupProps = ExtractPublicPropTypes<typeof xlsxEditorMenuGroupProps>;

/**
 * A named section of a panel: a visible heading and the rows under it, as one `role="group"`
 * whose accessible name is the heading.
 *
 * @public
 */
export const XlsxEditorMenuGroup = defineComponent({
  name: 'XlsxEditorMenuGroup',
  props: xlsxEditorMenuGroupProps,
  setup(props, { slots }) {
    const { t } = useTranslation();
    return () => {
      if (props.hidden) return null;
      const heading = props.label ?? (props.labelKey ? translateKey(t, props.labelKey) : undefined);
      return (
        <div
          role="group"
          {...(heading ? { 'aria-label': heading } : {})}
          class={['xlsx-menu-group', props.className]}
        >
          {heading ? (
            <span class="xlsx-menu-heading" aria-hidden="true">
              {heading}
            </span>
          ) : null}
          {slots.default?.()}
        </div>
      );
    };
  },
});

const xlsxEditorMenuSeparatorProps = {
  /** Appended after the `xlsx-menu-separator` class. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/** Props of `XlsxEditorMenu.Separator`. @public */
export type XlsxEditorMenuSeparatorProps = ExtractPublicPropTypes<
  typeof xlsxEditorMenuSeparatorProps
>;

/** A rule between groups of rows. @public */
export const XlsxEditorMenuSeparator = defineComponent({
  name: 'XlsxEditorMenuSeparator',
  props: xlsxEditorMenuSeparatorProps,
  setup(props) {
    return () => <div role="separator" class={['xlsx-menu-separator', props.className]} />;
  },
});

const SUBMENU_CARET = ['M504-480 320-664l56-56 240 240-240 240-56-56 184-184Z'];

const xlsxEditorMenuSubmenuProps = {
  /** The label, as you write it. Wins over `labelKey`. */
  label: { type: String, default: undefined },
  /** A catalog key for the label. */
  labelKey: { type: String, default: undefined },
  /** Material Symbols path data of the row's icon (view box `0 -960 960 960`). */
  paths: { type: Array as PropType<readonly string[] | null>, default: undefined },
  /** Appended after the `xlsx-menu-sub` class. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/** Props of `XlsxEditorMenu.Submenu`. @public */
export type XlsxEditorMenuSubmenuProps = ExtractPublicPropTypes<typeof xlsxEditorMenuSubmenuProps>;

/**
 * A row that opens a nested panel beside it. Opening a panel is no command, so the row stays
 * enabled, and each row of the panel answers for itself. It opens on a press, on hover, and
 * with ArrowRight (ArrowLeft right to left), Enter, or Space.
 *
 * @public
 */
export const XlsxEditorMenuSubmenu = defineComponent({
  name: 'XlsxEditorMenuSubmenu',
  props: xlsxEditorMenuSubmenuProps,
  setup(props, { slots }) {
    const { t } = useTranslation();
    const bar = inject(menuBarKey, null);
    const open = shallowRef(false);
    const panelId = useId();
    const trigger = shallowRef<HTMLButtonElement | null>(null);
    const focusEdge: { current: 'first' | 'last' | null } = { current: null };
    const close = (): void => {
      open.value = false;
      trigger.value?.focus();
    };
    return () => {
      const label = props.label ?? (props.labelKey ? translateKey(t, props.labelKey) : '');
      const openKey = bar?.rtl.value ? 'ArrowLeft' : 'ArrowRight';
      return (
        <div
          role="none"
          class={['xlsx-menu-sub', props.className]}
          onMouseenter={() => (open.value = true)}
          onMouseleave={() => (open.value = false)}
        >
          <button
            ref={trigger}
            type="button"
            role="menuitem"
            tabindex={-1}
            aria-haspopup="menu"
            aria-expanded={open.value}
            aria-controls={open.value ? panelId : undefined}
            class="xlsx-menu-row"
            onClick={() => (open.value = !open.value)}
            onKeydown={(e: KeyboardEvent) => {
              if (e.key === openKey || e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                e.stopPropagation();
                focusEdge.current = 'first';
                open.value = true;
              }
            }}
          >
            <span class="xlsx-menu-icon" aria-hidden="true">
              <ChromeIcon paths={props.paths ?? null} />
            </span>
            <span class="xlsx-menu-label">{label}</span>
            <ChromeIcon paths={SUBMENU_CARET} className="xlsx-menu-caret" />
          </button>
          {/* Mounted while closed, and hidden, so its rows keep their state between opens. */}
          <MenuPanel
            id={panelId}
            label={label}
            submenu
            hidden={!open.value}
            focusEdge={focusEdge}
            onClose={close}
          >
            {{ default: () => slots.default?.() }}
          </MenuPanel>
        </div>
      );
    };
  },
});

const xlsxEditorMenuEntryProps = {
  /** The registry entry (`CHROME_MENUS`) to render as its row. */
  entry: { type: Object as PropType<ChromeMenuEntry>, required: true },
} as const;

/** Props of `XlsxEditorMenu.Entry`: one registry entry as its row. @public */
export type XlsxEditorMenuEntryProps = ExtractPublicPropTypes<typeof xlsxEditorMenuEntryProps>;

/** One registry entry (`CHROME_MENUS`) as its row, for a host that arranges registry data. @public */
export const XlsxEditorMenuEntry = defineComponent({
  name: 'XlsxEditorMenuEntry',
  props: xlsxEditorMenuEntryProps,
  setup(props) {
    return () => renderRows([props.entry], new Map());
  },
});

/** The registry rows, each replaced by its host row where one names it. */
function renderRows(
  entries: readonly ChromeMenuEntry[],
  overrides: ReadonlyMap<string, VNode | null>
): VNode[] {
  return entries.map((entry, index) => {
    if (entry.kind === 'separator') return <XlsxEditorMenuSeparator key={`sep:${index}`} />;
    if (entry.kind === 'submenu')
      return (
        <XlsxEditorMenuSubmenu
          key={`sub:${entry.labelKey}`}
          labelKey={entry.labelKey}
          paths={entry.paths}
        >
          {{ default: () => renderRows(entry.items, overrides) }}
        </XlsxEditorMenuSubmenu>
      );
    const key = chromeMenuRowKey(entry);
    const override = overrides.get(key);
    return h(Fragment, { key: `row:${key}` }, [
      override ??
        h(XlsxEditorMenuItem, {
          slotId: entry.slot,
          ...(entry.labelKey ? { labelKey: entry.labelKey } : {}),
          ...(entry.shortcut ? { shortcut: entry.shortcut } : {}),
        }),
    ]);
  });
}

/** The rows of a panel: the registry rows with the host's rows merged, or only the host's. */
export function panelRows(
  entries: readonly ChromeMenuEntry[] | undefined,
  children: unknown,
  preset: boolean,
  title: string
): unknown {
  const unmatched = (slotId: string) =>
    warnToolbarProblems([menuHiddenRowProblem(slotId, title, VUE_MENU_PART_NAMES)]);
  if (!preset || !entries) {
    // With no registry rows, a hidden slot row removes nothing.
    for (const node of visibleNodes(children)) {
      const row = slotRowOf(node);
      if (row?.hidden) unmatched(row.slotId);
    }
    return children;
  }
  const arranged = arrangeMenuRows(entries, visibleNodes(children), slotRowOf, unmatched);
  return [...renderRows(arranged.entries, arranged.overrides), ...arranged.extras];
}

// ---------------------------------------------------------------------------------------------
// Panels and menus

/** The rows that a panel moves through. */
function rowsOf(panel: HTMLElement): HTMLElement[] {
  return [...panel.querySelectorAll<HTMLElement>(PANEL_ROWS)];
}

/** Focus the first or the last row of a panel. */
export function focusPanelEdge(panel: HTMLElement | null, edge: 'first' | 'last'): void {
  if (!panel) return;
  const rows = rowsOf(panel);
  (edge === 'first' ? rows[0] : rows.at(-1))?.focus();
}

/** One panel: `role="menu"` with the keys of the WAI-ARIA menu pattern. */
export const MenuPanel = defineComponent({
  name: 'XlsxMenuPanel',
  props: {
    id: { type: String, required: true },
    label: { type: String, required: true },
    submenu: { type: Boolean, default: false },
    /** The panel is closed: it stays mounted, hidden. */
    hidden: { type: Boolean, default: false },
    /** Set when the panel opens from the keyboard: that row takes the focus when it shows. */
    focusEdge: {
      type: Object as PropType<{ current: 'first' | 'last' | null }>,
      default: undefined,
    },
    /** For a bar menu's panel: the menu, whose keyboard open the bar asks for. */
    menuId: { type: String, default: undefined },
    /** The owner places the panel, as the context menu does at its point (`placeContextMenu`). */
    placed: { type: Boolean, default: false },
    onClose: { type: Function as PropType<() => void>, required: true },
  },
  setup(props, { slots }) {
    const bar = inject(menuBarKey, null);
    provide(panelKey, props.submenu);
    const panel = shallowRef<HTMLDivElement | null>(null);
    // When the panel shows, its rows are in place: a keyboard open moves the focus to one.
    watch(
      [panel, () => props.hidden],
      ([node, hidden]) => {
        if (!node || hidden) return;
        const edge = props.focusEdge?.current;
        if (edge) {
          props.focusEdge!.current = null;
          focusPanelEdge(node, edge);
        }
        const request = bar?.focusRequest.current;
        if (request && request.id === props.menuId) {
          bar!.focusRequest.current = null;
          focusPanelEdge(node, request.edge);
        }
      },
      { immediate: true, flush: 'post' }
    );
    // While the panel shows, it stays inside the viewport (`placeMenuPanel`).
    watch(
      [panel, () => props.hidden],
      ([node, hidden], _old, onCleanup) => {
        if (!node || hidden || props.placed) return;
        onCleanup(
          placeMenuPanel(node, {
            submenu: props.submenu,
            ...(props.menuId === MORE_MENU_ID ? { prefer: 'end' as const } : {}),
          })
        );
      },
      { immediate: true, flush: 'post' }
    );
    const onKeydown = (e: KeyboardEvent): void => {
      const el = panel.value;
      if (!el) return;
      const rows = rowsOf(el);
      const active = el.ownerDocument.activeElement as HTMLElement | null;
      const at = active ? rows.indexOf(active) : -1;
      const action = menuPanelKey(e, {
        rtl: bar?.rtl.value ?? false,
        submenu: props.submenu,
        onSubmenuRow: active?.getAttribute('aria-haspopup') === 'menu',
      });
      if (!action) return;
      if (action.kind === 'leave') {
        bar?.closeAll('none');
        return;
      }
      e.preventDefault();
      e.stopPropagation();
      switch (action.kind) {
        case 'step':
          rows[menuStep(rows.length, at, action.by)]?.focus();
          break;
        case 'edge':
          focusPanelEdge(el, action.edge);
          break;
        case 'typeahead': {
          const labels = rows.map((r) => r.querySelector('.xlsx-menu-label')?.textContent ?? '');
          const next = menuTypeahead(labels, at, action.char);
          if (next >= 0) rows[next]?.focus();
          break;
        }
        case 'close':
          props.onClose();
          break;
        case 'bar': {
          if (!bar) break;
          const order = bar.order.value;
          const from = order.indexOf(bar.openMenu.value ?? '');
          const next = order[menuStep(order.length, from, action.by)];
          if (next) bar.open(next, 'first');
          break;
        }
        case 'openSubmenu':
          active?.click();
          break;
      }
    };
    return () => (
      <div
        ref={panel}
        id={props.id}
        hidden={props.hidden}
        role="menu"
        aria-label={props.label}
        class={['xlsx-menu-panel', props.submenu ? 'xlsx-menu-panel-sub' : null]}
        data-popup=""
        onKeydown={onKeydown}
      >
        {slots.default?.()}
      </div>
    );
  },
});

const xlsxEditorMenuMenuProps = {
  /** The menu: a `ChromeMenuId` changes that built-in menu; another id adds a host menu. */
  id: { type: String as PropType<ChromeMenuId | (string & {})>, required: true },
  /** The trigger label, as you write it. Wins over `labelKey`. */
  label: { type: String, default: undefined },
  /** A catalog key for the trigger label. The default is the registry's label. */
  labelKey: { type: String, default: undefined },
  /** Appended after the `xlsx-menu` class. A host `class` also works. */
  className: { type: String, default: undefined },
  /** Removes the menu from the bar. */
  hidden: { type: Boolean, default: undefined },
  /**
   * `false` renders only the default slot as the panel. The default is `true`: the registry rows
   * of the menu, with a row that names a registry row replacing it in place, and the other rows
   * after the registry rows.
   */
  preset: { type: Boolean, default: true },
  /** For a host menu: the menu of the bar that it follows. Without it, the host menu goes last. */
  after: { type: String, default: undefined },
} as const;

/** Props of `XlsxEditorMenu.Menu`, the twin of React's `XlsxEditorMenuMenuProps`. @public */
export type XlsxEditorMenuMenuProps = ExtractPublicPropTypes<typeof xlsxEditorMenuMenuProps>;

/** A menu in the bar: its trigger and, while open, its panel. Also the "More" menu. */
export const BarMenu = defineComponent({
  name: 'XlsxBarMenu',
  props: {
    id: { type: String, required: true },
    text: { type: String, required: true },
    iconOnly: { type: Boolean, default: false },
    className: { type: String, default: undefined },
  },
  setup(props, { slots }) {
    const bar = inject(menuBarKey, null);
    const panelId = useId();
    const onKeydown = (e: KeyboardEvent): void => {
      if (!bar) return;
      const open = bar.openMenu.value === props.id;
      // Escape closes an open menu on its trigger; on a closed trigger, it leaves the bar for
      // the grid, as Escape leaves the toolbar.
      if (e.key === 'Escape') {
        e.preventDefault();
        bar.closeAll(open ? 'trigger' : 'grid');
        return;
      }
      const action = menuBarTriggerKey(e.key, bar.rtl.value);
      if (!action) return;
      e.preventDefault();
      if (action.kind === 'open') {
        bar.open(props.id, action.edge);
        return;
      }
      const order = bar.order.value;
      const target =
        action.kind === 'edge'
          ? order[action.edge === 'first' ? 0 : order.length - 1]
          : order[menuStep(order.length, order.indexOf(props.id), action.by)];
      if (!target) return;
      if (bar.openMenu.value !== null) bar.open(target, 'first');
      else {
        bar.setActiveMenu(target);
        bar.triggerOf(target)?.focus();
      }
    };
    return () => {
      const open = bar?.openMenu.value === props.id;
      const tabStop = bar ? bar.activeMenu.value === props.id : true;
      const icon = slots.icon?.();
      return (
        <div
          role="none"
          class={['xlsx-menu', props.className]}
          {...(props.id === MORE_MENU_ID ? { 'data-menu-more': '' } : { 'data-menu': props.id })}
        >
          <button
            type="button"
            role="menuitem"
            aria-haspopup="menu"
            aria-expanded={open}
            aria-controls={open ? panelId : undefined}
            tabindex={tabStop ? 0 : -1}
            class="xlsx-menu-trigger"
            {...(props.iconOnly ? { 'aria-label': props.text, title: props.text } : {})}
            {...(open ? { 'data-open': '' } : {})}
            onMousedown={guardToolbarMousedown}
            onClick={() => bar?.open(open ? null : props.id)}
            onMouseenter={() => {
              if (bar && bar.openMenu.value !== null && bar.openMenu.value !== props.id)
                bar.open(props.id);
            }}
            onFocus={() => bar?.setActiveMenu(props.id)}
            onKeydown={onKeydown}
          >
            {icon && visibleNodes(icon).length > 0 ? (
              <span class="xlsx-menu-trigger-icon" aria-hidden="true">
                {icon}
              </span>
            ) : null}
            {props.iconOnly ? null : props.text}
          </button>
          {open ? (
            <MenuPanel
              id={panelId}
              label={props.text}
              menuId={props.id}
              onClose={() => bar?.closeAll('trigger')}
            >
              {{ default: () => slots.default?.() }}
            </MenuPanel>
          ) : null}
        </div>
      );
    };
  },
});

/**
 * One menu of the bar: a trigger and the panel it opens. A press opens it and a second press
 * closes it. While a menu is open, the pointer over another trigger opens that menu. An icon
 * before the trigger label goes in the `icon` slot.
 *
 * @public
 */
export const XlsxEditorMenuMenu = defineComponent({
  name: 'XlsxEditorMenuMenu',
  props: xlsxEditorMenuMenuProps,
  setup(props, { slots, attrs }) {
    const placement = inject(
      menuPlacementKey,
      computed(() => null)
    );
    const { t } = useTranslation();
    return () => {
      if (props.hidden) return null;
      const where = placement.value;
      if (where && !where.ids.has(props.id)) return null;
      const registry = CHROME_MENUS.find((m) => m.id === props.id);
      const label =
        props.label ?? translateKey(t, props.labelKey ?? registry?.labelKey ?? props.id);
      const rows = () =>
        panelRows(registry?.entries, slots.default?.(), props.preset !== false, label);
      if (where?.kind === 'more')
        return <XlsxEditorMenuSubmenu label={label}>{{ default: rows }}</XlsxEditorMenuSubmenu>;
      const className = [props.className, text(attrs.class)].filter(Boolean).join(' ');
      return (
        <BarMenu id={props.id} text={label} {...(className ? { className } : {})}>
          {{ default: rows, ...(slots.icon ? { icon: slots.icon } : {}) }}
        </BarMenu>
      );
    };
  },
  inheritAttrs: false,
});

markPart(XlsxEditorMenuMenu, 'xlsxMenuGeneric');

const xlsxEditorMenuPartProps = {
  label: xlsxEditorMenuMenuProps.label,
  labelKey: xlsxEditorMenuMenuProps.labelKey,
  className: xlsxEditorMenuMenuProps.className,
  hidden: xlsxEditorMenuMenuProps.hidden,
  preset: xlsxEditorMenuMenuProps.preset,
  after: xlsxEditorMenuMenuProps.after,
} as const;

/** Props of a pinned menu (`XlsxEditorMenu.File`): `Menu.Menu` without its id. @public */
export type XlsxEditorMenuPartProps = ExtractPublicPropTypes<typeof xlsxEditorMenuPartProps>;

function defineMenu(id: ChromeMenuId): XlsxEditorMenuPart {
  const Part = defineComponent({
    name: `XlsxEditorMenu${id[0]!.toUpperCase()}${id.slice(1)}`,
    props: xlsxEditorMenuPartProps,
    setup(props, { slots, attrs }) {
      return () => h(XlsxEditorMenuMenu, { ...attrs, ...props, id }, slots);
    },
    inheritAttrs: false,
  });
  markPart(Part, 'xlsxMenu', id);
  return Part as unknown as XlsxEditorMenuPart;
}

/** A menu pinned to one registry id (`XlsxEditorMenu.File`). @public */
export type XlsxEditorMenuPart = DefineComponent<XlsxEditorMenuPartProps>;

export const MenuFile = defineMenu('file');
export const MenuEdit = defineMenu('edit');
export const MenuView = defineMenu('view');
export const MenuInsert = defineMenu('insert');
export const MenuFormat = defineMenu('format');

/** The "More" menu: the menus that do not fit, each as a submenu. */
export const MoreMenu = defineComponent({
  name: 'XlsxMenuMore',
  setup(_props, { slots }) {
    const { t } = useTranslation();
    return () => (
      <BarMenu id={MORE_MENU_ID} text={t('menu.more')} iconOnly>
        {{
          default: () => slots.default?.(),
          icon: () => <ChromeIcon paths={ICON_MORE} />,
        }}
      </BarMenu>
    );
  },
});

/** What the bar reads from a menu vnode: its id and props, or null for any other vnode. */
export function menuSpecOf(
  node: unknown
): { id: string; props: Record<string, unknown>; node: VNode } | null {
  if (!isVNode(node)) return null;
  const props = vnodeProps(node);
  const pinned = markerOf(node, 'xlsxMenu');
  if (typeof pinned === 'string') return { id: pinned, props, node };
  if (markerOf(node, 'xlsxMenuGeneric') === true && typeof props.id === 'string')
    return { id: props.id, props, node };
  return null;
}

export { text as menuText };
