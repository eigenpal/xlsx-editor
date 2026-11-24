/** @jsxImportSource react */
// The parts of the menu bar: one menu of the bar (`XlsxEditor.Menu.Menu`, and the pinned
// `.File`, `.Edit`, `.View`, `.Insert`, and `.Format`), and the rows of a panel (`.Item`, the named
// rows such as `.Cut`, `.Row`, `.Group`, `.Separator`, `.Submenu`, and `.Entry`). After
// docx-editor/packages/react/src/editor/menu/parts.tsx at ad686574 (EigenPal): a menu is a
// trigger and the panel it opens; a panel shows the registry rows of its menu
// (`CHROME_MENUS`), with each row child that names a registry row replacing it in place
// (`hidden` removes it) and the other children after the registry rows, and `preset={false}`
// renders only the children; a row of a slot reads its state from `toolbarCommandState`; a
// disabled row stays focusable with its reason; and the bar is one tab stop. The row rules are
// core's (`arrangeMenuRows`), as in PPTX Editor's menu parts
// (pptx-editor/packages/react/src/menu-parts.tsx at 526d3b23).
//
// Differences from DOCX Editor:
// - The menus, the rows, and the named rows of a workbook editor (`MENU_NAMED_PARTS` in core).
// - A row's shortcut is a key chord that the bar shows for the platform (`menuShortcutText`).
// - The keys of a panel are core's (`menuPanelKey`), with type-ahead, and ArrowRight and
//   ArrowLeft move to the next and the previous menu of the bar, as the WAI-ARIA menubar pattern
//   asks.
// - A disabled row gives its reason as its title and its accessible description, not as body
//   text.
// - `Menu.Menu` takes `after`, which places a host menu after a menu of the bar, as PPTX
//   Editor's `Menu.Entry` does. `.Entry` is DOCX Editor's: one registry entry as its row.
// - The markers are set with `markPart`, so they are no public member of the parts.

import {
  Children,
  Fragment,
  cloneElement,
  createContext,
  isValidElement,
  useContext,
  useId,
  useLayoutEffect,
  useRef,
  useState,
  type ReactElement,
  type ReactNode,
  type KeyboardEvent as ReactKeyboardEvent,
  type MutableRefObject,
} from 'react';
import {
  CHROME_MENUS,
  MENU_NAMED_PARTS,
  arrangeMenuRows,
  placeMenuPanel,
  MENU_MORE_ID,
  menuHiddenRowProblem,
  REACT_MENU_PART_NAMES,
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
  submenuRowSlots,
  focusMenuPanelEdge,
  followPressedTrigger,
  menuPanelRows,
  menuRowLabels,
  type ChromeMenuEntry,
  type ChromeMenuId,
  type ChromeSlotId,
  type MenuRowPartName,
} from '@xlsx-editor.dev/core/editor';
import { useTranslation, useXlsxEditor } from '../context.ts';
import { ChromeIcon, translateKey, useRowsState, useToolbarSlot } from '../toolbar/slot.tsx';
import { useNumberFormatDialogOpener } from '../toolbar/NumberFormatDialog.tsx';
import { warnToolbarProblems } from '../toolbar/host-parts.tsx';
import { ICON_MORE } from '../toolbar/icons.ts';

/** The id of the "More" menu, which holds the menus that do not fit. Never a host menu id. */
export const MORE_MENU_ID = MENU_MORE_ID;

/** What the menu bar gives its menus. */
export interface MenuBarState {
  readonly openMenu: string | null;
  /** Open a menu (null closes every menu). `edge` moves the keyboard to its first or last row. */
  readonly open: (id: string | null, edge?: 'first' | 'last') => void;
  /** The pending keyboard move into the panel that opens next. */
  readonly focusRequest: MutableRefObject<{ id: string; edge: 'first' | 'last' } | null>;
  /** The menu whose trigger is the bar's one tab stop. */
  readonly activeMenu: string | null;
  readonly setActiveMenu: (id: string) => void;
  /** The ids of the triggers in the bar, in order: for the arrow keys. */
  readonly order: readonly string[];
  readonly rtl: boolean;
  readonly apple: boolean;
  /**
   * Close every menu, then move the keyboard to the trigger, to the grid, or nowhere. `origin`
   * moves it to the trigger only when the bar held the keyboard as the menu opened (a keyboard
   * open); after a pointer open, the keyboard stays where it was, such as in the grid.
   */
  readonly closeAll: (focus: 'trigger' | 'grid' | 'none' | 'origin') => void;
  /** The trigger of a menu in the bar. */
  readonly triggerOf: (id: string) => HTMLElement | null;
}

export const MenuBarContext = createContext<MenuBarState | null>(null);

/** Where a menu renders: in the bar, or as a submenu of the "More" menu. */
export type MenuPlacement =
  | { readonly kind: 'bar'; readonly ids: ReadonlySet<string> }
  | { readonly kind: 'more'; readonly ids: ReadonlySet<string> };

export const MenuPlacementContext = createContext<MenuPlacement | null>(null);

/** The panel that holds a row: a submenu panel or a bar menu panel. */
interface PanelState {
  readonly submenu: boolean;
}

const PanelContext = createContext<PanelState>({ submenu: false });

/** Set a static marker that the bar and the panels read. Not a public member. */
function markPart(part: object, marker: string, value: unknown = true): void {
  (part as Record<string, unknown>)[marker] = value;
}

/** The marker of a part, or undefined. */
function markerOf(child: ReactNode, marker: string): unknown {
  if (!isValidElement(child)) return undefined;
  const type = child.type as unknown;
  if (typeof type !== 'function' && typeof type !== 'object') return undefined;
  return (type as Record<string, unknown>)[marker];
}

/** The children as a flat list, every Fragment read as its children, keys kept apart. */
export function flatMenuChildren(children: ReactNode, prefix = ''): ReactNode[] {
  return Children.toArray(children).flatMap((child) => {
    if (!isValidElement(child)) return child === '' ? [] : [child];
    if (child.type === Fragment)
      return flatMenuChildren(
        (child.props as { children?: ReactNode }).children,
        `${prefix}${String(child.key)}/`
      );
    return [prefix ? cloneElement(child, { key: `${prefix}${String(child.key)}` }) : child];
  });
}

// ---------------------------------------------------------------------------------------------
// Rows

/** Props for `XlsxEditor.Menu.Row`: one row of the host's own, not a chrome slot. @public */
export interface XlsxEditorMenuRowProps {
  /** The text of the row. Without it, the children are the text. */
  readonly label?: string;
  /** An icon before the label, hidden from assistive technology. */
  readonly icon?: ReactNode;
  /** The shortcut text at the end of the row, as you write it. */
  readonly shortcut?: string;
  /** The row is disabled: it stays focusable, and a press does nothing. */
  readonly disabled?: boolean;
  /** Why the row is disabled: its title and its accessible description. */
  readonly disabledReason?: string;
  /** The row shows a checked state (`menuitemcheckbox`, or `menuitemradio` with `selected`). */
  readonly active?: boolean;
  /**
   * The row is one of a one-of set, such as one alignment of several: its role is
   * `menuitemradio`, and `active` marks the one value. As DOCX Editor's prop of the same name.
   */
  readonly selected?: boolean;
  /** Runs on a press of an enabled row, after the menu closes. */
  readonly onSelect?: () => void;
  /** Appended after the `xlsx-menu-row` class. */
  readonly className?: string;
  readonly children?: ReactNode;
}

/** The markup of one row, for a slot row and a host row. */
function RowButton(props: {
  readonly label: ReactNode;
  readonly icon: ReactNode;
  readonly shortcut: string | undefined;
  readonly disabled: boolean;
  readonly reason: string | null;
  readonly title: string | undefined;
  readonly checked: boolean | undefined;
  /** A row of a one-of set: `menuitemradio` in place of `menuitemcheckbox`. */
  readonly radio?: boolean;
  readonly slot?: string;
  readonly className?: string;
  readonly onPress: () => void;
}) {
  const reasonId = useId();
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
      tabIndex={-1}
      className={`xlsx-menu-row${props.className ? ` ${props.className}` : ''}`}
      {...(props.checked === undefined ? {} : { 'aria-checked': props.checked })}
      {...(props.disabled ? { 'aria-disabled': true, 'data-disabled': '' } : {})}
      {...(props.reason ? { 'aria-describedby': reasonId } : {})}
      {...(props.title ? { title: props.title } : {})}
      {...(props.slot ? { 'data-slot': props.slot } : {})}
      onClick={() => {
        if (!props.disabled) props.onPress();
      }}
    >
      <span className="xlsx-menu-icon" aria-hidden="true">
        {props.icon}
      </span>
      <span className="xlsx-menu-label">{props.label}</span>
      {props.shortcut ? (
        <span className="xlsx-menu-shortcut" aria-hidden="true">
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

/**
 * One row of the host's own: a label, an icon, a shortcut text, and an action. Selecting it
 * closes the menu, then runs `onSelect`.
 *
 * @public
 */
export function XlsxEditorMenuRow(props: XlsxEditorMenuRowProps) {
  const bar = useContext(MenuBarContext);
  const disabled = props.disabled === true;
  const reason = disabled ? (props.disabledReason ?? null) : null;
  const text = props.label ?? props.children;
  return (
    <RowButton
      label={text}
      icon={props.icon ?? null}
      shortcut={props.shortcut}
      disabled={disabled}
      reason={reason}
      title={reason && typeof text === 'string' ? `${text}: ${reason}` : (reason ?? undefined)}
      checked={props.active === undefined ? (props.selected ? false : undefined) : props.active}
      radio={props.selected === true}
      {...(props.className ? { className: props.className } : {})}
      onPress={() => {
        // After a keyboard open, the keyboard goes back to the menu's trigger as the panel
        // closes; after a pointer open, it stays where it was. A command that moves it (to the
        // grid, or into a dialog) then takes it from there.
        bar?.closeAll('origin');
        props.onSelect?.();
      }}
    />
  );
}

/**
 * Props for `XlsxEditor.Menu.Item`: one chrome slot as a menu row. An interface, so a host can
 * extend it.
 *
 * @public
 */
export interface XlsxEditorMenuItemProps {
  /** The slot of the row, such as `'edit.copy'`. */
  readonly slotId: ChromeSlotId;
  /** A catalog key for the label in place of the slot's own label. */
  readonly labelKey?: string;
  /** A label in place of the slot's own label. Wins over `labelKey`. */
  readonly label?: string;
  /** The key chord of the row, such as `'Mod+C'`, shown for the platform. */
  readonly shortcut?: string;
  /** Appended after the `xlsx-menu-row` class. */
  readonly className?: string;
  /** Inside a panel, removes the registry row of the same slot. */
  readonly hidden?: boolean;
}

/**
 * One chrome slot as a live row: its label and icon from the registry, its state from
 * `toolbarCommandState`, and its command. Selecting it closes the menu and runs the slot.
 *
 * @public
 */
export function XlsxEditorMenuItem(props: XlsxEditorMenuItemProps) {
  if (props.hidden) return null;
  return <SlotRow {...props} />;
}

markPart(XlsxEditorMenuItem, 'xlsxMenuItem');

function SlotRow({ slotId, labelKey, label, shortcut, className }: XlsxEditorMenuItemProps) {
  const bar = useContext(MenuBarContext);
  const editor = useXlsxEditor();
  const { t } = useTranslation();
  const s = useToolbarSlot(slotId);
  const opener = useNumberFormatDialogOpener();
  const text = label ?? (labelKey ? translateKey(t, labelKey) : s.label);
  const chord = shortcut ?? chromeMenuShortcut(slotId);
  return (
    <RowButton
      label={text}
      icon={<ChromeIcon paths={s.control.paths} />}
      shortcut={
        chord
          ? menuShortcutText(
              chord,
              bar?.apple ?? false,
              menuKeyWords((k) => translateKey(t, k))
            )
          : undefined
      }
      disabled={!s.isEnabled}
      reason={s.description}
      title={s.title}
      checked={s.control.toggle ? s.isActive : undefined}
      radio={menuRowRole(slotId, s.control.toggle === true) === 'menuitemradio'}
      slot={slotId}
      {...(className ? { className } : {})}
      onPress={() => {
        // A row that keeps the keyboard (Open, Custom number format) gives it back to the trigger
        // after a keyboard open first, so a closed picker or dialog leaves it there; every other row gives it to the
        // grid after its command.
        bar?.closeAll(menuRowKeepsFocus(slotId) ? 'origin' : 'none');
        if (slotId === 'number.custom') {
          if (!opener?.open(null)) editor?.focus();
          return;
        }
        s.execute();
        if (!menuRowKeepsFocus(slotId)) editor?.focus();
      }}
    />
  );
}

/** Props for a named row (`XlsxEditor.Menu.Cut`): `Menu.Item` without its slot. @public */
export type XlsxEditorMenuNamedRowProps = Omit<XlsxEditorMenuItemProps, 'slotId'>;

/** A named row part, which stands for `Menu.Item` with its slot. */
export type XlsxEditorMenuNamedRow = (props: XlsxEditorMenuNamedRowProps) => ReactNode;

function defineNamedRow(name: MenuRowPartName): XlsxEditorMenuNamedRow {
  const slot = MENU_NAMED_PARTS[name].slot;
  const Part = (props: XlsxEditorMenuNamedRowProps) => (
    <XlsxEditorMenuItem {...props} slotId={slot} />
  );
  markPart(Part, 'xlsxMenuRowSlot', slot);
  return Part;
}

/** The named rows, one per row of core's `MENU_NAMED_PARTS`. */
export const MENU_NAMED_ROWS = Object.fromEntries(
  (Object.keys(MENU_NAMED_PARTS) as MenuRowPartName[]).map((name) => [name, defineNamedRow(name)])
) as Record<MenuRowPartName, XlsxEditorMenuNamedRow>;

/** The slot row that a child stands for, or null for any other child. */
export function slotRowOf(child: ReactNode): { slotId: string; hidden?: boolean } | null {
  if (!isValidElement(child)) return null;
  const props = child.props as { slotId?: unknown; hidden?: unknown };
  const named = markerOf(child, 'xlsxMenuRowSlot');
  if (typeof named === 'string') return { slotId: named, hidden: props.hidden === true };
  if (markerOf(child, 'xlsxMenuItem') === true && typeof props.slotId === 'string')
    return { slotId: props.slotId, hidden: props.hidden === true };
  return null;
}

/** Props for `XlsxEditor.Menu.Group`: a titled section of rows. @public */
export interface XlsxEditorMenuGroupProps {
  /** The heading, as you write it. Wins over `labelKey`. */
  readonly label?: string;
  /** A catalog key for the heading. */
  readonly labelKey?: string;
  /** Appended after the `xlsx-menu-group` class. */
  readonly className?: string;
  /** Renders nothing. */
  readonly hidden?: boolean;
  readonly children?: ReactNode;
}

/**
 * A named section of a panel: a visible heading and the rows under it, as one `role="group"`
 * whose accessible name is the heading.
 *
 * @public
 */
export function XlsxEditorMenuGroup(props: XlsxEditorMenuGroupProps) {
  const { t } = useTranslation();
  if (props.hidden) return null;
  const heading = props.label ?? (props.labelKey ? translateKey(t, props.labelKey) : undefined);
  return (
    <div
      role="group"
      {...(heading ? { 'aria-label': heading } : {})}
      className={`xlsx-menu-group${props.className ? ` ${props.className}` : ''}`}
    >
      {heading ? (
        <span className="xlsx-menu-heading" aria-hidden="true">
          {heading}
        </span>
      ) : null}
      {props.children}
    </div>
  );
}

/** Props for `XlsxEditor.Menu.Separator`. @public */
export interface XlsxEditorMenuSeparatorProps {
  /** Appended after the `xlsx-menu-separator` class. */
  readonly className?: string;
}

/** A rule between groups of rows. @public */
export function XlsxEditorMenuSeparator({ className }: XlsxEditorMenuSeparatorProps) {
  return (
    <div role="separator" className={`xlsx-menu-separator${className ? ` ${className}` : ''}`} />
  );
}

/** Props for `XlsxEditor.Menu.Submenu`: a row that opens a nested panel. @public */
export interface XlsxEditorMenuSubmenuProps {
  /** The label, as you write it. Wins over `labelKey`. */
  readonly label?: string;
  /** A catalog key for the label. */
  readonly labelKey?: string;
  /** Material Symbols path data of the row's icon (view box `0 -960 960 960`). */
  readonly paths?: readonly string[] | null;
  /** Appended after the `xlsx-menu-sub` class. */
  readonly className?: string;
  readonly children?: ReactNode;
}

/**
 * A row that opens a nested panel beside it. Opening a panel is no command, so each row of the
 * panel answers for itself, and the row is disabled, with the reason of its first row, only when
 * every row of the panel is a disabled slot row (`toolbarRowsState`). It opens on a press, on
 * hover, and with ArrowRight (ArrowLeft right to left), Enter, or Space.
 *
 * @public
 */
export function XlsxEditorMenuSubmenu(props: XlsxEditorMenuSubmenuProps) {
  const { t } = useTranslation();
  const bar = useContext(MenuBarContext);
  const [open, setOpen] = useState(false);
  const panelId = useId();
  const trigger = useRef<HTMLButtonElement>(null);
  const focusEdge = useRef<'first' | 'last' | null>(null);
  const text = props.label ?? (props.labelKey ? translateKey(t, props.labelKey) : '');
  const openKey = bar?.rtl ? 'ArrowLeft' : 'ArrowRight';
  const reasonId = useId();
  const rows = useRowsState(
    submenuRowSlots(flatMenuChildren(props.children), slotRowOf, isSeparatorRow)
  );
  const reason = rows.disabledReason;
  const show = (next: boolean) => setOpen(next && rows.isEnabled);
  const close = () => {
    setOpen(false);
    trigger.current?.focus();
  };
  return (
    <div
      role="none"
      className={`xlsx-menu-sub${props.className ? ` ${props.className}` : ''}`}
      onMouseEnter={() => show(true)}
      onMouseLeave={() => setOpen(false)}
    >
      <button
        ref={trigger}
        type="button"
        role="menuitem"
        tabIndex={-1}
        aria-haspopup="menu"
        aria-expanded={open}
        aria-controls={open ? panelId : undefined}
        className="xlsx-menu-row"
        {...(reason ? { 'aria-disabled': true, 'data-disabled': '' } : {})}
        {...(reason
          ? {
              'aria-describedby': reasonId,
              title: t('toolbar.disabledTitle', { label: text, reason }),
            }
          : {})}
        onClick={() => show(!open)}
        onKeyDown={(e) => {
          if (e.key === openKey || e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            e.stopPropagation();
            focusEdge.current = 'first';
            show(true);
          }
        }}
      >
        <span className="xlsx-menu-icon" aria-hidden="true">
          <ChromeIcon paths={props.paths ?? null} />
        </span>
        <span className="xlsx-menu-label">{text}</span>
        <ChromeIcon paths={SUBMENU_CARET} className="xlsx-menu-caret" />
        {reason ? (
          <span id={reasonId} hidden>
            {reason}
          </span>
        ) : null}
      </button>
      {/* Mounted while closed, and hidden, so its rows keep their state between opens. */}
      <MenuPanel
        id={panelId}
        label={text}
        submenu
        hidden={!open}
        focusEdge={focusEdge}
        onClose={close}
      >
        {props.children}
      </MenuPanel>
    </div>
  );
}

/** Whether a child of a panel is a separator, which no submenu state counts. */
const isSeparatorRow = (child: ReactNode): boolean =>
  isValidElement(child) && child.type === XlsxEditorMenuSeparator;

const SUBMENU_CARET = ['M504-480 320-664l56-56 240 240-240 240-56-56 184-184Z'];

/** Props for `XlsxEditor.Menu.Entry`: one registry entry as its row. @public */
export interface XlsxEditorMenuEntryProps {
  readonly entry: ChromeMenuEntry;
}

/** One registry entry (`CHROME_MENUS`) as its row, for a host that arranges registry data. @public */
export function XlsxEditorMenuEntry({ entry }: XlsxEditorMenuEntryProps) {
  return <>{renderRows([entry], new Map())}</>;
}

/** The registry rows, each replaced by its host row where one names it. */
function renderRows(
  entries: readonly ChromeMenuEntry[],
  overrides: ReadonlyMap<string, ReactNode | null>
): ReactNode[] {
  return entries.map((entry, index) => {
    if (entry.kind === 'separator') return <XlsxEditorMenuSeparator key={`sep:${index}`} />;
    if (entry.kind === 'submenu')
      return (
        <XlsxEditorMenuSubmenu
          key={`sub:${entry.labelKey}`}
          labelKey={entry.labelKey}
          paths={entry.paths}
        >
          {renderRows(entry.items, overrides)}
        </XlsxEditorMenuSubmenu>
      );
    const key = chromeMenuRowKey(entry);
    const override = overrides.get(key);
    return (
      <Fragment key={`row:${key}`}>
        {override ?? (
          <XlsxEditorMenuItem
            slotId={entry.slot}
            {...(entry.labelKey ? { labelKey: entry.labelKey } : {})}
            {...(entry.shortcut ? { shortcut: entry.shortcut } : {})}
          />
        )}
      </Fragment>
    );
  });
}

/** The rows of a panel: the registry rows with the host's rows merged, or only the host's. */
export function panelRows(
  entries: readonly ChromeMenuEntry[] | undefined,
  children: ReactNode,
  preset: boolean,
  title: string
): ReactNode {
  const unmatched = (slotId: string) =>
    warnToolbarProblems([menuHiddenRowProblem(slotId, title, REACT_MENU_PART_NAMES)]);
  if (!preset || !entries) {
    // With no registry rows, a hidden slot row removes nothing.
    for (const child of flatMenuChildren(children)) {
      const row = slotRowOf(child);
      if (row?.hidden) unmatched(row.slotId);
    }
    return children;
  }
  const arranged = arrangeMenuRows(entries, flatMenuChildren(children), slotRowOf, unmatched);
  return (
    <>
      {renderRows(arranged.entries, arranged.overrides)}
      {arranged.extras}
    </>
  );
}

// ---------------------------------------------------------------------------------------------
// Panels and menus

/** One panel: `role="menu"` with the keys of the WAI-ARIA menu pattern. */
export function MenuPanel(props: {
  readonly id: string;
  readonly label: string;
  readonly submenu?: boolean;
  /** The panel is closed: it stays mounted, hidden. */
  readonly hidden?: boolean;
  /** Set when the panel opens from the keyboard: that row takes the focus when it shows. */
  readonly focusEdge?: MutableRefObject<'first' | 'last' | null>;
  /** For a bar menu's panel: the menu, whose keyboard open the bar asks for. */
  readonly menuId?: string;
  /** The owner places the panel, as the context menu does at its point (`placeContextMenu`). */
  readonly placed?: boolean;
  readonly onClose: () => void;
  readonly className?: string;
  readonly children: ReactNode;
}) {
  const bar = useContext(MenuBarContext);
  const ref = useRef<HTMLDivElement | null>(null);
  const submenu = props.submenu === true;
  const hidden = props.hidden === true;
  // When the panel shows, its rows are in place: a keyboard open moves the focus to one.
  useLayoutEffect(() => {
    const node = ref.current;
    if (!node || hidden) return;
    const edge = props.focusEdge?.current;
    if (edge) {
      props.focusEdge!.current = null;
      focusMenuPanelEdge(node, edge);
    }
    const request = bar?.focusRequest.current;
    if (request && request.id === props.menuId) {
      bar!.focusRequest.current = null;
      focusMenuPanelEdge(node, request.edge);
    }
  });
  // While the panel shows, it stays inside the viewport (`placeMenuPanel`).
  const prefer = props.menuId === MORE_MENU_ID ? 'end' : undefined;
  const placed = props.placed === true;
  useLayoutEffect(() => {
    const node = ref.current;
    if (!node || hidden || placed) return undefined;
    return placeMenuPanel(node, { submenu, ...(prefer ? { prefer } : {}) });
  }, [hidden, submenu, prefer, placed]);
  const onKeyDown = (e: ReactKeyboardEvent<HTMLDivElement>) => {
    const panel = ref.current;
    if (!panel) return;
    const rows = menuPanelRows(panel);
    const active = panel.ownerDocument.activeElement as HTMLElement | null;
    const at = active ? rows.indexOf(active) : -1;
    const action = menuPanelKey(e, {
      rtl: bar?.rtl ?? false,
      submenu,
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
        focusMenuPanelEdge(panel, action.edge);
        break;
      case 'typeahead': {
        const next = menuTypeahead(menuRowLabels(rows), at, action.char);
        if (next >= 0) rows[next]?.focus();
        break;
      }
      case 'close':
        props.onClose();
        break;
      case 'bar':
        if (!bar) break;
        {
          const order = bar.order;
          const from = order.indexOf(bar.openMenu ?? '');
          const next = order[menuStep(order.length, from, action.by)];
          if (next) bar.open(next, 'first');
        }
        break;
      case 'openSubmenu':
        active?.click();
        break;
    }
  };
  return (
    <PanelContext.Provider value={{ submenu }}>
      <div
        ref={ref}
        id={props.id}
        hidden={hidden}
        role="menu"
        aria-label={props.label}
        className={`xlsx-menu-panel${submenu ? ' xlsx-menu-panel-sub' : ''}${props.className ? ` ${props.className}` : ''}`}
        data-popup=""
        onKeyDown={onKeyDown}
      >
        {props.children}
      </div>
    </PanelContext.Provider>
  );
}

/** Props for `XlsxEditor.Menu.Menu`, one menu of the bar. @public */
export interface XlsxEditorMenuMenuProps {
  /** The menu: a `ChromeMenuId` changes that built-in menu; another id adds a host menu. */
  readonly id: ChromeMenuId | (string & {});
  /** The trigger label, as you write it. Wins over `labelKey`. */
  readonly label?: string;
  /** A catalog key for the trigger label. The default is the registry's label. */
  readonly labelKey?: string;
  /** An icon before the trigger label, hidden from assistive technology. */
  readonly icon?: ReactNode;
  /** Appended after the `xlsx-menu` class. */
  readonly className?: string;
  /** Removes the menu from the bar. */
  readonly hidden?: boolean;
  /**
   * `false` renders only the children as the panel. The default is `true`: the registry rows of
   * the menu, with a row child that names a registry row replacing it in place, and the other
   * children after the registry rows.
   */
  readonly preset?: boolean;
  /** For a host menu: the menu of the bar that it follows. Without it, the host menu goes last. */
  readonly after?: string;
  readonly children?: ReactNode;
}

/** Props for a pinned menu (`XlsxEditor.Menu.File`): `Menu.Menu` without its id. @public */
export type XlsxEditorMenuPartProps = Omit<XlsxEditorMenuMenuProps, 'id'>;

/**
 * One menu of the bar: a trigger and the panel it opens. A press opens it and a second press
 * closes it. While a menu is open, the pointer over another trigger opens that menu.
 *
 * @public
 */
export function XlsxEditorMenuMenu(props: XlsxEditorMenuMenuProps) {
  const placement = useContext(MenuPlacementContext);
  const { t } = useTranslation();
  if (props.hidden) return null;
  if (placement && !placement.ids.has(props.id)) return null;
  const registry = CHROME_MENUS.find((m) => m.id === props.id);
  const text = props.label ?? translateKey(t, props.labelKey ?? registry?.labelKey ?? props.id);
  const rows = panelRows(registry?.entries, props.children, props.preset !== false, text);
  if (placement?.kind === 'more')
    return <XlsxEditorMenuSubmenu label={text}>{rows}</XlsxEditorMenuSubmenu>;
  return (
    <BarMenu
      id={props.id}
      text={text}
      icon={props.icon}
      {...(props.className ? { className: props.className } : {})}
    >
      {rows}
    </BarMenu>
  );
}

markPart(XlsxEditorMenuMenu, 'xlsxMenuGeneric');

/** A menu in the bar: its trigger and, while open, its panel. Also the "More" menu. */
export function BarMenu(props: {
  readonly id: string;
  readonly text: string;
  readonly icon?: ReactNode;
  /** The trigger shows only its icon, and the text is its name. */
  readonly iconOnly?: boolean;
  readonly className?: string;
  readonly children: ReactNode;
}) {
  const bar = useContext(MenuBarContext);
  const panelId = useId();
  const { id } = props;
  const open = bar?.openMenu === id;
  const tabStop = bar ? bar.activeMenu === id : true;
  const onKeyDown = (e: ReactKeyboardEvent<HTMLButtonElement>) => {
    if (!bar) return;
    // Escape closes an open menu on its trigger; on a closed trigger, it leaves the bar for the
    // grid, as Escape leaves the toolbar.
    if (e.key === 'Escape') {
      e.preventDefault();
      bar.closeAll(open ? 'trigger' : 'grid');
      return;
    }
    const action = menuBarTriggerKey(e.key, bar.rtl);
    if (!action) return;
    e.preventDefault();
    if (action.kind === 'open') {
      bar.open(id, action.edge);
      return;
    }
    const order = bar.order;
    const target =
      action.kind === 'edge'
        ? order[action.edge === 'first' ? 0 : order.length - 1]
        : order[menuStep(order.length, order.indexOf(id), action.by)];
    if (!target) return;
    if (bar.openMenu !== null) bar.open(target, 'first');
    else {
      bar.setActiveMenu(target);
      bar.triggerOf(target)?.focus();
    }
  };
  return (
    <div
      role="none"
      className={`xlsx-menu${props.className ? ` ${props.className}` : ''}`}
      {...(id === MORE_MENU_ID ? { 'data-menu-more': '' } : { 'data-menu': id })}
    >
      <button
        type="button"
        role="menuitem"
        aria-haspopup="menu"
        aria-expanded={open}
        aria-controls={open ? panelId : undefined}
        tabIndex={tabStop ? 0 : -1}
        className="xlsx-menu-trigger"
        {...(props.iconOnly ? { 'aria-label': props.text, title: props.text } : {})}
        {...(open ? { 'data-open': '' } : {})}
        onMouseDown={guardToolbarMousedown}
        onClick={(e) => {
          followPressedTrigger(e.currentTarget);
          bar?.open(open ? null : id);
        }}
        onMouseEnter={() => {
          if (bar && bar.openMenu !== null && bar.openMenu !== id) bar.open(id);
        }}
        onFocus={() => bar?.setActiveMenu(id)}
        onKeyDown={onKeyDown}
      >
        {props.icon ? (
          <span className="xlsx-menu-trigger-icon" aria-hidden="true">
            {props.icon}
          </span>
        ) : null}
        {props.iconOnly ? null : props.text}
      </button>
      {open ? (
        <MenuPanel
          id={panelId}
          label={props.text}
          menuId={id}
          onClose={() => bar?.closeAll('trigger')}
        >
          {props.children}
        </MenuPanel>
      ) : null}
    </div>
  );
}

/** The "More" menu: the menus that do not fit, each as a submenu. */
export function MoreMenu({ children }: { readonly children: ReactNode }) {
  const { t } = useTranslation();
  return (
    <BarMenu
      id={MORE_MENU_ID}
      text={t('menu.more')}
      icon={<ChromeIcon paths={ICON_MORE} />}
      iconOnly
    >
      {children}
    </BarMenu>
  );
}

/** A menu pinned to one registry id (`XlsxEditor.Menu.File`). @public */
export type XlsxEditorMenuPart = (props: XlsxEditorMenuPartProps) => ReactNode;

function defineMenu(id: ChromeMenuId): XlsxEditorMenuPart {
  const Part = (props: XlsxEditorMenuPartProps) => <XlsxEditorMenuMenu {...props} id={id} />;
  markPart(Part, 'xlsxMenu', id);
  return Part;
}

export const MenuFile = defineMenu('file');
export const MenuEdit = defineMenu('edit');
export const MenuView = defineMenu('view');
export const MenuInsert = defineMenu('insert');
export const MenuFormat = defineMenu('format');

/** What the bar reads from a menu child: its id and props, or null for any other child. */
export function menuSpecOf(
  child: ReactNode
): { id: string; props: XlsxEditorMenuPartProps; element: ReactElement } | null {
  if (!isValidElement(child)) return null;
  const pinned = markerOf(child, 'xlsxMenu');
  const props = child.props as XlsxEditorMenuMenuProps;
  if (typeof pinned === 'string') return { id: pinned, props, element: child };
  if (markerOf(child, 'xlsxMenuGeneric') === true && typeof props.id === 'string')
    return { id: props.id, props, element: child };
  return null;
}
