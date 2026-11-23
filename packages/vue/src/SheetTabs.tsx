/** @jsxImportSource vue */
// The sheet tabs: one tab for each visible sheet, an add button, a context menu on a tab, and
// inline rename. The Vue twin of the React adapter's `SheetTabs.tsx`; the rules are the same,
// and the header there explains them and names their sources in PPTX Editor. The state of each
// command comes from core's `sheetCommandState`, and the menu, the rename box, and the tab keys
// are core's (`openSheetMenu`, `listenSheetMenuDismiss`, `moveInSheetMenu`, `sheetRenameBox`,
// `tabKey`, `sheetListKey`), so both adapters act the same.
//
// Differences from React: the effects are watchers that run after the render (`flush: 'post'`),
// as React's effects run after the commit; the rename field installs its handle once, on mount,
// because its handlers read the field's text from a ref; and pixel values in `style` carry
// their unit.

import {
  computed,
  defineComponent,
  onBeforeUnmount,
  onMounted,
  shallowRef,
  useId,
  watch,
  type ExtractPublicPropTypes,
  type PropType,
} from 'vue';
import {
  RENAME_COMMIT_EVENT,
  disabledReasonText,
  focusAfterRelease,
  historyShortcutOf,
  inlineStep,
  isRightToLeft,
  keepGridFocus,
  listenSheetMenuDismiss,
  moveInSheetMenu,
  openSheetMenu,
  refusalText,
  sameSlice,
  scrollIntoNearestScroller,
  sheetCommandState,
  sheetListKey,
  sheetMenuLeft,
  sheetRenameBox,
  tabKey,
  tabsSlice,
  type EditorSheet,
  type RenameCommitDetail,
  type SheetCommand,
  type SheetCommandState,
  type SheetRenameBox,
  type ChromeTranslate,
} from '@xlsx-editor.dev/core/editor';
import { useEditorState, useXlsxEditor } from './context.ts';
import { AddIcon } from './icons.tsx';
import { usePartTranslation } from './locale.ts';

const sheetTabsProps = {
  /**
   * The label resolver of the tabs, their menu, and the rename field: a key and the values of its placeholders, as DOCX
   * Editor's `t` prop. A key that it leaves unresolved reads the root's `locale`. Keep it
   * stable, for example from `useChromeTranslate(overrides)`.
   */
  t: { type: Function as PropType<ChromeTranslate>, default: undefined },
} as const;

/** Props for `SheetTabs`, the twin of React's `SheetTabsProps`. @public */
export type SheetTabsProps = ExtractPublicPropTypes<typeof sheetTabsProps>;

/** An open tab menu: the sheet it acts on, and its left edge in the bar, in CSS pixels. */
interface TabMenu {
  readonly sheet: number;
  readonly left: number;
  /** The element that had focus when the menu opened: focus goes back to it on close. */
  readonly back: HTMLElement | null;
  /** Whether a key opened the menu: an action then focuses a tab. */
  readonly keyboard: boolean;
  /** Whether the bar lays out right to left, where the earlier sheet is on the right. */
  readonly rtl: boolean;
}

/** What the tab bar calls on an open rename field. */
interface RenameHandle {
  /** Commit the field as a blur does; false when the store refused the name. */
  commit(): boolean;
  /** End the rename with no write. */
  drop(): void;
}

/** One tab for each visible worksheet, in workbook order, with the sheet actions. @public */
export const SheetTabs = defineComponent({
  name: 'SheetTabs',
  props: sheetTabsProps,
  setup(props) {
    const { t, dir } = usePartTranslation('SheetTabs', () => props.t);
    const editor = useXlsxEditor();
    const snap = useEditorState(tabsSlice, sameSlice);
    const bar = shallowRef<HTMLDivElement | null>(null);
    const list = shallowRef<HTMLDivElement | null>(null);
    const menuRef = shallowRef<HTMLDivElement | null>(null);
    const addRef = shallowRef<HTMLButtonElement | null>(null);
    const menu = shallowRef<TabMenu | null>(null);
    const renaming = shallowRef<SheetRenameBox | null>(null);
    // The sheet whose tab takes focus after the next render (a menu action or a rename ended).
    const focusSheet = shallowRef<number | null>(null);
    const reasonId = useId();
    const viewReasonId = useId();
    const renameReasonId = useId();
    const tabReasonId = useId();
    // The open rename field: `commit` commits it as a blur does (false when the store refused
    // the name), and `drop` ends it with no write.
    const renameField: { current: RenameHandle | null } = { current: null };
    // Set when a press on a tab ended a rename that the store refused: the click, the
    // double-click, and the context menu of that press do nothing, so the reason stays in view.
    let refusedPress = false;
    // Whether the tab list had focus at the last press on a tab: the click focuses that tab.
    let listFocused = false;
    // Set when the store refused a name of the open rename: Escape then clears that reason.
    const renameRefused = shallowRef(false);
    // The last refusal of a rename, so a new rename clears only that reason.
    let renameRefusal: unknown = null;

    // The engine gives the state of each sheet command (`sheetCommandState`). View mode, or a
    // workbook protection that locks the sheet list, disables every sheet edit with its reason.
    const stateOf = (command: SheetCommand, sheet = -1): SheetCommandState =>
      sheetCommandState(snap.value, command, sheet);

    watch(
      focusSheet,
      (sheet) => {
        if (sheet === null) return;
        focusSheet.value = null;
        list.value?.querySelector<HTMLElement>(`[data-sheet-index="${sheet}"]`)?.focus();
      },
      { flush: 'post' }
    );

    // The tab of the shown sheet scrolls into view, so a new, chosen, or renamed sheet is never
    // hidden.
    const listKey = computed(() => sheetListKey(snap.value.sheets));
    watch(
      [() => snap.value.activeSheet, listKey],
      ([active]) => {
        // Only the tab list scrolls, never the host page (PPTX Editor
        // `scrollIntoNearestScroller`).
        const tab = list.value?.querySelector<HTMLElement>(`[data-sheet-index="${active}"]`);
        if (tab) scrollIntoNearestScroller(tab, list.value);
      },
      { flush: 'post', immediate: true }
    );

    // A change of the sheet list ends a rename: its position and its tab can be gone. A field
    // that held focus gives it to the grid, not to the page.
    watch(
      [listKey, editor],
      () => {
        // The field writes nothing when the focus change below blurs it.
        renameField.current?.drop();
        const field = bar.value?.ownerDocument.activeElement;
        if (field && bar.value?.contains(field) && field.classList.contains('xlsx-sheet-rename'))
          editor.value?.focus();
        renaming.value = null;
      },
      { flush: 'post' }
    );

    // A press outside the menu, a window blur, a resize, a scroll of the page, and a wheel
    // outside the menu close it (`listenSheetMenuDismiss`). The grid cancels `pointerdown`, so
    // core listens for that, in the capture phase, on the document that owns the bar.
    watch(
      menu,
      (open, _old, onCleanup) => {
        const el = bar.value;
        if (!open || !el) return;
        onCleanup(
          listenSheetMenuDismiss(
            el,
            (node) => menuRef.value?.contains(node) ?? false,
            () => (menu.value = null)
          )
        );
      },
      { flush: 'post' }
    );

    // The menu takes focus on its first item when it opens, and stays inside the bar. When it
    // closes with focus still in it, focus goes back to the element that had it at the open.
    watch(
      [menu, editor],
      ([open, ed], _old, onCleanup) => {
        const el = menuRef.value;
        if (!open || !el) return;
        // The bar of this render: the close reads it, not a later one.
        onCleanup(openSheetMenu(el, bar.value, open.left, open, ed));
      },
      { flush: 'post' }
    );

    const choose = (index: number): boolean => editor.value?.setActiveSheet(index) ?? false;
    const openMenu = (sheet: number, tab: HTMLElement, clientX?: number): void => {
      // A key opens the menu from its tab, which takes focus back; a right press keeps the
      // grid. A right press while the tab list has focus gives focus back to the pressed tab,
      // as a click on a tab does, so the focus never goes back to a tab of another sheet.
      const had = tab.ownerDocument.activeElement as HTMLElement | null;
      const inList = !!had && !!list.value?.contains(had);
      const back = clientX === undefined || inList ? tab : had;
      // A refused cell edit keeps its sheet shown: no menu opens for another sheet.
      if (!choose(sheet)) return;
      menu.value = {
        sheet,
        left: sheetMenuLeft(bar.value, tab, clientX),
        back,
        keyboard: clientX === undefined,
        rtl: bar.value !== null && isRightToLeft(bar.value),
      };
    };
    /** The rename box over the tab of the sheet named `name`, as the bar shows it now. */
    const renameBox = (name: string): SheetRenameBox | null =>
      sheetRenameBox(
        { bar: bar.value, list: list.value, add: addRef.value },
        snap.value.sheets,
        name
      );
    /**
     * Commit an open rename as a blur does, and give focus to the grid. False when the store
     * refused the name; true when no rename was open. With the document of a `press` on a
     * tab, a written name leaves the focus to the click, so the grid says only the sheet that
     * the click chooses. A press with no click, such as a drag off the tab, gives
     * the grid the focus after its release, so the focus does not stay on the page.
     */
    const commitRename = (press?: Document): boolean => {
      const open = renameField.current;
      if (!renaming.value || !open) return true;
      const written = open.commit();
      if (!press || !written) editor.value?.focus();
      else focusAfterRelease(press, () => editor.value?.focus());
      return written;
    };
    const startRename = (sheet: number): void => {
      // A rename that the engine disables starts no field (view mode, a workbook that locks its
      // sheets, or a part that refers to any sheet): the field would only show a refusal.
      if (!stateOf('rename', sheet).enabled) return;
      const ed = editor.value;
      // An open cell edit commits first. A refused one keeps its editor and its reason in
      // view, and no rename starts over it.
      if (ed && !ed.commitEdit().ok) return;
      // A new field clears the reason that an earlier rename left in view. Another reason
      // stays.
      const shown = ed?.snapshot().refusal;
      if (shown && shown === renameRefusal) ed?.dismissRefusal();
      renameRefusal = null;
      const name = snap.value.sheets[sheet]?.name;
      renameRefused.value = false;
      if (name !== undefined) renaming.value = renameBox(name);
    };
    /** Undo or redo from a focused tab or the menu; the tab of the shown sheet takes focus. */
    const stepHistory = (history: 'undo' | 'redo'): void => {
      const ed = editor.value;
      if (history === 'undo') ed?.undo();
      else ed?.redo();
      // The step can add, remove, or move tabs: the focus goes to the tab of the shown sheet.
      const shown = ed?.snapshot().activeSheet;
      if (shown !== undefined && shown >= 0) focusSheet.value = shown;
    };
    const onTabKeydown = (e: KeyboardEvent, sheet: number, tabs: readonly EditorSheet[]): void => {
      // Undo and redo reach the store history from a focused tab, as from the grid, so a sheet
      // delete or move from the tab menu can be undone with its key.
      const history = historyShortcutOf(e);
      if (history) {
        e.preventDefault();
        stepHistory(history);
        return;
      }
      if ((e.key === 'F10' && e.shiftKey) || e.key === 'ContextMenu') {
        e.preventDefault();
        openMenu(sheet, e.currentTarget as HTMLElement);
        return;
      }
      if (e.key === 'F2') {
        e.preventDefault();
        startRename(sheet);
        return;
      }
      // A chartsheet tab is disabled: the keys move between the worksheet tabs only.
      const sheets = tabs.filter((s) => s.worksheet);
      const position = sheets.findIndex((s) => s.index === sheet);
      // Along the reading direction: in a right-to-left bar, ArrowLeft moves to the next tab.
      const step = inlineStep(e.key, e.currentTarget as HTMLElement);
      const next =
        e.key === 'Home' ? 0 : e.key === 'End' ? sheets.length - 1 : step ? position + step : -1;
      if (position < 0 || next < 0 || next >= sheets.length) return;
      e.preventDefault();
      const target = sheets[next]!;
      // The focus moves only when the sheet changes, so it stays on the selected tab.
      if (!choose(target.index)) return;
      list.value?.querySelector<HTMLElement>(`[data-sheet-index="${target.index}"]`)?.focus();
    };
    const onMenuKeydown = (e: KeyboardEvent): void => {
      // Undo and redo close the menu and reach the store history, as on a focused tab: the step
      // can remove the sheet that the menu acts on.
      const history = historyShortcutOf(e);
      if (history) {
        e.preventDefault();
        e.stopPropagation();
        menu.value = null;
        stepHistory(history);
        return;
      }
      if (e.key === 'Escape' || e.key === 'Tab') {
        e.preventDefault();
        e.stopPropagation();
        menu.value = null;
        return;
      }
      if (menuRef.value && moveInSheetMenu(menuRef.value, e.key)) {
        e.preventDefault();
        e.stopPropagation();
      }
    };
    /**
     * A menu row, from its own command state: disabled rows stay focusable (`aria-disabled`)
     * and do nothing. The engine's reason is the tooltip and the accessible description of a
     * disabled row, as in the toolbar.
     */
    const item = (
      command: SheetCommand,
      label: string,
      state: SheetCommandState,
      run: () => void
    ) => {
      const disabled = !state.enabled;
      const reason =
        disabled && state.disabledReason ? disabledReasonText(state.disabledReason, t) : null;
      const id = `${reasonId}-${command}`;
      return [
        <button
          key={command}
          type="button"
          role="menuitem"
          class="xlsx-sheet-menu-item"
          tabindex={-1}
          aria-disabled={disabled || undefined}
          {...(reason
            ? { title: t('toolbar.disabledTitle', { label, reason }), 'aria-describedby': id }
            : {})}
          onMousedown={keepGridFocus}
          onClick={() => {
            if (!disabled) run();
          }}
        >
          {label}
        </button>,
        reason ? (
          <span key={id} id={id} hidden>
            {reason}
          </span>
        ) : null,
      ];
    };

    return () => {
      const s = snap.value;
      const tabs = s.sheets.filter((sheet) => sheet.visible);
      if (s.status !== 'ready' || tabs.length === 0) return null;
      const addState = stateOf('add');
      const sheetLock = addState.disabledReason
        ? disabledReasonText(addState.disabledReason, t)
        : null;
      const blocked = !addState.enabled;
      const open = menu.value;
      const sheet = open?.sheet ?? -1;
      // Move left and Move right go past the next visible tab: its workbook position (`to`).
      const renameState = stateOf('rename', sheet);
      const deleteState = stateOf('delete', sheet);
      // The commands move a sheet earlier or later in the workbook. In a right-to-left bar the
      // earlier sheet is on the right, so the words swap: the items name the side that they move to.
      const moveEarlier = open?.rtl ? 'sheetTabs.moveRight' : 'sheetTabs.moveLeft';
      const moveLater = open?.rtl ? 'sheetTabs.moveLeft' : 'sheetTabs.moveRight';
      const leftState = stateOf('moveLeft', sheet);
      const rightState = stateOf('moveRight', sheet);
      const keyboardMenu = open?.keyboard ?? false;
      const moveTo = (to: number): void => {
        menu.value = null;
        const r = editor.value?.moveSheet(sheet, to);
        if (keyboardMenu) focusSheet.value = r?.ok ? to : sheet;
      };
      const rename = renaming.value;
      // The reason of a refused name while its rename stays open: the field is invalid and
      // describes itself with it. The same text as the status message and the live region
      // (`refusalText`).
      const renameReason =
        rename && renameRefused.value && s.refusal
          ? refusalText(s.refusal, t, s.numberLocale)
          : null;
      // In view mode, or when the workbook protection locks the sheet list, the add button is
      // described by that reason, as the toolbar describes its disabled controls.
      const viewReason = sheetLock ?? '';
      // A tab whose rename the engine disables says why in its title and its description, so
      // F2 and a double-click that start no rename are explained. A tab of a sheet that is not
      // a worksheet is disabled whatever the mode: its reason is its own.
      const notWorksheet = disabledReasonText({ key: 'disabledReason.notWorksheet' }, t);
      const tabReasons = new Map<string, string>();
      const tabReason = (tab: EditorSheet): { text: string; id: string } | null => {
        const state = stateOf('rename', tab.index);
        const text = !tab.worksheet
          ? notWorksheet
          : !state.enabled && state.disabledReason
            ? disabledReasonText(state.disabledReason, t)
            : null;
        if (text === null) return null;
        // One hidden description for each distinct reason, outside the tab list.
        let id = tabReasons.get(text);
        if (id === undefined) tabReasons.set(text, (id = `${tabReasonId}-${tabReasons.size}`));
        return { text, id };
      };
      const reasonOfTab = tabs.map(tabReason);
      const addLabel = t('sheetTabs.add');

      return (
        <div class="xlsx-sheet-bar" dir={dir.value} ref={bar} onMousedown={keepGridFocus}>
          <div
            class="xlsx-sheet-tabs"
            role="tablist"
            aria-label={t('sheetTabs.ariaLabel')}
            ref={list}
            onScroll={() => {
              // The rename field follows its tab.
              if (renaming.value) renaming.value = renameBox(renaming.value.name);
            }}
          >
            {tabs.map((tab, i) => (
              <button
                key={tabKey(tabs, i)}
                type="button"
                role="tab"
                class={`xlsx-sheet-tab${rename?.name === tab.name ? ' xlsx-sheet-tab-renaming' : ''}`}
                data-sheet-index={tab.index}
                aria-selected={tab.index === s.activeSheet}
                {...(tab.worksheet
                  ? { 'aria-haspopup': 'menu', 'aria-expanded': open?.sheet === tab.index }
                  : { 'aria-disabled': true })}
                tabindex={tab.worksheet && tab.index === s.activeSheet ? 0 : -1}
                // While a rename is disabled, F2 and a double-click start no rename: the tab
                // describes itself with the reason, so a screen reader says why.
                aria-describedby={reasonOfTab[i]?.id}
                // A mouse user sees the reason in the tooltip, as a disabled toolbar control
                // shows it.
                title={
                  reasonOfTab[i]
                    ? t('toolbar.disabledTitle', { label: tab.name, reason: reasonOfTab[i].text })
                    : tab.name
                }
                onMousedown={(e: MouseEvent) => {
                  keepGridFocus(e);
                  const doc = (e.currentTarget as HTMLElement).ownerDocument;
                  listFocused = list.value?.contains(doc.activeElement) ?? false;
                  // The first press of a click commits an open rename; the second press of a
                  // double-click keeps the outcome of the first. After a written name, the
                  // click gives the focus, or the menu of a right press takes it.
                  if (e.detail <= 1) refusedPress = !commitRename(doc);
                }}
                onClick={(e: MouseEvent) => {
                  // A chartsheet tab is disabled: it takes no click.
                  if (refusedPress || !tab.worksheet) return;
                  choose(tab.index);
                  // The tab list keeps focus on the chosen tab; any other focus goes to the
                  // grid.
                  if (listFocused) (e.currentTarget as HTMLElement).focus();
                  else if (!renaming.value) editor.value?.focus();
                }}
                onDblclick={() => {
                  if (!refusedPress && tab.worksheet) startRename(tab.index);
                }}
                onContextmenu={(e: MouseEvent) => {
                  e.preventDefault();
                  // A disabled chartsheet tab takes no menu: none of its actions applies.
                  if (refusedPress || !tab.worksheet) return;
                  openMenu(tab.index, e.currentTarget as HTMLElement, e.clientX);
                }}
                onKeydown={(e: KeyboardEvent) => {
                  if (tab.worksheet) onTabKeydown(e, tab.index, tabs);
                }}
              >
                {tab.name}
              </button>
            ))}
          </div>
          {rename ? (
            <RenameField
              key={rename.name}
              name={rename.name}
              label={t('sheetTabs.renameInput')}
              box={rename}
              reasonId={renameReason ? renameReasonId : undefined}
              handle={renameField}
              focusGrid={() => editor.value?.focus()}
              onDone={(name: string | null, keyboard: boolean) => {
                const ed = editor.value;
                // The sheet by the name the rename started with: positions can change under it.
                const target = ed?.snapshot().sheets.find((x) => x.name === rename.name);
                let written = true;
                if (name !== null && target) {
                  const r = ed?.renameSheet(target.index, name);
                  written = r?.ok ?? false;
                  renameRefused.value = !written;
                  renameRefusal = written ? null : (ed?.snapshot().refusal ?? null);
                  // A refused Enter keeps the field open with its text. A refused blur ends the
                  // rename and keeps the old name. The reason shows in the status message.
                  if (!written && keyboard) return false;
                }
                renaming.value = null;
                if (keyboard && target) focusSheet.value = target.index;
                return written;
              }}
              onCancel={() => {
                // Only a reason that this rename made goes: another one stays in view.
                if (renameRefused.value) editor.value?.dismissRefusal();
                renameRefused.value = false;
              }}
            />
          ) : null}
          {blocked ? (
            <span id={viewReasonId} hidden>
              {viewReason}
            </span>
          ) : null}
          {[...tabReasons].map(([text, id]) => (
            <span key={id} id={id} hidden>
              {text}
            </span>
          ))}
          {renameReason ? (
            <span id={renameReasonId} hidden>
              {renameReason}
            </span>
          ) : null}
          <button
            ref={addRef}
            type="button"
            class="xlsx-sheet-add"
            aria-label={addLabel}
            title={
              blocked
                ? t('toolbar.disabledTitle', { label: addLabel, reason: viewReason })
                : addLabel
            }
            aria-disabled={blocked || undefined}
            aria-describedby={blocked ? viewReasonId : undefined}
            onMousedown={keepGridFocus}
            onKeydown={(e: KeyboardEvent) => {
              // Undo and redo reach the store history from the add button, as from a tab.
              const history = historyShortcutOf(e);
              if (!history) return;
              e.preventDefault();
              if (history === 'undo') editor.value?.undo();
              else editor.value?.redo();
            }}
            onClick={(e: MouseEvent) => {
              // An open rename commits first, as a blur does. Its field had focus: the grid
              // takes it. A refused name keeps its reason in view: no sheet is added over it.
              // In view mode, or when the workbook locks its sheets, the button is disabled.
              if (blocked) return;
              // The press keeps the focus where it was (`keepGridFocus`). When the tab list had
              // it, the tab of the new sheet takes it, as `choose` gives it to a clicked tab.
              const doc = (e.currentTarget as HTMLElement).ownerDocument;
              const fromList = list.value?.contains(doc.activeElement) ?? false;
              const ed = editor.value;
              if (commitRename() && ed?.addSheet().ok && fromList)
                focusSheet.value = ed.snapshot().activeSheet;
            }}
          >
            <AddIcon />
          </button>
          {open ? (
            <div
              ref={menuRef}
              class="xlsx-sheet-menu"
              role="menu"
              aria-label={t('sheetTabs.menu')}
              style={{ left: `${open.left}px` }}
              onMousedown={keepGridFocus}
              onContextmenu={(e: MouseEvent) => e.preventDefault()}
              onKeydown={onMenuKeydown}
            >
              {item('rename', t('sheetTabs.rename'), renameState, () => {
                menu.value = null;
                startRename(sheet);
              })}
              {item('delete', t('sheetTabs.delete'), deleteState, () => {
                menu.value = null;
                const ed = editor.value;
                ed?.deleteSheet(sheet);
                const shown = ed?.snapshot().activeSheet;
                if (keyboardMenu && shown !== undefined && shown >= 0) focusSheet.value = shown;
              })}
              {item('moveLeft', t(moveEarlier), leftState, () => moveTo(leftState.to!))}
              {item('moveRight', t(moveLater), rightState, () => moveTo(rightState.to!))}
            </div>
          ) : null}
        </div>
      );
    };
  },
});

/**
 * The rename field of one tab, over the tab. While a refused name keeps it open, it is invalid
 * and described by the reason (`reasonId`). `onDone` gets the trimmed new name, or null when
 * nothing is to be written (Escape, an empty name, or the same name), and whether a key ended
 * the rename. It returns false when the store refused the name: after Enter, the field stays
 * open with the text. `onCancel` runs when Escape ends the rename. `handle` gets the field's
 * `RenameHandle` while it is open. A `RENAME_COMMIT_EVENT` on the field commits it as a blur
 * does, gives focus to the grid (`focusGrid`), and sets the outcome in its detail.
 */
const RenameField = defineComponent({
  name: 'XlsxSheetRename',
  props: {
    name: { type: String, required: true },
    label: { type: String, required: true },
    box: {
      type: Object as PropType<{
        readonly left: number;
        readonly right?: number;
        readonly width: number;
        readonly maxWidth: number | undefined;
      }>,
      required: true,
    },
    /** The id of the reason of a refused name, while the field stays open after it. */
    reasonId: { type: String, default: undefined },
    handle: { type: Object as PropType<{ current: RenameHandle | null }>, required: true },
    focusGrid: { type: Function as PropType<() => void>, required: true },
    onDone: {
      type: Function as PropType<(name: string | null, keyboard: boolean) => boolean>,
      required: true,
    },
    onCancel: { type: Function as PropType<() => void>, required: true },
  },
  setup(props) {
    const value = shallowRef(props.name);
    const field = shallowRef<HTMLInputElement | null>(null);
    // Set when the rename has ended: the blur that follows Enter or Escape writes nothing.
    let ended = false;
    /** End the rename; false when the store refused the name. */
    const finish = (write: boolean, keyboard: boolean): boolean => {
      if (ended) return true;
      ended = true;
      const next = value.value.trim();
      const written = props.onDone(
        write && next !== '' && next !== props.name ? next : null,
        keyboard
      );
      if (!written && keyboard) {
        ended = false;
        field.value?.focus();
      }
      return written;
    };
    const onCommit = (e: Event): void => {
      (e as CustomEvent<RenameCommitDetail>).detail.written = finish(true, false);
      props.focusGrid();
    };
    onMounted(() => {
      field.value?.focus();
      field.value?.select();
      props.handle.current = {
        commit: () => finish(true, false),
        drop: () => {
          ended = true;
        },
      };
      field.value?.addEventListener(RENAME_COMMIT_EVENT, onCommit);
    });
    onBeforeUnmount(() => {
      props.handle.current = null;
      field.value?.removeEventListener(RENAME_COMMIT_EVENT, onCommit);
    });
    return () => {
      const { box } = props;
      return (
        <input
          ref={field}
          class="xlsx-sheet-rename"
          aria-label={props.label}
          aria-invalid={props.reasonId ? true : undefined}
          aria-describedby={props.reasonId}
          value={value.value}
          style={{
            ...(box.right === undefined
              ? { left: `${box.left}px` }
              : { left: 'auto', right: `${box.right}px` }),
            minWidth: `${box.width}px`,
            ...(box.maxWidth === undefined ? {} : { maxWidth: `${box.maxWidth}px` }),
          }}
          size={Math.max(4, Math.min(value.value.length + 1, 31))}
          onInput={(e: Event) => (value.value = (e.target as HTMLInputElement).value)}
          onBlur={() => finish(true, false)}
          onKeydown={(e: KeyboardEvent) => {
            // Every key stays in the field: no grid or tab key acts on it.
            e.stopPropagation();
            // A key that ends an IME composition, such as Enter that picks a candidate, belongs
            // to the composition: it neither commits nor cancels the rename (as
            // `cell-editor.ts` does).
            if (e.isComposing || e.keyCode === 229) return;
            if (e.key === 'Enter') {
              e.preventDefault();
              finish(true, true);
            } else if (e.key === 'Escape') {
              e.preventDefault();
              finish(false, true);
              props.onCancel();
            }
          }}
        />
      );
    };
  },
});
