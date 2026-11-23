/** @jsxImportSource react */
// The sheet tabs: one tab for each visible sheet, an add button, a context menu on a tab, and
// inline rename. Each action calls the editor instance; a refused one shows through the refusal
// channel (`StatusMessage`). In view mode, and when the workbook protection locks the sheet list
// (`snapshot.workbookProtection.allowStructure` is false), the add button and the edit rows of
// the menu are disabled with that reason, and no rename starts (double-click, F2, the menu); each
// tab is then described by the reason. The state of each command comes from `sheetCommandState`.
//
// The context menu follows PPTX Editor (eigenpal/pptx-editor) `origin/main` 79a5d22d,
// `packages/react/src/RailContextMenu.tsx` and its `ContextMenu.tsx`: it opens on a right
// press, Shift+F10, or the Menu key; the arrow keys, Home, and End move through the rows; Escape
// and Tab close it; the undo and redo keys close it and run the step; a press outside, a window
// blur, a resize, a scroll of the page, and a wheel outside the menu close it
// (`listenContextMenuDismiss`, and `listenAway` in `packages/core/src/editor/chrome-popover.ts`).
// A scroll that code makes inside the page, such as the grid or the tab list, keeps it open. On
// close, focus goes back to the element that had it when the menu opened, when focus is still in
// the menu (`ContextMenu.tsx:127-141`): the tab for a
// keyboard open, the grid for a right press. After Delete or a move from a keyboard open, the
// tab of the shown or moved sheet takes focus. Differences: the menu is a plain panel above the
// tab row, not a portal placed in the window, kept inside the bar, and it has four fixed rows.
// Move left and Move right move a sheet past the next visible tab, so a hidden sheet between
// two tabs never takes a move. Delete is disabled, with its reason, on the last visible sheet.
//
// Inline rename follows PPTX Editor `packages/react/src/RailSections.tsx:340-374` at the same
// commit (`commitName`): the field starts with the whole name selected; Enter or a blur writes
// the trimmed name only when it changed; Escape cancels, and the blur that follows writes
// nothing. A refused blur ends the rename and keeps the old name; only a refused Enter keeps
// the field open with its text, and Escape then clears its reason. Every key in the field
// stops there, so no grid key acts on it. Differences: the field lies over its tab outside the
// tab list, which holds only tabs, and follows the tab when the list scrolls. The rename names
// its sheet by the name it started with, and it ends when the sheet list changes, so it never
// renames another sheet. A press on the add button or on a tab commits an open rename as a
// blur does first; a refused one keeps its reason, and no sheet is added, chosen, or renamed
// by that press or the double-click that it starts. When a rename that holds focus ends this
// way, or because the sheet list changed, focus goes to the grid. A right press on a tab whose
// rename is refused opens no menu. A press on a toolbar control commits an open rename first
// (`rename-commit.ts`), and a refused one runs no command, nor does a press on a control that
// was disabled at the press. A menu that opened while a rename
// held focus gives focus, on close, to its tab (a key open) or the grid (a press, or an opener
// that was the page). The field is never wider than the room before the add button, so the
// button stays in reach, and it takes a name of any length: the store refuses a long one with
// its reason.
//
// A click on a tab gives focus to that tab when the tab list had focus, else to the grid. The
// arrow keys, Home, and End move between the worksheet tabs and skip a chartsheet tab. A
// chartsheet tab is disabled with `aria-disabled`, not the native attribute, so its press can
// keep the grid's focus; it takes no keyboard focus, no click, no rename, and no menu. A press
// on the empty part of the bar keeps the grid's focus too. The undo and redo keys on a focused
// tab reach the store history (`historyShortcutOf`), and the grid's live region announces a
// sheet delete. A key in the rename field that belongs to an IME composition ends nothing.
// While a refused name keeps the field open, the field is invalid and described by the reason.
//
// A rename starts only after an open cell edit commits: a refused commit keeps its editor and
// its reason. Escape clears only a reason that the rename made. The rename field stays over the
// tab list when its tab scrolls out of view.
//
// The tab of the shown sheet scrolls into view when the shown sheet or the sheet list changes,
// so a renamed tab that grows stays in view.

import {
  Fragment,
  useEffect,
  useId,
  useLayoutEffect,
  useRef,
  useState,
  type KeyboardEvent,
  type MouseEvent,
} from 'react';
import {
  RENAME_COMMIT_EVENT,
  disabledReasonText,
  focusAfterRelease,
  historyShortcutOf,
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
  type RenameCommitDetail,
  type SheetCommand,
  type SheetCommandState,
  type SheetRenameBox,
} from '@xlsx-editor.dev/core/editor';
import type { EditorSheet } from '@xlsx-editor.dev/core/editor';
import { useTranslation, useEditorState, useXlsxEditor } from './context.ts';
import { AddIcon } from './icons.tsx';

/** An open tab menu: the sheet it acts on, and its left edge in the bar, in CSS pixels. */
interface TabMenu {
  readonly sheet: number;
  readonly left: number;
  /** The element that had focus when the menu opened: focus goes back to it on close. */
  readonly back: HTMLElement | null;
  /** Whether a key opened the menu: an action then focuses a tab. */
  readonly keyboard: boolean;
}

/** One tab for each visible worksheet, in workbook order, with the sheet actions. @public */
export function SheetTabs() {
  const { t } = useTranslation();
  const editor = useXlsxEditor();
  const snap = useEditorState(tabsSlice, sameSlice);
  const bar = useRef<HTMLDivElement>(null);
  const list = useRef<HTMLDivElement>(null);
  const menuRef = useRef<HTMLDivElement>(null);
  const addRef = useRef<HTMLButtonElement>(null);
  const [menu, setMenu] = useState<TabMenu | null>(null);
  const [renaming, setRenaming] = useState<SheetRenameBox | null>(null);
  // The sheet whose tab takes focus after the next render (a menu action or a rename ended).
  const [focusSheet, setFocusSheet] = useState<number | null>(null);
  const reasonId = useId();
  const viewReasonId = useId();
  const renameReasonId = useId();
  const tabReasonId = useId();
  // The engine gives the state of each sheet command (`sheetCommandState`). View mode, or a
  // workbook protection that locks the sheet list, disables every sheet edit with its reason.
  const stateOf = (command: SheetCommand, sheet = -1) => sheetCommandState(snap, command, sheet);
  const addState = stateOf('add');
  const sheetLock = addState.disabledReason ? disabledReasonText(addState.disabledReason, t) : null;
  const blocked = !addState.enabled;

  useEffect(() => {
    if (focusSheet === null) return;
    setFocusSheet(null);
    list.current?.querySelector<HTMLElement>(`[data-sheet-index="${focusSheet}"]`)?.focus();
  }, [focusSheet]);

  // The tab of the shown sheet scrolls into view, so a new, chosen, or renamed sheet is never
  // hidden.
  const active = snap.activeSheet;
  const listKey = sheetListKey(snap.sheets);
  useEffect(() => {
    // Only the tab list scrolls, never the host page (PPTX Editor `scrollIntoNearestScroller`).
    const tab = list.current?.querySelector<HTMLElement>(`[data-sheet-index="${active}"]`);
    if (tab) scrollIntoNearestScroller(tab, list.current);
  }, [active, listKey]);

  // A change of the sheet list ends a rename: its position and its tab can be gone. A field
  // that held focus gives it to the grid, not to the page.
  // The open rename field: `commit` commits it as a blur does (false when the store refused
  // the name), and `drop` ends it with no write.
  const renameField = useRef<RenameHandle | null>(null);
  // Set when a press on a tab ended a rename that the store refused: the click, the
  // double-click, and the context menu of that press do nothing, so the reason stays in view.
  const refusedPress = useRef(false);
  // Whether the tab list had focus at the last press on a tab: the click focuses that tab.
  const listFocused = useRef(false);
  // Set when the store refused a name of the open rename: Escape then clears that reason.
  const renameRefused = useRef(false);
  // The last refusal of a rename, so a new rename clears only that reason.
  const renameRefusal = useRef<unknown>(null);
  useEffect(() => {
    // The field writes nothing when the focus change below blurs it.
    renameField.current?.drop();
    const field = bar.current?.ownerDocument.activeElement;
    if (field && bar.current?.contains(field) && field.classList.contains('xlsx-sheet-rename'))
      editor?.focus();
    setRenaming(null);
  }, [listKey, editor]);

  // A press outside the menu, a window blur, a resize, a scroll of the page, and a wheel
  // outside the menu close it (PPTX Editor `listenContextMenuDismiss` and `listenAway`). A
  // scroll that code makes inside the page, such as the grid that restores its scroll position
  // or the tab list that shows the chosen tab, does not close it. The grid cancels
  // `pointerdown`, so listen for that, in the capture phase, on the document that owns the bar.
  useEffect(() => {
    const el = bar.current;
    if (!menu || !el) return;
    return listenSheetMenuDismiss(
      el,
      (node) => menuRef.current?.contains(node) ?? false,
      () => setMenu(null)
    );
  }, [menu]);

  // The menu takes focus on its first item when it opens, and stays inside the bar. When it
  // closes with focus still in it, focus goes back to the element that had it at the open.
  useLayoutEffect(() => {
    const el = menuRef.current;
    if (!menu || !el) return;
    // The bar of this render: the close reads it, not a later one.
    return openSheetMenu(el, bar.current, menu.left, menu, editor);
  }, [menu, editor]);

  const tabs = snap.sheets.filter((s) => s.visible);
  if (snap.status !== 'ready' || tabs.length === 0) return null;

  const choose = (index: number): boolean => editor?.setActiveSheet(index) ?? false;
  const openMenu = (sheet: number, tab: HTMLElement, clientX?: number) => {
    // A key opens the menu from its tab, which takes focus back; a right press keeps the grid.
    // A right press while the tab list has focus gives focus back to the pressed tab, as a
    // click on a tab does, so the focus never goes back to a tab of another sheet.
    const had = tab.ownerDocument.activeElement as HTMLElement | null;
    const inList = !!had && !!list.current?.contains(had);
    const back = clientX === undefined || inList ? tab : had;
    // A refused cell edit keeps its sheet shown: no menu opens for another sheet.
    if (!choose(sheet)) return;
    setMenu({
      sheet,
      left: sheetMenuLeft(bar.current, tab, clientX),
      back,
      keyboard: clientX === undefined,
    });
  };
  /** The rename box over the tab of the sheet named `name`, as the bar shows it now. */
  const renameBox = (name: string): SheetRenameBox | null =>
    sheetRenameBox(
      { bar: bar.current, list: list.current, add: addRef.current },
      snap.sheets,
      name
    );
  /**
   * Commit an open rename as a blur does, and give focus to the grid. False when the store
   * refused the name; true when no rename was open. With the document of a `press` on a tab,
   * a written name leaves the focus to the click, so the grid says only the sheet that the
   * click chooses. A press with no click, such as a drag off the tab, gives
   * the grid the focus after its release, so the focus does not stay on the page.
   */
  const commitRename = (press?: Document): boolean => {
    const open = renameField.current;
    if (!renaming || !open) return true;
    const written = open.commit();
    if (!press || !written) editor?.focus();
    else focusAfterRelease(press, () => editor?.focus());
    return written;
  };
  const startRename = (sheet: number) => {
    // A rename that the engine disables starts no field (view mode, a workbook that locks its
    // sheets, or a part that refers to any sheet): the field would only show a refusal.
    if (!stateOf('rename', sheet).enabled) return;
    // An open cell edit commits first. A refused one keeps its editor and its reason in view,
    // and no rename starts over it.
    if (editor && !editor.commitEdit().ok) return;
    // A new field clears the reason that an earlier rename left in view. Another reason stays.
    const shown = editor?.snapshot().refusal;
    if (shown && shown === renameRefusal.current) editor?.dismissRefusal();
    renameRefusal.current = null;
    const name = snap.sheets[sheet]?.name;
    renameRefused.current = false;
    if (name !== undefined) setRenaming(renameBox(name));
  };
  const onTabKeyDown = (e: KeyboardEvent<HTMLButtonElement>, sheet: number) => {
    // Undo and redo reach the store history from a focused tab, as from the grid, so a sheet
    // delete or move from the tab menu can be undone with its key.
    const history = historyShortcutOf(e);
    if (history) {
      e.preventDefault();
      if (history === 'undo') editor?.undo();
      else editor?.redo();
      // The step can add, remove, or move tabs: the focus goes to the tab of the shown sheet.
      const shown = editor?.snapshot().activeSheet;
      if (shown !== undefined && shown >= 0) setFocusSheet(shown);
      return;
    }
    if ((e.key === 'F10' && e.shiftKey) || e.key === 'ContextMenu') {
      e.preventDefault();
      openMenu(sheet, e.currentTarget);
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
    const step = e.key === 'ArrowRight' ? 1 : e.key === 'ArrowLeft' ? -1 : 0;
    const next =
      e.key === 'Home' ? 0 : e.key === 'End' ? sheets.length - 1 : step ? position + step : -1;
    if (position < 0 || next < 0 || next >= sheets.length) return;
    e.preventDefault();
    const target = sheets[next]!;
    // The focus moves only when the sheet changes, so it stays on the selected tab.
    if (!choose(target.index)) return;
    list.current?.querySelector<HTMLElement>(`[data-sheet-index="${target.index}"]`)?.focus();
  };
  const onMenuKeyDown = (e: KeyboardEvent<HTMLDivElement>) => {
    // Undo and redo close the menu and reach the store history, as on a focused tab: the step
    // can remove the sheet that the menu acts on.
    const history = historyShortcutOf(e);
    if (history) {
      e.preventDefault();
      e.stopPropagation();
      setMenu(null);
      if (history === 'undo') editor?.undo();
      else editor?.redo();
      const shown = editor?.snapshot().activeSheet;
      if (shown !== undefined && shown >= 0) setFocusSheet(shown);
      return;
    }
    if (e.key === 'Escape' || e.key === 'Tab') {
      e.preventDefault();
      e.stopPropagation();
      setMenu(null);
      return;
    }
    if (menuRef.current && moveInSheetMenu(menuRef.current, e.key)) {
      e.preventDefault();
      e.stopPropagation();
    }
  };
  /**
   * A menu row, from its own command state: disabled rows stay focusable (`aria-disabled`) and
   * do nothing. The engine's reason is the tooltip and the accessible description of a disabled
   * row, as in the toolbar.
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
    return (
      <Fragment key={command}>
        <button
          type="button"
          role="menuitem"
          className="xlsx-sheet-menu-item"
          tabIndex={-1}
          aria-disabled={disabled || undefined}
          {...(reason
            ? { title: t('toolbar.disabledTitle', { label, reason }), 'aria-describedby': id }
            : {})}
          onMouseDown={keepGridFocus}
          onClick={() => {
            if (!disabled) run();
          }}
        >
          {label}
        </button>
        {reason ? (
          <span id={id} hidden>
            {reason}
          </span>
        ) : null}
      </Fragment>
    );
  };
  const sheet = menu?.sheet ?? -1;
  // Move left and Move right go past the next visible tab: its workbook position (`to`).
  const renameState = stateOf('rename', sheet);
  const deleteState = stateOf('delete', sheet);
  const leftState = stateOf('moveLeft', sheet);
  const rightState = stateOf('moveRight', sheet);
  const keyboardMenu = menu?.keyboard ?? false;
  const moveTo = (to: number) => {
    setMenu(null);
    const r = editor?.moveSheet(sheet, to);
    if (keyboardMenu) setFocusSheet(r?.ok ? to : sheet);
  };
  // The reason of a refused name while its rename stays open: the field is invalid and
  // describes itself with it.
  // The same text as the status message and the live region (`refusalText`).
  const renameReason =
    renaming && renameRefused.current && snap.refusal
      ? refusalText(snap.refusal, t, snap.numberLocale)
      : null;
  // In view mode, or when the workbook protection locks the sheet list, the add button is
  // described by that reason, as the toolbar describes its disabled controls.
  const viewReason = { id: viewReasonId, text: sheetLock ?? '' };
  // A tab whose rename the engine disables says why in its title and its description, so F2
  // and a double-click that start no rename are explained: view mode, a workbook that locks its
  // sheets, or a part that refers to any sheet (`sheetEditBlocked`). A tab of a sheet that is
  // not a worksheet is disabled whatever the mode: its reason is its own.
  const notWorksheet = disabledReasonText({ key: 'disabledReason.notWorksheet' }, t);
  const tabReasons = new Map<string, string>();
  const tabReason = (s: EditorSheet): { text: string; id: string } | null => {
    const state = stateOf('rename', s.index);
    const text = !s.worksheet
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
    <div className="xlsx-sheet-bar" ref={bar} onMouseDown={keepGridFocus}>
      <div
        className="xlsx-sheet-tabs"
        role="tablist"
        aria-label={t('sheetTabs.ariaLabel')}
        ref={list}
        onScroll={() => {
          // The rename field follows its tab.
          if (renaming) setRenaming(renameBox(renaming.name));
        }}
      >
        {tabs.map((s, i) => (
          <button
            key={tabKey(tabs, i)}
            type="button"
            role="tab"
            className={`xlsx-sheet-tab${renaming?.name === s.name ? ' xlsx-sheet-tab-renaming' : ''}`}
            data-sheet-index={s.index}
            aria-selected={s.index === snap.activeSheet}
            {...(s.worksheet
              ? { 'aria-haspopup': 'menu' as const, 'aria-expanded': menu?.sheet === s.index }
              : { 'aria-disabled': true })}
            tabIndex={s.worksheet && s.index === snap.activeSheet ? 0 : -1}
            // While a rename is disabled, F2 and a double-click start no rename: the tab
            // describes itself with the reason, so a screen reader says why.
            aria-describedby={reasonOfTab[i]?.id}
            // A mouse user sees the reason in the tooltip, as a disabled toolbar control shows it.
            title={
              reasonOfTab[i]
                ? t('toolbar.disabledTitle', { label: s.name, reason: reasonOfTab[i].text })
                : s.name
            }
            onMouseDown={(e) => {
              keepGridFocus(e);
              const doc = e.currentTarget.ownerDocument;
              listFocused.current = list.current?.contains(doc.activeElement) ?? false;
              // The first press of a click commits an open rename; the second press of a
              // double-click keeps the outcome of the first. After a written name, the click
              // gives the focus, or the menu of a right press takes it.
              if (e.detail <= 1) refusedPress.current = !commitRename(doc);
            }}
            onClick={(e) => {
              // A chartsheet tab is disabled: it takes no click.
              if (refusedPress.current || !s.worksheet) return;
              choose(s.index);
              // The tab list keeps focus on the chosen tab; any other focus goes to the grid.
              if (listFocused.current) e.currentTarget.focus();
              else if (!renaming) editor?.focus();
            }}
            onDoubleClick={() => {
              if (!refusedPress.current && s.worksheet) startRename(s.index);
            }}
            onContextMenu={(e: MouseEvent<HTMLButtonElement>) => {
              e.preventDefault();
              // A disabled chartsheet tab takes no menu: none of its actions applies.
              if (refusedPress.current || !s.worksheet) return;
              openMenu(s.index, e.currentTarget, e.clientX);
            }}
            onKeyDown={(e) => {
              if (s.worksheet) onTabKeyDown(e, s.index);
            }}
          >
            {s.name}
          </button>
        ))}
      </div>
      {renaming ? (
        <RenameField
          key={renaming.name}
          name={renaming.name}
          label={t('sheetTabs.renameInput')}
          box={renaming}
          reasonId={renameReason ? renameReasonId : undefined}
          handle={renameField}
          focusGrid={() => editor?.focus()}
          onDone={(name, keyboard) => {
            // The sheet by the name the rename started with: positions can change under it.
            const sheet = editor?.snapshot().sheets.find((s) => s.name === renaming.name);
            let written = true;
            if (name !== null && sheet) {
              const r = editor?.renameSheet(sheet.index, name);
              written = r?.ok ?? false;
              renameRefused.current = !written;
              renameRefusal.current = written ? null : (editor?.snapshot().refusal ?? null);
              // A refused Enter keeps the field open with its text. A refused blur ends the
              // rename and keeps the old name. The reason shows in the status message.
              if (!written && keyboard) return false;
            }
            setRenaming(null);
            if (keyboard && sheet) setFocusSheet(sheet.index);
            return written;
          }}
          onCancel={() => {
            // Only a reason that this rename made goes: another one stays in view.
            if (renameRefused.current) editor?.dismissRefusal();
            renameRefused.current = false;
          }}
        />
      ) : null}
      {blocked ? (
        <span id={viewReasonId} hidden>
          {viewReason.text}
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
        className="xlsx-sheet-add"
        aria-label={addLabel}
        title={
          blocked
            ? t('toolbar.disabledTitle', { label: addLabel, reason: viewReason.text })
            : addLabel
        }
        aria-disabled={blocked || undefined}
        aria-describedby={blocked ? viewReasonId : undefined}
        onMouseDown={keepGridFocus}
        onKeyDown={(e) => {
          // Undo and redo reach the store history from the add button, as from a tab.
          const history = historyShortcutOf(e);
          if (!history) return;
          e.preventDefault();
          if (history === 'undo') editor?.undo();
          else editor?.redo();
        }}
        onClick={(e) => {
          // An open rename commits first, as a blur does. Its field had focus: the grid takes
          // it. A refused name keeps its reason in view: no sheet is added over it. In view
          // mode, or when the workbook locks its sheets, the button is disabled.
          if (blocked) return;
          // The press keeps the focus where it was (`keepGridFocus`). When the tab list had it,
          // the tab of the new sheet takes it, as `choose` gives it to a clicked tab.
          const doc = e.currentTarget.ownerDocument;
          const fromList = list.current?.contains(doc.activeElement) ?? false;
          if (commitRename() && editor?.addSheet().ok && fromList)
            setFocusSheet(editor.snapshot().activeSheet);
        }}
      >
        <AddIcon />
      </button>
      {menu ? (
        <div
          ref={menuRef}
          className="xlsx-sheet-menu"
          role="menu"
          aria-label={t('sheetTabs.menu')}
          style={{ left: menu.left }}
          onMouseDown={keepGridFocus}
          onContextMenu={(e) => e.preventDefault()}
          onKeyDown={onMenuKeyDown}
        >
          {item('rename', t('sheetTabs.rename'), renameState, () => {
            setMenu(null);
            startRename(sheet);
          })}
          {item('delete', t('sheetTabs.delete'), deleteState, () => {
            setMenu(null);
            editor?.deleteSheet(sheet);
            const shown = editor?.snapshot().activeSheet;
            if (keyboardMenu && shown !== undefined && shown >= 0) setFocusSheet(shown);
          })}
          {item('moveLeft', t('sheetTabs.moveLeft'), leftState, () => moveTo(leftState.to!))}
          {item('moveRight', t('sheetTabs.moveRight'), rightState, () => moveTo(rightState.to!))}
        </div>
      ) : null}
    </div>
  );
}

/** What the tab bar calls on an open rename field. */
interface RenameHandle {
  /** Commit the field as a blur does; false when the store refused the name. */
  commit(): boolean;
  /** End the rename with no write. */
  drop(): void;
}

/**
 * The rename field of one tab, over the tab. While a refused name keeps it open, it is invalid
 * and described by the reason (`reasonId`). `onDone` gets the trimmed new name, or null when
 * nothing is to be written (Escape, an empty name, or the same name), and whether a key ended
 * the rename. It returns false when the store refused the name: after Enter, the field stays
 * open with the text. `onCancel` runs when Escape ends the rename. `handle` gets the field's
 * `RenameHandle` while it is open. A `RENAME_COMMIT_EVENT` on the field commits it as a blur
 * does, gives focus to the grid (`focusGrid`), and sets the outcome in its detail.
 */
function RenameField({
  name,
  label,
  box,
  reasonId,
  handle,
  focusGrid,
  onDone,
  onCancel,
}: {
  readonly name: string;
  readonly label: string;
  readonly box: {
    readonly left: number;
    readonly width: number;
    readonly maxWidth: number | undefined;
  };
  /** The id of the reason of a refused name, while the field stays open after it. */
  readonly reasonId: string | undefined;
  readonly handle: { current: RenameHandle | null };
  readonly focusGrid: () => void;
  readonly onDone: (name: string | null, keyboard: boolean) => boolean;
  readonly onCancel: () => void;
}) {
  const [value, setValue] = useState(name);
  const field = useRef<HTMLInputElement>(null);
  // Set when the rename has ended: the blur that follows Enter or Escape writes nothing.
  const ended = useRef(false);
  useLayoutEffect(() => {
    field.current?.focus();
    field.current?.select();
  }, []);
  /** End the rename; false when the store refused the name. */
  const finish = (write: boolean, keyboard: boolean): boolean => {
    if (ended.current) return true;
    ended.current = true;
    const next = value.trim();
    const written = onDone(write && next !== '' && next !== name ? next : null, keyboard);
    if (!written && keyboard) {
      ended.current = false;
      field.current?.focus();
    }
    return written;
  };
  useLayoutEffect(() => {
    handle.current = {
      commit: () => finish(true, false),
      drop: () => {
        ended.current = true;
      },
    };
    const el = field.current;
    const onCommit = (e: Event) => {
      (e as CustomEvent<RenameCommitDetail>).detail.written = finish(true, false);
      focusGrid();
    };
    el?.addEventListener(RENAME_COMMIT_EVENT, onCommit);
    return () => {
      handle.current = null;
      el?.removeEventListener(RENAME_COMMIT_EVENT, onCommit);
    };
  });
  return (
    <input
      ref={field}
      className="xlsx-sheet-rename"
      aria-label={label}
      aria-invalid={reasonId ? true : undefined}
      aria-describedby={reasonId}
      value={value}
      style={{ left: box.left, minWidth: box.width, maxWidth: box.maxWidth }}
      size={Math.max(4, Math.min(value.length + 1, 31))}
      onChange={(e) => setValue(e.target.value)}
      onBlur={() => finish(true, false)}
      onKeyDown={(e) => {
        // Every key stays in the field: no grid or tab key acts on it.
        e.stopPropagation();
        // A key that ends an IME composition, such as Enter that picks a candidate, belongs to
        // the composition: it neither commits nor cancels the rename (as `cell-editor.ts` does).
        if (e.nativeEvent.isComposing || e.keyCode === 229) return;
        if (e.key === 'Enter') {
          e.preventDefault();
          finish(true, true);
        } else if (e.key === 'Escape') {
          e.preventDefault();
          finish(false, true);
          onCancel();
        }
      }}
    />
  );
}
