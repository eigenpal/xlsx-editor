// The `editor` lane: the facade that the adapters render. The workbook editor opens a workbook,
// holds the view state, mounts the grid surface, and edits cells through the store; the
// adapters render chrome from its snapshot. It also re-exports the engine types and helpers
// that a host needs to name a cell.

export type { GridPoint, GridRange, GridSelection } from '../contracts/index.ts';
export {
  columnLetters,
  formatCellRef,
  parseCellRef,
  resolveLimits,
  MAX_COLS,
  MAX_ROWS,
  type CellAddress,
  type ResourceLimits,
} from '../store/index.ts';
export { pointsToPixels, pixelsToPoints, type FontSource } from '../layout/index.ts';
export { sameNumberLocale, type NumberLocale } from '../cascade/index.ts';
export { sameFontSource, type EditorFontSource } from './editor-fonts.ts';
export {
  resolveDarkChrome,
  subscribeSystemDark,
  systemPrefersDark,
  type ColorMode,
  type ResolvedColorMode,
} from './color-mode.ts';
export { createXlsxEditor, type XlsxEditorInstance } from './xlsx-editor.ts';
export type { DocumentChange, EditorEvents } from './editor-events.ts';
export type { SavedWorkbook, SelectionRangeInput } from './editor-instance.ts';
export {
  CUT_REFUSAL_CODES,
  EDIT_REFUSAL_CODES,
  SAVE_REFUSAL_CODES,
  parseCellInput,
  displayName,
  type CellEntryOptions,
  refusalText,
  subjectText,
  type RefusalSubject,
  type EditRefusal,
  type EditRefusalCode,
  type EditResult,
} from './cell-input.ts';
export {
  CLIPBOARD_TYPE,
  type ClipboardContent,
  type CopyResult,
  type PasteData,
} from './clipboard.ts';
export { historyShortcutOf, type HistoryShortcutKeys } from './history-shortcut.ts';
export type { CellEditMode, EditMove, EditTextOptions } from './cell-editor.ts';
export { MAX_GO_TO_LENGTH } from './go-to.ts';
export { MAX_SELECTION_RANGES } from './selection.ts';
export { MAX_CONTENT_FIELD_CHARS, contentFieldText } from './cell-editor.ts';
export {
  composesKey,
  editMoveOf,
  fieldHistoryKey,
  refusesEditText,
  refusesReference,
  type FieldKeys,
} from './field-keys.ts';
export type { CellInput, SheetProtection, WorkbookProtection } from '../store/index.ts';
export {
  STRUCTURE_SUBJECT_KINDS,
  type StructureSubject,
  type StructureSubjectKind,
} from '../store/index.ts';
export {
  MAX_COLUMN_WIDTH,
  MAX_FORMAT_NEW_CELLS,
  MAX_RESIZE_NEW_ROWS,
  MAX_ROW_HEIGHT,
  PATCH_BORDER_STYLES,
  PATCH_HORIZONTAL,
  PATCH_VERTICAL,
  type FormatPatch,
  type PatchBorderStyle,
  type PatchHorizontal,
  type PatchVertical,
} from '../store/index.ts';
export type { FormatToggle, MergeOptions } from '../automation/index.ts';
export {
  EMPTY_SNAPSHOT,
  MAX_ZOOM,
  MIN_ZOOM,
  type ActiveCellFormat,
  type EditorFailure,
  type PartPastLimit,
  type EditorMode,
  type EditorSheet,
  type EditorSnapshot,
  type EditorTranslate,
  type WorkbookKind,
  type XlsxEditorOptions,
} from './editor-model.ts';
export {
  loadSubstituteFonts,
  reportFontError,
  substituteFontFailure,
  type FontReport,
  type SubstituteFontFailure,
} from './substitute-fonts.ts';
export {
  editorHostDocument,
  editorHostElement,
  editorHostWindow,
  hostDocumentOf,
  hostWindowOf,
  isElementNode,
  isHTMLElementNode,
  type HostWindow,
} from './host-realm.ts';
export {
  CHROME_GROUPS,
  CHROME_SLOT_IDS,
  chromeControlForSlot,
  chromeSlotId,
  defaultChromeGroups,
  type ChromeControl,
  type ChromeControlKind,
  type ChromeControlShape,
  type ChromeGroup,
  type ChromeGroupId,
  type ChromeSlotId,
  type SlotOf,
} from './chrome-controls.ts';
export {
  AUTO_COLOR,
  CLEAR_CHIP_HEX,
  COLOR_PICKER_STANDARD_COLORS,
  COLOR_PICKER_THEME_SLOTS,
  colorPickHex,
  colorPickLabel,
  colorPickerGrid,
  customHexDraft,
  focusColorSwatch,
  hexFromText,
  moveColorSwatchFocus,
  workbookThemeHexes,
  type ColorPick,
  type ColorPickerGrid,
  type ColorPickerThemeSlot,
  type ColorSwatch,
} from './color-picker.ts';
export {
  COLOR_CHOICES,
  CUSTOM_NUMBER_FORMAT_PREFIX,
  DISABLED_REASON_KEYS,
  TOOLBAR_HINT_KEYS,
  type ToolbarHintKey,
  FONT_FAMILY_CHOICES,
  FONT_SIZE_CHOICES,
  NO_FILL,
  numberFormatValue,
  runToolbarCommand,
  disabledReasonText,
  toolbarCommandState,
  toolbarCommandStates,
  toolbarRowsState,
  type ChromeDisabledReason,
  type ChromeHint,
  type DisabledReasonKey,
  type ToolbarCommandState,
  type ToolbarHost,
} from './toolbar-commands.ts';
export { mergeCommandState, type MergeCommand, type MergeCommandState } from './merge-state.ts';
export { isDialogBackdropPoint, trapTabWithin } from './dialog-keys.ts';
export {
  NUMBER_FORMAT_CODE_PROBLEM_KEYS,
  numberFormatCodeProblem,
  type NumberFormatCodeProblem,
} from './number-format-code.ts';
export {
  MAX_LISTED_NUMBER_FORMATS,
  NUMBER_FORMAT_SAMPLE,
  numberFormatDialogApplyStep,
  numberFormatDialogCurrent,
  numberFormatDialogStart,
  previewNumberFormat,
  workbookNumberFormatCodes,
  type NumberFormatDialogApplyStep,
  type NumberFormatDialogStart,
  type NumberFormatPreview,
} from './number-format-entry.ts';
export {
  numberFormatChoiceCode,
  numberFormatChoices,
  type NumberFormatChoice,
  type NumberFormatChoiceValue,
} from './number-format-commands.ts';
export {
  numberFormatListView,
  type NumberFormatListEntry,
  type NumberFormatListView,
} from './number-format-list.ts';
export { type DecimalSteps } from './decimal-selection.ts';
export {
  sheetCommandState,
  type SheetCommand,
  type SheetCommandInput,
  type SheetCommandState,
} from './sheet-command-state.ts';
export {
  editorInstanceScope,
  hasOpenNestedPopup,
  listenForOutsidePress,
  listenForPopupEscape,
  listenForPopupFocusLeave,
  type OutsidePressOptions,
  type PopupEscapeOptions,
  type PopupFocusLeaveOptions,
} from './popup-escape.ts';
export {
  MENU_PANEL_ROWS,
  enterMenuPanel,
  enterToolbarPopup,
  focusMenuPanelEdge,
  followPressedTrigger,
  itemStartingWith,
  menuPanelOf,
  menuPanelRows,
  menuRowLabels,
  listenForPointerPopupKeys,
  listenForToolbarPopupKeys,
  popupEntryOf,
  type PointerPopupKeysOptions,
  type PopupEntry,
  type PopupEntryKey,
  type PopupEntryResult,
} from './popup-keys.ts';
export { ICON_PATHS as CHROME_ICON_PATHS } from './icon-paths.ts';
export { relabelEditor } from './cell-editing.ts';
export { scrollIntoNearestScroller } from './scroll-nearest.ts';
// Adapter support: the framework-neutral parts of the React and Vue chrome. Not a stable API.
export {
  editorCommandSlice,
  editorRowsSlice,
  formulaSlice,
  gridSetupProblem,
  packageFailureText,
  pastLimitText,
  sameEditorCommandSlice,
  sameSlice,
  slotWords,
  statusSlice,
  statusView,
  tabsSlice,
  watchGridSetup,
  zoomSlice,
  type EditorCommandSlice,
  type GridSetupFacts,
  type SlotWords,
  type StatusNotice,
  type StatusView,
} from './chrome-status.ts';
export { inlineStep, isRightToLeft, markChromeDirection } from './chrome-direction.ts';
export {
  chromeTranslator,
  partTranslator,
  warnPartTranslateError,
  type ChromePartTranslate,
  type ChromeTranslate,
} from './chrome-translate.ts';
export {
  cycleRegion,
  guardToolbarMousedown,
  keepGridFocus,
  type ChromePress,
  type RegionKey,
} from './chrome-focus.ts';
export {
  RENAME_COMMIT_EVENT,
  blockNextClick,
  commitFocusedRename,
  focusAfterRelease,
  renameFocused,
  type RenameCommitDetail,
} from './rename-commit.ts';
export {
  insertFieldText,
  pressFieldBox,
  showCaretLine,
  stepFieldHistory,
  type FieldBoxPress,
  type FieldHistoryKey,
} from './formula-field.ts';
export {
  focusPopupItem,
  groupsThatFit,
  isHexColor,
  colorBarTone,
  listenSaveKey,
  measureToolbar,
  moveBarFocus,
  movePopupFocus,
  observeToolbar,
  parseFontSize,
  placePopup,
  popupTriggerAction,
  TOOLBAR_FIXED_ATTRIBUTE,
  TOOLBAR_PANEL_ITEMS,
  toolbarCollapse,
  toolbarGroups,
  type FocusMoveKey,
  type ToolbarFitGroup,
} from './toolbar-layout.ts';
export {
  arrangeToolbar,
  arrangeToolbarGroups,
  collapseOrder,
  notInToolbar,
  TOOLBAR_COLLAPSE_ORDER,
  TOOLBAR_PINNED_GROUPS,
  toolbarGroupHeading,
  toolbarAsChildProblem,
  toolbarGroupPriority,
  toolbarSlotAdvice,
  type ArrangedToolbarGroup,
  type ToolbarArrangement,
  type ToolbarGroupRequest,
  type ToolbarHostGroupPlacement,
  type ToolbarPartNames,
  type ToolbarSlotRequest,
} from './toolbar-arrange.ts';
export {
  isToolbarButtonSlot,
  takesToolbarIcon,
  TOOLBAR_NAMED_PARTS,
  type ToolbarPartName,
} from './toolbar-named-parts.ts';
export {
  SHEET_MENU_ITEM,
  listenSheetMenuDismiss,
  moveInSheetMenu,
  openSheetMenu,
  sheetListKey,
  sheetMenuLeft,
  sheetRenameBox,
  tabKey,
  type SheetBarParts,
  type SheetMenuOpening,
  type SheetRenameBox,
} from './sheet-tab-parts.ts';
export {
  emptySnapshotOf,
  followModeProp,
  watchFontErrors,
  whenOpenSettles,
  type ModeProp,
} from './root-support.ts';
export {
  CHROME_MENUS,
  chromeMenuRowKey,
  chromeMenuShortcut,
  chromeMenuSlots,
  isApplePlatform,
  MENU_KEY_WORD_KEYS,
  menuKeyWords,
  menuRowKeepsFocus,
  menuRowRole,
  menuShortcutText,
  tidySeparators,
  type ChromeMenu,
  type ChromeMenuEntry,
  type ChromeMenuId,
  type ChromeMenuItemEntry,
  type ChromeMenuSeparatorEntry,
  type ChromeMenuSubmenuEntry,
  type MenuKeyWords,
} from './chrome-menus.ts';
export {
  MENU_FIXED_ATTRIBUTE,
  MENU_MORE_ID,
  MENU_NAMED_PARTS,
  REACT_MENU_PART_NAMES,
  VUE_MENU_PART_NAMES,
  arrangeMenuBar,
  arrangeMenuRows,
  menuHiddenRowProblem,
  menuRowInBarProblem,
  menuBarArrow,
  menuBarTriggerKey,
  menuPanelKey,
  menuPartName,
  measureMenuBar,
  menuStep,
  submenuRowSlots,
  menusThatFit,
  observeMenuBar,
  menuTypeahead,
  type ArrangedBarMenu,
  type ArrangedMenuRows,
  type MenuBarSpec,
  type MenuBarTriggerKey,
  type MenuNamedRow,
  type MenuPanelKey,
  type MenuPartNames,
  type MenuRowPartName,
  type MenuSlotRowSpec,
} from './menu-arrangement.ts';
export { editorClipboardReader, systemClipboardReader } from './menu-commands.ts';
export {
  gridContextMenuEntries,
  type GridContextMenuRequest,
  type GridContextMenuTargetKind,
} from './grid-context-menu.ts';
export { openPickedWorkbook } from './file-open.ts';
export {
  MENU_PANEL_MARGIN,
  contextMenuPlacement,
  menuPanelPlacement,
  placeContextMenu,
  placeMenuPanel,
  type MenuBox,
  type MenuPanelPlacement,
} from './menu-placement.ts';
