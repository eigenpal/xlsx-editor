/**
 * `@xlsx-editor.dev/vue`: the Vue 3 adapter for XLSX Editor, the twin of
 * `@xlsx-editor.dev/react`.
 *
 * Provider-first, after PPTX Editor's and DOCX Editor's Vue adapters: `XlsxEditorRoot` owns the
 * editor instance, `XlsxEditorViewport` scopes the styles, and `XlsxEditorContent` attaches the
 * grid surface. Every piece of chrome reads the editor through composables;
 * `XlsxEditorToolbar` renders the chrome registry. `XlsxEditor` composes them all.
 * `scripts/parity/parity.contract.json` lists which React members Vue pairs.
 *
 * @packageDocumentation
 * @public
 */

import { XlsxEditor as XlsxEditorDefault } from './XlsxEditor.ts';
import { XlsxEditorContent, XlsxEditorRoot, XlsxEditorViewport } from './XlsxEditorRoot.ts';
import { StatusMessage, ZoomControl } from './chrome.tsx';
import { FormulaBar } from './FormulaBar.tsx';
import { SheetTabs } from './SheetTabs.tsx';
import { XlsxEditorToolbar } from './toolbar/Toolbar.tsx';

/**
 * The type of `XlsxEditor`: the composed editor, and its parts as statics for a custom layout.
 * It is written out, as React's `XlsxEditorNamespace` is, so the declarations name each part
 * and do not expand the component instance type.
 *
 * @public
 */
export type XlsxEditorNamespace = typeof XlsxEditorDefault & {
  readonly Root: typeof XlsxEditorRoot;
  readonly Viewport: typeof XlsxEditorViewport;
  readonly Content: typeof XlsxEditorContent;
  readonly Toolbar: typeof XlsxEditorToolbar;
  readonly FormulaBar: typeof FormulaBar;
  readonly SheetTabs: typeof SheetTabs;
  readonly ZoomControl: typeof ZoomControl;
  readonly StatusMessage: typeof StatusMessage;
};

/**
 * The composed editor, with its parts for a custom layout, the twin of React's `XlsxEditor`.
 * Type a template ref as `XlsxEditorRef`, for example `useTemplateRef<XlsxEditorRef>('editor')`.
 * Each part is also exported by its own name. @public
 */
export const XlsxEditor: XlsxEditorNamespace = Object.assign(XlsxEditorDefault, {
  Root: XlsxEditorRoot,
  Viewport: XlsxEditorViewport,
  Content: XlsxEditorContent,
  Toolbar: XlsxEditorToolbar,
  FormulaBar,
  SheetTabs,
  ZoomControl,
  StatusMessage,
});

export type { XlsxEditorProps } from './XlsxEditor.ts';
export {
  useEditorColorMode,
  type UseEditorColorModeReturn,
  useScopeClassName,
  type UseScopeClassNameReturn,
} from './color-mode.ts';
export type { ColorMode, ResolvedColorMode } from '@xlsx-editor.dev/core/editor';
export type { SheetProtection, WorkbookProtection } from '@xlsx-editor.dev/core/editor';
export { XlsxEditorContent, XlsxEditorRoot, XlsxEditorViewport } from './XlsxEditorRoot.ts';
export { XlsxEditorToolbar } from './toolbar/Toolbar.tsx';
export { XlsxEditorToolbarGroup, XlsxEditorToolbarSlot } from './toolbar/host-parts.tsx';
export type {
  XlsxEditorToolbarGroupProps,
  XlsxEditorToolbarSlotProps,
} from './toolbar/host-parts.tsx';
export { StatusMessage, ZoomControl } from './chrome.tsx';
export { FormulaBar } from './FormulaBar.tsx';
export { SheetTabs } from './SheetTabs.tsx';
export type { XlsxEditorToolbarNamespace, XlsxEditorToolbarProps } from './toolbar/Toolbar.tsx';
export type {
  XlsxEditorContentProps,
  XlsxEditorFonts,
  XlsxEditorRef,
  XlsxEditorRootProps,
  XlsxEditorViewportProps,
} from './XlsxEditorRoot.ts';
export {
  translatorFor,
  useEditorSnapshot,
  type UseEditorSnapshotReturn,
  useEditorState,
  type UseEditorStateReturn,
  useTranslation,
  type UseTranslationReturn,
  useViewModeWaits,
  type UseViewModeWaitsReturn,
  useXlsxEditor,
  type UseXlsxEditorReturn,
  xlsxEditorKey,
  xlsxTranslationKey,
} from './context.ts';
export {
  CLIPBOARD_TYPE,
  composesKey,
  contentFieldText,
  CUSTOM_NUMBER_FORMAT_PREFIX,
  editMoveOf,
  fieldHistoryKey,
  MAX_COLUMN_WIDTH,
  MAX_CONTENT_FIELD_CHARS,
  MAX_GO_TO_LENGTH,
  MAX_LISTED_NUMBER_FORMATS,
  MAX_RESIZE_NEW_ROWS,
  MAX_ROW_HEIGHT,
  mergeCommandState,
  NUMBER_FORMAT_CODE_PROBLEM_KEYS,
  NUMBER_FORMAT_SAMPLE,
  numberFormatChoiceCode,
  numberFormatChoices,
  numberFormatCodeProblem,
  numberFormatDialogApplyStep,
  numberFormatDialogCurrent,
  numberFormatDialogStart,
  numberFormatListView,
  numberFormatValue,
  previewNumberFormat,
  refusalText,
  refusesEditText,
  refusesReference,
  sheetCommandState,
  trapTabWithin,
  workbookNumberFormatCodes,
} from '@xlsx-editor.dev/core/editor';
export { useEditorCommand, type UseEditorCommandReturn } from './useEditorCommand.ts';
export type {
  ActiveCellFormat,
  CellEditMode,
  ChromeDisabledReason,
  ChromeGroup,
  ChromeGroupId,
  ChromeHint,
  ChromeSlotId,
  ClipboardContent,
  CopyResult,
  DecimalSteps,
  EditMove,
  EditorFailure,
  EditorMode,
  EditorSheet,
  EditorSnapshot,
  EditRefusal,
  EditRefusalCode,
  EditResult,
  EditTextOptions,
  FieldKeys,
  FontSource,
  GridPoint,
  GridRange,
  GridSelection,
  MergeCommand,
  MergeCommandState,
  MergeOptions,
  NumberFormatChoice,
  NumberFormatChoiceValue,
  NumberFormatCodeProblem,
  NumberFormatDialogApplyStep,
  NumberFormatDialogStart,
  NumberFormatListEntry,
  NumberFormatListView,
  NumberFormatPreview,
  NumberLocale,
  PartPastLimit,
  PasteData,
  SavedWorkbook,
  SheetCommand,
  SheetCommandInput,
  SheetCommandState,
  ToolbarCommandState,
  ToolbarHintKey,
  WorkbookKind,
  XlsxEditorInstance,
} from '@xlsx-editor.dev/core/editor';
export type { Translations } from '@xlsx-editor.dev/i18n';
