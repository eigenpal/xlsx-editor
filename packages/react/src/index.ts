/**
 * `@xlsx-editor.dev/react`: the React adapter for XLSX Editor.
 *
 * Provider-first, after PPTX Editor's adapter: `XlsxEditor.Root` owns the editor instance,
 * `.Viewport` scopes the styles, and `.Content` attaches the grid surface. Every piece of chrome
 * is a hook consumer; `.Toolbar` renders the chrome registry. `XlsxEditor` composes them all.
 *
 * @packageDocumentation
 * @public
 */

import type { ForwardRefExoticComponent, RefAttributes } from 'react';
import { XlsxEditor as XlsxEditorDefault, type XlsxEditorProps } from './XlsxEditor.tsx';
import type { XlsxEditorRef } from './XlsxEditorRoot.tsx';
import { XlsxEditorContent, XlsxEditorRoot, XlsxEditorViewport } from './XlsxEditorRoot.tsx';
import { StatusMessage, ZoomControl } from './chrome.tsx';
import { FormulaBar } from './FormulaBar.tsx';
import { SheetTabs } from './SheetTabs.tsx';
import { XlsxEditorToolbar } from './toolbar/Toolbar.tsx';

/**
 * The type of `XlsxEditor`: the composed editor, and its parts as statics for a custom layout.
 * Each part is also exported by its own name.
 *
 * @public
 */
export interface XlsxEditorNamespace extends ForwardRefExoticComponent<
  XlsxEditorProps & RefAttributes<XlsxEditorRef>
> {
  readonly Root: typeof XlsxEditorRoot;
  readonly Viewport: typeof XlsxEditorViewport;
  readonly Content: typeof XlsxEditorContent;
  readonly Toolbar: typeof XlsxEditorToolbar;
  readonly FormulaBar: typeof FormulaBar;
  readonly SheetTabs: typeof SheetTabs;
  readonly ZoomControl: typeof ZoomControl;
  readonly StatusMessage: typeof StatusMessage;
}

/** The composed editor, with its parts for a custom layout. @public */
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

export type { XlsxEditorProps } from './XlsxEditor.tsx';
export {
  useEditorColorMode,
  type UseEditorColorModeReturn,
  useScopeClassName,
} from './color-mode.ts';
export type { ColorMode, ResolvedColorMode } from '@xlsx-editor.dev/core/editor';
export type { SheetProtection, WorkbookProtection } from '@xlsx-editor.dev/core/editor';
export { XlsxEditorContent, XlsxEditorRoot, XlsxEditorViewport } from './XlsxEditorRoot.tsx';
export { XlsxEditorToolbar } from './toolbar/Toolbar.tsx';
export { XlsxEditorToolbarGroup, XlsxEditorToolbarSlot } from './toolbar/host-parts.tsx';
export type {
  XlsxEditorToolbarGroupProps,
  XlsxEditorToolbarSlotProps,
} from './toolbar/host-parts.tsx';
export {
  ToolbarAction,
  type ToolbarActionProps,
  ToolbarSeparator,
  type ToolbarSeparatorProps,
} from './toolbar/action.tsx';
export type {
  ToolbarNamedParts,
  ToolbarPartComponent,
  ToolbarPartProps,
  ToolbarSlotPartComponent,
  ToolbarSlotPartProps,
} from './toolbar/named-parts.tsx';
export {
  StatusMessage,
  type StatusMessageProps,
  ZoomControl,
  type ZoomControlProps,
} from './chrome.tsx';
export { FormulaBar, type FormulaBarProps } from './FormulaBar.tsx';
export { SheetTabs, type SheetTabsProps } from './SheetTabs.tsx';
export {
  LocaleProvider,
  type LocaleProviderProps,
  useChromeDir,
  useChromeTranslate,
} from './locale.tsx';
export type { XlsxEditorToolbarNamespace, XlsxEditorToolbarProps } from './toolbar/Toolbar.tsx';
export type {
  XlsxEditorContentProps,
  XlsxEditorFonts,
  XlsxEditorRef,
  XlsxEditorRootProps,
  XlsxEditorViewportProps,
} from './XlsxEditorRoot.tsx';
export {
  TranslationContext,
  translatorFor,
  useEditorSnapshot,
  useEditorState,
  useTranslation,
  useViewModeWaits,
  useXlsxEditor,
  XlsxEditorContext,
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
export { type EditorCommandState, useEditorCommand } from './useEditorCommand.ts';
export { useEditorEvent } from './customization-hooks.ts';
export type {
  ActiveCellFormat,
  CellEditMode,
  ChromeDisabledReason,
  ChromeGroup,
  ChromeGroupId,
  ChromeHint,
  ChromeSlotId,
  ChromeTranslate,
  ClipboardContent,
  CopyResult,
  DecimalSteps,
  DocumentChange,
  EditMove,
  EditorEvents,
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
