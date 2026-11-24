/** @jsxImportSource react */
// The composed editor: the menu bar, the toolbar, the formula bar, the grid, and a footer with
// the sheet tabs and the zoom. The `menu` prop shows or hides the menu bar, as DOCX Editor's
// `menu` prop does (docx-editor/packages/react/src/components/DocxEditor.tsx at ad686574). The
// `contextMenu` prop mounts the grid context menu, as DOCX Editor's prop of the same name.

import { forwardRef, type CSSProperties } from 'react';
import type { SavedWorkbook } from '@xlsx-editor.dev/core/editor';
import {
  XlsxEditorContent,
  XlsxEditorRoot,
  XlsxEditorViewport,
  type XlsxEditorRef,
  type XlsxEditorRootProps,
} from './XlsxEditorRoot.tsx';
import { StatusMessage, ZoomControl } from './chrome.tsx';
import { FormulaBar } from './FormulaBar.tsx';
import { SheetTabs } from './SheetTabs.tsx';
import { XlsxEditorToolbar } from './toolbar/Toolbar.tsx';
import { useChromeDir } from './locale.tsx';
import { XlsxEditorMenu, type XlsxEditorMenuProps } from './menu/Menu.tsx';
import { XlsxEditorContextMenu, type XlsxEditorContextMenuProps } from './menu/ContextMenu.tsx';

/** Props for `XlsxEditor`: the root's props and the viewport's class. @public */
export interface XlsxEditorProps extends Omit<XlsxEditorRootProps, 'children'> {
  /** Appended after the `xlsx-editor` and `xlsx-editor-viewport` classes. */
  readonly className?: string;
  /** Inline style of the viewport, for example its height. */
  readonly style?: CSSProperties;
  /**
   * Receives the workbook as bytes of the opened package's kind (`snapshot().workbookKind`: a
   * workbook or a template, macro-enabled when `snapshot().macroEnabled` is true) when the
   * toolbar's save control is pressed, or when the user presses Control or Command and S in the
   * editor, also with the save control hidden. Without it, the save control is disabled with the
   * engine's reason, and the browser keeps the shortcut. A refused save calls nothing: the snapshot
   * holds the refusal with `command: 'save'`. The bytes arrive after the press, because a save runs
   * in short tasks. `saved.stillOpen` is false when another open, `close()`, or `destroy()` ended
   * the workbook during the save, and `saved.revision` is the revision of the bytes.
   */
  readonly onSave?: (bytes: Uint8Array<ArrayBuffer>, saved: SavedWorkbook) => void;
  /**
   * Called once when a press of the save control, or of Control or Command and S in the editor,
   * starts a delivery to `onSave`. It gets the promise of the result, so you can show that the
   * save runs. A press that joins the running save of an earlier press and a refused save do
   * not call it. A press that joins a save from `saveInTasks` calls it, because that save has no
   * delivery yet. The promise resolves to `null` when the save is refused later.
   */
  readonly onSaveStart?: (saving: Promise<SavedWorkbook | null>) => void;
  /**
   * Whether your page opens another workbook now, while the editor still shows the old one. While
   * it is true, the save control is disabled and Control or Command and S starts no save. See the
   * `opening` prop of `XlsxEditor.Toolbar`.
   */
  readonly opening?: boolean;
  /**
   * The menu bar above the toolbar: `true` (the default) shows it, `false` leaves it out, and an
   * object shows it with those props, which win over the ones that the editor passes (`onSave`,
   * `onSaveStart`, and `opening`).
   *
   * The default File > Open replaces the open workbook with the picked file and does not ask
   * first, also when the workbook has unsaved edits. To ask first, pass your own `onOpen` in the
   * object; `snapshot().revision` tells you whether the workbook changed since your last save.
   * To remove the row, set `menu` to `false` and compose the menu bar with a hidden `Open` row.
   */
  readonly menu?: boolean | XlsxEditorMenuProps;
  /**
   * The grid context menu: `true` (the default) mounts it, `false` leaves it out, so a secondary
   * press on the grid shows the browser's own menu, and an object mounts it with those props.
   */
  readonly contextMenu?: boolean | XlsxEditorContextMenuProps;
}

function Body({
  onSave,
  onSaveStart,
  opening,
  menu = true,
  contextMenu = true,
}: Pick<XlsxEditorProps, 'onSave' | 'onSaveStart' | 'opening' | 'menu' | 'contextMenu'>) {
  const dir = useChromeDir();
  return (
    <>
      {menu !== false ? (
        <XlsxEditorMenu
          {...(onSave ? { onSave } : {})}
          {...(onSaveStart ? { onSaveStart } : {})}
          {...(opening === undefined ? {} : { opening })}
          {...(typeof menu === 'object' ? menu : {})}
        />
      ) : null}
      <XlsxEditorToolbar
        {...(onSave ? { onSave } : {})}
        {...(onSaveStart ? { onSaveStart } : {})}
        {...(opening === undefined ? {} : { opening })}
      />
      <FormulaBar />
      <StatusMessage />
      <XlsxEditorContent />
      <div className="xlsx-footer" dir={dir}>
        <SheetTabs />
        <ZoomControl />
      </div>
      {contextMenu !== false ? (
        <XlsxEditorContextMenu {...(typeof contextMenu === 'object' ? contextMenu : {})} />
      ) : null}
    </>
  );
}

/** A workbook editor with its chrome. Takes a `ref` to an `XlsxEditorRef`. @public */
export const XlsxEditor = forwardRef<XlsxEditorRef, XlsxEditorProps>(function XlsxEditor(
  { className, style, onSave, onSaveStart, opening, menu, contextMenu, ...props },
  ref
) {
  return (
    <XlsxEditorRoot ref={ref} {...props}>
      <XlsxEditorViewport {...(className ? { className } : {})} {...(style ? { style } : {})}>
        <Body
          {...(onSave ? { onSave } : {})}
          {...(onSaveStart ? { onSaveStart } : {})}
          {...(opening === undefined ? {} : { opening })}
          {...(menu === undefined ? {} : { menu })}
          {...(contextMenu === undefined ? {} : { contextMenu })}
        />
      </XlsxEditorViewport>
    </XlsxEditorRoot>
  );
});
