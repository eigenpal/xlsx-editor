/** @jsxImportSource react */
// The composed editor: the toolbar, the formula bar, the grid, and a footer with the sheet tabs
// and the zoom.

import { forwardRef, type CSSProperties } from 'react';
import type { EditorSnapshot, SavedWorkbook } from '@xlsx-editor.dev/core/editor';
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
import { useEditorState } from './context.ts';
import { XlsxEditorToolbar } from './toolbar/Toolbar.tsx';
import { useChromeDir } from './locale.tsx';

/** Props for `XlsxEditor`: the root's props and the viewport's class. @public */
export interface XlsxEditorProps extends Omit<XlsxEditorRootProps, 'children'> {
  /** Appended after the `xlsx-editor` and `xlsx-editor-viewport` classes. */
  readonly className?: string;
  /** Inline style of the viewport, for example its height. */
  readonly style?: CSSProperties;
  /**
   * Receives the workbook as bytes of the opened package's kind (`snapshot().workbookKind`: a
   * workbook or a template) when the toolbar's save control is pressed, or when the user presses
   * Control or Command and S in the editor, also with the save control hidden. Without it, the
   * save control is disabled with the engine's reason, and the browser keeps the shortcut. A
   * refused save calls nothing: the snapshot holds the refusal with `command: 'save'`. The bytes
   * arrive after the press, because a save runs in short tasks. `saved.stillOpen` is false when
   * another open, `close()`, or `destroy()` ended the workbook during the save, and
   * `saved.revision` is the revision of the bytes.
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
}

const statusOf = (snapshot: EditorSnapshot) => snapshot.status;

function Body({
  onSave,
  onSaveStart,
  opening,
}: Pick<XlsxEditorProps, 'onSave' | 'onSaveStart' | 'opening'>) {
  // The status alone: the body renders again only when it changes, not on each key of an edit.
  const status = useEditorState(statusOf);
  const dir = useChromeDir();
  return (
    <>
      <XlsxEditorToolbar
        {...(onSave ? { onSave } : {})}
        {...(onSaveStart ? { onSaveStart } : {})}
        {...(opening === undefined ? {} : { opening })}
      />
      <FormulaBar />
      <StatusMessage />
      <XlsxEditorContent className={status === 'ready' ? undefined : 'xlsx-canvas-idle'} />
      <div className="xlsx-footer" dir={dir}>
        <SheetTabs />
        <ZoomControl />
      </div>
    </>
  );
}

/** A workbook editor with its chrome. Takes a `ref` to an `XlsxEditorRef`. @public */
export const XlsxEditor = forwardRef<XlsxEditorRef, XlsxEditorProps>(function XlsxEditor(
  { className, style, onSave, onSaveStart, opening, ...props },
  ref
) {
  return (
    <XlsxEditorRoot ref={ref} {...props}>
      <XlsxEditorViewport {...(className ? { className } : {})} {...(style ? { style } : {})}>
        <Body
          {...(onSave ? { onSave } : {})}
          {...(onSaveStart ? { onSaveStart } : {})}
          {...(opening === undefined ? {} : { opening })}
        />
      </XlsxEditorViewport>
    </XlsxEditorRoot>
  );
});
