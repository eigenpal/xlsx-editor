/** @jsxImportSource react */
// The workbook panel beside the chat: a compact header with the workbook's own actions, the
// editor's toolbar, the formula bar, the grid, and a footer with the sheet tabs and the zoom.
// Every part is a public part of `@xlsx-editor.dev/react`; the header's buttons are
// `XlsxEditor.Toolbar.Action`s, so they get the editor's tooltip, focus guard, and disabled
// state with no private class name. Adapted from pptx-editor/examples/chat/src/DeckPanel.tsx at
// e0ac8351 (EigenPal).
//
// Differences from PPTX Editor:
// - A grid scrolls in both directions and has no slide to fit, so the panel has one layout at
//   every width: the toolbar moves its groups into its More panel as the panel narrows.
// - The header has a View only switch, which drives the root's `mode` prop. Expand changes its
//   label to Show chat and has no pressed state, so a screen reader hears one state, not two.
// - The toolbar leaves out Open and Save: the chat owns where a workbook comes from, and the
//   header's Download saves. Control+S in the grid still downloads, through `onSave`.
// - A phone-wide panel drops the zoom control, so the sheet tabs keep the footer, as the demo
//   does. A narrow toolbar needs nothing from the host: the bar moves the font family field into
//   its More panel first, and keeps size, bold, italic, underline, and color.

import { useLayoutEffect, useRef, useState, type ReactNode, type RefObject } from 'react';
import {
  Icon,
  XlsxEditor,
  useChromeDir,
  useEditorState,
  type ChromeSlotId,
  type EditorSnapshot,
} from '@xlsx-editor.dev/react';
import { PATHS } from './icon-paths.ts';
import { TEXT } from './text.ts';

/** The toolbar slots that the panel never shows: the chat owns open and save. */
const HIDDEN: readonly ChromeSlotId[] = ['file.open', 'file.save'];
/** Under this panel width, the footer drops the zoom control. */
const PHONE = 480;

/** The width of an element, from a resize observer, and 0 before the first layout. */
function useWidth(): [RefObject<HTMLDivElement | null>, number] {
  const box = useRef<HTMLDivElement>(null);
  const [width, setWidth] = useState(0);
  useLayoutEffect(() => {
    const el = box.current;
    if (!el) return;
    const observer = new ResizeObserver(([entry]) => setWidth(entry?.contentRect.width ?? 0));
    observer.observe(el);
    return () => observer.disconnect();
  }, []);
  return [box, width];
}

const sheetCount = (s: EditorSnapshot): number => s.sheets.filter((sheet) => sheet.visible).length;

function HeaderAction({
  label,
  icon,
  active,
  disabled,
  onSelect,
}: {
  label: string;
  icon: ReactNode;
  active?: boolean;
  /** Off until the workbook is ready; the reason is the tooltip. */
  disabled?: boolean;
  onSelect: () => void;
}) {
  return (
    <XlsxEditor.Toolbar.Action
      label={label}
      className="book__action"
      icon={icon}
      onSelect={onSelect}
      {...(active === undefined ? {} : { active })}
      {...(disabled ? { disabled, disabledReason: TEXT.notReady } : {})}
    />
  );
}

export function WorkbookPanel({
  title,
  expanded,
  sheet,
  mode,
  onClose,
  onExpand,
  onToggleMode,
  onDownload,
  onSave,
}: {
  title: string;
  expanded: boolean;
  /** The panel covers the whole window (a phone): Back replaces Expand and Close. */
  sheet: boolean;
  mode: 'edit' | 'view';
  onClose: () => void;
  onExpand: () => void;
  onToggleMode: () => void;
  onDownload: () => void;
  onSave: (bytes: Uint8Array<ArrayBuffer>) => void;
}) {
  const count = useEditorState(sheetCount);
  const edited = useEditorState((s) => s.revision > 0);
  const ready = useEditorState((s) => s.status === 'ready');
  const dir = useChromeDir();
  const viewOnly = mode === 'view';
  const [box, width] = useWidth();
  const phone = width > 0 && width < PHONE;
  return (
    // `chat-theme` scopes the chat's token map (styles.css). The panel `<section>` of App.tsx is
    // the named region, so the viewport adds none (`label={false}`).
    <XlsxEditor.Viewport className="book chat-theme" label={false}>
      <header className="book__head">
        {sheet ? (
          <HeaderAction label={TEXT.back} icon={<Icon paths={PATHS.back} />} onSelect={onClose} />
        ) : null}
        <span className="book__glyph" aria-hidden="true">
          <Icon paths={PATHS.table} />
        </span>
        <div className="book__title">
          <span className="book__name">{title}</span>
          <span className="book__meta">
            {count ? TEXT.sheets(count) : TEXT.workbook}
            {viewOnly ? ` · ${TEXT.viewOnly}` : ''}
            {/* The word swaps in place; the dot and the word fade together. */}
            <span className="book__edited" data-shown={(edited && !viewOnly) || undefined}>
              {' · '}
              {TEXT.edited}
            </span>
          </span>
        </div>
        <div className="book__actions">
          <HeaderAction
            label={TEXT.lock}
            icon={<Icon paths={PATHS.lock} />}
            active={viewOnly}
            disabled={!ready}
            onSelect={onToggleMode}
          />
          <HeaderAction
            label={TEXT.download}
            icon={<Icon name="file_download" />}
            disabled={!ready}
            onSelect={onDownload}
          />
          {sheet ? null : (
            <>
              <HeaderAction
                label={expanded ? TEXT.collapse : TEXT.expand}
                icon={<Icon paths={expanded ? PATHS.collapse : PATHS.expand} />}
                onSelect={onExpand}
              />
              <HeaderAction
                label={TEXT.closeWorkbook}
                icon={<Icon paths={PATHS.close} />}
                onSelect={onClose}
              />
            </>
          )}
        </div>
      </header>
      <div ref={box} className="book__dock">
        <XlsxEditor.Toolbar className="book__toolbar" hidden={HIDDEN} onSave={onSave} />
      </div>
      <div className="book__sheet">
        <XlsxEditor.FormulaBar />
        <XlsxEditor.StatusMessage />
        <XlsxEditor.Content />
        <div className="xlsx-footer" dir={dir}>
          <XlsxEditor.SheetTabs />
          {phone ? null : <XlsxEditor.ZoomControl />}
        </div>
      </div>
      <XlsxEditor.ContextMenu />
    </XlsxEditor.Viewport>
  );
}
