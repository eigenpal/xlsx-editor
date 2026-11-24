/** @jsxImportSource react */
// A chat app with a workbook panel, as a host would build it: the chat on the left, and the
// editor in a panel that slides in from the right when you open the workbook card. One
// `XlsxEditor.Root` wraps both, so the chat reads the workbook (the live card) and acts on it
// (go to a sheet, select a range, fill a cell, download) through the public API. Adapted from
// pptx-editor/examples/chat/src/App.tsx at e0ac8351 (EigenPal).
//
// Differences from PPTX Editor:
// - The actions are workbook commands of the editor instance, and the View only switch drives
//   the root's `mode` prop.
// - The download saves with `saveInTasks`, so a large workbook does not hold the page.

import { useCallback, useEffect, useRef, useState, type CSSProperties } from 'react';
import { flushSync } from 'react-dom';
import {
  XlsxEditor,
  refusalText,
  useEditorEvent,
  useTranslation,
  useXlsxEditor,
  type EditorMode,
  type EditResult,
  type XlsxEditorInstance,
} from '@xlsx-editor.dev/react';
import {
  FIRST_WORKBOOK_TEXT,
  reply,
  type AssistantAction,
  type WorkbookContext,
} from './assistant.ts';
import { Composer, EmptyState, Thread, type Message } from './Chat.tsx';
import { WorkbookPanel } from './WorkbookPanel.tsx';
import { downloadBlob } from '../../vite/src/download.ts';
import { usePanelLayout } from './panel-layout.ts';
import { Icon } from './icons.tsx';
import { TEXT } from './text.ts';

const SAMPLE_URL = `${import.meta.env.BASE_URL}sample.xlsx`;
/** The sheet that the chat shows first: a table, not the demo's welcome page. */
const FIRST_SHEET = 'Sales';
const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/** Saves the workbook bytes as `<name>.xlsx`, with the demo's checked download link. */
const downloadWorkbook = (bytes: Uint8Array<ArrayBuffer>, name: string): void =>
  downloadBlob(new Blob([bytes], { type: XLSX_TYPE }), `${name}.xlsx`);

/** The conversation that is there when the page opens, so the workbook card shows at once. */
const SEED: readonly Message[] = [
  { id: 1, role: 'user', text: TEXT.seedPrompt },
  { id: 2, role: 'assistant', text: FIRST_WORKBOOK_TEXT, workbook: true },
];

/** What the assistant sees of the editor: the sheets, the shown sheet, and the selection. */
function contextOf(editor: XlsxEditorInstance | null, hasWorkbook: boolean): WorkbookContext {
  const s = editor?.snapshot();
  return {
    hasWorkbook,
    sheets: (s?.sheets ?? [])
      .filter((sheet) => sheet.visible && sheet.worksheet)
      .map(({ index, name }) => ({ index, name })),
    activeSheet: s?.sheets[s.activeSheet]?.name ?? null,
    selection: s?.selectionRef ?? null,
  };
}

function ChatApp({
  loadWorkbook,
  mode,
  setMode,
}: {
  loadWorkbook: () => void;
  mode: EditorMode;
  setMode: (mode: EditorMode) => void;
}) {
  const editor = useXlsxEditor();
  const { t } = useTranslation();
  const [messages, setMessages] = useState<readonly Message[]>(SEED);
  const [chatTitle, setChatTitle] = useState<string>(TEXT.seedTitle);
  const [thinking, setThinking] = useState<string | null>(null);
  const [streaming, setStreaming] = useState(false);
  const layout = usePanelLayout();
  const { open, setOpen } = layout;
  const card = useRef<HTMLButtonElement>(null);
  const composer = useRef<HTMLTextAreaElement>(null);
  const panel = useRef<HTMLElement>(null);
  const hasWorkbook = messages.some((m) => m.workbook);
  const nextId = useRef(SEED.length + 1);
  /** The reply that is on its way: New chat cancels it, so it never lands in the new chat. */
  const pending = useRef<number | null>(null);
  const cancelReply = useCallback(() => {
    if (pending.current !== null) window.clearTimeout(pending.current);
    pending.current = null;
  }, []);
  useEffect(() => cancelReply, [cancelReply]);

  const closePanel = useCallback(() => {
    // Focus inside the panel goes back to the card that opened it.
    const active = document.activeElement;
    const inside = !active || active === document.body || active.closest('.shell__panel');
    // Render the close first: an expanded panel or a sheet keeps the chat inert until then,
    // and an inert element does not take focus.
    flushSync(() => setOpen(false));
    if (inside) (card.current ?? composer.current)?.focus({ preventScroll: true });
  }, [setOpen]);

  const download = useCallback(() => {
    if (!editor) return;
    void editor.saveInTasks().then((saved) => {
      if (saved?.stillOpen) downloadWorkbook(saved.bytes, TEXT.workbookName);
    });
  }, [editor]);
  const saveFromKeys = useCallback(
    (bytes: Uint8Array<ArrayBuffer>) => downloadWorkbook(bytes, TEXT.workbookName),
    []
  );
  const toggleMode = useCallback(() => setMode(mode === 'view' ? 'edit' : 'view'), [mode, setMode]);

  /** Runs the assistant's action, and returns the editor's result when it is an edit. */
  const run = useCallback(
    (action: AssistantAction | undefined): EditResult | null => {
      if (!action || !editor) return null;
      switch (action.kind) {
        case 'open':
          setOpen(true);
          return null;
        case 'sheet':
          setOpen(true);
          editor.setActiveSheet(action.index);
          return null;
        case 'select':
          setOpen(true);
          return editor.selectReference(action.ref);
        case 'write': {
          setOpen(true);
          const at = editor.selectReference(action.ref);
          return at.ok ? editor.enterValue(action.text) : at;
        }
        case 'format':
          setOpen(true);
          return editor.toggleFormat(action.toggle);
        case 'addSheet':
          setOpen(true);
          return editor.addSheet();
        case 'undo':
          return editor.undo();
        case 'redo':
          return editor.redo();
        case 'zoom':
          setOpen(true);
          editor.zoomStep(action.direction);
          return null;
        case 'mode':
          setMode(action.mode);
          return null;
        case 'download':
          download();
          return null;
      }
    },
    [editor, download, setOpen, setMode]
  );

  const send = useCallback(
    (text: string) => {
      const answer = reply(text, contextOf(editor, hasWorkbook));
      if (!messages.length) setChatTitle(text.length > 34 ? `${text.slice(0, 32)}…` : text);
      setMessages((m) => [...m, { id: nextId.current++, role: 'user', text }]);
      setThinking(answer.workbook ? TEXT.thinking : TEXT.working);
      setStreaming(true);
      // A new workbook starts from fresh bytes while the assistant "thinks".
      if (answer.workbook) loadWorkbook();
      cancelReply();
      pending.current = window.setTimeout(
        () => {
          pending.current = null;
          setThinking(null);
          // The editor has the last word: a refused edit replaces the planned answer with the
          // editor's own reason, in the editor's catalog.
          const result = run(answer.action);
          const refused =
            result && !result.ok
              ? TEXT.refused(refusalText(result.refusal, t, editor?.snapshot().numberLocale))
              : null;
          setMessages((m) => [
            ...m,
            {
              id: nextId.current++,
              role: 'assistant',
              text: refused ?? answer.text,
              stream: true,
              ...(answer.workbook ? { workbook: true } : {}),
            },
          ]);
        },
        answer.workbook ? 1100 : 450
      );
    },
    [editor, hasWorkbook, messages.length, loadWorkbook, run, cancelReply, t]
  );

  // Each open shows the table first. A host picks the first sheet; the file keeps its own.
  useEditorEvent('change', (change) => {
    if (change.source !== 'load' || !editor) return;
    const sheet = editor.snapshot().sheets.find((s) => s.name === FIRST_SHEET);
    if (sheet) editor.setActiveSheet(sheet.index);
  });

  const onStreamed = useCallback(() => setStreaming(false), []);

  const newChat = useCallback(() => {
    cancelReply();
    setOpen(false);
    setMessages([]);
    setChatTitle(TEXT.newChat);
    setThinking(null);
    setStreaming(false);
    composer.current?.focus();
  }, [setOpen, cancelReply]);

  // A sheet covers the chat, so the focus goes into it: to Back, the way out.
  useEffect(() => {
    if (!open || !layout.sheet) return;
    const frame = requestAnimationFrame(() =>
      // A pointer opened the sheet: no ring on Back. A keyboard user still gets one on Tab.
      panel.current
        ?.querySelector<HTMLElement>('.book__head button')
        ?.focus({ focusVisible: false } as FocusOptions)
    );
    return () => cancelAnimationFrame(frame);
  }, [open, layout.sheet]);

  // Expand makes the chat inert: focus that was in it moves to the panel's Show chat button.
  useEffect(() => {
    if (!open || !layout.expanded) return;
    const active = document.activeElement;
    if (active && active !== document.body && !active.closest('.shell__chat')) return;
    panel.current
      ?.querySelector<HTMLElement>(`.book__head [aria-label="${TEXT.collapse}"]`)
      ?.focus();
  }, [open, layout.expanded]);

  // Escape closes the panel from the chat, the panel header, or the divider. Escape in the
  // grid, the formula bar, the toolbar, or a menu belongs to the editor: it cancels a cell edit
  // or closes a list.
  useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent) => {
      // An input method keeps the Escape that ends its composition.
      if (e.key !== 'Escape' || e.defaultPrevented || e.isComposing || e.keyCode === 229) return;
      const target = e.target as Element | null;
      const chrome = target?.closest?.('.book__head, .shell__handle');
      if (!chrome && target?.closest?.('.shell__panel')) return;
      closePanel();
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [open, closePanel]);

  const shellStyle = { '--panel-w': `${layout.panelWidth}px` } as CSSProperties;
  return (
    <div
      ref={layout.shell}
      className="shell"
      style={shellStyle}
      data-open={open || undefined}
      data-expanded={layout.expanded || undefined}
      data-sheet={layout.sheet || undefined}
      data-resizing={layout.resizing || undefined}
    >
      <main
        className="shell__chat"
        inert={(open && (layout.expanded || layout.sheet)) || undefined}
      >
        <header className="chat-head">
          {/* The empty state has its own heading; a thread makes the app name the page heading. */}
          {messages.length ? (
            <h1 className="chat-head__brand">{TEXT.appName}</h1>
          ) : (
            <span className="chat-head__brand">{TEXT.appName}</span>
          )}
          {messages.length ? (
            <>
              <span className="chat-head__sep" aria-hidden="true">
                /
              </span>
              <span className="chat-head__title">{chatTitle}</span>
            </>
          ) : null}
          <button type="button" className="chat-head__new" onClick={newChat}>
            <Icon name="editSquare" />
            <span className="chat-head__new-label">{TEXT.newChat}</span>
          </button>
        </header>
        {messages.length ? (
          <Thread
            messages={messages}
            thinking={thinking}
            workbookTitle={TEXT.workbookTitle}
            workbookOpen={open}
            // The card of an open workbook moves the focus to the grid, ready to type.
            onOpenWorkbook={() => (open ? editor?.focus() : setOpen(true))}
            onStreamed={onStreamed}
            cardRef={card}
          />
        ) : (
          <EmptyState onPick={send} />
        )}
        <Composer ref={composer} busy={!!thinking || streaming} onSend={send} />
      </main>
      <div className="shell__spacer" aria-hidden="true" />
      <section
        ref={panel}
        className="shell__panel"
        aria-label={TEXT.workbookTitle}
        inert={!open || undefined}
      >
        <div
          className="shell__handle"
          role="separator"
          aria-orientation="vertical"
          aria-label={TEXT.resize}
          aria-valuemin={Math.round(layout.minFraction * 100)}
          aria-valuemax={100}
          aria-valuenow={Math.round(layout.fraction * 100)}
          tabIndex={0}
          {...layout.handleProps}
        />
        {/* Mounted while closed, so the workbook is ready the moment the panel opens. */}
        <WorkbookPanel
          title={TEXT.workbookTitle}
          expanded={layout.expanded}
          sheet={layout.sheet}
          mode={mode}
          onClose={closePanel}
          onExpand={layout.toggleExpanded}
          onToggleMode={toggleMode}
          onDownload={download}
          onSave={saveFromKeys}
        />
      </section>
    </div>
  );
}

export function App() {
  const [bytes, setBytes] = useState<ArrayBuffer | undefined>();
  const [mode, setMode] = useState<EditorMode>('edit');
  const loadWorkbook = useCallback(() => {
    void fetch(SAMPLE_URL)
      .then((response) => {
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        return response.arrayBuffer();
      })
      .then(setBytes)
      .catch(() => setBytes(undefined));
  }, []);
  useEffect(loadWorkbook, [loadWorkbook]);
  return (
    <XlsxEditor.Root
      {...(bytes ? { document: bytes } : {})}
      mode={mode}
      colorMode="system"
      // The browser check (`scripts/verify-chat.ts`) reads the editor in development builds. A
      // host does not put the editor on `window`.
      {...(import.meta.env.DEV
        ? {
            onReady: (instance: XlsxEditorInstance) => {
              (window as unknown as { __xlsxEditor?: unknown }).__xlsxEditor = instance;
            },
          }
        : {})}
    >
      <ChatApp loadWorkbook={loadWorkbook} mode={mode} setMode={setMode} />
    </XlsxEditor.Root>
  );
}
