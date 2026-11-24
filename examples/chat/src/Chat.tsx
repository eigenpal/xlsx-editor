/** @jsxImportSource react */
// The chat column: the thread, the workbook card that an assistant message attaches, and the
// composer. The card reads the editor state live, so a new sheet, a sheet switch, or an edit in
// the panel shows in the chat too. Adapted from pptx-editor/examples/chat/src/Chat.tsx at
// e0ac8351 (EigenPal).
//
// Differences from PPTX Editor:
// - The card shows the sheet tabs of the workbook, with the shown sheet marked, in place of a
//   slide thumbnail: the editor has no thumbnail of a sheet.
// - The icons are the chat's own inline SVG (`icons.tsx`): the React adapter exports no `Icon`.

import {
  forwardRef,
  memo,
  useCallback,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  type FormEvent,
  type KeyboardEvent,
  type Ref,
} from 'react';
import { useEditorState, type EditorSnapshot } from '@xlsx-editor.dev/react';
import { Icon } from './icons.tsx';
import { TEXT } from './text.ts';

export interface Message {
  readonly id: number;
  readonly role: 'user' | 'assistant';
  readonly text: string;
  /** The assistant attached the workbook under its text. */
  readonly workbook?: boolean;
  /** The text streams in; false for messages that were there before the page loaded. */
  readonly stream?: boolean;
}

const reducedMotion = (): boolean => matchMedia('(prefers-reduced-motion: reduce)').matches;

/** Reveal `text` a few words at a time, as a model streams it; `done` runs at the end. */
function useStreamed(text: string, stream: boolean, done: () => void): string {
  const [shown, setShown] = useState(stream && !reducedMotion() ? 0 : text.length);
  const finished = useRef(done);
  finished.current = done;
  useEffect(() => {
    if (shown >= text.length) {
      finished.current();
      return;
    }
    // Break at spaces, so a word never shows half-typed.
    const timer = setTimeout(() => {
      const next = text.indexOf(' ', shown + 6);
      setShown(next === -1 ? text.length : next);
    }, 28);
    return () => clearTimeout(timer);
  }, [shown, text]);
  return text.slice(0, shown);
}

/** The most tabs that the card shows; the rest is a count. */
const CARD_TABS = 4;

/** The names of the sheets that have a tab, in one string, so the selector result compares by value. */
const tabNames = (s: EditorSnapshot): string =>
  s.sheets
    .filter((sheet) => sheet.visible)
    .map((sheet) => sheet.name)
    .join('\n');

const activeName = (s: EditorSnapshot): string | null => s.sheets[s.activeSheet]?.name ?? null;

/** The workbook's sheet tabs, as a strip on the card, from the live editor state. */
function SheetStrip() {
  const names = useEditorState(tabNames);
  const active = useEditorState(activeName);
  const ready = useEditorState((s) => s.status === 'ready');
  const tabs = names ? names.split('\n') : [];
  const shown = tabs.slice(0, CARD_TABS);
  return (
    <span className="card__sheet" data-ready={ready || undefined} aria-hidden="true">
      <span className="card__grid" />
      <span className="card__tabs">
        {shown.map((name) => (
          <span key={name} className="card__tab" data-active={name === active || undefined}>
            {name}
          </span>
        ))}
        {tabs.length > CARD_TABS ? (
          <span className="card__tab card__tab--more">
            {TEXT.moreSheets(tabs.length - CARD_TABS)}
          </span>
        ) : null}
      </span>
    </span>
  );
}

const sheetCount = (s: EditorSnapshot): number => s.sheets.filter((sheet) => sheet.visible).length;

export const WorkbookCard = forwardRef<
  HTMLButtonElement,
  { title: string; open: boolean; onOpen: () => void }
>(function WorkbookCard({ title, open, onOpen }, ref) {
  const count = useEditorState(sheetCount);
  const edited = useEditorState((s) => s.revision > 0);
  return (
    <button
      ref={ref}
      type="button"
      className="card"
      data-open={open || undefined}
      aria-label={TEXT.openWorkbookLabel(title)}
      aria-expanded={open}
      onClick={onOpen}
    >
      <SheetStrip />
      <span className="card__foot">
        <span className="card__glyph" aria-hidden="true">
          <Icon name="table" />
        </span>
        <span className="card__body">
          <span className="card__title">{title}</span>
          <span className="card__meta">
            {count ? TEXT.sheets(count) : TEXT.workbook}
            {edited ? ` · ${TEXT.edited}` : ''}
          </span>
        </span>
        <span className="card__cta">{open ? TEXT.workbookOpen : TEXT.openWorkbook}</span>
      </span>
    </button>
  );
});

const MessageView = memo(function MessageView({
  message,
  workbookTitle,
  workbookOpen,
  onOpenWorkbook,
  onStreamed,
  cardRef,
}: {
  message: Message;
  workbookTitle: string;
  workbookOpen: boolean;
  onOpenWorkbook: () => void;
  onStreamed: (id: number) => void;
  cardRef: Ref<HTMLButtonElement>;
}) {
  const [streamed, setStreamed] = useState(!message.stream);
  const text = useStreamed(message.text, !!message.stream, () => {
    if (streamed) return;
    setStreamed(true);
    onStreamed(message.id);
  });
  if (message.role === 'user')
    return (
      <div className="msg msg--user">
        <span className="msg__avatar msg__avatar--user" aria-hidden="true">
          {TEXT.youInitial}
        </span>
        <div className="msg__content">
          <span className="msg__name">{TEXT.you}</span>
          <p className="msg__text">{message.text}</p>
        </div>
      </div>
    );
  return (
    <div className="msg msg--assistant">
      <BrandMark />
      <div className="msg__content">
        <span className="msg__name">{TEXT.appName}</span>
        <p className="msg__text">{text}</p>
        {message.workbook && streamed ? (
          <WorkbookCard
            ref={cardRef}
            title={workbookTitle}
            open={workbookOpen}
            onOpen={onOpenWorkbook}
          />
        ) : null}
      </div>
    </div>
  );
});

/** The assistant's mark: a square with the spark, in the accent color. */
export function BrandMark({ busy = false }: { busy?: boolean }) {
  return (
    <span className="brand-mark" data-busy={busy || undefined} aria-hidden="true">
      <Icon name="spark" />
    </span>
  );
}

export function Thinking({ label }: { label: string }) {
  return (
    <div className="msg msg--assistant" role="status">
      <BrandMark busy />
      <div className="msg__content">
        <span className="msg__name">{TEXT.appName}</span>
        <p className="msg__text msg__thinking">{label}</p>
      </div>
    </div>
  );
}

export function Thread({
  messages,
  thinking,
  workbookTitle,
  workbookOpen,
  onOpenWorkbook,
  onStreamed,
  cardRef,
}: {
  messages: readonly Message[];
  thinking: string | null;
  workbookTitle: string;
  workbookOpen: boolean;
  onOpenWorkbook: () => void;
  onStreamed: (id: number) => void;
  cardRef: Ref<HTMLButtonElement>;
}) {
  const scroller = useRef<HTMLDivElement>(null);
  const pinned = useRef(true);
  // A screen reader hears each reply once, whole, when it has streamed in.
  const [spoken, setSpoken] = useState('');
  const streamed = useCallback(
    (id: number) => {
      setSpoken(messages.find((m) => m.id === id)?.text ?? '');
      onStreamed(id);
    },
    [messages, onStreamed]
  );
  // The thread follows new text while the reader is at the bottom, and stays put once they
  // scroll up to read.
  useLayoutEffect(() => {
    const el = scroller.current;
    if (!el) return;
    const follow = () => {
      if (pinned.current) el.scrollTop = el.scrollHeight;
    };
    follow();
    const observer = new ResizeObserver(follow);
    for (const child of Array.from(el.children)) observer.observe(child);
    return () => observer.disconnect();
  }, [messages.length, thinking]);
  // The newest workbook card is the one that opens the panel.
  const lastWorkbook = messages.reduce((at, m, i) => (m.workbook ? i : at), -1);
  return (
    <div
      ref={scroller}
      className="thread"
      onScroll={(e) => {
        const el = e.currentTarget;
        pinned.current = el.scrollHeight - el.scrollTop - el.clientHeight < 48;
      }}
    >
      <div className="thread__inner">
        {messages.map((m, i) => (
          <MessageView
            key={m.id}
            message={m}
            workbookTitle={workbookTitle}
            workbookOpen={workbookOpen && i === lastWorkbook}
            onOpenWorkbook={onOpenWorkbook}
            onStreamed={streamed}
            cardRef={i === lastWorkbook ? cardRef : null}
          />
        ))}
        {thinking ? <Thinking label={thinking} /> : null}
        <p className="sr-only" role="status">
          {spoken}
        </p>
      </div>
    </div>
  );
}

export function EmptyState({ onPick }: { onPick: (text: string) => void }) {
  return (
    <div className="empty">
      <BrandMark />
      <h1 className="empty__title">{TEXT.greeting}</h1>
      <p className="empty__lead">{TEXT.greetingLead}</p>
      <div className="empty__ideas">
        {TEXT.suggestions.map((s, i) => (
          <button
            key={s}
            type="button"
            className="idea"
            style={{ animationDelay: `${i * 50}ms` }}
            onClick={() => onPick(s)}
          >
            <span className="idea__text">{s}</span>
            <Icon name="chevronRight" />
          </button>
        ))}
      </div>
    </div>
  );
}

export const Composer = forwardRef<
  HTMLTextAreaElement,
  { busy: boolean; onSend: (text: string) => void }
>(function Composer({ busy, onSend }, ref) {
  const [value, setValue] = useState('');
  const ready = value.trim().length > 0 && !busy;
  const submit = (e?: FormEvent) => {
    e?.preventDefault();
    if (!ready) return;
    onSend(value.trim());
    setValue('');
  };
  const onKeyDown = (e: KeyboardEvent<HTMLTextAreaElement>) => {
    // Enter sends; Shift+Enter breaks the line; an input method composition keeps its Enter.
    if (e.key === 'Enter' && !e.shiftKey && !e.nativeEvent.isComposing) submit(e);
  };
  return (
    <form className="composer" onSubmit={submit}>
      <div className="composer__box">
        <textarea
          ref={ref}
          className="composer__input"
          rows={1}
          value={value}
          placeholder={TEXT.composerPlaceholder}
          aria-label={TEXT.composerLabel}
          onChange={(e) => setValue(e.target.value)}
          onKeyDown={onKeyDown}
        />
        <div className="composer__bar">
          <span className="composer__hint">{TEXT.composerHint}</span>
          <button type="submit" className="composer__send" disabled={!ready}>
            {TEXT.send}
            <Icon name="arrowUp" />
          </button>
        </div>
      </div>
      <p className="composer__note">{TEXT.disclaimer}</p>
    </form>
  );
});
