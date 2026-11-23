/** @jsxImportSource react */
// The editor chrome: the zoom control and the status messages (the formula bar is in
// `FormulaBar.tsx`). Each part reads its slice of the editor snapshot
// (`useEditorState`), so it renders only when that slice changes, and calls the editor; none
// holds state of its own. Chrome `mousedown` keeps the keyboard on the grid
// (AGENTS.md "Conventions", focus stealing). Each part takes DOCX Editor's `t` prop and sets the
// chrome direction on its element (`locale.tsx`).

import { useContext, type ReactElement } from 'react';
import {
  MAX_ZOOM,
  MIN_ZOOM,
  keepGridFocus,
  sameSlice,
  statusSlice,
  statusView,
  zoomSlice,
  type ChromeTranslate,
} from '@xlsx-editor.dev/core/editor';
import {
  ModeWaitsContext,
  useLayoutPart,
  useTranslation,
  useEditorState,
  useXlsxEditor,
} from './context.ts';
import { ZoomInIcon, ZoomOutIcon } from './icons.tsx';
import { useChromeDir, withChromeTranslate } from './locale.tsx';

/** Props for `XlsxEditor.ZoomControl`. @public */
export interface ZoomControlProps {
  /**
   * The label resolver of the control: a key and the values of its placeholders, as DOCX
   * Editor's `t` prop. A key that it leaves unresolved reads the root's `locale`. Keep it stable,
   * for example from `useChromeTranslate(overrides)`.
   */
  readonly t?: ChromeTranslate;
}

/** Props for `XlsxEditor.StatusMessage`. @public */
export interface StatusMessageProps {
  /**
   * The label resolver of the messages: a key and the values of its placeholders, as DOCX
   * Editor's `t` prop. A key that it leaves unresolved reads the root's `locale`. Keep it stable,
   * for example from `useChromeTranslate(overrides)`.
   */
  readonly t?: ChromeTranslate;
}

function ZoomControlPart() {
  const { t } = useTranslation();
  const dir = useChromeDir();
  const editor = useXlsxEditor();
  const snap = useEditorState(zoomSlice, sameSlice);
  if (snap.status !== 'ready') return null;
  const level = `${Math.round(snap.zoom * 100)}%`;
  // At a limit the button stays focusable with `aria-disabled` and its press does nothing, as
  // the toolbar's controls do: a `disabled` button would drop the focus to the page.
  const atMin = snap.zoom <= MIN_ZOOM;
  const atMax = snap.zoom >= MAX_ZOOM;
  return (
    <div
      className="xlsx-zoom"
      dir={dir}
      role="group"
      aria-label={t('zoom.zoomLevel')}
      onMouseDown={keepGridFocus}
    >
      <button
        type="button"
        aria-label={t('zoom.zoomOut')}
        title={t('zoom.zoomOut')}
        aria-disabled={atMin}
        onClick={() => {
          if (!atMin) editor?.zoomStep(-1);
        }}
      >
        <ZoomOutIcon />
      </button>
      <button
        type="button"
        className="xlsx-zoom-level"
        aria-label={t('zoom.resetAt', { level })}
        title={t('zoom.reset')}
        onClick={() => editor?.setZoom(1)}
      >
        {level}
      </button>
      <button
        type="button"
        aria-label={t('zoom.zoomIn')}
        title={t('zoom.zoomIn')}
        aria-disabled={atMax}
        onClick={() => {
          if (!atMax) editor?.zoomStep(1);
        }}
      >
        <ZoomInIcon />
      </button>
    </div>
  );
}

/** Zoom out, the zoom level (press to reset to 100%), and zoom in. @public */
export const ZoomControl: (props: ZoomControlProps) => ReactElement = withChromeTranslate(
  ZoomControlPart,
  'ZoomControl'
);

/**
 * One notice of the status strip: a live status node whose text is the full text while `on`.
 * `short` shows the brief text in its place, visually only: the full text stays in the node,
 * hidden from view, and the brief text is hidden from assistive technology.
 */
function Notice(props: { on: boolean; short: boolean; text: string; brief: string }) {
  const { on, short, text, brief } = props;
  return (
    <div
      className="xlsx-notice"
      role="status"
      {...(on ? { title: text } : {})}
      {...(on && short ? { 'data-short': '' } : {})}
    >
      {on ? (
        <>
          <span className="xlsx-notice-full">{text}</span>
          <span className="xlsx-notice-brief" aria-hidden="true">
            {brief}
          </span>
        </>
      ) : null}
    </div>
  );
}

function StatusMessagePart() {
  const { t } = useTranslation();
  const dir = useChromeDir();
  const snap = useEditorState(statusSlice, sameSlice);
  const modeWaits = useContext(ModeWaitsContext);
  useLayoutPart('statusMessages');
  // The status strip is present while a workbook is ready, one line high and empty when there
  // is no message, so nothing covers the grid. A refusal shows in full on its own row, and the
  // strip grows while it shows. When more than one message shows, a notice shows its short
  // text and keeps its full text in `title`. The short text is visual only (`Notice`), so a
  // notice's live text never changes while its state stays the same, and a screen reader
  // announces it once. The stale notice, the protection of the shown sheet, the conditional
  // formatting that the grid can leave out, the parts past a limit of the read, and a wait of
  // the `mode` prop stay mounted while the workbook is open (`statusView`).
  const view = statusView(snap, modeWaits, t);
  if (view.kind === 'strip')
    return (
      <div className="xlsx-status-strip" dir={dir}>
        {view.refusal ? <div className="xlsx-refusal">{view.refusal}</div> : null}
        {view.notices.map((n) => (
          <Notice key={n.id} on={n.on} short={n.short} text={n.text} brief={n.brief} />
        ))}
      </div>
    );
  // Without a grid, the live region announced each refusal already, so it is plain text.
  return (
    <div className="xlsx-message" dir={dir}>
      <div role={view.role}>{view.text}</div>
      {view.refusals.map((text) => (
        <div key={text} className="xlsx-message-refusal">
          {text}
        </div>
      ))}
    </div>
  );
}

/**
 * The state of the editor when it shows no grid, the reason the last edit was refused, the
 * stale-results notice, and the notice for conditional formatting that the grid can leave out
 * (`conditionalHidden`). The engine's live region announces the refusal, so the refusal here
 * has no live role: it is shown, not announced twice. @public
 */
export const StatusMessage: (props: StatusMessageProps) => ReactElement = withChromeTranslate(
  StatusMessagePart,
  'StatusMessage'
);
