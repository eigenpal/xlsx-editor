/** @jsxImportSource vue */
// The editor chrome: the zoom control and the status messages (the formula bar is in
// `FormulaBar.tsx`). The Vue twin of the React adapter's `chrome.tsx`. Each part reads its
// slice of the editor snapshot (`useEditorState`), so it renders only when that slice changes,
// and calls the editor; none holds state of its own. Chrome `mousedown` keeps the keyboard on
// the grid (AGENTS.md "Conventions", focus stealing).

import { defineComponent } from 'vue';
import {
  MAX_ZOOM,
  MIN_ZOOM,
  keepGridFocus,
  sameSlice,
  statusSlice,
  statusView,
  zoomSlice,
} from '@xlsx-editor.dev/core/editor';
import {
  useEditorState,
  useLayoutPart,
  useTranslation,
  useViewModeWaits,
  useXlsxEditor,
} from './context.ts';
import { ZoomInIcon, ZoomOutIcon } from './icons.tsx';

/** Zoom out, the zoom level (press to reset to 100%), and zoom in. @public */
export const ZoomControl = defineComponent({
  name: 'ZoomControl',
  setup() {
    const { t } = useTranslation();
    const editor = useXlsxEditor();
    const snap = useEditorState(zoomSlice, sameSlice);
    return () => {
      const { status, zoom } = snap.value;
      if (status !== 'ready') return null;
      const level = `${Math.round(zoom * 100)}%`;
      // At a limit the button stays focusable with `aria-disabled` and its press does nothing,
      // as the toolbar's controls do: a `disabled` button would drop the focus to the page.
      const atMin = zoom <= MIN_ZOOM;
      const atMax = zoom >= MAX_ZOOM;
      return (
        <div
          class="xlsx-zoom"
          role="group"
          aria-label={t('zoom.zoomLevel')}
          onMousedown={keepGridFocus}
        >
          <button
            type="button"
            aria-label={t('zoom.zoomOut')}
            title={t('zoom.zoomOut')}
            aria-disabled={atMin}
            onClick={() => {
              if (!atMin) editor.value?.zoomStep(-1);
            }}
          >
            <ZoomOutIcon />
          </button>
          <button
            type="button"
            class="xlsx-zoom-level"
            aria-label={t('zoom.resetAt', { level })}
            title={t('zoom.reset')}
            onClick={() => editor.value?.setZoom(1)}
          >
            {level}
          </button>
          <button
            type="button"
            aria-label={t('zoom.zoomIn')}
            title={t('zoom.zoomIn')}
            aria-disabled={atMax}
            onClick={() => {
              if (!atMax) editor.value?.zoomStep(1);
            }}
          >
            <ZoomInIcon />
          </button>
        </div>
      );
    };
  },
});

/**
 * One notice of the status strip: a live status node whose text is the full text while `on`.
 * `short` shows the brief text in its place, visually only: the full text stays in the node,
 * hidden from view, and the brief text is hidden from assistive technology.
 */
function Notice(props: { on: boolean; short: boolean; text: string; brief: string }) {
  const { on, short, text, brief } = props;
  return (
    <div
      class="xlsx-notice"
      role="status"
      {...(on ? { title: text } : {})}
      {...(on && short ? { 'data-short': '' } : {})}
    >
      {on
        ? [
            <span class="xlsx-notice-full">{text}</span>,
            <span class="xlsx-notice-brief" aria-hidden="true">
              {brief}
            </span>,
          ]
        : null}
    </div>
  );
}

/**
 * The state of the editor when it shows no grid, the reason the last edit was refused, the
 * stale-results notice, and the notice for conditional formatting that the grid can leave out
 * (`conditionalHidden`). The engine's live region announces the refusal, so the refusal here
 * has no live role: it is shown, not announced twice. @public
 */
export const StatusMessage = defineComponent({
  name: 'StatusMessage',
  setup() {
    const { t } = useTranslation();
    const snap = useEditorState(statusSlice, sameSlice);
    const modeWaits = useViewModeWaits();
    useLayoutPart('statusMessages');
    // The status strip is present while a workbook is ready, one line high and empty when there
    // is no message, so nothing covers the grid. A refusal shows in full on its own row, and the
    // strip grows while it shows. When more than one message shows, a notice shows its short
    // text and keeps its full text in `title`. The short text is visual only (`Notice`), so a
    // notice's live text never changes while its state stays the same, and a screen reader
    // announces it once. The stale notice, the protection of the shown sheet, the conditional
    // formatting that the grid can leave out, the parts past a limit of the read, and a wait of
    // the `mode` prop stay mounted while the workbook is open (`statusView`).
    return () => {
      const view = statusView(snap.value, modeWaits.value, t);
      if (view.kind === 'strip')
        return (
          <div class="xlsx-status-strip">
            {view.refusal ? <div class="xlsx-refusal">{view.refusal}</div> : null}
            {view.notices.map((n) => (
              <Notice key={n.id} on={n.on} short={n.short} text={n.text} brief={n.brief} />
            ))}
          </div>
        );
      // Without a grid, the live region announced each refusal already, so it is plain text.
      return (
        <div class="xlsx-message">
          <div role={view.role}>{view.text}</div>
          {view.refusals.map((text) => (
            <div key={text} class="xlsx-message-refusal">
              {text}
            </div>
          ))}
        </div>
      );
    };
  },
});
