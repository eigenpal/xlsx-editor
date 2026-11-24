// The adapter-neutral part of the status chrome: the snapshot slices that the zoom control, the
// status strip, the formula bar, and the sheet tabs render, and the texts of a failed open and
// of the parts past a limit. Moved from the React adapter's `chrome.tsx`, so the React and Vue
// adapters render the same slices and say the same words. Adapter support, not a stable API.

import { displayName, refusalText, type EditorTranslate } from './cell-input.ts';
import type { ChromeSlotId } from './chrome-controls.ts';
import type { EditorFailure, EditorSnapshot, PartPastLimit } from './editor-model.ts';
import { disabledReasonText, toolbarCommandState, type ToolbarHost } from './toolbar-commands.ts';

/** Whether two slices hold the same values, key by key. @internal */
export function sameSlice<T extends object>(a: T, b: T): boolean {
  for (const k of Object.keys(a) as (keyof T)[]) if (!Object.is(a[k], b[k])) return false;
  return true;
}

/** The part of the snapshot that the zoom control renders. @internal */
export const zoomSlice = (s: EditorSnapshot) => ({ status: s.status, zoom: s.zoom });

/** The part of the snapshot that the status strip and the status message render. @internal */
export const statusSlice = (s: EditorSnapshot) => ({
  status: s.status,
  failure: s.failure?.code,
  failureLimit: s.failure?.limit,
  failurePart: s.failure?.part,
  partsPastLimit: s.partsPastLimit,
  refusal: s.refusal,
  editDropped: s.editDropped,
  numberLocale: s.numberLocale,
  stale: s.stale,
  protected: s.protection !== null,
  conditionalHidden: s.conditionalHidden,
});

/** The part of the snapshot that the formula bar renders. @internal */
export const formulaSlice = (s: EditorSnapshot) => ({
  status: s.status,
  ref: s.selectionRef,
  text: s.editText ?? s.activeCell?.content ?? '',
  editing: s.editing,
  view: s.mode === 'view',
  refusal: s.refusal,
  numberLocale: s.numberLocale,
});

/** The part of the snapshot that the sheet tabs render. @internal */
export const tabsSlice = (s: EditorSnapshot) => ({
  status: s.status,
  sheets: s.sheets,
  activeSheet: s.activeSheet,
  refusal: s.refusal,
  numberLocale: s.numberLocale,
  mode: s.mode,
  workbookProtection: s.workbookProtection,
  sheetEditRefusal: s.sheetEditRefusal,
});

/** The text of each limit of the read, as `docs/site/content/workbook-limits.mdx` gives it. */
const LIMIT_KEYS = {
  'zip.maxEntries': 'status.limits.zipMaxEntries',
  'zip.maxTotalBytes': 'status.limits.zipMaxTotalBytes',
  'zip.maxRatio': 'status.limits.zipMaxRatio',
  'xml.maxBytes': 'status.limits.xmlMaxBytes',
  'xml.maxElements': 'status.limits.xmlMaxElements',
  'xml.maxDepth': 'status.limits.xmlMaxDepth',
  'workbook.maxSheetNameLength': 'status.limits.workbookMaxSheetNameLength',
} as const satisfies Record<NonNullable<EditorFailure['limit']>, string>;

/** The notice of a part past a limit, by what the workbook shows without it. */
const PAST_LIMIT_KEYS = {
  styles: 'status.pastLimit.styles',
  theme: 'status.pastLimit.theme',
  drawing: 'status.pastLimit.drawing',
  relationships: 'status.pastLimit.relationships',
  other: 'status.pastLimit.other',
} as const satisfies Record<PartPastLimit['kind'], string>;

/**
 * The limits that a smaller copy of the file can pass. A smaller copy does not help with a
 * compression ratio or with nested levels, so the message gives no next step for them.
 */
const SMALLER_COPY_HELPS: ReadonlySet<NonNullable<EditorFailure['limit']>> = new Set([
  'zip.maxEntries',
  'zip.maxTotalBytes',
  'xml.maxBytes',
  'xml.maxElements',
]);

/**
 * The message of a load that failed with the code `package`. A file past a limit is too large
 * to open, as DOCX Editor says it, and the message names the part and the limit. @internal
 */
export function packageFailureText(
  failure: Pick<EditorFailure, 'limit' | 'part'>,
  t: EditorTranslate
): string {
  if (!failure.limit) return t('status.failed.package');
  const limit = t(LIMIT_KEYS[failure.limit]);
  const text = failure.part
    ? t('status.failed.tooLargePart', { part: displayName(failure.part), limit })
    : t('status.failed.tooLarge', { limit });
  return SMALLER_COPY_HELPS.has(failure.limit) ? `${text} ${t('status.failed.smallerCopy')}` : text;
}

/**
 * The notice for the parts that stay as bytes past a limit: the first part, and a count of the
 * others. @internal
 */
export function pastLimitText(parts: readonly PartPastLimit[], t: EditorTranslate): string {
  const [first, ...others] = parts;
  if (!first) return '';
  const text = t(PAST_LIMIT_KEYS[first.kind], { part: displayName(first.part) });
  return others.length > 0
    ? `${text} ${t('status.pastLimit.more', { count: others.length })}`
    : text;
}

/** What the grid measures after it attaches. @internal */
export interface GridSetupFacts {
  /** The computed `overflow-y` of the grid scroller. */
  readonly overflowY: string;
}

/**
 * The setup mistake that the facts show, or `null`. An empty `overflowY` means that the host
 * computes no styles, so it shows no mistake. `part` and `stylesheet` name the grid part and the
 * stylesheet of the adapter that asks. @internal
 */
export function gridSetupProblem(
  { overflowY }: GridSetupFacts,
  part = 'XlsxEditor.Content',
  stylesheet = '@xlsx-editor.dev/react/styles.css'
): string | null {
  if (overflowY === '' || overflowY === 'auto' || overflowY === 'scroll') return null;
  return `${part}: the grid has no editor styles. Import ${stylesheet} once.`;
}

/**
 * Warn in development, once the document has loaded, when the grid in `element` has no editor
 * styles (`gridSetupProblem`). A stylesheet that loads late gives no warning. Returns the
 * cancel function. @internal
 */
export function watchGridSetup(element: HTMLElement, part: string, stylesheet: string): () => void {
  const view = element.ownerDocument.defaultView;
  if (!view) return () => {};
  let frame = 0;
  const check = () => {
    frame = view.requestAnimationFrame(() => {
      const scroller = element.querySelector<HTMLElement>('.xlsx-scroller');
      if (!scroller) return;
      const problem = gridSetupProblem(
        { overflowY: view.getComputedStyle(scroller).overflowY },
        part,
        stylesheet
      );
      if (problem) console.warn(problem);
    });
  };
  if (element.ownerDocument.readyState === 'complete') check();
  else view.addEventListener('load', check, { once: true });
  return () => {
    view.removeEventListener('load', check);
    view.cancelAnimationFrame(frame);
  };
}

/** The state of one chrome slot that `useEditorCommand` shows, with its words translated. @internal */
export interface EditorCommandSlice {
  readonly active: boolean;
  readonly enabled: boolean;
  readonly disabledReason: string | null;
  readonly hint: string | null;
  readonly value: string | null;
}

/**
 * The state of `slot` in `snapshot` (`toolbarCommandState`), with the disabled reason and the
 * hint of an enabled command translated. @internal
 */
export function editorCommandSlice(
  snapshot: EditorSnapshot,
  slot: ChromeSlotId,
  host: ToolbarHost,
  t: EditorTranslate
): EditorCommandSlice {
  const state = toolbarCommandState(snapshot, slot, host);
  const reason = state.disabledReason;
  return {
    active: state.active,
    enabled: state.enabled,
    disabledReason: reason ? disabledReasonText(reason, t) : null,
    hint: !reason && state.hint ? t(state.hint.key) : null,
    value: state.value ?? null,
  };
}

/** Whether two slot states show the same thing. @internal */
export function sameEditorCommandSlice(a: EditorCommandSlice, b: EditorCommandSlice): boolean {
  return (
    a.active === b.active &&
    a.enabled === b.enabled &&
    a.disabledReason === b.disabledReason &&
    a.hint === b.hint &&
    a.value === b.value
  );
}

/** The words of one toolbar slot. @internal */
export interface SlotWords {
  /** The accessible name. */
  readonly label: string;
  /** The disabled reason or the hint: the accessible description, or null. */
  readonly description: string | null;
  /**
   * The tooltip: the label, and the disabled reason after it when there is one; else the label
   * with the hint on a second line.
   */
  readonly title: string;
}

/** The words of a slot with the label key `labelKey` and the state `command`. @internal */
export function slotWords(
  labelKey: string,
  command: Pick<EditorCommandSlice, 'disabledReason' | 'hint'>,
  t: EditorTranslate
): SlotWords {
  const label = t(labelKey);
  return {
    label,
    description: command.disabledReason ?? command.hint,
    title: command.disabledReason
      ? t('toolbar.disabledTitle', { label, reason: command.disabledReason })
      : command.hint
        ? t('toolbar.hintTitle', { label, hint: command.hint })
        : label,
  };
}

/**
 * One notice of the status strip. While `on`, its live text is the full `text`; `short` shows
 * the brief text in its place, visually only. @internal
 */
export interface StatusNotice {
  readonly id: 'stale' | 'protected' | 'conditionalHidden' | 'pastLimit' | 'viewWaits';
  readonly on: boolean;
  readonly short: boolean;
  readonly text: string;
  readonly brief: string;
}

/**
 * What `StatusMessage` shows: the strip of a ready workbook (a refusal on its own row, and the
 * notices), or the message of an editor with no grid (its text, and the refusals under it).
 * @internal
 */
export type StatusView =
  | {
      readonly kind: 'strip';
      readonly refusal: string | null;
      readonly notices: readonly StatusNotice[];
    }
  | {
      readonly kind: 'message';
      /** `alert` for a failed open, else `status`. */
      readonly role: 'alert' | 'status';
      readonly text: string;
      readonly refusals: readonly string[];
    };

/**
 * The status strip of `snap` (`statusSlice`), with `modeWaits` when a change of the `mode` prop
 * to view mode waits for the open cell edit. While a workbook is ready, the strip shows every
 * notice, empty when it is off, so a screen reader announces each one when its text arrives.
 * When more than one message shows, a notice shows its short text. Without a grid, a refusal
 * shows under the message: a refused save, or an open edit that a load or `close()` dropped.
 * @internal
 */
export function statusView(
  snap: ReturnType<typeof statusSlice>,
  modeWaits: boolean,
  t: EditorTranslate
): StatusView {
  if (snap.status === 'ready') {
    const refusal = snap.refusal ? refusalText(snap.refusal, t, snap.numberLocale) : null;
    const pastLimit = snap.partsPastLimit.length > 0;
    const shown = [
      refusal !== null,
      snap.stale,
      snap.protected,
      snap.conditionalHidden,
      pastLimit,
      modeWaits,
    ];
    // A notice is short when any message before it shows.
    const before = (at: number) => shown.slice(0, at).some(Boolean);
    const notices: StatusNotice[] = [
      {
        id: 'stale',
        on: snap.stale,
        short: before(1),
        text: t('status.stale'),
        brief: t('status.staleShort'),
      },
      {
        id: 'protected',
        on: snap.protected,
        short: before(2),
        text: t('status.protected'),
        brief: t('status.protectedShort'),
      },
      {
        id: 'conditionalHidden',
        on: snap.conditionalHidden,
        short: before(3),
        text: t('status.conditionalHidden'),
        brief: t('status.conditionalHiddenShort'),
      },
      {
        id: 'pastLimit',
        on: pastLimit,
        short: before(4),
        text: pastLimitText(snap.partsPastLimit, t),
        brief: t('status.pastLimit.short'),
      },
      {
        id: 'viewWaits',
        on: modeWaits,
        short: before(5),
        text: t('status.viewWaits'),
        brief: t('status.viewWaitsShort'),
      },
    ];
    return { kind: 'strip', refusal, notices };
  }
  // The drop of typed text stays shown when a later command is refused with `loading`.
  const refusals = [
    ...(snap.editDropped && snap.refusal?.code !== 'edit-dropped'
      ? [refusalText({ code: 'edit-dropped' }, t, snap.numberLocale)]
      : []),
    ...(snap.refusal ? [refusalText(snap.refusal, t, snap.numberLocale)] : []),
  ];
  if (snap.status === 'failed') {
    const code = snap.failure;
    const text =
      code === 'fonts'
        ? t('status.failed.fonts')
        : code === 'no-worksheet'
          ? t('status.failed.noWorksheet')
          : code === 'read'
            ? t('status.failed.read')
            : packageFailureText({ limit: snap.failureLimit, part: snap.failurePart }, t);
    return { kind: 'message', role: 'alert', text, refusals };
  }
  return {
    kind: 'message',
    role: 'status',
    text: snap.status === 'loading' ? t('status.loading') : t('status.empty'),
    refusals,
  };
}
