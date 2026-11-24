/** @jsxImportSource vue */
// One chrome slot as the toolbar sees it: the Vue twin of the React adapter's
// `toolbar/slot.tsx`. The public `useEditorCommand` state and command, and the slot's control
// and label. The adapter keeps no command state of its own: every part of the toolbar reads a
// slot through this composable, so the bar, the menus, the overflow panel, and a host's own
// control show the same answer, from the one source `toolbarCommandState`.
//
// Differences from React: `useToolbarSlot` returns a computed view of the words beside the
// command's refs, and `ChromeIcon` is a function component.

import {
  computed,
  inject,
  shallowRef,
  toValue,
  type ComputedRef,
  type MaybeRefOrGetter,
} from 'vue';
import {
  chromeControlForSlot,
  editorRowsSlice,
  slotWords,
  type ChromeControl,
  type ChromeSlotId,
  type ToolbarHost,
} from '@xlsx-editor.dev/core/editor';
import { useEditorState, useTranslation, useTranslator, useXlsxEditor } from '../context.ts';
import {
  toolbarHostKey,
  translateKey,
  useEditorCommand,
  type UseEditorCommandReturn,
} from '../useEditorCommand.ts';

export { toolbarHostKey, translateKey };

/** The words of one slot. */
export interface ToolbarSlotWords {
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

/** The state, words, and command of one slot. */
export interface ToolbarSlot extends UseEditorCommandReturn {
  readonly slot: ChromeSlotId;
  readonly control: ChromeControl;
  readonly words: ComputedRef<ToolbarSlotWords>;
  /** Give the keyboard back to the grid, and keep the selection. */
  focusGrid(): void;
}

/** One slot of the chrome registry. */
export function useToolbarSlot(slot: ChromeSlotId): ToolbarSlot {
  const editor = useXlsxEditor();
  const { t } = useTranslation();
  const command = useEditorCommand(slot);
  const control = chromeControlForSlot(slot)!;
  return {
    ...command,
    words: computed(() =>
      slotWords(
        control.labelKey,
        { disabledReason: command.disabledReason.value, hint: command.hint.value },
        t
      )
    ),
    slot,
    control,
    focusGrid: () => editor.value?.focus(),
  };
}

/** The state of a control that opens rows: enabled, and the reason when disabled. */
export interface RowsState {
  readonly isEnabled: ComputedRef<boolean>;
  readonly disabledReason: ComputedRef<string | null>;
}

const noHost = shallowRef<ToolbarHost>({});

/**
 * The state of a control that only opens the rows `slots`, such as a group menu or a submenu:
 * disabled with the reason of its first row when every row is disabled (`toolbarRowsState`).
 */
export function useRowsState(slots: MaybeRefOrGetter<readonly ChromeSlotId[]>): RowsState {
  const host = inject(toolbarHostKey, noHost);
  const t = useTranslator();
  const slice = useEditorState(
    (snapshot) => editorRowsSlice(snapshot, toValue(slots), host.value, t.value),
    (a, b) => a.enabled === b.enabled && a.disabledReason === b.disabledReason,
    () => [toValue(slots).join('\n'), host.value, t.value]
  );
  return {
    isEnabled: computed(() => slice.value.enabled),
    disabledReason: computed(() => slice.value.disabledReason),
  };
}

/**
 * Registry icon paths as inline SVG (Material Symbols view box), hidden from assistive
 * technology.
 */
export function ChromeIcon(props: {
  readonly paths: readonly string[] | null;
  readonly className?: string;
}) {
  if (!props.paths) return null;
  return (
    <svg class={props.className} viewBox="0 -960 960 960" aria-hidden="true" focusable="false">
      {props.paths.map((d) => (
        <path key={d} d={d} />
      ))}
    </svg>
  );
}
