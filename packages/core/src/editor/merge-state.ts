// The state of the merge commands, from the editor snapshot alone: enabled, or the reason when
// disabled, with the keys that a toolbar control state uses. The merge commands have no chrome
// slot: a later change adds the slots, and `toolbarCommandState` then reads this state for
// them, so this state stays the one source of the merge commands' enabled state. A host that
// offers the commands in controls of its own reads this state.

import type { EditorSnapshot } from './editor-model.ts';
import type { ChromeDisabledReason } from './toolbar-commands.ts';

/**
 * A merge command: merge, merge and center, or unmerge (`mergeCells` and `unmergeCells`).
 *
 * @public
 */
export type MergeCommand = 'merge' | 'mergeCenter' | 'unmerge';

/**
 * The state of one merge command: enabled, or the reason when it is disabled.
 *
 * @public
 */
export interface MergeCommandState {
  readonly enabled: boolean;
  /**
   * The engine's reason when disabled, as a catalog key: `loading`, `editorNotReady`, `viewMode`,
   * `composing`, `mergeLocked`, `mergeOneCell`, `mergeOverlap`, `stylesUnread`, `rangeTooLarge`,
   * or `noMergedRange` (`mergeCommandState` gives the order).
   */
  readonly disabledReason: ChromeDisabledReason | null;
  /**
   * For an enabled merge or merge and center: the first cell, in reading order, whose value the
   * merge would drop, or null. An open cell edit counts as it would be after its commit, which
   * the merge makes first. When it is set, the command refuses with `merge-drops-values`: ask the
   * user first, then call `mergeCells({ dropValues: true })`.
   */
  readonly dropsValuesAt: string | null;
}

/**
 * The state of a merge command, from the snapshot alone. It first gives the reasons of
 * `toolbarCommandState`, in its order, and they must stay equal: a workbook that loads or is not
 * ready, view mode, and a composing input method. Then it gives the reasons of the command: a
 * protected sheet (`mergeLocked`); for a merge, a selection of one cell (`mergeOneCell`), then
 * one that overlaps a merged range (`mergeOverlap`); for a merge and center, also a styles part
 * that stays as bytes (`stylesUnread`) and a range too large to format (`rangeTooLarge`), as the
 * format controls; and for an unmerge, a selection with no merged range (`noMergedRange`). A
 * merge over values outside the top-left cell stays enabled, with `dropsValuesAt`: the host asks
 * the user before the command, then merges with `dropValues`.
 *
 * @public
 */
export function mergeCommandState(
  snapshot: EditorSnapshot,
  command: MergeCommand
): MergeCommandState {
  const off = (key: ChromeDisabledReason['key']): MergeCommandState => ({
    enabled: false,
    disabledReason: { key },
    dropsValuesAt: null,
  });
  if (snapshot.status === 'loading') return off('disabledReason.loading');
  if (snapshot.status !== 'ready' || !snapshot.selection)
    return off('disabledReason.editorNotReady');
  if (snapshot.mode === 'view') return off('disabledReason.viewMode');
  if (snapshot.editing && snapshot.composing) return off('disabledReason.composing');
  if (snapshot.protection) return off('disabledReason.mergeLocked');
  if (command === 'unmerge') {
    if (!snapshot.selectionMerged) return off('disabledReason.noMergedRange');
  } else {
    // The order of the merge op: one cell first, then a merged range.
    const { start, end } = snapshot.selection.range;
    if (start.row === end.row && start.col === end.col) return off('disabledReason.mergeOneCell');
    if (snapshot.selectionMerged) return off('disabledReason.mergeOverlap');
    if (command === 'mergeCenter' && snapshot.partsPastLimit.some((p) => p.kind === 'styles'))
      return off('disabledReason.stylesUnread');
    if (command === 'mergeCenter' && snapshot.formatRefusal)
      return off('disabledReason.rangeTooLarge');
  }
  const dropsValuesAt = command === 'unmerge' ? null : snapshot.selectionDropsValuesAt;
  return { enabled: true, disabledReason: null, dropsValuesAt };
}
