// The state of the sheet commands of the sheet tabs, DOM-free: the add button and the Rename,
// Delete, Move left, and Move right rows of the sheet menu. As `toolbarCommandState` is the one
// source of a toolbar control's enabled state, this is the one source of a sheet command's
// state, and the adapters only render it. After PPTX Editor (eigenpal/pptx-editor) `origin/main`
// 79a5d22d, `packages/core/src/editor/menu-arrangement.ts` `commandRowReason`, which the menu
// rows of `packages/react/src/RailContextMenu.tsx` read. Difference: the state is a pure
// function of the snapshot, not a dry run of the command, because the sheet ops refuse most
// cases only after a scan of the workbook, and the menu shows only the reasons that the
// snapshot holds.

import type { EditorSnapshot } from './editor-model.ts';
import type { ChromeDisabledReason } from './toolbar-commands.ts';

/**
 * A command of the sheet tabs: `add` the add button, and `rename`, `delete`, `moveLeft`, and
 * `moveRight` the rows of the sheet menu.
 *
 * @public
 */
export type SheetCommand = 'add' | 'rename' | 'delete' | 'moveLeft' | 'moveRight';

/**
 * The state of one sheet command.
 *
 * @public
 */
export interface SheetCommandState {
  readonly enabled: boolean;
  /**
   * The engine's reason when the command is disabled, as a catalog key: `viewMode`,
   * `workbookLocked`, `notWorksheet`, `sheetEditBlocked` (with the part that refuses it as its
   * `about`), `sheetOrderBlocked` (macros), or `lastSheet`. Null when the command is enabled,
   * and for a move at the end of the tabs or a sheet that has no tab, except in a workbook that
   * holds macros, where a move is disabled with `sheetOrderBlocked`.
   */
  readonly disabledReason: ChromeDisabledReason | null;
  /**
   * The workbook position that a move takes the sheet to: the position of the next visible tab
   * in that direction, so a hidden sheet never takes a move. Set only on an enabled move.
   */
  readonly to?: number;
}

/**
 * What the state reads from the editor snapshot. The reason of a rename or a delete comes from
 * each sheet (`EditorSheet.editRefusal`). `sheetEditRefusal`, the reason of the sheets of the
 * opened file, applies to a sheet that does not set it.
 *
 * @public
 */
export type SheetCommandInput = Pick<
  EditorSnapshot,
  'mode' | 'sheets' | 'workbookProtection' | 'sheetEditRefusal'
>;

const disabled = (key: ChromeDisabledReason['key'] | null): SheetCommandState => ({
  enabled: false,
  disabledReason: key ? { key } : null,
});
const ENABLED: SheetCommandState = { enabled: true, disabledReason: null };

/**
 * The state of a sheet command on the sheet at workbook position `sheet` (ignored by `add`). View
 * mode disables every command with `viewMode`, and a workbook protection that locks the sheet list
 * with `workbookLocked`. A sheet that is not a worksheet (a chartsheet, a macro sheet, a dialog
 * sheet, or a sheet whose part is missing) has every command but `add` disabled with
 * `notWorksheet`, except that macros come first for the moves. Delete needs another visible sheet
 * (`lastSheet`). A chart, a pivot table, a sheet that is not a worksheet, or a workbook or package
 * relationship outside the allowlist (for example data connections) disables Rename and Delete of
 * each sheet of the opened file with `sheetEditBlocked`, as the store refuses them with
 * `structure-blocked`. A sheet that the editor added in the session stays enabled, unless such
 * content can refer to it. The reason of each sheet is its `EditorSheet.editRefusal`. Macros
 * block every sheet, and also disable Add and the moves with `sheetOrderBlocked`.
 *
 * @public
 */
export function sheetCommandState(
  snapshot: SheetCommandInput,
  command: SheetCommand,
  sheet: number
): SheetCommandState {
  if (snapshot.mode === 'view') return disabled('disabledReason.viewMode');
  if (snapshot.workbookProtection?.allowStructure === false)
    return disabled('disabledReason.workbookLocked');
  // Macro code can read a sheet by its position, so a workbook that holds macros takes no add
  // and no move (`sheetOrderBlocked`), as the store refuses them.
  const ordered = command === 'add' || command === 'moveLeft' || command === 'moveRight';
  if (ordered && snapshot.sheetEditRefusal?.about?.kind === 'macros')
    return {
      enabled: false,
      disabledReason: { key: 'disabledReason.sheetOrderBlocked', about: { kind: 'macros' } },
    };
  if (command === 'add') return ENABLED;
  const tabs = snapshot.sheets.filter((s) => s.visible);
  const at = tabs.findIndex((s) => s.index === sheet);
  if (at < 0) return disabled(null);
  // A sheet that is not a worksheet takes no rename, delete, or move from the tabs: the store
  // refuses a rename and a delete for it, and its disabled tab opens no menu, so the tabs move
  // no such sheet (design D4). The reason lets a host menu and the tab title say why.
  if (!tabs[at]!.worksheet) return disabled('disabledReason.notWorksheet');
  const own = tabs[at]!.editRefusal;
  const blocked = own === undefined ? snapshot.sheetEditRefusal : own;
  // Every part blocker names what refuses the edit (`about`). The fallback words never show a
  // raw part path, which is file text.
  const blockedState = (): SheetCommandState => ({
    enabled: false,
    disabledReason: {
      key: 'disabledReason.sheetEditBlocked',
      about: blocked?.about ?? { kind: 'workbook-content' },
    },
  });
  switch (command) {
    case 'rename':
      return blocked ? blockedState() : ENABLED;
    case 'delete':
      if (tabs.length < 2) return disabled('disabledReason.lastSheet');
      return blocked ? blockedState() : ENABLED;
    case 'moveLeft':
      return at > 0 ? { ...ENABLED, to: tabs[at - 1]!.index } : disabled(null);
    case 'moveRight':
      return at < tabs.length - 1 ? { ...ENABLED, to: tabs[at + 1]!.index } : disabled(null);
  }
}
