// The named toolbar parts (`XlsxEditor.Toolbar.Bold`, ...) and the slot that each one stands for,
// in one table that both adapters read, so the React and Vue parts have the same names. After
// `TOOLBAR_NAMED_PARTS` of pptx-editor/packages/core/src/editor/toolbar-layout.ts at 70052cf1
// (EigenPal). DOCX Editor lists its named parts by hand in each adapter
// (docx-editor/packages/react/src/editor/toolbar/parts.tsx at ad686574); a table in core keeps
// the two adapters from drifting, which is why this file takes the PPTX Editor shape.
//
// The names are DOCX Editor's where the idea is the same (`Undo`, `Bold`, `FontFamily`,
// `AlignLeft`, `ClearFormatting`, `Save`). The other names come from the label of the slot.
// Differences from PPTX Editor: the slots of the workbook toolbar, and a test in place of a
// compile-time check that each slot of `CHROME_GROUPS` has exactly one name.
//
// Adapter support, not a stable API: `@internal`.

import { chromeControlForSlot, type ChromeSlotId } from './chrome-controls.ts';

/**
 * The named toolbar parts and the slot that each one stands for. `Toolbar.Bold` is shorthand
 * for `Toolbar.Slot slotId="font.bold"`, with the same props. @internal
 */
export const TOOLBAR_NAMED_PARTS = {
  Undo: 'history.undo',
  Redo: 'history.redo',
  FontFamily: 'font.family',
  FontSize: 'font.size',
  Bold: 'font.bold',
  Italic: 'font.italic',
  Underline: 'font.underline',
  FontColor: 'font.color',
  FillColor: 'fill.color',
  OutsideBorders: 'borders.outline',
  AllBorders: 'borders.all',
  NoBorder: 'borders.none',
  AlignLeft: 'alignment.left',
  AlignCenter: 'alignment.center',
  AlignRight: 'alignment.right',
  AlignTop: 'alignment.top',
  AlignMiddle: 'alignment.middle',
  AlignBottom: 'alignment.bottom',
  WrapText: 'alignment.wrap',
  NumberFormat: 'number.format',
  MoreNumberFormats: 'number.custom',
  IncreaseDecimals: 'number.increaseDecimals',
  DecreaseDecimals: 'number.decreaseDecimals',
  InsertRows: 'cells.insertRows',
  InsertColumns: 'cells.insertColumns',
  DeleteRows: 'cells.deleteRows',
  DeleteColumns: 'cells.deleteColumns',
  ClearContents: 'clear.contents',
  ClearFormatting: 'clear.formats',
  ClearAll: 'clear.all',
  Save: 'file.save',
} as const satisfies Record<string, ChromeSlotId>;

/** The name of a named toolbar part. @internal */
export type ToolbarPartName = keyof typeof TOOLBAR_NAMED_PARTS;

/**
 * Whether the named part of `slot` is a button part, which takes `icon`, `asChild`, and icon
 * children, as DOCX Editor's `ToolbarPartProps`. The other parts are value controls (a list, a
 * palette, a dialog opener), which take only `hidden` and `className`, as DOCX Editor's
 * `ToolbarSlotPartProps`. @internal
 */
export function isToolbarButtonSlot(slot: string): boolean {
  const control = chromeControlForSlot(slot);
  return control !== null && (control.shape ?? 'icon') === 'icon';
}
