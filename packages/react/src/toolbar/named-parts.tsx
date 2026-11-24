/** @jsxImportSource react */
// The named toolbar parts (`XlsxEditor.Toolbar.Bold`, `XlsxEditor.Toolbar.FontFamily`, ...),
// after DOCX Editor's `definePart` (docx-editor/packages/react/src/editor/toolbar/parts.tsx at
// ad686574) and PPTX Editor's toolbar-named-parts.tsx (pptx-editor/packages/react/src at
// 70052cf1). Each part is a `Toolbar.Slot` with its slot pinned, and carries the slot as a
// static (`xlsxSlot`), so the toolbar recognizes it among its children and puts it in the
// slot's place. The table of names lives in core (`TOOLBAR_NAMED_PARTS`), so the Vue twin has
// the same names.
//
// Differences from DOCX Editor: the parts come from the core table, not one line each; the
// marker static stays out of the declared type; and the value parts (a list, a palette, the
// dialog opener) take `hidden` and `className` only, as DOCX Editor's `ToolbarSlotPartProps`.

import type { ReactElement, ReactNode } from 'react';
import { TOOLBAR_NAMED_PARTS, type ChromeSlotId } from '@xlsx-editor.dev/core/editor';
import { XlsxEditorToolbarSlot, markNamedPart, namedPartSlotProps } from './host-parts.tsx';

/**
 * Props for the named button parts (`Toolbar.Bold`, `Toolbar.Undo`, ...), as DOCX Editor's
 * `ToolbarPartProps`: the slot is pinned, and the children are icon content (`icon`, else the
 * children, else the registry icon), not a replacement. `<Toolbar.Bold hidden />` is
 * `<Toolbar.Slot slotId="font.bold" hidden />`. @public
 */
export interface ToolbarPartProps {
  /** Removes the control where it stands. Its command keeps working from keys and the API. */
  readonly hidden?: boolean;
  /** An icon in place of the registry icon, about 18 pixels. */
  readonly icon?: ReactNode;
  /** Render the control as the single child element, with its state and press merged in. */
  readonly asChild?: boolean;
  /** Appended after the class of the control. */
  readonly className?: string;
  /** Content for the control's row in the "More" panel, when its group leaves the bar. */
  readonly overflowContent?: () => ReactNode;
  /** Icon content when there is no `icon`, or the element for `asChild`. */
  readonly children?: ReactNode;
}

/**
 * Props for the named value parts (`Toolbar.FontFamily`, `Toolbar.FontColor`, ...), as DOCX
 * Editor's `ToolbarSlotPartProps`: a value control is not one button, so it takes no `icon`,
 * `asChild`, or children. @public
 */
export interface ToolbarSlotPartProps {
  /** Removes the control where it stands. Its command keeps working from keys and the API. */
  readonly hidden?: boolean;
  /** Appended after the class of the control's trigger. */
  readonly className?: string;
  /** Content for the control's row in the "More" panel, when its group leaves the bar. */
  readonly overflowContent?: () => ReactNode;
}

/** A named button part. @public */
export type ToolbarPartComponent = (props: ToolbarPartProps) => ReactElement;

/** A named value part. @public */
export type ToolbarSlotPartComponent = (props: ToolbarSlotPartProps) => ReactElement;

/**
 * Every named toolbar part, attached to `XlsxEditor.Toolbar` as statics: one for each slot of
 * the preset toolbar (`defaultChromeGroups()`). The slots that only the menu bar shows have menu
 * row parts instead. Button parts take `ToolbarPartProps`, value parts `ToolbarSlotPartProps`, as
 * in DOCX Editor. @public
 */
export interface ToolbarNamedParts {
  /** `history.undo`. */
  readonly Undo: ToolbarPartComponent;
  /** `history.redo`. */
  readonly Redo: ToolbarPartComponent;
  /** `font.family`. */
  readonly FontFamily: ToolbarSlotPartComponent;
  /** `font.size`. */
  readonly FontSize: ToolbarSlotPartComponent;
  /** `font.bold`. */
  readonly Bold: ToolbarPartComponent;
  /** `font.italic`. */
  readonly Italic: ToolbarPartComponent;
  /** `font.underline`. */
  readonly Underline: ToolbarPartComponent;
  /** `font.color`. */
  readonly FontColor: ToolbarSlotPartComponent;
  /** `fill.color`. */
  readonly FillColor: ToolbarSlotPartComponent;
  /** `borders.outline`. */
  readonly OutsideBorders: ToolbarPartComponent;
  /** `borders.all`. */
  readonly AllBorders: ToolbarPartComponent;
  /** `borders.none`. */
  readonly NoBorder: ToolbarPartComponent;
  /** `alignment.left`. */
  readonly AlignLeft: ToolbarPartComponent;
  /** `alignment.center`. */
  readonly AlignCenter: ToolbarPartComponent;
  /** `alignment.right`. */
  readonly AlignRight: ToolbarPartComponent;
  /** `alignment.top`. */
  readonly AlignTop: ToolbarPartComponent;
  /** `alignment.middle`. */
  readonly AlignMiddle: ToolbarPartComponent;
  /** `alignment.bottom`. */
  readonly AlignBottom: ToolbarPartComponent;
  /** `alignment.wrap`. */
  readonly WrapText: ToolbarPartComponent;
  /** `number.format`. */
  readonly NumberFormat: ToolbarSlotPartComponent;
  /** `number.custom`: opens the number format dialog. */
  readonly MoreNumberFormats: ToolbarSlotPartComponent;
  /** `number.increaseDecimals`. */
  readonly IncreaseDecimals: ToolbarPartComponent;
  /** `number.decreaseDecimals`. */
  readonly DecreaseDecimals: ToolbarPartComponent;
  /** `cells.insertRows`. */
  readonly InsertRows: ToolbarPartComponent;
  /** `cells.insertColumns`. */
  readonly InsertColumns: ToolbarPartComponent;
  /** `cells.deleteRows`. */
  readonly DeleteRows: ToolbarPartComponent;
  /** `cells.deleteColumns`. */
  readonly DeleteColumns: ToolbarPartComponent;
  /** `clear.contents`. */
  readonly ClearContents: ToolbarPartComponent;
  /** `clear.formats`. */
  readonly ClearFormatting: ToolbarPartComponent;
  /** `clear.all`. */
  readonly ClearAll: ToolbarPartComponent;
  /** `file.save`. */
  readonly Save: ToolbarPartComponent;
}

function definePart(slot: ChromeSlotId) {
  const Part = (props: Record<string, unknown>) => (
    <XlsxEditorToolbarSlot {...namedPartSlotProps(slot, props)} />
  );
  markNamedPart(Part, slot);
  return Part;
}

/** The named parts, built from core's table. */
export const TOOLBAR_PARTS = Object.fromEntries(
  Object.entries(TOOLBAR_NAMED_PARTS).map(([name, slot]) => [name, definePart(slot)])
) as unknown as ToolbarNamedParts;

// Each name of core's table has a typed part here, and each typed part a name there: a name
// added on one side only fails the build.
type MissingPart = Exclude<keyof typeof TOOLBAR_NAMED_PARTS, keyof ToolbarNamedParts>;
type ExtraPart = Exclude<keyof ToolbarNamedParts, keyof typeof TOOLBAR_NAMED_PARTS>;
const everyPartTyped: [MissingPart | ExtraPart] extends [never] ? true : MissingPart | ExtraPart =
  true;
void everyPartTyped;
