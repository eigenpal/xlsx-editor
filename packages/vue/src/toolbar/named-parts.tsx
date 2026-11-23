/** @jsxImportSource vue */
// The named toolbar parts (`XlsxEditorToolbar.Bold`, `XlsxEditorToolbar.FontFamily`, ...). The
// Vue twin of the React adapter's `toolbar/named-parts.tsx`, after DOCX Editor's `definePart`
// (docx-editor/packages/vue/src/editor/toolbar/parts.ts at ad686574); the header there lists the
// rules. Each part is an `XlsxEditorToolbar.Slot` with its slot pinned, and the names come from
// core (`TOOLBAR_NAMED_PARTS`).
//
// Differences from React: a part is a component with runtime props, as DOCX Editor's Vue parts;
// the icon content of a button part is its default slot, or the `icon` prop; and a template
// `class` joins `className`.

import { defineComponent, h, type DefineComponent, type PropType, type VNodeChild } from 'vue';
import {
  TOOLBAR_NAMED_PARTS,
  isToolbarButtonSlot,
  type ChromeSlotId,
} from '@xlsx-editor.dev/core/editor';
import {
  XlsxEditorToolbarSlot,
  markNamedPart,
  namedPartSlotProps,
  vnodeFlag,
  type XlsxEditorToolbarSlotProps,
} from './host-parts.tsx';

/**
 * Props of the named button parts, the twin of React's `ToolbarPartProps`: the slot is pinned,
 * and the default slot is icon content (`icon`, else the default slot, else the registry
 * icon), not a replacement. @public
 */
export interface ToolbarPartProps {
  /** Removes the control where it stands. Its command keeps working from keys and the API. */
  readonly hidden?: boolean;
  /** An icon in place of the registry icon, about 18 pixels: a vnode or a string. */
  readonly icon?: VNodeChild;
  /** Render the control as the single element of the default slot. */
  readonly asChild?: boolean;
  /** Appended after the class of the control. A host `class` also works. */
  readonly className?: string;
  /** Content for the control's row in the "More" panel, when its group leaves the bar. */
  readonly overflowContent?: () => VNodeChild;
}

/** Props of the named value parts, the twin of React's `ToolbarSlotPartProps`. @public */
export interface ToolbarSlotPartProps {
  /** Removes the control where it stands. Its command keeps working from keys and the API. */
  readonly hidden?: boolean;
  /** Appended after the class of the control's trigger. A host `class` also works. */
  readonly className?: string;
  /** Content for the control's row in the "More" panel, when its group leaves the bar. */
  readonly overflowContent?: () => VNodeChild;
}

/** A named button part. @public */
export type ToolbarPartComponent = DefineComponent<ToolbarPartProps>;

/** A named value part. @public */
export type ToolbarSlotPartComponent = DefineComponent<ToolbarSlotPartProps>;

/**
 * Every named toolbar part, attached to `XlsxEditorToolbar` as statics: one for each slot of
 * `CHROME_GROUPS`. The twin of React's `ToolbarNamedParts`. @public
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

function definePart(name: string, slot: ChromeSlotId) {
  const button = isToolbarButtonSlot(slot);
  const Part = defineComponent({
    name: `XlsxEditorToolbar${name}`,
    props: {
      hidden: { type: Boolean, default: undefined },
      className: { type: String, default: undefined },
      overflowContent: { type: Function as PropType<() => VNodeChild>, default: undefined },
      ...(button
        ? {
            icon: { type: [Object, String, Array] as PropType<VNodeChild>, default: undefined },
            asChild: { type: Boolean, default: undefined },
          }
        : {}),
    },
    setup(props, { attrs, slots }) {
      return () => {
        const own = props as Record<string, unknown>;
        const content = slots.default?.();
        const slotProps = namedPartSlotProps(slot, { ...own, class: attrs.class }, content);
        const { children, class: hostClass, ...rest } = slotProps;
        const className =
          [rest.className, hostClass].filter((c) => typeof c === 'string' && c).join(' ') ||
          undefined;
        const partProps = {
          ...rest,
          slotId: slot,
          className,
          hidden: vnodeFlag(rest.hidden) || undefined,
        } as unknown as XlsxEditorToolbarSlotProps;
        return h(
          XlsxEditorToolbarSlot,
          partProps,
          vnodeFlag(rest.asChild) ? { default: () => children } : undefined
        );
      };
    },
  });
  markNamedPart(Part, slot);
  return Part;
}

/** The named parts, built from core's table. */
export const TOOLBAR_PARTS = Object.fromEntries(
  Object.entries(TOOLBAR_NAMED_PARTS).map(([name, slot]) => [name, definePart(name, slot)])
) as unknown as ToolbarNamedParts;

// Each name of core's table has a typed part here, and each typed part a name there: a name
// added on one side only fails the build.
type MissingPart = Exclude<keyof typeof TOOLBAR_NAMED_PARTS, keyof ToolbarNamedParts>;
type ExtraPart = Exclude<keyof ToolbarNamedParts, keyof typeof TOOLBAR_NAMED_PARTS>;
const everyPartTyped: [MissingPart | ExtraPart] extends [never] ? true : MissingPart | ExtraPart =
  true;
void everyPartTyped;
