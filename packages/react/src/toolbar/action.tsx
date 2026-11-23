/** @jsxImportSource react */
// Host-owned toolbar controls: `XlsxEditor.Toolbar.Action` and `XlsxEditor.Toolbar.Separator`,
// after docx-editor/packages/react/src/editor/toolbar/ToolbarAction.tsx and `ToolbarSeparator`
// of packages/react/src/editor/toolbar/parts.tsx at ad686574 (EigenPal).
//
// `Toolbar.Slot` and the named parts are bound to a slot of the chrome registry: the engine owns
// their label, icon, and enabled state. A host's own action ("Send for review") has no slot.
// `Toolbar.Action` gives it the toolbar's button class, the state attributes, and the mousedown
// guard, so a host does not copy private class names, and a press never moves the keyboard
// from the grid. It is not a way to fake a slot: the enabled and pressed state are the host's.
//
// Differences from DOCX Editor:
// - A disabled action stays focusable (`aria-disabled`, not `disabled`), and its reason is the
//   tooltip and the accessible description, as every control of this toolbar
//   (`toolbar/parts.tsx`) and PPTX Editor's `ToolbarAction`
//   (pptx-editor/packages/react/src/toolbar-action.tsx at 70052cf1) do.
// - Without an `icon` or children, the action shows its `label` as text, as in PPTX Editor.
// - In a row of the "More" panel, a press keeps the panel open, as the built-in rows of this
//   toolbar do. DOCX Editor's panel row closes the panel.
// - A held Enter runs the action once, as the built-in buttons of this toolbar.
// - `asChild` with children that are not one element renders the packaged button, with a
//   development warning, as in PPTX Editor.
// - The separator turns horizontal in a row of the "More" panel, where the controls stand in a
//   column.

import {
  useContext,
  useEffect,
  useId,
  type KeyboardEvent,
  type MouseEvent,
  type ReactNode,
} from 'react';
import { guardToolbarMousedown, toolbarAsChildProblem } from '@xlsx-editor.dev/core/editor';
import { asChildElement, renderAsChild } from './as-child.tsx';
import { PartLayoutContext, warnToolbarProblems } from './host-parts.tsx';

/** Props for `XlsxEditor.Toolbar.Action`. @public */
export interface ToolbarActionProps {
  /**
   * The accessible name and the tooltip. A resolved string, not a catalog key: the action is
   * yours, so your own catalog translates it.
   */
  readonly label: string;
  /** Icon content, about 18 pixels, as the packaged icons. */
  readonly icon?: ReactNode;
  /** The pressed state, for an action that toggles: `aria-pressed` and `data-active`. */
  readonly active?: boolean;
  /**
   * Turns the action off: the button stays focusable (`aria-disabled`, `data-disabled`), and a
   * press does not call `onSelect`.
   */
  readonly disabled?: boolean;
  /** Why the action is disabled: the tooltip and the accessible description. */
  readonly disabledReason?: string;
  /** Runs your action. */
  readonly onSelect?: () => void;
  /**
   * Merge the behavior onto the single child element instead of rendering a `<button>`. The
   * child's own value wins for each prop but the class, the style, and the handlers, so its own
   * `title` or `aria-describedby` hides the disabled reason.
   */
  readonly asChild?: boolean;
  /** Appended after the class of the button. */
  readonly className?: string;
  /** Visible content when there is no `icon`, or the element for `asChild`. */
  readonly children?: ReactNode;
}

/**
 * A host-owned toolbar action, styled and behaved like the packaged controls. Inside a
 * `Toolbar.Group` of the preset toolbar, it moves into the "More" panel with its group, where it
 * renders as a row with its icon and `label`. A direct child of the preset toolbar stands in the
 * block after the groups and never collapses. Under `preset={false}`, it renders where you put
 * it.
 *
 * @example
 * ```tsx
 * <XlsxEditor.Toolbar>
 *   <XlsxEditor.Toolbar.Group id="review" label="Review" priority={15}>
 *     <XlsxEditor.Toolbar.Action label="Send for review" icon={<SendIcon />} onSelect={send} />
 *   </XlsxEditor.Toolbar.Group>
 * </XlsxEditor.Toolbar>
 * ```
 *
 * @public
 */
export function ToolbarAction(props: ToolbarActionProps) {
  const { label, icon, active, disabled, disabledReason, onSelect, asChild, className, children } =
    props;
  const layout = useContext(PartLayoutContext);
  const reasonId = useId();
  const child = asChild ? asChildElement(children) : null;
  const badChild = !!asChild && !child;
  useEffect(() => {
    if (badChild) warnToolbarProblems([toolbarAsChildProblem('XlsxEditor.Toolbar.Action', label)]);
  }, [badChild, label]);
  const reason = disabled && disabledReason ? disabledReason : null;
  const row = layout === 'row';
  const shared = {
    className: [row ? 'xlsx-toolbar-row' : 'xlsx-toolbar-button', className]
      .filter(Boolean)
      .join(' '),
    'aria-label': label,
    // The same presence attributes as the registry controls, so one stylesheet selects both.
    ...(active ? { 'data-active': '' } : {}),
    ...(disabled ? { 'data-disabled': '', 'aria-disabled': 'true' as const } : {}),
    ...(active !== undefined ? { 'aria-pressed': active } : {}),
    ...(reason ? { 'aria-describedby': reasonId } : {}),
    title: reason ?? label,
    onMouseDown: guardToolbarMousedown,
    onKeyDown: (e: KeyboardEvent<HTMLElement>) => {
      if (e.key === 'Enter' && e.repeat) e.preventDefault();
    },
    // `onSelect` takes no argument: the event never reaches it. A disabled action also stops
    // the default of its element, so an `asChild` link does not navigate.
    onClick: (e: MouseEvent) => {
      if (disabled) e.preventDefault();
      else onSelect?.();
    },
  };
  const note = reason ? (
    <span id={reasonId} hidden>
      {reason}
    </span>
  ) : null;
  if (child)
    return (
      <>
        {renderAsChild(child, shared)}
        {note}
      </>
    );
  const own = asChild ? undefined : children;
  const content = icon ?? (own === null || own === false ? undefined : own);
  return (
    <button type="button" {...shared}>
      {row ? (
        <>
          {content !== undefined ? <span aria-hidden="true">{content}</span> : null}
          <span>{label}</span>
        </>
      ) : (
        (content ?? <span>{label}</span>)
      )}
      {note}
    </button>
  );
}

/** Props for `XlsxEditor.Toolbar.Separator`. @public */
export interface ToolbarSeparatorProps {
  /** Appended after the class of the separator. */
  readonly className?: string;
}

/**
 * A rule between controls: vertical in the bar, horizontal in a row of the "More" panel. Use it
 * under `preset={false}`, or inside a `Toolbar.Group`; the preset bar draws its own line
 * between two groups.
 *
 * @public
 */
export function ToolbarSeparator({ className }: ToolbarSeparatorProps) {
  const layout = useContext(PartLayoutContext);
  return (
    <div
      role="separator"
      aria-orientation={layout === 'row' ? 'horizontal' : 'vertical'}
      className={['xlsx-toolbar-separator', className].filter(Boolean).join(' ')}
    />
  );
}
