/** @jsxImportSource vue */
// Host-owned toolbar controls: `XlsxEditorToolbar.Action` and `XlsxEditorToolbar.Separator`. The
// Vue twin of the React adapter's `toolbar/action.tsx`, after
// docx-editor/packages/vue/src/editor/toolbar/ToolbarAction.ts and `ToolbarSeparator` of
// packages/vue/src/editor/toolbar/parts.ts at ad686574 (EigenPal); the header there lists the
// rules and the differences from DOCX Editor.
//
// Differences from React: the press is the `select` event (`@select`), as in DOCX Editor's Vue
// adapter, in place of the `onSelect` prop; `icon` is a prop that takes a vnode or a string;
// the content is the default slot; and a template `class` joins `className`.

import { defineComponent, inject, useId, type PropType, type VNodeChild } from 'vue';
import { guardToolbarMousedown, toolbarAsChildProblem } from '@xlsx-editor.dev/core/editor';
import { asChildElement, renderAsChild } from './as-child.tsx';
import { partLayoutKey, visibleNodes, warnToolbarProblems } from './host-parts.tsx';

const toolbarActionProps = {
  /**
   * The accessible name and the tooltip. A resolved string, not a catalog key: the action is
   * yours, so your own catalog translates it.
   */
  label: { type: String, required: true },
  /** Icon content, about 18 pixels, as the packaged icons: a vnode or a string. */
  icon: { type: [Object, String, Array] as PropType<VNodeChild>, default: undefined },
  /** The pressed state, for an action that toggles: `aria-pressed` and `data-active`. */
  active: { type: Boolean, default: undefined },
  /**
   * Turns the action off: the button stays focusable (`aria-disabled`, `data-disabled`), and a
   * press emits no `select`.
   */
  disabled: { type: Boolean, default: undefined },
  /** Why the action is disabled: the tooltip and the accessible description. */
  disabledReason: { type: String, default: undefined },
  /**
   * Merge the behavior onto the single element of the default slot instead of a `<button>`. The
   * element's own value wins for each prop but the class, the style, and the handlers, so its
   * own `title` or `aria-describedby` hides the disabled reason.
   */
  asChild: { type: Boolean, default: undefined },
  /** Appended after the class of the button. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/** Props of `XlsxEditorToolbar.Action`, the twin of React's `ToolbarActionProps`. @public */
export interface ToolbarActionProps {
  /** The accessible name and the tooltip: a resolved string from your own catalog. */
  readonly label: string;
  /** Icon content, about 18 pixels: a vnode or a string. */
  readonly icon?: VNodeChild;
  /** The pressed state, for an action that toggles: `aria-pressed` and `data-active`. */
  readonly active?: boolean;
  /** Turns the action off: the button stays focusable (`aria-disabled`). */
  readonly disabled?: boolean;
  /** Why the action is disabled: the tooltip and the accessible description. */
  readonly disabledReason?: string;
  /** Runs your action: the `select` event. */
  readonly onSelect?: () => void;
  /** Merge the behavior onto the single element of the default slot. Its own props win. */
  readonly asChild?: boolean;
  /** Appended after the class of the button. */
  readonly className?: string;
}

/**
 * A host-owned toolbar action, styled and behaved like the packaged controls. Inside an
 * `XlsxEditorToolbar.Group` of the preset toolbar, it moves into the "More" panel with its
 * group, where it renders as a row with its icon and `label`. In the default slot of the preset
 * toolbar, it stands in the block after the groups and never collapses. Under
 * `:preset="false"`, it renders where you put it.
 *
 * @example
 * ```ts
 * h(XlsxEditorToolbar, null, () => [
 *   h(XlsxEditorToolbar.Group, { id: 'review', label: 'Review', priority: 15 }, () => [
 *     h(XlsxEditorToolbar.Action, { label: 'Send for review', onSelect: send }),
 *   ]),
 * ]);
 * ```
 *
 * @public
 */
export const ToolbarAction = defineComponent({
  name: 'XlsxToolbarAction',
  inheritAttrs: false,
  props: toolbarActionProps,
  emits: { select: () => true },
  setup(props, { attrs, emit, slots }) {
    const layout = inject(partLayoutKey, 'bar');
    const reasonId = useId();
    return () => {
      const children = slots.default?.();
      const child = props.asChild ? asChildElement(children) : null;
      if (props.asChild && !child)
        warnToolbarProblems([toolbarAsChildProblem('XlsxEditorToolbar.Action', props.label)]);
      const reason = props.disabled && props.disabledReason ? props.disabledReason : null;
      const row = layout === 'row';
      // The host's other attributes land on the button, or on the element of `asChild`, also
      // when a disabled reason renders beside it.
      const { class: hostClass, ...hostAttrs } = attrs;
      const shared = {
        ...hostAttrs,
        class: [
          row ? 'xlsx-toolbar-row' : 'xlsx-toolbar-button',
          props.className,
          typeof hostClass === 'string' ? hostClass : undefined,
        ]
          .filter(Boolean)
          .join(' '),
        'aria-label': props.label,
        // The same presence attributes as the registry controls, so one stylesheet selects both.
        ...(props.active ? { 'data-active': '' } : {}),
        ...(props.disabled ? { 'data-disabled': '', 'aria-disabled': true } : {}),
        ...(props.active !== undefined ? { 'aria-pressed': props.active } : {}),
        ...(reason ? { 'aria-describedby': reasonId } : {}),
        title: reason ?? props.label,
        onMousedown: guardToolbarMousedown,
        onKeydown: (e: KeyboardEvent) => {
          if (e.key === 'Enter' && e.repeat) e.preventDefault();
        },
        // A disabled action also stops the default of its element, so an `asChild` link does
        // not navigate.
        onClick: (e: MouseEvent) => {
          if (props.disabled) e.preventDefault();
          else emit('select');
        },
      };
      const note = reason ? (
        <span id={reasonId} hidden>
          {reason}
        </span>
      ) : null;
      if (child) return [renderAsChild(child, shared), note];
      const own = props.asChild ? [] : visibleNodes(children);
      const content: VNodeChild | undefined = props.icon ?? (own.length > 0 ? own : undefined);
      return (
        <button type="button" {...shared}>
          {row ? (
            <>
              {content !== undefined ? <span aria-hidden="true">{content}</span> : null}
              <span>{props.label}</span>
            </>
          ) : (
            (content ?? <span>{props.label}</span>)
          )}
          {note}
        </button>
      );
    };
  },
});

const toolbarSeparatorProps = {
  /** Appended after the class of the separator. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/** Props of `XlsxEditorToolbar.Separator`, the twin of React's `ToolbarSeparatorProps`. @public */
export interface ToolbarSeparatorProps {
  /** Appended after the class of the separator. */
  readonly className?: string;
}

/**
 * A rule between controls: vertical in the bar, horizontal in a row of the "More" panel. Use it
 * under `:preset="false"`, or inside an `XlsxEditorToolbar.Group`; the preset bar draws its own
 * line between two groups.
 *
 * @public
 */
export const ToolbarSeparator = defineComponent({
  name: 'XlsxToolbarSeparator',
  props: toolbarSeparatorProps,
  setup(props) {
    const layout = inject(partLayoutKey, 'bar');
    return () => (
      <div
        role="separator"
        aria-orientation={layout === 'row' ? 'horizontal' : 'vertical'}
        class={['xlsx-toolbar-separator', props.className]}
      />
    );
  },
});
