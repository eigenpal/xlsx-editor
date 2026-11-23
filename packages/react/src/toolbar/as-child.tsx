/** @jsxImportSource react */
// `asChild`: a host element renders a toolbar control, with the control's attributes and
// handlers merged in. A copy of docx-editor/packages/react/src/editor/toolbar/Slot.tsx at
// 592db9fd (EigenPal), as a function in place of a component. Differences:
// - `asChildElement` reads the one element child of a part, and returns null for any other
//   children, so the part can warn and fall back to its own button, as
//   pptx-editor/packages/react/src/toolbar-action.tsx at 70052cf1 (EigenPal) does.
// - The child's ref is read from its props in React 19 and from the element in React 18.
// - An element that is not a `<button>` takes the role of a button, a place in the Tab order,
//   and Enter and Space as a press (`buttonRole`), so the keyboard reaches a host `<span>` or
//   `<div>` as it reaches the packaged buttons.
// The merge rules are DOCX Editor's: the class names join, the styles merge with the child's
// value winning, the handlers run the child's first and skip the control's when the child
// called `preventDefault()`, the refs fan out to both, and the child's value wins for each
// other prop.

import {
  Children,
  cloneElement,
  isValidElement,
  type KeyboardEvent,
  type ReactElement,
  type ReactNode,
  type Ref,
  version,
} from 'react';

type AnyProps = Record<string, unknown>;

function composeHandlers(childHandler: unknown, slotHandler: unknown): unknown {
  if (typeof slotHandler !== 'function') return childHandler;
  if (typeof childHandler !== 'function') return slotHandler;
  return (event: { defaultPrevented: boolean }) => {
    (childHandler as (e: unknown) => void)(event);
    if (!event.defaultPrevented) (slotHandler as (e: unknown) => void)(event);
  };
}

function mergeProps(slotProps: AnyProps, childProps: AnyProps): AnyProps {
  const merged: AnyProps = { ...slotProps, ...childProps };
  for (const key of Object.keys(slotProps)) {
    if (/^on[A-Z]/.test(key)) merged[key] = composeHandlers(childProps[key], slotProps[key]);
    else if (key === 'className')
      merged[key] = [slotProps[key], childProps[key]].filter(Boolean).join(' ');
    else if (key === 'style')
      merged[key] = { ...(slotProps[key] as object), ...(childProps[key] as object) };
  }
  return merged;
}

/** The one element child of an `asChild` part, or null for none, several, or text. */
export function asChildElement(children: ReactNode): ReactElement | null {
  const all = Children.toArray(children);
  const [only] = all;
  return all.length === 1 && isValidElement(only) ? only : null;
}

/**
 * The props that make an element that is not a button act as one: the button role, a place in
 * the Tab order, and Enter or Space as a click.
 */
function buttonRole(child: ReactElement): AnyProps {
  if (child.type === 'button') return { type: 'button' };
  return {
    role: 'button',
    tabIndex: 0,
    onKeyDown: (event: KeyboardEvent<HTMLElement>) => {
      if (event.key !== 'Enter' && event.key !== ' ') return;
      event.preventDefault();
      if (!event.repeat) event.currentTarget.click();
    },
  };
}

const REF_IN_PROPS = Number.parseInt(version, 10) >= 19;

/** Render `child` with the control's props merged in, as DOCX Editor's `Slot` does. */
export function renderAsChild(child: ReactElement, slotProps: AnyProps): ReactElement {
  const own = buttonRole(child);
  // The control's own key handler runs after the button role's, both after the child's.
  const props: AnyProps = { ...own, ...slotProps };
  if (own.onKeyDown && slotProps.onKeyDown)
    props.onKeyDown = composeHandlers(own.onKeyDown, slotProps.onKeyDown);
  const merged = mergeProps(props, child.props as AnyProps);
  // React 19 keeps the ref in the props; React 18 keeps it on the element, and React 19 warns
  // when code reads it there.
  const childRef = REF_IN_PROPS
    ? (child.props as { ref?: Ref<unknown> }).ref
    : (child as ReactElement & { ref?: Ref<unknown> }).ref;
  const slotRef = slotProps.ref as Ref<unknown> | undefined;
  if (childRef || slotRef) {
    merged.ref = (node: unknown) => {
      for (const ref of [childRef, slotRef]) {
        if (typeof ref === 'function') ref(node);
        else if (ref && typeof ref === 'object') (ref as { current: unknown }).current = node;
      }
    };
  }
  return cloneElement(child, merged);
}
