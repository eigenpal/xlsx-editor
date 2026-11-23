/** @jsxImportSource vue */
// `asChild`: a host element renders a toolbar control, with the control's attributes and
// handlers merged in. The Vue twin of the React adapter's `toolbar/as-child.tsx`, after
// docx-editor/packages/vue/src/editor/toolbar/Slot.ts at 592db9fd (EigenPal); the header there
// lists the rules. Differences from DOCX Editor: a function over the vnode in place of a
// component, so the control's ref and the child's ref stay the child's own (no ref is
// forwarded), with the same FULL_PROPS patch flag on the clone; `asChildElement` returns null for children that are not one element, so the
// part can warn and fall back to its own button; and an element that is not a `<button>`
// takes the role of a button, a place in the Tab order, and Enter and Space as a press, as in
// React.

import { Comment, Fragment, Text, cloneVNode, isVNode, mergeProps, type VNode } from 'vue';

type AnyProps = Record<string, unknown>;

function composeHandlers(childHandler: unknown, slotHandler: unknown): unknown {
  if (typeof slotHandler !== 'function') return childHandler;
  if (typeof childHandler !== 'function') return slotHandler;
  return (event: Event) => {
    (childHandler as (e: Event) => void)(event);
    if (!event.defaultPrevented) (slotHandler as (e: Event) => void)(event);
  };
}

/** The class value of a prop: a string, or the joined names of an array or an object. */
function classText(value: unknown): string {
  if (typeof value === 'string') return value;
  if (Array.isArray(value)) return value.map(classText).filter(Boolean).join(' ');
  if (value && typeof value === 'object')
    return Object.entries(value as Record<string, unknown>)
      .filter(([, on]) => !!on)
      .map(([name]) => name)
      .join(' ');
  return '';
}

function mergeSlotProps(slotProps: AnyProps, childProps: AnyProps): AnyProps {
  const merged = mergeProps(slotProps, childProps) as AnyProps;
  for (const key of Object.keys(slotProps)) {
    if (/^on[A-Z]/.test(key)) merged[key] = composeHandlers(childProps[key], slotProps[key]);
    else if (key === 'class')
      merged[key] = [classText(slotProps[key]), classText(childProps[key])]
        .filter(Boolean)
        .join(' ');
    else if (key === 'style')
      merged[key] = { ...(slotProps[key] as object), ...(childProps[key] as object) };
  }
  // The child's own value wins for each other prop.
  for (const [key, value] of Object.entries(childProps))
    if (!/^on[A-Z]/.test(key) && key !== 'class' && key !== 'style') merged[key] = value;
  return merged;
}

/** The one element vnode of `content`, or null for none, several, or text. */
export function asChildElement(content: unknown): VNode | null {
  const nodes: VNode[] = [];
  let text = false;
  const walk = (node: unknown): void => {
    if (Array.isArray(node)) node.forEach(walk);
    else if (isVNode(node)) {
      if (node.type === Fragment) walk(node.children);
      else if (node.type === Text) text ||= String(node.children ?? '').trim() !== '';
      else if (node.type !== Comment) nodes.push(node);
    } else if (typeof node === 'string' || typeof node === 'number')
      text ||= String(node).trim() !== '';
  };
  walk(content);
  return nodes.length === 1 && !text ? nodes[0]! : null;
}

/**
 * The props that make an element that is not a button act as one: the button role, a place in
 * the Tab order, and Enter or Space as a click.
 */
function buttonRole(child: VNode): AnyProps {
  if (child.type === 'button') return { type: 'button' };
  return {
    role: 'button',
    tabindex: 0,
    onKeydown: (event: KeyboardEvent) => {
      if (event.key !== 'Enter' && event.key !== ' ') return;
      event.preventDefault();
      if (!event.repeat) (event.currentTarget as HTMLElement).click();
    },
  };
}

/** `child` with the control's props merged in, as DOCX Editor's `Slot` merges them. */
export function renderAsChild(child: VNode, slotProps: AnyProps): VNode {
  const own = buttonRole(child);
  const props: AnyProps = { ...own, ...slotProps };
  if (own.onKeydown && slotProps.onKeydown)
    props.onKeydown = composeHandlers(own.onKeydown, slotProps.onKeydown);
  const merged = mergeSlotProps(props, (child.props ?? {}) as AnyProps);
  // Extra props, even empty ones, make `cloneVNode` add the FULL_PROPS patch flag, as DOCX
  // Editor's `Slot` gets it from its extra `ref`. Without it, the compiler patch flag of a
  // template child (such as CLASS only) stays, and the patch never updates the merged state
  // attributes or handlers.
  const cloned = cloneVNode(child, {});
  // `merged` already holds the child's props: merging it into the clone again would repeat the
  // child's class and run the child's handlers twice.
  cloned.props = merged;
  return cloned;
}
