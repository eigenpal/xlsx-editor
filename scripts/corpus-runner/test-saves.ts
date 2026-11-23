// Faulty writers for the tests of the corpus tasks (`census.test.ts`, `validate-xsd.test.ts`).
// Each one saves a changed package, so a test can show that a task finds the change.

import { saveXlsx, type Element, type XlsxPackage } from '../../packages/core/src/store/index.ts';

/** Whether some text of `el` or its descendants holds `text`. */
function holdsText(el: Element, text: string): boolean {
  return el.children.some((child) =>
    child.type === 'text' ? child.value.includes(text) : holdsText(child, text)
  );
}

/**
 * A writer that sets `@attribute` of every `element` to `value` before it saves. With `when`, it
 * changes only a package that holds that text, such as the text of an edit, and saves any
 * other package as it is.
 */
export function saveWithAttribute(
  element: string,
  attribute: string,
  value: string,
  when?: string
): (pkg: XlsxPackage) => Uint8Array {
  const walk = (el: Element): Element => ({
    ...el,
    attrs:
      el.local === element
        ? el.attrs.map((a) => (a.local === attribute ? { ...a, value } : a))
        : el.attrs,
    children: el.children.map((child) => (child.type === 'element' ? walk(child) : child)),
  });
  return (pkg) => {
    const edited = [...pkg.parts.values()].some(
      (part) => part.kind === 'xml' && holdsText(part.xml.root, when ?? '')
    );
    if (when !== undefined && !edited) return saveXlsx(pkg);
    const parts = new Map(
      [...pkg.parts].map(([name, part]) => [
        name,
        part.kind === 'xml' ? { ...part, xml: { ...part.xml, root: walk(part.xml.root) } } : part,
      ])
    );
    return saveXlsx({ ...pkg, parts } as XlsxPackage);
  };
}

/** A writer whose output is not a workbook. */
export const saveGarbage = (): Uint8Array => new TextEncoder().encode('not a zip');
