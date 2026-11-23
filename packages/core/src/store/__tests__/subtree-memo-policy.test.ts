// After DOCX Editor's `packages/core/src/store/__tests__/subtree-memo-policy.test.ts` at
// 6c11927b. Differences: the nodes are worksheet cells, and the answers are the cell child checks
// of a value edit (`holdsWrappedValue`), not the paragraph id and move range reads.

import { expect, test } from 'bun:test';
import { parseXmlPart, type Element, type Node } from '../package/tree.ts';
import { keepsSubtreeMemo, WIDE_SUBTREE_CHILDREN } from '../package/subtree-memo-policy.ts';
import { holdsWrappedValue } from '../store/cells.ts';

const X = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const element = (node: Node | undefined) => node as Element;

/** The cells of one row, from `cells` markup. */
function cellsOf(cells: string): Element[] {
  const xml =
    `<worksheet xmlns="${X}" xmlns:q="urn:q"><sheetData><row r="1">${cells}</row>` +
    '</sheetData></worksheet>';
  const read = parseXmlPart('/xl/worksheets/sheet1.xml', xml);
  if (!read.ok) throw Error(read.reason);
  const data = element(read.part.root.children.find((n) => n.type === 'element'));
  return element(data.children[0]).children.map(element);
}

test('only wide nodes and nodes with an element below their children keep an entry', () => {
  const leaves = '<q:b/>'.repeat(WIDE_SUBTREE_CHILDREN - 1);
  const [narrow, inline, wide] = cellsOf(
    `<c r="A1"><v>1</v></c><c r="B1" t="inlineStr"><is><t>x</t></is></c>` +
      `<c r="C1"><v>1</v><q:w>${leaves}<q:i/></q:w></c>`
  );
  const value = element(narrow!.children[0]);
  expect(keepsSubtreeMemo(value.children[0]!)).toBe(false);
  expect(keepsSubtreeMemo(value)).toBe(false);
  expect(keepsSubtreeMemo(narrow!)).toBe(false);
  expect(keepsSubtreeMemo(element(inline!.children[0]))).toBe(false);
  expect(keepsSubtreeMemo(inline!)).toBe(true);
  // A wide text-only child also bounds the policy scan.
  const wideValue = {
    ...value,
    children: Array.from({ length: WIDE_SUBTREE_CHILDREN }, () => value.children[0]!),
  };
  expect(keepsSubtreeMemo({ ...narrow!, children: [wideValue] })).toBe(true);
  expect(keepsSubtreeMemo(element(wide!.children[1]))).toBe(true);
  expect(keepsSubtreeMemo(wide!)).toBe(true);
});

test('answers over nodes without an entry stay equal to a fresh read', () => {
  // The wrapper holds only leaves, so it keeps no entry; its cell keeps one.
  const [plain, wrapped, empty] = cellsOf(
    '<c r="A1"><v>1</v></c><c r="B1"><q:w><v>2</v></q:w></c><c r="C1"><q:w/></c>'
  );
  expect(keepsSubtreeMemo(element(wrapped!.children[0]))).toBe(false);
  expect(keepsSubtreeMemo(plain!)).toBe(false);
  expect(keepsSubtreeMemo(empty!)).toBe(false);
  for (let pass = 0; pass < 3; pass += 1) {
    const answers = [plain, wrapped, empty].map((cell) => holdsWrappedValue(cell!));
    const fresh = [plain, wrapped, empty].map((cell) => holdsWrappedValue(structuredClone(cell!)));
    expect(answers).toEqual(fresh);
    expect(answers).toEqual([false, true, false]);
  }
});
