// The demo's heading switch. It rewrites `x:sheetView/@showRowColHeaders` in a copy of the
// workbook bytes: on, every worksheet view shows its row and column headings, so whole rows and
// columns can be selected by their headings in a file that hides them; off, each view gets back
// the value that the opened file had. Save also puts those values back, so the switch never
// changes the saved file. It reads and writes the bytes with the engine's package API; the
// editor then opens the copy as any other file.

import {
  readXlsxSteps,
  saveXlsxSteps,
  workbookIndex,
  xmlPartOf,
  type Element,
  type Node,
  type PackagePart,
  type PackageResult,
  type XlsxPackage,
} from '@xlsx-editor.dev/core/store';

/** The status that the demo shows after the headings option goes on. */
export const HEADINGS_SHOWN = 'All headings are shown.';
/** The status that the demo shows after the headings option goes off. */
export const HEADINGS_FROM_FILE = 'Headings show as the file sets them.';

/**
 * The `showRowColHeaders` value of each worksheet view, in document order, by worksheet part
 * name; `null` when unset.
 */
export type HeadingFlags = ReadonlyMap<string, readonly (string | null)[]>;

const isView = (n: Node): n is Element => n.type === 'element' && n.local === 'sheetView';
const isFlag = (a: Element['attrs'][number]) => a.ns === '' && a.local === 'showRowColHeaders';

function withFlag(view: Element, value: string | null): Element {
  const attrs = view.attrs.filter((a) => !isFlag(a));
  return {
    ...view,
    attrs: value === null ? attrs : [...attrs, { ns: '', local: 'showRowColHeaders', value }],
  };
}

/** Each worksheet part of `pkg` with the root element that it has. */
function worksheets(pkg: XlsxPackage) {
  const index = workbookIndex(pkg);
  return (index?.sheets ?? []).flatMap((sheet) => {
    const part = sheet.kind === 'worksheet' ? xmlPartOf(pkg, sheet.partName) : undefined;
    return part ? [part] : [];
  });
}

/** The `x:sheetView` elements of a worksheet root, in document order. */
const viewsOf = (root: Element): Element[] =>
  root.children.flatMap((n) =>
    n.type === 'element' && n.local === 'sheetViews' ? n.children.filter(isView) : []
  );

/** The heading flag of each worksheet view of `read`, to restore later. */
export function headingFlags(read: PackageResult): HeadingFlags {
  const flags = new Map<string, (string | null)[]>();
  if (!read.ok) return flags;
  for (const part of worksheets(read.pkg))
    flags.set(
      part.partName,
      viewsOf(part.root).map((view) => view.attrs.find(isFlag)?.value ?? null)
    );
  return flags;
}

/**
 * The bytes with every worksheet view showing its headings (`show` true), or with the flags of
 * `restore` put back (`show` false), in steps. Null when the bytes do not read, because bytes
 * with the wrong flags must not reach a file or a reload.
 */
export function* withHeadingsSteps(
  bytes: Uint8Array,
  show: boolean,
  restore: HeadingFlags
): Generator<void, Uint8Array | null> {
  const read = yield* readXlsxSteps(bytes);
  return read.ok ? yield* packageWithHeadingsSteps(read.pkg, show, restore) : null;
}

/** {@link withHeadingsSteps} for a workbook that is already read. */
export function packageWithHeadingsSteps(
  pkg: XlsxPackage,
  show: boolean,
  restore: HeadingFlags
): Generator<void, Uint8Array<ArrayBuffer>> {
  const parts = new Map<string, PackagePart>(pkg.parts);
  for (const part of worksheets(pkg)) {
    const was = restore.get(part.partName) ?? [];
    // The index of each view in document order, as `headingFlags` counts them.
    let at = 0;
    const root: Element = {
      ...part.root,
      children: part.root.children.map((n) =>
        n.type === 'element' && n.local === 'sheetViews'
          ? {
              ...n,
              children: n.children.map((v) => {
                if (!isView(v)) return v;
                const value = show ? '1' : (was[at] ?? null);
                at += 1;
                return withFlag(v, value);
              }),
            }
          : n
      ),
    };
    parts.set(part.partName, { kind: 'xml', partName: part.partName, xml: { ...part, root } });
  }
  return saveXlsxSteps({ parts });
}
