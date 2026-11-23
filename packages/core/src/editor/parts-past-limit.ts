// The parts that the read kept as bytes because they passed an XML part limit. The workbook opens
// without them, a save keeps their bytes, and the chrome shows a notice that names them.

import {
  CONTENT_TYPE,
  contentTypeOf,
  partPastLimit,
  type PackagePart,
  workbookIndex,
  type XlsxPackage,
} from '../store/index.ts';
import { xmlLimitOf } from '../automation/open-diagnostics.ts';
import type { PartPastLimit } from './editor-model.ts';

const KINDS = new Map<string, PartPastLimit['kind']>([
  [CONTENT_TYPE.styles, 'styles'],
  [CONTENT_TYPE.theme, 'theme'],
  [CONTENT_TYPE.drawing, 'drawing'],
]);

const NONE: readonly PartPastLimit[] = Object.freeze([]);
const cache = new WeakMap<XlsxPackage, readonly PartPastLimit[]>();
/**
 * The last list of each first part. An edit makes a new package but keeps the object of each
 * part that it does not change, so an equal list keeps the identity of the earlier array.
 */
const lastByFirstPart = new WeakMap<PackagePart, readonly PartPastLimit[]>();

const sameList = (a: readonly PartPastLimit[], b: readonly PartPastLimit[]): boolean =>
  a.length === b.length &&
  a.every((p, i) => p.part === b[i]!.part && p.kind === b[i]!.kind && p.limit === b[i]!.limit);

/**
 * The parts of `pkg` that the read kept as bytes past a limit, by part name. Read once for each
 * package. The list keeps its identity while its contents do not change.
 */
export function partsPastLimit(pkg: XlsxPackage): readonly PartPastLimit[] {
  const hit = cache.get(pkg);
  if (hit) return hit;
  const found: { entry: PartPastLimit; part: PackagePart }[] = [];
  let index: ReturnType<typeof workbookIndex> | null | undefined;
  for (const part of pkg.parts.values()) {
    const reason = partPastLimit(part);
    const limit = reason === undefined ? undefined : xmlLimitOf(reason);
    if (!limit) continue;
    const name = part.partName;
    let kind = name.toLowerCase().endsWith('.rels')
      ? 'relationships'
      : KINDS.get(contentTypeOf(pkg, name)?.trim().toLowerCase() ?? '');
    // A styles or a theme part changes the workbook only when the workbook part links to it, as
    // the store and the cascade read it. Another part of the same type gets the general kind.
    if (kind === 'styles' || kind === 'theme') {
      index ??= workbookIndex(pkg) ?? null;
      const used = kind === 'styles' ? index?.stylesPart : index?.themePart;
      if (used !== name) kind = 'other';
    }
    if (kind) found.push({ entry: Object.freeze({ part: name, kind, limit }), part });
  }
  found.sort((a, b) => (a.entry.part < b.entry.part ? -1 : a.entry.part > b.entry.part ? 1 : 0));
  let list = NONE;
  const first = found[0]?.part;
  if (first) {
    const fresh = found.map((f) => f.entry);
    const last = lastByFirstPart.get(first);
    list = last && sameList(last, fresh) ? last : Object.freeze(fresh);
    lastByFirstPart.set(first, list);
  }
  cache.set(pkg, list);
  return list;
}
