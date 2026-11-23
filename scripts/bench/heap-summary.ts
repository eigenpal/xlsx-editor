// The heap growth by constructor, for the memory bench with `MEMORY_SNAPSHOTS=1`. A copy of
// `Summary`, `heapSummary`, and `summaryDiff` of PPTX Editor's `e2e/memory.bench.ts` at d4502552
// (EigenPal), first in f448820c. Differences: `Summary`, `heapSummary`, and
// `summaryDiff` are exported, so `browser-bench.ts` can use them.

import type { CDPSession } from '@playwright/test';

export type Summary = Map<string, { count: number; size: number }>;

/**
 * A heap snapshot, summed by constructor name (detached DOM nodes keep their own name). The
 * snapshot of a large heap is longer than the longest string, so it stays bytes: the node list
 * is read number by number, and only the meta data and the string table become JSON.
 */
export async function heapSummary(cdp: CDPSession): Promise<Summary> {
  const chunks: Buffer[] = [];
  const onChunk = (p: { chunk: string }): void => void chunks.push(Buffer.from(p.chunk));
  cdp.on('HeapProfiler.addHeapSnapshotChunk', onChunk);
  await cdp.send('HeapProfiler.collectGarbage');
  await cdp.send('HeapProfiler.takeHeapSnapshot', { reportProgress: false });
  cdp.off('HeapProfiler.addHeapSnapshotChunk', onChunk);
  const bytes = Buffer.concat(chunks);
  chunks.length = 0;
  const at = (marker: string): number => {
    const i = bytes.indexOf(marker);
    if (i < 0) throw new Error(`no ${marker} in the heap snapshot`);
    return i + marker.length;
  };
  const nodesAt = at('"nodes":[');
  const meta = JSON.parse(
    bytes.subarray(bytes.indexOf('{'), bytes.lastIndexOf(',', nodesAt)).toString() + '}'
  ) as { snapshot: { meta: { node_fields: string[]; node_types: [string[]] } } };
  const stringsAt = at('"strings":[');
  const strings = JSON.parse(
    `[${bytes.subarray(stringsAt, bytes.lastIndexOf(']')).toString()}]`
  ) as string[];
  const fields = meta.snapshot.meta.node_fields;
  const types = meta.snapshot.meta.node_types[0];
  const [iType, iName, iSize] = ['type', 'name', 'self_size'].map((f) => fields.indexOf(f));
  const out: Summary = new Map();
  const node: number[] = [];
  let value = -1;
  for (let i = nodesAt; i < bytes.length; i += 1) {
    const c = bytes[i]!;
    if (c >= 48 && c <= 57) {
      value = (value < 0 ? 0 : value * 10) + (c - 48);
      continue;
    }
    if (value >= 0) node.push(value);
    value = -1;
    if (node.length === fields.length) {
      const type = types[node[iType!]!]!;
      if (type !== 'hidden' && type !== 'synthetic') {
        const name =
          type === 'object' || type === 'native' || type === 'closure'
            ? strings[node[iName!]!]!
            : `(${type})`;
        const entry = out.get(name) ?? { count: 0, size: 0 };
        entry.count += 1;
        entry.size += node[iSize!]!;
        out.set(name, entry);
      }
      node.length = 0;
    }
    if (c === 93) break; // The `]` that ends the node list.
  }
  return out;
}

export function summaryDiff(before: Summary, after: Summary, top = 15): string[] {
  const rows = [...after].map(([name, a]) => {
    const b = before.get(name) ?? { count: 0, size: 0 };
    return { name, count: a.count - b.count, size: a.size - b.size };
  });
  const detached = rows.filter((r) => r.name.startsWith('Detached ') && r.count > 0);
  // The browser's own layout objects (`blink::`, `cppgc::`) count apart from the page's objects.
  const native = (r: { name: string }): boolean => /^(blink|cppgc|v8)::/.test(r.name);
  const blink = rows.filter(native).reduce((kb, r) => kb + r.size / 1024, 0);
  const grown = rows
    .filter((r) => !native(r))
    .sort((x, y) => y.size - x.size)
    .slice(0, top);
  return [
    `  ${blink.toFixed(0).padStart(8)} KB          (browser objects: layout, style, text shaping)`,
    ...grown.map(
      (r) =>
        `  ${(r.size / 1024).toFixed(0).padStart(8)} KB ${String(r.count).padStart(7)}  ${r.name.slice(0, 80)}`
    ),
    ...(detached.length
      ? [`  detached: ${detached.map((r) => `${r.name} +${r.count}`).join(', ')}`]
      : []),
  ];
}
