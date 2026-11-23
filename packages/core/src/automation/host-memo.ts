// What one batch of the automation host resolved, and the work counters of its tests. Taken out
// of `host.ts` to keep that file under its line cap; `host.ts` is the only user.

import type { SheetEntry } from '../store/index.ts';
import type { Element } from '../store/package/tree.ts';
import type { lockReader } from '../store/store/sheet-protection.ts';
import type { AutomationCellValue } from './protocol.ts';

/**
 * What the planner of one batch resolved, by element and by cell format index. The batch is
 * planned against the tree as it was opened, and the tree is immutable, so each entry holds for
 * the whole batch. A repeated read of a cell costs a lookup, never a new scan of its text nodes
 * or attributes.
 */
export interface ReadMemo {
  /** The sheet name text of the addresses of each worksheet entry, from `addressSheetText`. */
  readonly sheets: Map<SheetEntry, string>;
  /** The value of each cell element that a read gave. */
  readonly values: Map<Element, AutomationCellValue>;
  /** The cell format index text (`x:c/@s`) of each cell element, or undefined with none. */
  readonly styles: Map<Element, string | undefined>;
  /** The cell format of each `x:row` element, from `rowFormat`. */
  readonly rows: Map<Element, number | undefined>;
  /** The number format code of each cell format index text, or undefined when it is unknown. */
  readonly codes: Map<string | undefined, string | undefined>;
  /**
   * Whether each worksheet root holds `x:sheetProtection` (`CT_SheetProtection`). The writes of
   * a batch are planned against the tree as it was opened, so one scan of a root holds for each
   * write to it. DOCX Editor keeps the block controls of a story root the same way
   * (`blockControlsOf`, `packages/core/src/automation/plan-span-controls.ts`).
   */
  readonly guarded: Map<Element, boolean>;
  /** Whether each cell format index locks a cell (`x:protection/@locked`). */
  readonly locks: Map<number, boolean>;
  /** The lock reader of the cell formats, made at the first write to a protected worksheet. */
  locker?: ReturnType<typeof lockReader>;
}

export function readMemo(): ReadMemo {
  return {
    sheets: new Map(),
    values: new Map(),
    styles: new Map(),
    rows: new Map(),
    codes: new Map(),
    guarded: new Map(),
    locks: new Map(),
  };
}

// Work counters for tests, as `cellChildTestRecorder` (`store/store/cells.ts`): a count is exact
// on every machine, a wall-clock limit is not.
export const hostWork = {
  sheets: 0,
  names: 0,
  folds: 0,
  values: 0,
  styles: 0,
  rows: 0,
  codes: 0,
  guarded: 0,
  locks: 0,
  searches: 0,
};

/**
 * @internal Counts the work of a batch: the sheet name texts of addresses made, the selector
 * names looked up and folded, the cell values resolved, the `x:c/@s` and `x:row` formats read,
 * the cell format indexes resolved, the worksheet roots scanned for protection, the cell format locks read, and
 * the store passes that search for a refused write.
 */
export function hostWorkTestRecorder(): Readonly<typeof hostWork> & { reset(): void } {
  return {
    get sheets() {
      return hostWork.sheets;
    },
    get names() {
      return hostWork.names;
    },
    get folds() {
      return hostWork.folds;
    },
    get guarded() {
      return hostWork.guarded;
    },
    get locks() {
      return hostWork.locks;
    },
    get searches() {
      return hostWork.searches;
    },
    get values() {
      return hostWork.values;
    },
    get styles() {
      return hostWork.styles;
    },
    get rows() {
      return hostWork.rows;
    },
    get codes() {
      return hostWork.codes;
    },
    reset() {
      for (const key of Object.keys(hostWork) as (keyof typeof hostWork)[]) hostWork[key] = 0;
    },
  };
}

/** The value of `key` in `map`, made by `make` the first time. */
export function memoized<K, V>(map: Map<K, V>, key: K, make: () => V): V {
  if (map.has(key)) return map.get(key) as V;
  const value = make();
  map.set(key, value);
  return value;
}
