// The document store: the single source of truth and the only write path.
//
// One transaction = one atomic publication = at most one history entry (DOCX Editor's
// TreeDocumentStore contract, through pptx-editor/packages/core/src/store/store/
// document-store.ts at 1943d2c9). `transact` stages every op against a working package; nothing
// is visible until all ops apply and the touched parts pass their checks, so a refused op leaves
// the revision, the package, and every subscriber exactly as they were.
//
// Differences from PPTX Editor, each for a reason:
// - A history entry records inverse ops, never package snapshots (design D2). Undo applies the
//   inverse ops as a new transaction, and its own inverse becomes the redo entry. An inverse
//   restores whole rows, so it can overwrite a later change in the same row by another writer;
//   the collaboration change decides the position model and the inverse granularity (design D10).
// - No `remote`, `projection`, or `awareness` origin, no `merged` flag, and no `block`: those
//   serve collaboration and progressive open, which this change does not have (design D10).
// - No history groups. A history entry keeps the selection before and after its transaction
//   when the caller supplies them, as DOCX Editor's package history does
//   (docx-editor/packages/core/src/store/store/package-history-selection.ts): undo gives back
//   the selection before, and redo the selection after. The store copies and freezes them, and
//   never reads them. Three differences from DOCX Editor: the caller passes the selections as
//   transaction options, not through the context callbacks `selectionBefore` and
//   `selectionAfter`; an omitted `selectionAfter` stays omitted, with no default from the caret;
//   and a selection keeps its anchor and its focus corner, which DOCX Editor's `SelectionMark`
//   (docx-editor/packages/core/src/store/store/tree-store.ts) does not keep, because a grid
//   range needs the corner that Shift+Arrow extends.
// - No `replaceParts` and no constructor options: nothing replaces whole parts, and listener
//   errors go to the host's `reportError`.
// - A transaction outside the history (`system`, or `recordsHistory: false`) clears both
//   stacks: the inverse ops of older entries were made against a package that no longer
//   holds, and one of them could restore a shorter styles part under a cell that this
//   transaction pointed past it (a repair prompt). The side effects of designs D3 and D5 (the
//   stale flag and the dropped calculation chain) are not separate transactions: `applyOps`
//   adds them to the transaction that needs them, with no inverse, so one user command is one
//   revision and one undo step, and undo never brings the stale chain back or clears the flag.
// The store holds no session state in the tree: history lives here, and so do the part names of
// the package that the store opened. A sheet whose part is not among them is a sheet that an add
// op of the session made, which the sheet ops can rename and delete in a workbook whose content
// refuses that for the sheets of the file (`sheetEditBlocker`).
//
// Listeners are isolated: one that throws does not stop delivery to the others, and never makes
// a committed transaction look failed.

import type { XlsxPackage } from '../package/package.ts';
import { applyOps, type HistoryOp, type OpRejection, type WorkbookOp } from './ops.ts';
import { partNamesOf } from '../structure/blockers.ts';
import { touchedPartViolation } from './part-checks.ts';
import { reportListenerError } from './listener-error.ts';
import { sameNode } from './calc.ts';
import { sheetByName, sheetNameKey, workbookIndex } from '../workbook/workbook-index.ts';
import type { RangeBox } from './clear-range.ts';

/** What made a transaction. */
export type TransactOrigin = 'user' | 'undo' | 'redo' | 'system';

/** What a committed transaction changed. */
export interface StoreChange {
  readonly fromRevision: number;
  readonly toRevision: number;
  readonly origin: TransactOrigin;
  /** Part names whose content changed. Layout re-does only what these touch. */
  readonly parts: readonly string[];
  readonly label?: string;
  /**
   * The ops as committed, with the side-effect ops that the transaction added at the end. An
   * undo or a redo commits inverse ops (`InverseOp`), which `transact` does not take.
   */
  readonly ops: readonly HistoryOp[];
}

/** The outcome of a transaction. */
export type TransactResult =
  | { readonly ok: true; readonly revision: number }
  | { readonly ok: false; readonly rejection: OpRejection };

/** Options for one transaction. */
export interface TransactOptions {
  /** A short label for the history entry (`set value`). */
  readonly label?: string;
  /** What made the transaction; `user` when omitted. */
  readonly origin?: TransactOrigin;
  /**
   * Whether this transaction enters the undo stack. Omitted: `true` for `user`, `false` for
   * `system`. A transaction outside the history clears the undo and redo stacks. The stale
   * flag is not such a transaction: it is a side effect of the user transaction.
   */
  readonly recordsHistory?: boolean;
  /**
   * The revision that the caller read the references of its ops against. When the store has
   * moved on since, the transaction is refused with `stale-revision`, because a reference can
   * now name another cell (a structural edit moved it). Omitted: the current revision.
   */
  readonly baseRevision?: number;
  /**
   * The selection before the transaction, kept with its history entry:
   * {@link WorkbookDocumentStore.selectionForUndo} gives it back. Omitted: undo of the entry
   * names no selection.
   */
  readonly selectionBefore?: HistorySelection;
  /**
   * The selection after the transaction, kept with its history entry:
   * {@link WorkbookDocumentStore.selectionForRedo} gives it back. Omitted: redo of the entry
   * names no selection. Its sheet is a worksheet after the ops, so a rename or an add can name
   * the new sheet.
   */
  readonly selectionAfter?: HistorySelection;
  /**
   * The name of the sheet that the transaction edits, kept with its history entry also when
   * it keeps no selection: {@link WorkbookDocumentStore.sheetForUndo} gives it back, so a host
   * can name the sheet that an undo or a redo changes. It is a worksheet before or after the
   * ops: a delete names the deleted sheet, and an add the added one.
   */
  readonly sheet?: string;
}

/** A cell of a {@link HistorySelection}, 0-based. */
export interface HistoryCell {
  readonly row: number;
  readonly col: number;
}

/**
 * A selection that a history entry keeps: a sheet by name, a range, the active cell, the
 * anchor that the range extends from, and the focus corner that Shift+Arrow moves.
 */
export interface HistorySelection {
  readonly sheet: string;
  /** The range that holds the active cell. */
  readonly range: RangeBox;
  readonly active: HistoryCell;
  readonly anchor: HistoryCell;
  readonly focus: HistoryCell;
  /**
   * Every range of a selection of two or more ranges, in the order of selection, or absent for
   * one range. `range` is `ranges[activeRange]`.
   */
  readonly ranges?: readonly RangeBox[];
  /** The position in `ranges` of the range that holds the active cell. Present with `ranges`. */
  readonly activeRange?: number;
}

const frozenCell = (c: HistoryCell): HistoryCell => Object.freeze({ row: c.row, col: c.col });

/**
 * The most ranges that a history selection keeps: the editor's `MAX_SELECTION_RANGES`, which the
 * store lane cannot import. A test keeps the two equal. A host transaction can name more.
 */
const HISTORY_RANGES = 64;

/** A frozen copy of a selection, so no caller can change what the history keeps. */
function frozenSelection(s: HistorySelection): HistorySelection {
  const box = ([top, left, bottom, right]: RangeBox): RangeBox =>
    Object.freeze([top, left, bottom, right] as const);
  const kept = s.ranges?.slice(0, HISTORY_RANGES);
  const ranges = kept && kept.length > 1 ? kept : undefined;
  // `range` is the active range; past the kept ranges, the last kept range is active.
  const at = !ranges
    ? 0
    : Number.isInteger(s.activeRange) && s.activeRange! >= 0 && s.activeRange! < ranges.length
      ? s.activeRange!
      : ranges.length - 1;
  // When the active range moved, the active cell moves to its first corner, so it stays inside.
  const holds = (
    [top, left, bottom, right]: RangeBox,
    { row, col }: { row: number; col: number }
  ) => row >= top && row <= bottom && col >= left && col <= right;
  const moved = ranges !== undefined && !holds(ranges[at]!, s.active);
  const corner = moved ? { row: ranges[at]![0], col: ranges[at]![1] } : undefined;
  return Object.freeze({
    sheet: s.sheet,
    range: box(ranges ? ranges[at]! : s.range),
    active: frozenCell(corner ?? s.active),
    anchor: frozenCell(corner ?? s.anchor),
    focus: frozenCell(corner ?? s.focus),
    ...(ranges ? { ranges: Object.freeze(ranges.map(box)), activeRange: at } : {}),
  });
}

/** What a transaction changed that can change what the text of a formula refers to. */
interface ReferenceEdit {
  /** Rows or columns moved on a sheet. */
  readonly axis: boolean;
  /** The names of the sheets that were renamed, moved, or deleted. */
  readonly sheets: readonly string[];
  /** A sheet was added, renamed, moved, or deleted: a span of sheets can hold other sheets. */
  readonly order: boolean;
}

/** The reference edit of `ops`, or undefined when they hold none. */
function referenceEditOf(ops: readonly WorkbookOp[]): ReferenceEdit | undefined {
  let axis = false;
  let order = false;
  const sheets: string[] = [];
  for (const op of ops) {
    if (
      op.op === 'insertRows' ||
      op.op === 'deleteRows' ||
      op.op === 'insertColumns' ||
      op.op === 'deleteColumns'
    )
      axis = true;
    else if (op.op === 'renameSheet') sheets.push(op.sheet, op.name);
    else if (op.op === 'moveSheet' || op.op === 'deleteSheet') sheets.push(op.sheet);
    order ||= op.op === 'addSheet' || sheets.length > 0;
  }
  return axis || order ? { axis, sheets, order } : undefined;
}

interface HistoryEntry {
  /** The ops that undo the entry, in the order to apply them. */
  readonly inverse: readonly HistoryOp[];
  /** What an undo or a redo of the entry changes again for the text of a formula. */
  readonly references?: ReferenceEdit;
  readonly label?: string;
  readonly sheet?: string;
  readonly selectionBefore?: HistorySelection;
  readonly selectionAfter?: HistorySelection;
}

/** The sheet and the selections of an entry, the selections copied, when they are present. */
function selectionsOf(s: {
  readonly selectionBefore?: HistorySelection;
  readonly selectionAfter?: HistorySelection;
  readonly sheet?: string;
}): Pick<HistoryEntry, 'sheet' | 'selectionBefore' | 'selectionAfter'> {
  return {
    ...(s.sheet !== undefined ? { sheet: s.sheet } : {}),
    ...(s.selectionBefore ? { selectionBefore: frozenSelection(s.selectionBefore) } : {}),
    ...(s.selectionAfter ? { selectionAfter: frozenSelection(s.selectionAfter) } : {}),
  };
}

/**
 * The first sheet name of the history options that is not a worksheet name, or undefined. The
 * selection before names a sheet of the package before the ops, and the selection after a
 * sheet of the package after them (a rename or an add gives a new name). `sheet`, the sheet
 * that the ops change, can be in either: a delete removes it, and an add creates it.
 */
function unknownHistorySheet(
  before: XlsxPackage,
  after: XlsxPackage,
  o: TransactOptions
): string | undefined {
  const indexes = [workbookIndex(before), workbookIndex(after)];
  const worksheet = (name: string, at: readonly (0 | 1)[]) =>
    at.some((i) => {
      const index = indexes[i];
      return !!index && sheetByName(index, name)?.kind === 'worksheet';
    });
  const checks: [string | undefined, readonly (0 | 1)[]][] = [
    [o.sheet, [0, 1]],
    [o.selectionBefore?.sheet, [0]],
    [o.selectionAfter?.sheet, [1]],
  ];
  return checks.find(([name, at]) => name !== undefined && !worksheet(name, at))?.[0];
}

const HISTORY_LIMIT = 200;

/** Whether every named part holds the same content in `a` and `b` (bytes by identity). */
function sameParts(a: XlsxPackage, b: XlsxPackage, names: readonly string[]): boolean {
  return names.every((name) => {
    const x = a.parts.get(name);
    const y = b.parts.get(name);
    if (!x || !y) return x === y;
    if (x.kind === 'xml' && y.kind === 'xml') return sameNode(x.xml.root, y.xml.root);
    return x.kind === 'binary' && y.kind === 'binary' && x.bytes === y.bytes;
  });
}

/** Part names whose part object differs between two packages. */
export function changedParts(a: XlsxPackage, b: XlsxPackage): string[] {
  const out: string[] = [];
  for (const [name, part] of b.parts) if (a.parts.get(name) !== part) out.push(name);
  for (const name of a.parts.keys()) if (!b.parts.has(name)) out.push(name);
  return out;
}

/** The transactional document store of one workbook. */
export class WorkbookDocumentStore {
  #pkg: XlsxPackage;
  readonly #openedParts: ReadonlySet<string>;
  #revision = 0;
  #undo: HistoryEntry[] = [];
  #redo: HistoryEntry[] = [];
  #listeners = new Set<(change: StoreChange) => void>();
  #axisEdits = 0;
  #sheetEdits = new Map<string, number>();
  #orderEdits = 0;

  constructor(pkg: XlsxPackage) {
    this.#pkg = pkg;
    this.#openedParts = partNamesOf(pkg);
  }

  /** The current package. Reference-stable until the next commit. */
  get package(): XlsxPackage {
    return this.#pkg;
  }

  /** Monotonic revision, bumped by every commit (including undo and redo). */
  get revision(): number {
    return this.#revision;
  }

  /**
   * How many committed edits can have changed what a formula that names `sheets` refers to: each
   * insert or delete of rows or columns, and each rename, move, or delete of a sheet of `sheets`.
   * With `span`, for a formula with a reference across a span of sheets (`Jan:Mar!A1`), each add,
   * rename, move, or delete of any sheet counts: the span can then hold other sheets, and the
   * formula does not name them. An undo and a redo of such an edit count too, so the number never
   * goes back.
   */
  referenceEdits(sheets: readonly string[], span = false): number {
    let count = this.#axisEdits + (span ? this.#orderEdits : 0);
    for (const name of sheets) count += this.#sheetEdits.get(sheetNameKey(name)) ?? 0;
    return count;
  }

  #countReferenceEdit(edit: ReferenceEdit | undefined): void {
    if (!edit) return;
    if (edit.axis) this.#axisEdits += 1;
    if (edit.order) this.#orderEdits += 1;
    for (const name of new Set(edit.sheets.map(sheetNameKey)))
      this.#sheetEdits.set(name, (this.#sheetEdits.get(name) ?? 0) + 1);
  }

  get canUndo(): boolean {
    return this.#undo.length > 0;
  }

  get canRedo(): boolean {
    return this.#redo.length > 0;
  }

  /**
   * The selection before the entry that {@link undo} reverts next, as a copy, or null when its
   * transaction supplied none. A host restores it after the undo.
   */
  selectionForUndo(): HistorySelection | null {
    const s = this.#undo[this.#undo.length - 1]?.selectionBefore;
    return s ? frozenSelection(s) : null;
  }

  /** The label of the entry that {@link undo} reverts next (`renameSheet`), or null. */
  labelForUndo(): string | null {
    return this.#undo[this.#undo.length - 1]?.label ?? null;
  }

  /** The label of the entry that {@link redo} applies next, or null. */
  labelForRedo(): string | null {
    return this.#redo[this.#redo.length - 1]?.label ?? null;
  }

  /**
   * The sheet of the entry that {@link undo} reverts next, or null when it names none. For a
   * rename, it is the name of the sheet before the rename: the name that the undo gives back.
   */
  sheetForUndo(): string | null {
    return this.#undo[this.#undo.length - 1]?.sheet ?? null;
  }

  /**
   * The sheet of the entry that {@link redo} applies next, or null when it names none. For a
   * rename, it is the name of the sheet before the rename, as {@link sheetForUndo} gives.
   */
  sheetForRedo(): string | null {
    return this.#redo[this.#redo.length - 1]?.sheet ?? null;
  }

  /**
   * The selection after the entry that {@link redo} applies next, as a copy, or null when its
   * transaction supplied none. A host restores it after the redo.
   */
  selectionForRedo(): HistorySelection | null {
    const s = this.#redo[this.#redo.length - 1]?.selectionAfter;
    return s ? frozenSelection(s) : null;
  }

  /** Validate ops, and the parts they leave, without committing. */
  dryRun(ops: readonly WorkbookOp[]): TransactResult {
    const r = this.#stage(ops, false);
    return r.ok ? { ok: true, revision: this.#revision } : r;
  }

  #stage(
    ops: readonly HistoryOp[],
    history: boolean
  ):
    | {
        readonly ok: true;
        readonly pkg: XlsxPackage;
        readonly parts: string[];
        readonly inverse: readonly HistoryOp[];
        readonly effects: readonly WorkbookOp[];
      }
    | { readonly ok: false; readonly rejection: OpRejection } {
    const r = applyOps(this.#pkg, ops, history, this.#openedParts);
    if (!r.ok) return r;
    const parts = changedParts(this.#pkg, r.pkg);
    const violation = touchedPartViolation(this.#pkg, r.pkg, parts, history);
    if (violation) return { ok: false, rejection: { code: 'invalid-part', detail: violation } };
    return { ok: true, pkg: r.pkg, parts, inverse: r.inverse, effects: r.effects };
  }

  /** Apply ops atomically and publish one change. */
  transact(ops: readonly WorkbookOp[], options: TransactOptions = {}): TransactResult {
    const origin = options.origin ?? 'user';
    if (options.baseRevision !== undefined && options.baseRevision !== this.#revision)
      return {
        ok: false,
        rejection: {
          code: 'stale-revision',
          detail: `the ops were read at revision ${options.baseRevision}, and the store is at ${this.#revision}`,
        },
      };
    // Inverse ops come only from the history: a caller cannot write parts or rows through them.
    const r = ops.length === 0 ? null : this.#stage(ops, false);
    if (r && !r.ok) return r;
    // The history names sheets by name: a name that is not a worksheet of the workbook would
    // send a host to a sheet that does not exist.
    const unknown = unknownHistorySheet(this.#pkg, r?.pkg ?? this.#pkg, options);
    if (unknown !== undefined)
      return {
        ok: false,
        rejection: {
          code: 'unknown-sheet',
          detail: `${JSON.stringify(unknown)} is not a worksheet`,
        },
      };
    if (!r) return { ok: true, revision: this.#revision };
    // A transaction that changes nothing commits nothing: no revision and no history entry.
    if (r.pkg === this.#pkg) return { ok: true, revision: this.#revision };
    const records = options.recordsHistory ?? origin === 'user';
    const references = referenceEditOf(ops);
    this.#countReferenceEdit(references);
    if (records) {
      this.#undo.push({
        inverse: r.inverse,
        ...(references ? { references } : {}),
        ...(options.label ? { label: options.label } : {}),
        ...selectionsOf(options),
      });
      if (this.#undo.length > HISTORY_LIMIT) this.#undo.shift();
      this.#redo = [];
    } else {
      // Older inverse ops do not apply on top of a change that has no inverse in the history.
      this.#undo = [];
      this.#redo = [];
    }
    return this.#publish(r.pkg, origin, r.parts, [...ops, ...r.effects], options.label);
  }

  /**
   * Undo the newest history entry. When its inverse no longer applies, the store clears both
   * stacks, because no older entry can apply either, and returns the reason.
   */
  undo(): TransactResult {
    return this.#step(this.#undo, this.#redo, 'undo');
  }

  /** Redo the newest undone entry, with the same refusal rule as {@link undo}. */
  redo(): TransactResult {
    return this.#step(this.#redo, this.#undo, 'redo');
  }

  #step(from: HistoryEntry[], to: HistoryEntry[], origin: 'undo' | 'redo'): TransactResult {
    const entry = from[from.length - 1];
    if (!entry)
      return {
        ok: false,
        rejection: { code: `nothing-to-${origin}`, detail: `nothing to ${origin}` },
      };
    const r = this.#stage(entry.inverse, true);
    if (!r.ok) {
      this.#undo = [];
      this.#redo = [];
      return r;
    }
    from.pop();
    to.push({
      inverse: r.inverse,
      ...(entry.references ? { references: entry.references } : {}),
      ...(entry.label ? { label: entry.label } : {}),
      ...selectionsOf(entry),
    });
    // A step that leaves every part as it was moves the entry and publishes nothing: no
    // revision, so a host's unsaved check stays true to the content. The store keeps the new
    // package, equal in content, so the moved entry's inverse applies to the nodes it names.
    // A reader that still holds the old package, or references read at this revision, stays
    // valid: every part holds the same content, so each cell reference names the same cell,
    // and no listener runs, because nothing that a listener reads has changed.
    if (sameParts(this.#pkg, r.pkg, r.parts)) {
      this.#pkg = r.pkg;
      return { ok: true, revision: this.#revision };
    }
    this.#countReferenceEdit(entry.references);
    return this.#publish(r.pkg, origin, r.parts, [...entry.inverse, ...r.effects], entry.label);
  }

  /** Subscribe to committed changes. Returns the unsubscribe function. */
  subscribe(listener: (change: StoreChange) => void): () => void {
    this.#listeners.add(listener);
    return () => this.#listeners.delete(listener);
  }

  #publish(
    pkg: XlsxPackage,
    origin: TransactOrigin,
    parts: readonly string[],
    ops: readonly HistoryOp[],
    label: string | undefined
  ): TransactResult {
    const fromRevision = this.#revision;
    this.#pkg = pkg;
    this.#revision += 1;
    const change: StoreChange = {
      fromRevision,
      toRevision: this.#revision,
      origin,
      parts,
      ops,
      ...(label ? { label } : {}),
    };
    for (const listener of [...this.#listeners]) {
      try {
        listener(change);
      } catch (error) {
        reportListenerError(error);
      }
    }
    return { ok: true, revision: this.#revision };
  }
}
