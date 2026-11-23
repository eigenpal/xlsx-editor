// The format commands, DOM-free: a format patch, a clear, or a toggle, lowered to store ops
// (design D2, "Ops carry intent"). The editor and the fixture edits of the save check call the
// same lowering, so a command and its saved fixture write the same ops.
//
// The store checks every patch value (`store/store/format-patch.ts`): a custom number format
// code must be one that the number format parser reads, or the op is refused with
// `unreadable-number-format`, and a code that equals a built-in format code of ECMA-376 Part 1,
// 18.8.30, is written as that built-in id. The lowering here applies the same check, so a
// command is refused before it reaches the store.

import type { ResolvedCellFormat } from '../cascade/index.ts';
import {
  lowerNumberFormat,
  type FormatPatch,
  type OpRejection,
  type WorkbookOp,
} from '../store/index.ts';

/** Ops for a command, or why the command is refused before the store sees it. */
export type CommandOps =
  | { readonly ok: true; readonly ops: readonly WorkbookOp[] }
  | { readonly ok: false; readonly rejection: OpRejection };

/**
 * The `applyFormat` op of a patch on a range of a sheet. A number format code that the parser
 * cannot read is refused here; the op keeps the code, and the store lowers it with the
 * workbook's own `x:numFmt` records (`lowerNumberFormat`).
 */
export function formatOps(sheet: string, ref: string, patch: FormatPatch): CommandOps {
  const lowered = lowerNumberFormat(patch);
  if ('problem' in lowered)
    return { ok: false, rejection: { code: 'unreadable-number-format', detail: lowered.problem } };
  return { ok: true, ops: [{ op: 'applyFormat', sheet, ref, patch }] };
}

/** What a clear removes: values and formulas, cell formats, or both. */
export type ClearKind = 'contents' | 'formats' | 'all';

/** The ops of a clear of a range. `all` is one transaction of both clears. */
export function clearOps(sheet: string, ref: string, kind: ClearKind): readonly WorkbookOp[] {
  const values: WorkbookOp = { op: 'clearValues', sheet, ref };
  const formats: WorkbookOp = { op: 'clearFormats', sheet, ref };
  return kind === 'contents' ? [values] : kind === 'formats' ? [formats] : [values, formats];
}

/** The font properties that a shortcut toggles. */
export type FormatToggle = 'bold' | 'italic' | 'underline';

/**
 * The patch that toggles a font property, from the resolved format of the active cell: the
 * same format that the editor shows as the toolbar state, so the state and the command agree.
 */
export function togglePatch(format: ResolvedCellFormat, which: FormatToggle): FormatPatch {
  switch (which) {
    case 'bold':
      return { bold: !format.font.bold };
    case 'italic':
      return { italic: !format.font.italic };
    case 'underline':
      return { underline: format.font.underline === 'none' };
  }
}
