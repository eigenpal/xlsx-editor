// Apply inverse ops as undo and redo apply them: the ops, then the checks of the parts they
// touched. `transact` refuses inverse ops, so the tests of their own checks go through here.

import type { XlsxPackage } from '../package/package.ts';
import { applyOps, type HistoryOp, type OpRejection } from '../store/ops.ts';
import { changedParts } from '../store/document-store.ts';
import { touchedPartViolation } from '../store/part-checks.ts';

export function applyHistory(
  pkg: XlsxPackage,
  ops: readonly HistoryOp[]
):
  | { ok: true; pkg: XlsxPackage; inverse: readonly HistoryOp[] }
  | { ok: false; rejection: OpRejection } {
  const r = applyOps(pkg, ops, true);
  if (!r.ok) return r;
  const violation = touchedPartViolation(pkg, r.pkg, changedParts(pkg, r.pkg));
  if (violation) return { ok: false, rejection: { code: 'invalid-part', detail: violation } };
  return { ok: true, pkg: r.pkg, inverse: r.inverse };
}
