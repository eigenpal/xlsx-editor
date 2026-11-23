// The decimal step of the active cell, as the command treats that cell: a test oracle, kept out
// of the engine. The tests compare it with the ops of the command and with the state of the
// slots, which read the whole selection.
import type { EditorSnapshot } from '../editor-model.ts';
import {
  decimalStepOf,
  isGeneralCode,
  stepGeneralText,
  type DecimalStepResult,
} from '../number-format-commands.ts';
import { decimalStepsDetail } from '../decimal-selection.ts';

/**
 * The result of `decimalStepPreview`: the result of `decimalStepOf`, or `noPlaces` for a
 * decrease on a single `General` cell with no number or with no shown places.
 */
export type DecimalStepPreview =
  DecimalStepResult | { readonly ok: false; readonly problem: 'noPlaces' };

/**
 * A preview of the decimal step of the active cell, as the command treats that cell: the new code,
 * or the problem. A `General` number steps from the text that the grid shows
 * (`decimalStepsDetail(decimalSteps).shownText`), as the command does, and else from its full text.
 * An empty `General` cell gives `0.0` on an increase and `noPlaces` on a decrease when it is the
 * only selected cell. A cell that holds text, a boolean, or an error, and
 * among more cells a cell with no number (an empty position too, whatever its code) or a cell
 * that a merged range covers, gives `noNumber`
 * (`activeWritten` false), because the step never writes it. Null with no active cell
 * format. The state of the slots reads the whole selection (`decimalSteps`).
 * On more cells, a `General` number at the active
 * cell is the first of its column, so its step is the step that every `General` number of that
 * column takes when it succeeds; else each `General` number steps from its own places. A `General`
 * number in exponent form steps alone. On a single `General` number that shows no places, a
 * decrease gives `noPlaces`, as `decimalSteps.decrease` does.
 */
export function decimalStepPreview(
  snapshot: EditorSnapshot,
  step: 1 | -1
): DecimalStepPreview | null {
  const f = snapshot.activeFormat?.numberFormat;
  if (!f) return null;
  const {
    shownText: text,
    activeWritten,
    activeNumber: value,
  } = decimalStepsDetail(snapshot.decimalSteps);
  const single = snapshot.selectionSingle;
  // Among more selected cells, a cell with no number, whatever its code, or a cell that a
  // merged range covers, is never written: `activeWritten` is false for each.
  if (!activeWritten) return { ok: false, problem: 'noNumber' };
  const general = f.hasCode !== false && isGeneralCode(f.code);
  if (general && value === null && step < 0) return { ok: false, problem: 'noPlaces' };
  const fitted =
    general && text !== null && value !== null
      ? stepGeneralText(step, text, snapshot.numberLocale.decimal)
      : undefined;
  const result = fitted ?? decimalStepOf(f, step, value);
  // The only selected cell, a `General` number that shows no places: the reason speaks of the
  // cell, as `decimalSteps.decrease` does.
  if (single && general && !result.ok && result.problem === 'noDecimals')
    return { ok: false, problem: 'noPlaces' };
  return result;
}
