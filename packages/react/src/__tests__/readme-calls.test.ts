// The calls that the React guides in `packages/react/docs/` show, written as the guides write
// them, so typecheck proves them:
// `refusalText(refusal, t)` and `disabledReasonText(reason, t)` take the adapter's `TFunction`.

import { describe, expect, test } from 'bun:test';
import {
  disabledReasonText,
  refusalText,
  type ChromeDisabledReason,
  type EditRefusal,
} from '@xlsx-editor.dev/core/editor';
import type { TFunction } from '@xlsx-editor.dev/i18n';
import { translatorFor } from '../index.ts';

describe('React guide calls', () => {
  test('refusalText and disabledReasonText take the adapter translator', () => {
    const t: TFunction = translatorFor(undefined);
    const refusal: EditRefusal = {
      code: 'structure-blocked',
      about: { kind: 'defined-name', name: 'Rates' },
    };
    expect(refusalText(refusal, t)).toContain('the defined name Rates');
    const reason: ChromeDisabledReason = { key: 'disabledReason.nothingToUndo' };
    expect(disabledReasonText(reason, t)).toBe('There is nothing to undo.');
  });
});
