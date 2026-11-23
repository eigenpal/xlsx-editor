// The decimal step of codes with `?` padding beside a literal: the room of the
// zero section, the padding of other sections, and the steps that would scale the number. A
// walk of steps returns the code where the step keeps no approximation.

import { describe, expect, test } from 'bun:test';
import type { NumberLocale } from '../../cascade/index.ts';
import { numberFormatChoiceCode } from '../index.ts';
import { numberFormatPlaces } from '../number-format-commands.ts';
import { stepDecimals } from '../number-format-commands.ts';

const EN: NumberLocale = { decimal: '.', group: ',' };

describe('stepDecimals with padding', () => {
  test('accounting keeps the room of the zero section in step with its places', () => {
    const accounting = numberFormatChoiceCode('accounting', EN)!;
    const codes = [accounting];
    for (const step of [-1, -1, 1, 1] as const) {
      const r = stepDecimals(codes[codes.length - 1]!, step);
      expect(r.ok).toBe(true);
      if (r.ok) codes.push(r.code);
    }
    expect(codes).toEqual([
      accounting,
      '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)',
      '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
      '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)',
      accounting,
    ]);
    expect(stepDecimals(codes[2]!, -1)).toEqual({ ok: false, problem: 'noDecimals' });
    expect(stepDecimals(accounting, 1)).toEqual({
      ok: true,
      code: '_("$"* #,##0.000_);_("$"* \\(#,##0.000\\);_("$"* "-"???_);_(@_)',
    });
    // A bare "-" with no fill and no `?` is no room: it stays as it is.
    expect(stepDecimals('0;-0;"-"', 1)).toEqual({ ok: true, code: '0.0;-0.0;"-"' });
    expect(stepDecimals('0.0;-0.0;"-"?', -1)).toEqual({ ok: true, code: '0;-0;"-"' });
  });

  test('a zero section with a bare literal stays as it is, and the steps return the code', () => {
    const walk = (code: string, steps: readonly (1 | -1)[]): string => {
      let at = code;
      for (const step of steps) {
        const r = stepDecimals(at, step);
        if (!r.ok) throw new Error(`${at} does not step ${step}: ${r.problem}`);
        at = r.code;
      }
      return at;
    };
    const bare = '0.00;-0.00;"-"';
    expect(stepDecimals(bare, 1)).toEqual({ ok: true, code: '0.000;-0.000;"-"' });
    expect(walk(bare, [1, -1])).toBe(bare);
    expect(walk(bare, [-1, -1, 1, 1])).toBe(bare);
    expect(stepDecimals('#,##0.0_);(#,##0.0);"-"_)', 1)).toEqual({
      ok: true,
      code: '#,##0.00_);(#,##0.00);"-"_)',
    });
    // The accounting code goes down to no places and back with no loss.
    const accounting = numberFormatChoiceCode('accounting', EN)!;
    expect(walk(accounting, [-1, -1, 1, 1])).toBe(accounting);
    // A section of a literal and `?` padding keeps its padding, with no point, when another
    // section has a decimal point.
    expect(stepDecimals('0.00E+00;"-"??', 1)).toEqual({ ok: true, code: '0.000E+00;"-"??' });
    expect(walk('0.00E+00;"-"??', [-1, 1])).toBe('0.00E+00;"-"??');
    // The first section too.
    expect(stepDecimals('"-"??;0.00', 1)).toEqual({ ok: true, code: '"-"??;0.000' });
    expect(walk('"-"??;0.00', [-1, 1])).toBe('"-"??;0.00');
    // With no point in any section, `?` padding takes and gives back a point.
    expect(stepDecimals('??" kg";-??" kg"', 1)).toEqual({ ok: true, code: '??.0" kg";-??.0" kg"' });
    expect(walk('??" kg";-??" kg"', [1, 1, -1, -1])).toBe('??" kg";-??" kg"');
    expect(stepDecimals('??" kg";-??" kg";0', 1)).toEqual({
      ok: true,
      code: '??.0" kg";-??.0" kg";0.0',
    });
    expect(walk('"$"??;"$"??;"-"??', [1, -1])).toBe('"$"??;"$"??;"-"??');
    expect(walk('"-"??;"-"??', [1, -1])).toBe('"-"??;"-"??');
    expect(stepDecimals('"-"??;"-"??', -1)).toEqual({ ok: false, problem: 'noDecimals' });
    // Only a digit placeholder outside such a run, in a section other than the zero section,
    // keeps the run as padding.
    expect(walk('0.00;"-"??', [-1, -1])).toBe('0;"-"??');
    expect(walk('0.00;"-"??', [-1, -1, 1, 1])).toBe('0.00;"-"??');
    expect(walk('0.00E+00;"-"??', [-1, -1, 1, 1])).toBe('0.00E+00;"-"??');
    expect(walk('"-"??;0.00', [-1, -1, 1, 1])).toBe('"-"??;0.00');
    expect(stepDecimals('0;"-"??', 1)).toEqual({ ok: true, code: '0.0;"-"??' });
    expect(stepDecimals('"-"??;0', 1)).toEqual({ ok: true, code: '"-"??;0.0' });
    expect(walk('??" kg";-??" kg";0', [1, 1, -1, -1])).toBe('??" kg";-??" kg";0');
    expect(walk('"-"??', [1, -1])).toBe('"-"??');
    // A zero section with its own digit placeholders counts as a number section.
    expect(walk('"a"??;"b"??;0.00', [-1, 1])).toBe('"a"??;"b"??;0.00');
    expect(walk('"a"??;"b"??;0.00', [-1, -1, 1, 1])).toBe('"a"??;"b"??;0.00');
    // An empty zero room takes a `?` only when the sections before it had no places, as after a
    // decrease to no places. Else it stays empty, so the steps return in either order.
    expect(stepDecimals('"a";-0.00;* "-"', 1)).toEqual({ ok: true, code: '"a";-0.000;* "-"' });
    const emptyRoom = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"_);_(@_)';
    expect(walk(emptyRoom, [1, -1])).toBe(emptyRoom);
    expect(walk(emptyRoom, [-1, 1])).toBe(emptyRoom);
    expect(walk(emptyRoom, [-1, -1, 1, 1])).toBe(emptyRoom.replace('"-"_', '"-"??_'));
    const noPlaces = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)';
    expect(walk(noPlaces, [1, -1])).toBe(noPlaces);
    // In the zero section, a dash as `\\-` or `-` before the run is a literal too.
    const escaped = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \\-??_);_(@_)';
    expect(stepDecimals(escaped, 1)).toEqual({
      ok: true,
      code: escaped.replaceAll('0.00', '0.000').replace('-??', '-???'),
    });
    expect(walk(escaped, [-1, -1, 1, 1])).toBe(escaped);
    expect(stepDecimals('0.00;-0.00;-??', 1)).toEqual({ ok: true, code: '0.000;-0.000;-???' });
    // Outside the zero section, `-??` shows a negative number, and steps as a number.
    expect(stepDecimals('??;-??', 1)).toEqual({ ok: true, code: '??.0;-??.0' });
    // A text section's point does not count: `@.` has no digit placeholder.
    expect(stepDecimals('"-"??;@.', 1)).toEqual({ ok: true, code: '"-"??.0;@.' });
    // A code whose only number is the room of its zero section steps that room.
    expect(stepDecimals('"pos";"neg";"-"??', 1)).toEqual({ ok: true, code: '"pos";"neg";"-"???' });
    expect(stepDecimals('"pos";"neg";"-"??', -1)).toEqual({ ok: true, code: '"pos";"neg";"-"?' });
    // A section with a point and no digit placeholder stays as it is.
    expect(stepDecimals('0.00;.', 1)).toEqual({ ok: true, code: '0.000;.' });
    expect(stepDecimals('0.00;.,', -1)).toEqual({ ok: true, code: '0.0;.,' });
    // A trailing point does not come back: `0.` steps up to `0.0` and down to `0`.
    expect(walk('0.', [1, -1])).toBe('0');
    // So does a fill between `"-"` and its `?`: the `?` come back after the fill.
    expect(walk('0.00;-0.00;"-"* ??', [-1, -1, 1, 1])).toBe('0.00;-0.00;"-"* ??');
    // Two runs of `?` in the zero section never merge: they stay as they are.
    expect(stepDecimals('0.0;-0.0;"a"?"b"?', 1)).toEqual({ ok: true, code: '0.00;-0.00;"a"?"b"?' });
    expect(stepDecimals('0.0;-0.0;"a"?"b"?', -1)).toEqual({ ok: true, code: '0;-0;"a"?"b"?' });
    // Beside another literal, a full decrease loses the run, as the TSDoc of stepDecimals says.
    expect(walk('0.00;-0.00;"x"??', [-1, -1])).toBe('0;-0;"x"');
    // A run of `?` before the literal is no padding, also in the zero section: it steps as a
    // number, as `??" kg"` does.
    expect(walk('0.00;-0.00;??"x"', [1])).toBe('0.000;-0.000;??.0"x"');
    expect(walk('0.00;-0.00;??"-"', [1])).toBe('0.000;-0.000;??.0"-"');
    // The room of the zero section follows its first `"-"`, also with a literal after the `?`.
    const euro = '_-* #,##0.00\\ "€"_-;\\-* #,##0.00\\ "€"_-;_-* "-"??\\ "€"_-;_-@_-';
    expect(walk(euro, [-1, -1])).toBe('_-* #,##0\\ "€"_-;\\-* #,##0\\ "€"_-;_-* "-"\\ "€"_-;_-@_-');
    expect(walk(euro, [-1, -1, 1, 1])).toBe(euro);
    expect(numberFormatPlaces(euro)).toBe(2);
    // A code whose only number is the room keeps its last `?` when no increase finds the room
    // again: a room with no fill, or with no `"-"`.
    expect(stepDecimals('"pos";"neg";"-"?', -1)).toEqual({ ok: false, problem: 'noDecimals' });
    expect(stepDecimals('"pos";"neg";"x"?', -1)).toEqual({ ok: false, problem: 'noDecimals' });
    expect(walk('"pos";"neg";* "-"?', [-1, 1])).toBe('"pos";"neg";* "-"?');
  });

  test('a step that would scale the number or change its separators is refused', () => {
    // The decrease of `0.00,0` would leave `0.00,`, whose comma scales the number by 1000.
    expect(stepDecimals('0.00,0', -1)).toEqual({ ok: false, problem: 'unreadable' });
    expect(stepDecimals('#,##0,.0', -1)).toEqual({ ok: false, problem: 'unreadable' });
    // A comma that already scales the number stays, with the same scale.
    expect(stepDecimals('0.0,', -1)).toEqual({ ok: true, code: '0,' });
    expect(stepDecimals('#,##0,', 1)).toEqual({ ok: true, code: '#,##0.0,' });
  });
});
