// The user text of refusals (`refusalText`): one whole-sentence message of the catalog for each
// code, a save message for each code that a save can return, a locked message by the
// protection rule, no placeholder in any text, and a bounded length.

import { describe, expect, test } from 'bun:test';
import en from '../../../../i18n/en.json';
import {
  CUT_REFUSAL_CODES,
  EDIT_REFUSAL_CODES,
  SAVE_REFUSAL_CODES,
  disabledReasonText,
  refusalText,
} from '../index.ts';
import { displayName } from '../cell-input.ts';
import { STRUCTURE_SUBJECT_KINDS } from '../../store/index.ts';
import { storeOf } from '../../store/__tests__/structure-helpers.ts';

/**
 * The English catalog, with parameters filled in. A placeholder without a parameter stays as
 * `{name}`, so a test sees it.
 */
function english(key: string, params?: Readonly<Record<string, string | number>>): string {
  let node: unknown = en;
  for (const part of key.split('.')) node = (node as Record<string, unknown>)[part];
  return String(node).replace(/\{(\w+)\}/g, (all, k: string) => String(params?.[k] ?? all));
}

describe('refusalText', () => {
  test('each refusal is one whole message: a save, a locked command, and a load', () => {
    const texts = [
      refusalText({ code: 'not-editable', command: 'save' }, english),
      refusalText({ code: 'loading', command: 'save' }, english),
      refusalText({ code: 'composing', command: 'save' }, english),
      refusalText({ code: 'unsupported-function', subject: 'IFNA', command: 'save' }, english),
      refusalText({ code: 'nothing-to-undo', command: 'save' }, english),
      refusalText({ code: 'loading' }, english),
      // `detail` is for logs: it never changes the text.
      refusalText({ code: 'not-editable', detail: 'loading' }, english),
      refusalText({ code: 'locked' }, english),
      refusalText({ code: 'locked', command: 'format', lock: 'format' }, english),
      refusalText({ code: 'locked', command: 'clear', lock: 'format' }, english),
      refusalText({ code: 'locked', command: 'clear', lock: 'cell' }, english),
      refusalText({ code: 'locked', command: 'clearAll', lock: 'format' }, english),
      refusalText({ code: 'locked', command: 'clearAll', lock: 'cell' }, english),
      refusalText({ code: 'locked', command: 'structure', lock: 'structure' }, english),
    ];
    expect(texts).toEqual([
      'The workbook was not saved: no workbook is open.',
      'The workbook was not saved: a workbook is loading.',
      'The workbook was not saved: finish the text that the input method composes first.',
      'The workbook was not saved: the formula of the cell edit calls IFNA, which the editor cannot write. Use a predefined function.',
      'The workbook was not saved: the open cell edit cannot be stored.',
      'The command was not applied: a workbook is loading.',
      'The edit was not applied: no workbook is open for editing.',
      'The edit was not applied: the sheet is protected, and the cell is locked.',
      'The format was not applied: the sheet is protected, and formatting cells is not allowed.',
      'The formats were not cleared: the sheet is protected, and formatting cells is not allowed.',
      'The cells were not cleared: the sheet is protected, and a cell in the selection is locked.',
      'The cells were not cleared: the sheet is protected, and formatting cells is not allowed.',
      'The cells were not cleared: the sheet is protected, and a cell in the selection is locked.',
      'The rows or columns were not changed: the sheet is protected, and inserting or deleting them is not allowed.',
    ]);
    for (const text of texts)
      for (const sentence of text.split(/(?<=\.) /))
        expect(sentence.split(' ').length).toBeLessThanOrEqual(20);
  });

  test('a shared master refused for the formula of one cell names both cells', () => {
    const base = {
      code: 'shared-formula',
      subject: 'T!A2:A3',
      causeCell: 'T!A3',
      masterCell: 'T!A2',
    } as const;
    const cause =
      'T!A2 is a shared formula master, and the editor cannot rewrite the formula of T!A3.';
    const texts = {
      edit: `The edit was not applied: ${cause} Edit another cell.`,
      save: `The workbook was not saved: ${cause} Cancel the cell edit, then save.`,
      cut: `The cells were not cut: ${cause} Cut a range without that master.`,
      clear: `The cells were not cleared: ${cause} Clear a range without that master.`,
      clearAll: `The cells were not cleared: ${cause} To remove only the formats, use Clear formats.`,
      paste: `The cells were not pasted: ${cause} Paste to a range without that master.`,
      merge: `The cells were not merged: ${cause} Merge a range without that master.`,
    };
    for (const [command, text] of Object.entries(texts))
      expect(
        refusalText(
          { ...base, ...(command === 'edit' ? {} : { command: command as 'save' }) },
          english
        )
      ).toBe(text);
    // A master whose formula cannot be read says so, not that it cannot be rewritten.
    const unread = { ...base, causeCell: 'T!A2', causeUnreadable: true } as const;
    expect(refusalText({ ...unread, command: 'clear' }, english)).toBe(
      'The cells were not cleared: T!A2 is a shared formula master whose formula cannot be read. Clear a range without that master.'
    );
    // A cause with no master shows the message of a refusal that names nothing.
    expect(refusalText({ code: 'shared-formula', causeCell: 'T!A3' }, english)).toBe(
      refusalText({ code: 'shared-formula' }, english)
    );
  });

  test('a formula range refusal names the cell or range that blocks it, when it has one', () => {
    expect(refusalText({ code: 'shared-formula', subject: 'T!A3' }, english)).toBe(
      'The edit was not applied: it would change the formula range at T!A3, which the editor cannot rewrite or remove. Edit another cell.'
    );
    expect(refusalText({ code: 'shared-formula' }, english)).toBe(
      'The edit was not applied: it would change a formula range that the editor cannot rewrite or remove. Edit another cell.'
    );
    expect(refusalText({ code: 'shared-formula', command: 'save' }, english)).toBe(
      'The workbook was not saved: the cell edit would change a formula range that the editor cannot rewrite. Cancel the cell edit, then save.'
    );
    expect(
      refusalText({ code: 'shared-formula', command: 'save', subject: 'T!A2:A3' }, english)
    ).toBe(
      'The workbook was not saved: the cell edit would change the formula range at T!A2:A3, which the editor cannot rewrite. Cancel the cell edit, then save.'
    );
    // A cut names the cut, not an edit.
    expect(
      refusalText({ code: 'shared-formula', command: 'cut', subject: 'T!A2:A3' }, english)
    ).toBe(
      'The cells were not cut: clearing them would change the formula range at T!A2:A3, which the editor cannot rewrite or remove. Cut a range without cells of that range.'
    );
    expect(refusalText({ code: 'shared-formula', command: 'cut' }, english)).toBe(
      'The cells were not cut: clearing them would change a formula range that the editor cannot rewrite or remove. Cut a range without cells of that range.'
    );
    // A clear and a paste name their command too. Clear all points to Clear formats.
    expect(
      refusalText({ code: 'shared-formula', command: 'clear', subject: 'T!A2:A3' }, english)
    ).toBe(
      'The cells were not cleared: clearing them would change the formula range at T!A2:A3, which the editor cannot rewrite or remove. Clear a range without cells of that range.'
    );
    expect(refusalText({ code: 'shared-formula', command: 'clear' }, english)).toBe(
      'The cells were not cleared: clearing them would change a formula range that the editor cannot rewrite or remove. Clear a range without cells of that range.'
    );
    expect(
      refusalText({ code: 'shared-formula', command: 'clearAll', subject: 'T!A2:A3' }, english)
    ).toBe(
      'The cells were not cleared: clearing them would change the formula range at T!A2:A3, which the editor cannot rewrite or remove. To remove only the formats, use Clear formats.'
    );
    expect(refusalText({ code: 'shared-formula', command: 'clearAll' }, english)).toBe(
      'The cells were not cleared: clearing them would change a formula range that the editor cannot rewrite or remove. To remove only the formats, use Clear formats.'
    );
    expect(
      refusalText({ code: 'shared-formula', command: 'paste', subject: 'T!A2:A3' }, english)
    ).toBe(
      'The cells were not pasted: the paste would change the formula range at T!A2:A3, which the editor cannot rewrite or remove. Paste to a range without cells of that range.'
    );
    expect(refusalText({ code: 'shared-formula', command: 'paste' }, english)).toBe(
      'The cells were not pasted: the paste would change a formula range that the editor cannot rewrite or remove. Paste to a range without cells of that range.'
    );
  });

  test('a formula inside mc:AlternateContent names its cell and says why', () => {
    const about = { kind: 'wrapped-formula', name: 'S!B2' } as const;
    const named =
      'The edit was not applied: cell S!B2 holds its formula in a form that the editor cannot ' +
      'update. The edit can change that formula.';
    // A row or column edit refuses with `structure-blocked`, and a sheet move or delete that
    // changes a sheet range of the formula with `sheet-range-blocked`.
    for (const code of ['structure-blocked', 'sheet-range-blocked'] as const) {
      const text = refusalText({ code, subject: 'S!B2', about }, english);
      expect(text).toBe(named);
      expect(text.split(/\s+/).length).toBeLessThanOrEqual(26);
      expect(refusalText({ code, about: { kind: 'wrapped-formula' } }, english)).toBe(
        'The edit was not applied: a cell holds its formula in a form that the editor cannot ' +
          'update. The edit can change that formula.'
      );
    }
  });

  test('the store refusal of a sheet move over a wrapped sheet range gives that message', () => {
    // T!B2 keeps `SUM(S:U!A1)` only in the branches of `mc:AlternateContent`.
    const f = '<f>SUM(S:U!A1)</f><v>1</v>';
    const wrapped =
      '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
      `<mc:Choice Requires="mc">${f}</mc:Choice><mc:Fallback>${f}</mc:Fallback></mc:AlternateContent>`;
    const one = (name: string) => ({ name, rows: '<row r="1"><c r="A1"><v>1</v></c></row>' });
    for (const op of [
      { op: 'moveSheet', sheet: 'V', index: 1 },
      { op: 'deleteSheet', sheet: 'U' },
    ] as const) {
      const s = storeOf({
        sheets: [
          one('S'),
          { name: 'T', rows: `<row r="2"><c r="B2">${wrapped}</c></row>` },
          one('U'),
          one('V'),
        ],
      });
      const r = s.transact([op]);
      expect(r.ok).toBe(false);
      if (r.ok) continue;
      expect(r.rejection.code).toBe('sheet-range-blocked');
      expect(refusalText(r.rejection, english)).toBe(
        'The edit was not applied: cell T!B2 holds its formula in a form that the editor cannot ' +
          'update. The edit can change that formula.'
      );
    }
  });

  test('a move or a delete of a sheet that is not a worksheet says so, not a reference', () => {
    // The editor refuses a move, and the store a delete, with these refusals.
    const move = {
      code: 'structure-blocked',
      subject: 'Chart1',
      about: { kind: 'not-worksheet', name: 'Chart1' },
      detail: 'Chart1: the sheet is not a worksheet that the editor can move',
    } as const;
    const remove = { ...move, detail: 'the sheet is not a worksheet that the editor can remove' };
    const text =
      'The sheet Chart1 was not moved or deleted: it is not a worksheet, or its part is missing. ' +
      'The editor can move or delete only a worksheet.';
    expect(refusalText(move, english)).toBe(text);
    expect(refusalText(remove, english)).toBe(text);
    expect(
      refusalText({ code: 'structure-blocked', about: { kind: 'not-worksheet' } }, english)
    ).toBe(
      'The sheet was not moved or deleted: it is not a worksheet, or its part is missing. ' +
        'The editor can move or delete only a worksheet.'
    );
  });

  test('a resize over a drawing that cannot be read says so, and other resize refusals do not', () => {
    const unread = 'The size was not applied: the editor cannot read the drawing of this sheet';
    // The editor adds a generic subject to a resize refusal with none; the cause still wins.
    for (const subject of [undefined, english('editing.anchoredObject')])
      expect(
        refusalText({ code: 'resize-anchor', subject, causeUnreadable: true }, english)
      ).toStartWith(unread);
    expect(refusalText({ code: 'resize-anchor', subject: 'Pic' }, english)).toBe(
      english('editing.refused.resize-anchor', { subject: 'Pic' })
    );
  });

  test('each save code has a whole-sentence save message, and every formula code one opening', () => {
    const save = (en as { editing: { saveRefused: Record<string, string> } }).editing.saveRefused;
    expect(Object.keys(save).sort()).toEqual([...SAVE_REFUSAL_CODES, 'other'].sort());
    for (const text of Object.values(save)) {
      expect(text.startsWith('The workbook was not saved: ')).toBe(true);
      for (const sentence of text.split(/(?<=\.) /))
        expect(sentence.split(' ').length).toBeLessThanOrEqual(20);
    }
    const refused = (en as { editing: { refused: Record<string, string> } }).editing.refused;
    for (const code of [
      'unreadable-formula',
      'formula-syntax',
      'r1c1-formula',
      'unsupported-function',
      'unknown-formula-reference',
    ])
      expect(refused[code]!.startsWith('The formula was not applied: ')).toBe(true);
  });

  test('each cut code has a whole-sentence cut message, and any other cut code a generic one', () => {
    const cut = (en as { editing: { cutRefused: Record<string, string> } }).editing.cutRefused;
    expect(Object.keys(cut).sort()).toEqual([...CUT_REFUSAL_CODES, 'other'].sort());
    for (const text of Object.values(cut))
      expect(text.startsWith('The cells were not cut: ')).toBe(true);
    // A copy code of a cut names the cut, not a copy; a code with no cut text is generic.
    expect(
      refusalText({ code: 'copy-unreadable-formula', command: 'cut', subject: 'S!A2' }, english)
    ).toBe(
      'The cells were not cut: the formula in S!A2 cannot be read. Fix the formula, or cut a range without it.'
    );
    expect(refusalText({ code: 'bad-reference', command: 'cut' }, english)).toBe(cut.other!);
  });

  test('a refused merge names the merge and never gives the next step of a cell edit', () => {
    const refusals = [
      { code: 'shared-formula', subject: 'T!D2:D3' },
      { code: 'shared-formula' },
      { code: 'shared-formula', causeCell: 'T!A3', masterCell: 'T!A2' },
      { code: 'shared-formula', causeCell: 'T!A2', masterCell: 'T!A2', causeUnreadable: true },
      { code: 'cell-alternate-content', subject: 'T!D2' },
      { code: 'cell-alternate-content' },
    ] as const;
    for (const refusal of refusals) {
      const text = refusalText({ ...refusal, command: 'merge' }, english);
      expect(text).toStartWith('The cells were not merged: ');
      expect(text).toEndWith('.');
      expect(text).toContain('Merge a range without');
      expect(text).not.toContain('Edit another cell');
    }
  });

  test('a refused commit leaves the cell editor open, so it says to cancel the cell edit', () => {
    const refusals = [
      { code: 'shared-formula', subject: 'T!D2:D3' },
      { code: 'shared-formula' },
      { code: 'shared-formula', causeCell: 'T!A3', masterCell: 'T!A2' },
      { code: 'shared-formula', causeCell: 'T!A2', masterCell: 'T!A2', causeUnreadable: true },
      { code: 'cell-alternate-content', subject: 'T!D2' },
      { code: 'cell-alternate-content' },
    ] as const;
    for (const refusal of refusals) {
      const text = refusalText({ ...refusal, command: 'commit' }, english);
      expect(text).toStartWith('The edit was not applied: ');
      expect(text).toEndWith(' Cancel the cell edit.');
      expect(text).not.toContain('nother cell');
    }
    // No other code that a commit can return sends the user to another cell.
    for (const code of EDIT_REFUSAL_CODES)
      for (const subject of [{}, { subject: 'T!A1' }])
        expect(refusalText({ code, ...subject, command: 'commit' }, english)).not.toContain(
          'Edit another cell'
        );
  });

  test('no refusal text shows a placeholder, with or without a subject', () => {
    const variants = [
      {},
      { command: 'save' as const },
      { command: 'cut' as const },
      { command: 'format' as const, lock: 'format' as const },
      { command: 'clear' as const, lock: 'cell' as const },
      { command: 'structure' as const, lock: 'structure' as const },
    ];
    for (const code of EDIT_REFUSAL_CODES)
      for (const variant of variants)
        for (const subject of [{}, { subject: 'X' }, { about: { kind: 'table' as const } }]) {
          const text = refusalText({ code, ...variant, ...subject }, english);
          expect([code, text.includes('{') || text.includes('}')]).toEqual([code, false]);
        }
    expect(refusalText({ code: 'unsupported-function' }, english)).toBe(
      'The formula was not applied: the editor cannot write a call to one of its functions. Use a predefined function.'
    );
  });

  test('every refusal text, also as a save, stays under 280 characters with the longest name', () => {
    // The status strip shows a refusal in full, so the catalog bounds its height: a name is
    // cut to 64 characters, and the longest composed text stays under this length.
    const name = 'X'.repeat(200);
    const subjects = [
      { subject: name },
      ...STRUCTURE_SUBJECT_KINDS.map((kind) => ({ about: { kind, name } })),
    ];
    let longest = 0;
    for (const code of EDIT_REFUSAL_CODES)
      for (const save of [false, true])
        for (const subject of subjects) {
          const refusal = { code, ...subject, ...(save ? { command: 'save' as const } : {}) };
          longest = Math.max(longest, refusalText(refusal, english).length);
        }
    // The texts that name a formula range and the cell that blocks it hold two names, for each
    // command, and the toolbar reasons that name both.
    const commands = [
      undefined,
      'save',
      'cut',
      'clear',
      'clearAll',
      'paste',
      'merge',
      'commit',
    ] as const;
    for (const command of commands)
      for (const causeUnreadable of [undefined, true] as const) {
        const refusal = { code: 'shared-formula' as const, causeCell: name, masterCell: name };
        const text = refusalText({ ...refusal, command, causeUnreadable }, english);
        expect(text).not.toContain('X'.repeat(65));
        longest = Math.max(longest, text.length);
      }
    for (const key of ['formulaCell', 'formulaUnreadable'])
      for (const reason of [key, `${key}ClearAll`]) {
        const params = { masterCell: displayName(name), causeCell: displayName(name) };
        const text = disabledReasonText(
          { key: `disabledReason.${reason}`, params } as never,
          english
        );
        expect(text).toContain(params.masterCell);
        longest = Math.max(longest, text.length);
      }
    expect(longest).toBeGreaterThan(200);
    expect(longest).toBeLessThan(280);
  });

  test('each sentence of the catalog has 26 words or fewer', () => {
    // Global audience: one idea in each sentence, and no sentence over 26 words.
    const long: string[] = [];
    const walk = (node: unknown, key: string) => {
      if (typeof node === 'string') {
        for (const sentence of node.split(/(?<=[.!?])\s+/))
          if (sentence.split(/\s+/).length > 26) long.push(`${key}: ${sentence}`);
      } else if (node && typeof node === 'object')
        for (const [k, v] of Object.entries(node)) walk(v, key ? `${key}.${k}` : k);
    };
    walk(en, '');
    expect(long).toEqual([]);
  });

  test('a raw refusal subject loses bidirectional controls and is cut to 64 characters', () => {
    const subject = `\u202EIFNA${'X'.repeat(3001)}`;
    const text = refusalText({ code: 'unsupported-function', subject }, english);
    expect(text).not.toContain('\u202E');
    const name = text.slice(text.indexOf('call to ') + 8, text.indexOf('. Use'));
    expect([...name].length).toBe(64);
    expect(name.startsWith('IFNAX')).toBe(true);
    expect(name.endsWith('…')).toBe(true);
  });

  test('a display name loses the C0, U+007F, and C1 control characters', () => {
    expect(displayName('a\u0000b\u001fc\u007fd\u0080e\u0085f\u009bg\u009fh')).toBe('abcdefgh');
    // The character after the C1 block stays.
    expect(displayName('a b')).toBe('a b');
  });
});
