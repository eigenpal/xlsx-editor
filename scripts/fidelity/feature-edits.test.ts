// Each recipe edits a fixture of its support matrix row through the public editor API, saves,
// reopens, and finds the edit again. Its check must fail on the unedited copy, a second save must
// give the same package, and the edit must change only the parts it owns. A copy that the save
// check opened with no repair prompt keeps its hash: a new hash needs a new save check
// (CONTRIBUTING.md, "Save check").
//
// Same form as pptx-editor/scripts/fidelity/feature-edits.test.ts at 30dc48ea (EigenPal).

import { describe, expect, test } from 'bun:test';
import { supportFeatures } from '../../docs/site/data/support-matrix.ts';
import {
  featureRecordProblems,
  NO_RECORD,
  readFeatureRecords,
} from '../save-check/feature-records.ts';
import { added, point, removed, runFeatureEdit, type FeatureEdit } from './feature-edit-kit.ts';
import { FEATURE_EDITS } from './feature-edits.ts';

const records = readFeatureRecords();

describe('feature edits survive save and reopen', () => {
  for (const [row, recipe] of Object.entries(FEATURE_EDITS)) {
    test(row, async () => {
      const run = await runFeatureEdit(recipe);
      expect(run.problems).toEqual([]);
      // A recorded copy keeps its hash and its result. `check:support-matrix` decides which rows
      // need a record; a stale one names its remedy there.
      if (Object.hasOwn(records, row)) expect(featureRecordProblems(row, run, records)).toEqual([]);
    });
  }
});

describe('the feature edit records', () => {
  test('every recipe and every record names a row of the support matrix', () => {
    const rows = new Set(supportFeatures.map((r) => r.id));
    const named = [...Object.keys(FEATURE_EDITS), ...Object.keys(records)];
    expect(named.filter((id) => !rows.has(id))).toEqual([]);
  });

  test('every record names a recipe', () => {
    expect(Object.keys(records).filter((id) => !Object.hasOwn(FEATURE_EDITS, id))).toEqual([]);
  });

  test('a missing, stale, or failed record is reported', () => {
    const run = { bytes: new Uint8Array([1]), sheetCount: 1, problems: [] };
    expect(featureRecordProblems('row', run, {})).toEqual([NO_RECORD]);
    expect(
      featureRecordProblems('row', run, { row: { sha256: '0', result: 'repair-prompt' } })
    ).toEqual([
      'the edited copy changed since the check opened it (stale)',
      'the edited copy: repair-prompt',
    ]);
    expect(featureRecordProblems('row', { problems: ['no edit'] }, {})).toEqual([
      'the recipe fails: no edit',
    ]);
  });

  test('a recipe whose check passes on the unedited copy fails', async () => {
    const idle: FeatureEdit = {
      fixture: 'basic-values',
      edit: () => {},
      check: () => undefined,
    };
    const run = await runFeatureEdit(idle);
    expect(run.problems.join('\n')).toContain('unedited');
  });

  test('an added or removed rule does not let an existing part change', async () => {
    const recipe = FEATURE_EDITS['editing.font-style']!;
    for (const rule of [added(/^\/xl\/styles\.xml$/), removed(/^\/xl\/styles\.xml$/)]) {
      const run = await runFeatureEdit({ ...recipe, parts: [rule] });
      expect(run.problems).toEqual(['the edit changes /xl/styles.xml, which it does not own']);
    }
  });

  test('a pattern rule lets only the parts that it matches change', async () => {
    const recipe = FEATURE_EDITS['editing.font-style']!;
    const run = await runFeatureEdit({ ...recipe, parts: [/^\/xl\/workbook\.xml$/] });
    expect(run.problems).toEqual(['the edit changes /xl/styles.xml, which it does not own']);
  });

  test('an added rule lets a new part in', async () => {
    // `colors-fonts` stores inline strings, so a typed text adds the shared string table.
    const recipe: FeatureEdit = {
      fixture: 'colors-fonts',
      edit: (t) => t.run(t.editor.enterValue('typed', point('C1'))),
      check: (d) => (d.cellText('Colors', 'C1') === 'typed' ? undefined : 'no text'),
    };
    const rules = [/^\/xl\/_rels\/workbook\.xml\.rels$/, /^\/\[Content_Types\]\.xml$/];
    expect(
      (await runFeatureEdit({ ...recipe, parts: [added(/^\/xl\/sharedStrings\.xml$/), ...rules] }))
        .problems
    ).toEqual([]);
    expect((await runFeatureEdit({ ...recipe, parts: rules })).problems).toEqual([
      'the edit adds /xl/sharedStrings.xml, which it does not own',
    ]);
  });

  test('an edit that changes a part it does not own fails', async () => {
    const recipe = FEATURE_EDITS['editing.font-style']!;
    const run = await runFeatureEdit({ ...recipe, parts: [] });
    expect(run.problems.join('\n')).toContain('/xl/styles.xml');
  });
});
