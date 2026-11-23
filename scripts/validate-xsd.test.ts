import { afterEach, describe, expect, test } from 'bun:test';
import { copyFileSync, mkdtempSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { saveXlsx } from '../packages/core/src/store/index.ts';
import { CORPUS_TMP_ENV } from './corpus-runner/runner.ts';
import { saveGarbage, saveWithAttribute } from './corpus-runner/test-saves.ts';
import { buildXlsx, c, row } from './fixtures/builder.ts';
import {
  corpusProblem,
  planCorpus,
  runXmllint,
  schemaMessage,
  validateCorpus,
  validateCorpusWorkbook,
  xmllintGroups,
  type CorpusPlan,
} from './validate-xsd.ts';

const X = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';

const dirs: string[] = [];
const scratch = (): string => {
  const dir = mkdtempSync(join(tmpdir(), 'xlsx-xsd-test-'));
  dirs.push(dir);
  return dir;
};
afterEach(() => {
  for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true });
});

/** Two worksheets. `rowHeight` is the `defaultRowHeight` of the first; `protect` locks sheets. */
const workbook = (rowHeight: string, protect: readonly [boolean, boolean] = [false, false]) =>
  buildXlsx({
    sheets: (['First', 'Second'] as const).map((name, i) => ({
      name,
      rows: row(1, [c('A1', 'text')]),
      formatPr: i === 0 ? `defaultRowHeight="${rowHeight}"` : undefined,
      beforeMerges: protect[i] ? '<sheetProtection sheet="1"/>' : undefined,
    })),
  });

describe('a corpus schema message', () => {
  test('keeps the element, the attribute, and the type, and drops the value', () => {
    expect(
      schemaMessage(
        `element row: Schemas validity error : Element '${X}row', attribute 'r': 'O'Brien Secret' is not a valid value of the atomic type 'xs:unsignedInt'.`
      )
    ).toBe(
      `Element '${X}row', attribute 'r': '<value>' is not a valid value of the atomic type 'xs:unsignedInt'.`
    );
  });

  test('keeps a type name with its namespace', () => {
    expect(
      schemaMessage(
        `Element '${X}color', attribute 'rgb': 'Secret' is not a valid value of the atomic type '${X}ST_UnsignedIntHex'.`
      )
    ).toBe(
      `Element '${X}color', attribute 'rgb': '<value>' is not a valid value of the atomic type '${X}ST_UnsignedIntHex'.`
    );
  });

  test('keeps the facet and drops the value and the allowed set', () => {
    expect(
      schemaMessage(
        `element c: Schemas validity error : Element '${X}c', attribute 't': [facet 'enumeration'] The value 'Pri'vate' is not an element of the set {'b', 'n'}.`
      )
    ).toBe(
      `Element '${X}c', attribute 't': [facet 'enumeration'] The value '<value>' is not an element of the set {<values>}.`
    );
  });

  test('keeps no part of a value that holds a quote and a brace', () => {
    expect(
      schemaMessage(
        `Element '${X}c', attribute 't': [facet 'enumeration'] The value 'x' Salary of Bob {' z' is not an element of the set {'b', 'n'}.`
      )
    ).toBe(
      `Element '${X}c', attribute 't': [facet 'enumeration'] The value '<value>' is not an element of the set {<values>}.`
    );
  });

  test('drops the rest of a value that held a line break', () => {
    expect(
      schemaMessage(`Element '${X}c', attribute 'r': [facet 'pattern'] The value 'Secret`)
    ).toBe(`Element '${X}c', attribute 'r': [facet 'pattern'] The value '<value>'`);
  });

  test('keeps a message with no value, and the names that the schemas declare', () => {
    const order = `Element '${X}row': This element is not expected. Expected is ( ${X}sheetData ).`;
    expect(schemaMessage(`element row: Schemas validity error : ${order}`)).toBe(order);
    for (const message of [
      // `ht` is only an attribute in the schemas.
      `Element '${X}row': The attribute 'ht' is not allowed.`,
      `Element '${X}c': The attribute 'r' is not allowed.`,
      `Element '${X}c': The attribute 'r' is required but missing.`,
    ])
      expect(schemaMessage(message)).toBe(message);
  });

  test('writes (name) for a name that the author chose', () => {
    expect(
      schemaMessage(
        `element SalaryOfAda: Schemas validity error : Element '${X}SalaryOfAda': This element is not expected. Expected is ( ${X}sheetData ).`
      )
    ).toBe(`Element '${X}(name)': This element is not expected. Expected is ( ${X}sheetData ).`);
    expect(schemaMessage(`Element '${X}c': The attribute 'AdaNote' is not allowed.`)).toBe(
      `Element '${X}c': The attribute '(name)' is not allowed.`
    );
    // A prefix is the author's too.
    expect(schemaMessage(`Element '${X}c': The attribute 'ada:r' is not allowed.`)).toBe(
      `Element '${X}c': The attribute '(name)' is not allowed.`
    );
    expect(
      schemaMessage(
        `Element '${X}c', attribute 'AdaNote': 'Secret' is not a valid value of the atomic type 'xs:string'.`
      )
    ).toBe(
      `Element '${X}c', attribute '(name)': '<value>' is not a valid value of the atomic type 'xs:string'.`
    );
  });

  test('names a namespace that the engine does not know as other', () => {
    expect(
      schemaMessage(
        "Element '{urn:secret-client}note': This element is not expected. Expected is ( a )."
      )
    ).toBe("Element '{other}(name)': This element is not expected. Expected is ( a ).");
  });

  test('drops a message of another form that holds a quote', () => {
    expect(schemaMessage("Something 'Secret' else")).toBe('(other schema message)');
  });

  test('takes the element only at the start of the message', () => {
    expect(schemaMessage(`Secret text Element '${X}c': LEAKED 'value'`)).toBe(
      '(other schema message)'
    );
  });
});

describe('the xmllint calls', () => {
  test('take one schema and at most 100 files each', () => {
    const files = [
      ...Array.from({ length: 250 }, (_, i) => ({ schema: 'sml.xsd', i })),
      ...Array.from({ length: 3 }, (_, i) => ({ schema: 'dml.xsd', i })),
    ];
    expect(xmllintGroups(files).map((g) => [g.schema, g.group.length])).toEqual([
      ['sml.xsd', 100],
      ['sml.xsd', 100],
      ['sml.xsd', 50],
      ['dml.xsd', 3],
    ]);
  });

  test('fail, never pass, when xmllint is missing or writes too much', () => {
    expect(() => runXmllint(['--version'], { command: 'xmllint-not-installed' })).toThrow(
      'xmllint did not run: install libxml2.'
    );
    expect(() => runXmllint(['--version'], { maxBuffer: 16 })).toThrow(
      'xmllint wrote more than 16 bytes of messages.'
    );
    expect(runXmllint(['--version'])).toContain('libxml');
  });
});

describe('the corpus task of validate:xsd', () => {
  test('finds no new message in the saved and edited copies of a valid workbook', () => {
    expect(validateCorpusWorkbook(workbook('15'))).toMatchObject({
      open: 'ok',
      original: 0,
      saved: 0,
      edited: 0,
      editSheet: 1,
      editRefused: 0,
      structure: 'ok',
      newMessages: [],
    });
  });

  test('counts a new message of a kind that the original already has, without its value', () => {
    const original = workbook('Secret one');
    expect(validateCorpusWorkbook(original)).toMatchObject({ original: 1, saved: 0, edited: 0 });
    const result = validateCorpusWorkbook(
      original,
      saveWithAttribute('sheetFormatPr', 'defaultRowHeight', 'Secret two')
    );
    expect(result).toMatchObject({ original: 1, saved: 1, edited: 1 });
    expect(result.newMessages).toEqual([
      `Element '${X}sheetFormatPr', attribute 'defaultRowHeight': '<value>' is not a valid value of the atomic type 'xs:double'.`,
    ]);
    expect(corpusProblem(result)).toBe('saved: 1 new, edited: 1 new');
  });

  test('validates the edited copy, not the original again', () => {
    // The writer breaks only a package that holds the text of the edit session.
    const save = saveWithAttribute('sheetFormatPr', 'defaultRowHeight', 'Secret', 'Edited & ');
    expect(validateCorpusWorkbook(workbook('15'), save)).toMatchObject({
      original: 0,
      saved: 0,
      edited: 1,
    });
  });

  test('writes its work files into the folder of the run and removes them', () => {
    const root = scratch();
    const before = process.env[CORPUS_TMP_ENV];
    process.env[CORPUS_TMP_ENV] = root;
    let seen: string[] = [];
    try {
      validateCorpusWorkbook(workbook('15'), (pkg) => {
        seen = readdirSync(root);
        return saveXlsx(pkg);
      });
    } finally {
      if (before === undefined) delete process.env[CORPUS_TMP_ENV];
      else process.env[CORPUS_TMP_ENV] = before;
    }
    expect(seen).toHaveLength(1);
    expect(seen[0]).toStartWith('xlsx-xsd-corpus-');
    expect(readdirSync(root)).toEqual([]);
  });

  test('records copies that do not reopen', () => {
    const result = validateCorpusWorkbook(workbook('15'), saveGarbage);
    expect(result).toMatchObject({ saved: 'unreadable', edited: 'unreadable' });
    expect(corpusProblem(result)).toBe('a saved copy does not reopen');
  });

  test('edits the first worksheet that takes the edits, and fails when none does', () => {
    expect(validateCorpusWorkbook(workbook('15', [true, false]))).toMatchObject({
      editSheet: 2,
      edited: 0,
      structure: 'ok',
    });
    const locked = validateCorpusWorkbook(workbook('15', [true, true]));
    expect(locked).toMatchObject({ edited: 'skipped', structure: 'none' });
    expect(corpusProblem(locked)).toBe('the edit check was skipped: no worksheet takes the edits');
  });

  test('records a structural session that the engine refuses', () => {
    // An element of an unknown namespace can refer to cells, so the engine refuses a row insert.
    const blocked = buildXlsx({
      sheets: [
        {
          name: 'First',
          rows: row(1, [c('A1', 'text')]),
          afterMerges: '<ext:note xmlns:ext="urn:example:ext"/>',
        },
      ],
    });
    const result = validateCorpusWorkbook(blocked);
    expect(result).toMatchObject({ open: 'ok', edited: 0, structure: 'structure-blocked' });
    expect(corpusProblem(result)).toBeUndefined();
  });

  test('records a refused open as a code', () => {
    expect(validateCorpusWorkbook(new TextEncoder().encode('not a zip'))).toEqual({
      open: 'refused',
      refusal: 'zip:inflate-error',
    });
  });
});

describe('the verdict of validate:xsd:corpus', () => {
  const pass = {
    open: 'ok',
    parts: 5,
    original: 2,
    saved: 0,
    edited: 0,
    structure: 'ok',
  } as const;

  test('passes a workbook with no new message, also when its structure is blocked', () => {
    expect(corpusProblem(pass)).toBeUndefined();
    expect(corpusProblem({ ...pass, structure: 'structure-blocked' })).toBeUndefined();
  });

  test('fails every result that is not a pass', () => {
    expect(corpusProblem({ stopped: 'memory' })).toBe('stopped (memory)');
    expect(corpusProblem({ error: 'boom' })).toBe('task error: boom');
    expect(corpusProblem({ open: 'refused', refusal: 'zip' })).toBe('refused at open (zip)');
    expect(corpusProblem({})).toBe('the line is not a result of the task');
    expect(corpusProblem({ ...pass, parts: 0 })).toBe('no part was validated');
    expect(corpusProblem({ ...pass, structure: 'locked' })).toBe(
      'the structural session is refused (locked)'
    );
    expect(corpusProblem({ ...pass, edited: 2 })).toBe('saved: 0 new, edited: 2 new');
    expect(corpusProblem({ ...pass, saved: 1 })).toBe('saved: 1 new, edited: 0 new');
  });
});

describe('validate:xsd:corpus', () => {
  const plan = (dir: string, extra: string[] = []): CorpusPlan => {
    const planned = planCorpus(['--dir', dir, '--out', join(scratch(), 'out'), ...extra]);
    if ('refused' in planned) throw new Error(planned.refused);
    return planned;
  };

  test('refuses wrong flags, a folder with no workbook, and earlier results', () => {
    expect(planCorpus([], {})).toMatchObject({ usage: true });
    const file = join(scratch(), 'file');
    writeFileSync(file, '');
    expect(planCorpus(['--dir', scratch(), '--out', join(file, 'out')])).toEqual({
      refused: '--out must be a folder that can be made and written.',
      usage: true,
    });
    expect(planCorpus(['--dir', scratch(), '--out', scratch()])).toEqual({
      refused: 'The corpus folder holds no .xlsx file. A check of nothing is a gap.',
      usage: false,
    });
    const dir = scratch();
    copyFileSync('fixtures/basic-values.xlsx', join(dir, 'a.xlsx'));
    const out = scratch();
    writeFileSync(join(out, 'results.jsonl'), '{}\n');
    expect(planCorpus(['--dir', dir, '--out', out])).toEqual({
      refused: `The output folder ${out} holds the results of an earlier run. Add --fresh to replace them.`,
      usage: false,
    });
  });

  test('passes a folder of valid workbooks and fails a folder with a refused one', async () => {
    const dir = scratch();
    copyFileSync('fixtures/basic-values.xlsx', join(dir, 'a.xlsx'));
    expect(await validateCorpus(plan(dir))).toBe(true);
    writeFileSync(join(dir, 'b.xlsx'), 'not a zip');
    expect(await validateCorpus(plan(dir))).toBe(false);
  }, 60_000);
});
