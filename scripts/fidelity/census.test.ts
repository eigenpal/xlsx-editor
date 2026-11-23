import { describe, expect, test } from 'bun:test';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { buildXlsx, c, row } from '../fixtures/builder.ts';
import { refusalCode } from '../corpus-runner/open-refusal.ts';
import { saveGarbage, saveWithAttribute } from '../corpus-runner/test-saves.ts';
import { ROOT } from '../corpus-runner/runner.ts';
import { censusWorkbook, contentTypeName, planCensus, renderCensus } from './census.ts';

/** Words that only the workbook holds. The census must not repeat any of them. */
const PRIVATE = [
  'Quarterly Secret',
  'Ada Example',
  '"Account "0.00',
  'Private Font Name',
  'priv',
  'PayrollSecret',
  'FalconTarget',
];

const privateWorkbook = (): Uint8Array =>
  buildXlsx({
    sheets: [
      {
        name: 'Quarterly Secret',
        rows: [
          row(1, [c('A1', 'Ada Example'), c('B1', 42, { s: 1 }), c('C1', 3, { f: 'B1/14' })]),
          row(2, [c('A2', true), c('B2', '#N/A', { t: 'e' })]),
        ].join(''),
        merges: ['A3:B3'],
        // An element of an unknown namespace, one that the author named in the main one, and
        // markup compatibility with a name of the author in its namespace.
        afterMerges: `<priv:note xmlns:priv="urn:example:priv"/><FalconTarget/><ref/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:ext="urn:example:ext" Requires="ext"/><mc:Fallback/><mc:FalconChoice/></mc:AlternateContent>`,
      },
      { name: 'Hidden Notes', state: 'hidden', rows: row(1, [c('A1', 'Ada Example')]) },
    ],
    styles: {
      numFmts: ['<numFmt numFmtId="164" formatCode="&quot;Account &quot;0.00"/>'],
      fonts: ['<font><sz val="11"/><name val="Private Font Name"/></font>'],
      cellXfs: ['<xf numFmtId="164" fontId="1" applyNumberFormat="1"/>'],
    },
    // A content type that the author chose: it holds private text in a well-formed value.
    extraParts: [['xl/custom/data.bin', 'application/x-PayrollSecret', new Uint8Array([1, 2])]],
  });

describe('the census of one workbook', () => {
  test('counts sheets, cells, formulas, cell types, and features', () => {
    const census = censusWorkbook(privateWorkbook());
    expect(census).toMatchObject({
      open: 'ok',
      sheets: 2,
      worksheets: 2,
      cells: 6,
      formulas: 1,
      edit: 'ok',
      // The element of an unknown namespace can refer to cells, so a row insert is refused.
      structure: 'structure-blocked',
      save: 'same',
      editedSave: 'same',
    });
    expect(census.cellTypes).toEqual({ inlineStr: 2, n: 2, b: 1, e: 1 });
    expect(census.worksheetElements?.['x:mergeCells']).toBe(1);
    // An element of an unknown namespace counts under `other`, never under its prefix or name.
    expect(census.worksheetElements?.['other:(name)']).toBe(1);
    // A name in the main namespace that the schema does not declare as an element is the
    // author's choice, also when the schema has it as an attribute (`ref`).
    expect(census.worksheetElements?.['x:(name)']).toBe(2);
    // Every element counts, at any depth, and markup compatibility keeps only its own names.
    expect(census.worksheetElements).toMatchObject({
      'x:row': 3,
      'x:c': 6,
      'x:f': 1,
      'mc:AlternateContent': 1,
      'mc:Choice': 1,
      'mc:Fallback': 1,
      'mc:(name)': 1,
    });
    expect(census.workbookElements?.['x:sheets']).toBe(1);
  });

  test('holds no sheet name, text, number format, font name, prefix, or content type of the workbook', () => {
    const census = censusWorkbook(privateWorkbook());
    const text = JSON.stringify(census);
    for (const word of PRIVATE) expect(text).not.toContain(word);
    expect(census.contentTypes?.['(other)']).toBe(1);
    expect(
      census.contentTypes?.[
        'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
      ]
    ).toBe(2);
  });

  test('names only the content types that the format defines, and no other text', () => {
    for (const type of [
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings',
      'application/vnd.openxmlformats-officedocument.drawing+xml',
      'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
      'application/vnd.openxmlformats-officedocument.theme+xml',
      'application/vnd.openxmlformats-officedocument.vmlDrawing',
      'application/vnd.openxmlformats-package.core-properties+xml',
      'application/octet-stream',
      'image/png',
    ])
      expect(contentTypeName(type)).toBe(type);
    // Content types compare without case, and the census writes the form of its list.
    expect(contentTypeName(' APPLICATION/XML ')).toBe('application/xml');
    expect(contentTypeName(undefined)).toBe('(none)');
    for (const type of [
      'application/x-PayrollSecret',
      'application/vnd.openxmlformats-officedocument.AcmeMergerPayroll+xml',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.Secret+xml',
      'Secret/application/xml',
      'application/xml/Secret',
      'text/Secret',
    ])
      expect(contentTypeName(type)).toBe('(other)');
  });

  test('records refused edit probes on a protected sheet, and no edited save', () => {
    const locked = buildXlsx({
      sheets: [
        {
          name: 'Locked',
          rows: row(1, [c('A1', 'text')]),
          beforeMerges: '<sheetProtection sheet="1"/>',
        },
      ],
    });
    expect(censusWorkbook(locked)).toMatchObject({
      edit: 'locked',
      structure: 'locked',
      save: 'same',
      editedSave: 'none',
    });
  });

  test('saves the edited package for the edited check', () => {
    // The writer breaks only a package that holds the text of the edit probe.
    const save = saveWithAttribute('sheetFormatPr', 'defaultRowHeight', '20', 'Census & ');
    expect(censusWorkbook(privateWorkbook(), save)).toMatchObject({
      save: 'same',
      editedSave: 'changed',
    });
  });

  test('records a saved copy that changes or does not reopen', () => {
    const changed = censusWorkbook(
      privateWorkbook(),
      saveWithAttribute('sheetFormatPr', 'defaultRowHeight', '20')
    );
    expect(changed).toMatchObject({ save: 'changed', editedSave: 'changed' });
    expect(censusWorkbook(privateWorkbook(), saveGarbage)).toMatchObject({
      save: 'unreadable',
      editedSave: 'unreadable',
    });
  });

  test('records a refused open as a code', () => {
    expect(censusWorkbook(new TextEncoder().encode('not a zip'))).toEqual({
      open: 'refused',
      refusal: 'zip:inflate-error',
    });
  });

  test('keeps no file-derived text in a refusal code', () => {
    expect(refusalCode({ reason: 'not-a-workbook', detail: 'Secret/type' })).toBe('not-a-workbook');
    expect(refusalCode({ reason: 'zip', detail: 'Secret Name' })).toBe('zip');
    expect(refusalCode({ reason: 'xml', partName: '/xl/Secret.xml', detail: 'too-deep' })).toBe(
      'xml:too-deep'
    );
  });

  test('censuses committed fixtures of three kinds with no open refusal and no loss on save', () => {
    for (const name of ['basic-values', 'preserved-parts', 'structure']) {
      const census = censusWorkbook(new Uint8Array(readFileSync(`fixtures/${name}.xlsx`)));
      expect(census.open).toBe('ok');
      expect(census.save).toBe('same');
    }
  });
});

describe('the census report', () => {
  const lines = [
    { id: 'workbook1', ...censusWorkbook(privateWorkbook()) },
    { id: 'workbook2', open: 'refused' as const, refusal: 'zip:inflate-error' },
    { id: 'workbook3', stopped: 'memory' } as never,
  ];

  test('gives totals only for a private corpus', () => {
    const report = renderCensus(lines, 'Corpus', false);
    expect(report).toContain('| `opened` | 1 |');
    expect(report).toContain('| `refused at open` | 1 |');
    expect(report).toContain('| `stopped (memory)` | 1 |');
    expect(report).not.toContain('Per workbook');
    expect(report).not.toContain('workbook1');
    for (const word of PRIVATE) expect(report).not.toContain(word);
  });

  test('adds a row for each workbook of the fixture set', () => {
    const report = renderCensus(lines, 'Fixture set', true);
    expect(report).toContain('## Per workbook');
    expect(report).toContain('| workbook2 | refused (zip:inflate-error) |');
  });
});

describe('the plan of a census', () => {
  const scratch = (): string => mkdtempSync(join(tmpdir(), 'xlsx-census-test-'));

  test('names the fixture set, and only the fixture set', () => {
    const out = scratch();
    const dir = scratch();
    try {
      writeFileSync(join(dir, 'a.xlsx'), '');
      expect(planCensus(['--dir', join(ROOT, 'fixtures'), '--out', out])).toMatchObject({
        fixtureSet: true,
        fresh: false,
      });
      expect(planCensus(['--dir', dir, '--out', out, '--fresh'])).toMatchObject({
        fixtureSet: false,
        fresh: true,
      });
    } finally {
      rmSync(out, { recursive: true, force: true });
      rmSync(dir, { recursive: true, force: true });
    }
  });

  test('refuses wrong flags with the usage text, and an empty folder without it', () => {
    const dir = scratch();
    try {
      const file = join(dir, 'file');
      writeFileSync(file, '');
      expect(planCensus([], {})).toMatchObject({ usage: true });
      expect(planCensus(['--dir', dir, '--out', join(file, 'out')])).toEqual({
        refused: '--out must be a folder that can be made and written.',
        usage: true,
      });
      expect(planCensus(['--dir', join(dir, 'empty'), '--out', dir])).toMatchObject({
        usage: true,
      });
      expect(planCensus(['--dir', dir, '--out', join(dir, 'out')])).toEqual({
        refused: 'The corpus folder holds no .xlsx file. A census of nothing is a gap.',
        usage: false,
      });
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});
