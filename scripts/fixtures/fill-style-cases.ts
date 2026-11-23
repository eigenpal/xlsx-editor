// The fixtures `fills-patterns` and `cell-styles`: pattern and gradient fills, and cell formats
// that omit a record or set their own under a cell style format (ECMA-376 Part 1, 18.8.45 and
// 18.8.10). Every value is valid against the schema, so each file opens with no repair prompt.

import { c, row, type FixtureWorkbook } from './builder.ts';

/** The non-solid values of `ST_PatternType` (Part 1, 18.18.55), in schema order. */
export const PATTERN_TYPES = [
  'darkGray',
  'mediumGray',
  'lightGray',
  'gray125',
  'gray0625',
  'darkHorizontal',
  'darkVertical',
  'darkDown',
  'darkUp',
  'darkGrid',
  'darkTrellis',
  'lightHorizontal',
  'lightVertical',
  'lightDown',
  'lightUp',
  'lightGrid',
  'lightTrellis',
] as const;

const RED = 'FFFF0000';
const BLUE = 'FF0000FF';
const YELLOW = 'FFFFFF00';
const GREEN = 'FF00B050';

const pattern = (type: string, fg?: string, bg?: string): string =>
  `<fill><patternFill patternType="${type}">${fg ? `<fgColor rgb="${fg}"/>` : ''}${
    bg ? `<bgColor rgb="${bg}"/>` : ''
  }</patternFill></fill>`;

const stops = (list: readonly (readonly [number, string])[]): string =>
  list.map(([p, color]) => `<stop position="${p}"><color rgb="${color}"/></stop>`).join('');

const TWO_STOPS = stops([
  [0, RED],
  [1, BLUE],
]);

const linear = (degree: number, body = TWO_STOPS): string =>
  `<fill><gradientFill degree="${degree}">${body}</gradientFill></fill>`;

const path = (l: number, r: number, t: number, b: number, body = TWO_STOPS): string =>
  `<fill><gradientFill type="path" left="${l}" right="${r}" top="${t}" bottom="${b}">${body}</gradientFill></fill>`;

/** One fill case: the cells it covers, and its fill. */
export interface FillCase {
  readonly label: string;
  readonly refs: readonly string[];
  readonly fill: string;
}

/** The case row of index `i`: rows 2, 4, 6, ... with a narrow gap row between cases. */
const caseRow = (i: number): number => 2 + 2 * i;

export const FILL_CASES: readonly FillCase[] = [
  ...PATTERN_TYPES.map((t, i) => ({
    label: `${t} red on yellow`,
    refs: [`B${caseRow(i)}`],
    fill: pattern(t, RED, YELLOW),
  })),
  ...PATTERN_TYPES.map((t, i) => ({
    label: `${t} blue, no background`,
    refs: [`D${caseRow(i)}`],
    fill: pattern(t, BLUE),
  })),
  ...PATTERN_TYPES.map((t, i) => ({
    label: `${t} no foreground, yellow`,
    refs: [`F${caseRow(i)}`],
    fill: pattern(t, undefined, YELLOW),
  })),
  // The same pattern in adjacent cells and at several offsets: the tile phase.
  ...[0, 1, 2].map((i) => ({
    label: `lightGrid adjacent ${i}`,
    refs: [`H${caseRow(i)}`, `I${caseRow(i)}`],
    fill: pattern('lightGrid', RED, YELLOW),
  })),
  ...[3, 4, 5].map((i) => ({
    label: `darkTrellis adjacent ${i}`,
    refs: [`H${caseRow(i)}`, `I${caseRow(i)}`],
    fill: pattern('darkTrellis', RED, YELLOW),
  })),
  { label: 'darkGrid merged', refs: ['H14'], fill: pattern('darkGrid', RED, YELLOW) },
  { label: 'linear 0 merged', refs: ['H18'], fill: linear(0) },
  { label: 'linear 0', refs: ['K2'], fill: linear(0) },
  { label: 'linear 45', refs: ['K4'], fill: linear(45) },
  { label: 'linear 90', refs: ['K6'], fill: linear(90) },
  { label: 'linear 135', refs: ['K8'], fill: linear(135) },
  { label: 'linear 270', refs: ['K10'], fill: linear(270) },
  { label: 'path center', refs: ['K12'], fill: path(0.5, 0.5, 0.5, 0.5) },
  { label: 'path corner', refs: ['K14'], fill: path(0, 0, 0, 0) },
  { label: 'path off center', refs: ['K16'], fill: path(0.2, 0.4, 0.3, 0.6) },
  {
    label: 'linear 0 three stops',
    refs: ['K18'],
    fill: linear(
      0,
      stops([
        [0, RED],
        [0.3, GREEN],
        [1, BLUE],
      ])
    ),
  },
  {
    label: 'path center three stops',
    refs: ['K20'],
    fill: path(
      0.5,
      0.5,
      0.5,
      0.5,
      stops([
        [0, RED],
        [0.3, GREEN],
        [1, BLUE],
      ])
    ),
  },
];

// The reference export draws each row of a merged range with a gradient as its own image,
// flipped vertically. A linear gradient at 0 degrees looks the same either way, so it is the only
// merged gradient here: any other would measure the export, not the fill.
const FILL_MERGES = ['H14:I16', 'H18:I20'];

const colKey = (ref: string): number => ref.charCodeAt(0);
const rowOf = (ref: string): number => Number(ref.slice(1));

/** The workbook of `fills-patterns`: one cell format for each case, with no gridlines. */
export function fillsPatternsWorkbook(): FixtureWorkbook {
  const fills: string[] = [];
  const cellXfs: string[] = [];
  const cells = new Map<number, string[]>();
  FILL_CASES.forEach((k, i) => {
    fills.push(k.fill);
    cellXfs.push(
      `<xf numFmtId="0" fontId="0" fillId="${i + 2}" borderId="0" xfId="0" applyFill="1"/>`
    );
    for (const ref of k.refs) {
      const list = cells.get(rowOf(ref)) ?? [];
      list.push(c(ref, undefined, { s: i + 1 }));
      cells.set(rowOf(ref), list);
    }
  });
  const last = caseRow(PATTERN_TYPES.length - 1);
  const rows: string[] = [];
  for (let r = 1; r <= last; r++) {
    const list = (cells.get(r) ?? []).sort((a, b) => colKey(a.slice(6)) - colKey(b.slice(6)));
    const height = r % 2 === 0 ? 24 : 6;
    rows.push(row(r, list, `ht="${height}" customHeight="1"`));
  }
  const width = (min: number, max: number, w: number) =>
    `<col min="${min}" max="${max}" width="${w}" customWidth="1"/>`;
  return {
    styles: { fills, cellXfs },
    sheets: [
      {
        name: 'Fills',
        cols: `<cols>${width(1, 1, 2)}${width(2, 2, 11)}${width(3, 3, 2)}${width(4, 4, 11)}${width(5, 5, 2)}${width(6, 6, 11)}${width(7, 7, 2)}${width(8, 9, 11)}${width(10, 10, 2)}${width(11, 11, 11)}</cols>`,
        rows: rows.join(''),
        merges: FILL_MERGES,
      },
    ],
  };
}

/** A record of a cell format that a cell style format can also give. */
type StyleRecord = 'font' | 'fill' | 'border' | 'numFmt' | 'alignment';

/** One cell style case: the cell format and what it tests. */
export interface StyleResolutionCase {
  readonly label: string;
  readonly ref: string;
  readonly xf: string;
}

/** The style format 1 records, set explicitly on each case for the records it does not test. */
const STYLE_IDS = { font: ['fontId', '1'], fill: ['fillId', '2'], border: ['borderId', '1'], numFmt: ['numFmtId', '10'] } as const; // prettier-ignore
const APPLY: Record<StyleRecord, string> = {
  font: 'applyFont',
  fill: 'applyFill',
  border: 'applyBorder',
  numFmt: 'applyNumberFormat',
  alignment: 'applyAlignment',
};
const OWN_ID: Record<Exclude<StyleRecord, 'alignment'>, string> = {
  font: '0',
  fill: '0',
  border: '0',
  numFmt: '0',
};

type State = 'omit' | 'own';

function styleXf(
  record: StyleRecord,
  state: State,
  apply: '' | '0' | '1',
  xfId: string | undefined = '1',
  ownAlignment = '<alignment horizontal="left"/>'
): string {
  const attrs: string[] = [];
  for (const r of ['numFmt', 'font', 'fill', 'border'] as const) {
    const [name, styleValue] = STYLE_IDS[r];
    if (r !== record) attrs.push(`${name}="${styleValue}"`, `${APPLY[r]}="1"`);
    else if (state === 'own') attrs.push(`${name}="${OWN_ID[r]}"`);
  }
  if (xfId !== undefined) attrs.push(`xfId="${xfId}"`);
  if (apply) attrs.push(`${APPLY[record]}="${apply}"`);
  if (record !== 'alignment') attrs.push('applyAlignment="1"');
  const alignment =
    record !== 'alignment'
      ? '<alignment horizontal="center"/>'
      : state === 'own'
        ? ownAlignment
        : '';
  return `<xf ${attrs.join(' ')}>${alignment}</xf>`;
}

const RECORDS: readonly StyleRecord[] = ['font', 'fill', 'border', 'numFmt', 'alignment'];

export const STYLE_RESOLUTION_CASES: readonly StyleResolutionCase[] = (() => {
  const list: Omit<StyleResolutionCase, 'ref'>[] = [];
  for (const record of RECORDS) {
    for (const state of ['omit', 'own'] as const)
      for (const apply of ['', '0', '1'] as const)
        list.push({
          label: `${record} ${state} apply ${apply || 'absent'}`,
          xf: styleXf(record, state, apply),
        });
    list.push({
      label: `${record} omit under style apply 0`,
      xf: styleXf(record, 'omit', '', '2'),
    });
  }
  list.push({
    label: 'alignment own wrap only',
    xf: styleXf('alignment', 'own', '', '1', '<alignment wrapText="1"/>'),
  });
  list.push({
    label: 'fill omit, no xfId',
    xf: '<xf numFmtId="0" fontId="0" borderId="0"/>',
  });
  return list.map((k, i) => ({ ...k, ref: `B${i + 1}` }));
})();

/**
 * The workbook of `cell-styles`: the case cell, 0.5, in column B. The labels stay in
 * `STYLE_RESOLUTION_CASES`, so that the reference PDF holds few glyphs.
 */
export function cellStylesWorkbook(): FixtureWorkbook {
  const allOn =
    'applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"';
  const allOff =
    'applyNumberFormat="0" applyFont="0" applyFill="0" applyBorder="0" applyAlignment="0"';
  const style = (apply: string) =>
    `<xf numFmtId="10" fontId="1" fillId="2" borderId="1" ${apply}><alignment horizontal="center"/></xf>`;
  return {
    styles: {
      fonts: [
        '<font><b/><sz val="11"/><color rgb="FFC00000"/><name val="Carlito"/><family val="2"/><scheme val="minor"/></font>',
      ],
      fills: [
        `<fill><patternFill patternType="solid"><fgColor rgb="${GREEN}"/><bgColor indexed="64"/></patternFill></fill>`,
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFFC000"/><bgColor indexed="64"/></patternFill></fill>',
      ],
      borders: [
        '<border><left/><right/><top/><bottom style="thin"><color rgb="FF0070C0"/></bottom><diagonal/></border>',
      ],
      normalStyleXf: '<xf numFmtId="0" fontId="0" fillId="3" borderId="0"/>',
      cellStyleXfs: [style(allOn), style(allOff)],
      cellStyles: [
        '<cellStyle name="Fixture Style" xfId="1"/>',
        '<cellStyle name="Fixture Style Off" xfId="2"/>',
      ],
      cellXfs: STYLE_RESOLUTION_CASES.map((k) => k.xf),
    },
    sheets: [
      {
        name: 'Styles',
        cols: '<cols><col min="1" max="1" width="30" customWidth="1"/><col min="2" max="2" width="16" customWidth="1"/></cols>',
        rows: STYLE_RESOLUTION_CASES.map((k, i) => row(i + 1, [c(k.ref, 0.5, { s: i + 1 })])).join(
          ''
        ),
      },
    ],
  };
}
