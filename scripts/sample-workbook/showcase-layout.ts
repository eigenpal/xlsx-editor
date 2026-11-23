// The last sheets of the sample workbook: text and fonts, grid layout with a picture and
// shapes, and a schedule of dates. All names and figures are made up for the sample.

import type { FixtureSheet } from '../fixtures/builder.ts';
import { PICTURE_SIZE, landscapePng } from './showcase-image.ts';
import {
  COLORS,
  THIN,
  cellWriter,
  columns,
  dateSerial,
  frozenPane,
  sheetRow,
  type CellStyle,
  type StringTable,
  type StyleTable,
} from './showcase-kit.ts';

const groupStyle: CellStyle = {
  font: { bold: true, size: 12, color: COLORS.green },
  fill: COLORS.white,
  border: { bottom: 'thin' },
};

/** Text: font families, sizes, styles, rich text, alignment, wrap, rotation, and fills. */
export function buildTextSheet(styles: StyleTable, strings: StringTable): FixtureSheet {
  const c = cellWriter(strings);
  const group = styles.xf(groupStyle);
  const label = styles.xf({ font: { color: COLORS.muted } });
  const rows: string[] = [];
  // Text that overflows into its neighbor cells sits in a merged range with a white fill
  // instead, so no gridline crosses it.
  const merges: string[] = [];
  let r = 1;
  const section = (title: string) => {
    rows.push(
      sheetRow(r, [c(`A${r}`, title, group), c(`B${r}`, undefined, group)], { height: 22 })
    );
    merges.push(`A${r}:B${r}`);
    r += 1;
  };

  section('Font families');
  for (const name of ['Calibri', 'Cambria', 'Arial', 'Times New Roman', 'Courier New']) {
    rows.push(
      sheetRow(r, [
        c(`A${r}`, name, label),
        c(
          `B${r}`,
          'Sphinx of black quartz, judge my vow. 0123456789',
          styles.xf({ font: { name }, fill: COLORS.white })
        ),
      ])
    );
    merges.push(`B${r}:D${r}`);
    r += 1;
  }

  r += 1;
  section('Sizes and styles');
  for (const size of [8, 11, 14, 18]) {
    rows.push(
      sheetRow(r, [
        c(`A${r}`, `${size} pt`, label),
        c(`B${r}`, `Text at ${size} points`, styles.xf({ font: { size } })),
      ])
    );
    r += 1;
  }
  const looks: readonly (readonly [string, CellStyle['font']])[] = [
    ['Bold', { bold: true }],
    ['Italic', { italic: true }],
    ['Underline', { underline: 'single' }],
    ['Double underline', { underline: 'double' }],
    ['Strikethrough', { strike: true }],
    ['Color', { color: COLORS.green, bold: true }],
  ];
  for (const [name, font] of looks) {
    rows.push(
      sheetRow(r, [c(`A${r}`, name, label), c(`B${r}`, `${name} text`, styles.xf({ font }))])
    );
    r += 1;
  }
  const rich = strings.rich([
    ['Rich text mixes ', ''],
    ['bold', '<b/><sz val="11"/><color theme="1"/><rFont val="Calibri"/>'],
    [', ', ''],
    ['italic', '<i/><sz val="11"/><color theme="1"/><rFont val="Calibri"/>'],
    [', and ', ''],
    ['colored', `<b/><sz val="11"/><color rgb="FF${COLORS.green}"/><rFont val="Calibri"/>`],
    [' runs in one cell.', ''],
  ]);
  rows.push(sheetRow(r, [c(`A${r}`, 'Rich text', label), c(`B${r}`, { rich })]));
  r += 2;

  section('Alignment');
  const box = (align: CellStyle['align']) => styles.xf({ border: THIN, align });
  rows.push(
    sheetRow(
      r,
      [
        c(`A${r}`, 'In a tall row', label),
        c(`B${r}`, 'Left and top', box({ horizontal: 'left', vertical: 'top' })),
        c(`C${r}`, 'Center and middle', box({ horizontal: 'center', vertical: 'center' })),
        c(`D${r}`, 'Right and bottom', box({ horizontal: 'right', vertical: 'bottom' })),
      ],
      { height: 42 }
    )
  );
  r += 1;
  rows.push(
    sheetRow(r, [
      c(`A${r}`, 'Indent', label),
      c(`B${r}`, 'Indent 1', box({ horizontal: 'left', indent: 1 })),
      c(`C${r}`, 'Indent 2', box({ horizontal: 'left', indent: 2 })),
      c(`D${r}`, 'Indent 3', box({ horizontal: 'left', indent: 3 })),
    ])
  );
  r += 1;
  const across = styles.xf({ align: { horizontal: 'centerContinuous' }, fill: COLORS.greenLight });
  rows.push(
    sheetRow(r, [
      c(`A${r}`, 'Center across', label),
      c(`B${r}`, 'Centered across three cells', across),
      c(`C${r}`, undefined, across),
      c(`D${r}`, undefined, across),
    ])
  );
  r += 1;
  const merged = styles.xf({
    font: { bold: true, color: COLORS.white },
    fill: COLORS.green,
    align: { horizontal: 'center', vertical: 'center' },
  });
  rows.push(
    sheetRow(
      r,
      [
        c(`A${r}`, 'Merged cells', label),
        c(`B${r}`, 'One merged range over B:D and two rows', merged),
        c(`C${r}`, undefined, merged),
        c(`D${r}`, undefined, merged),
      ],
      { height: 18 }
    ),
    sheetRow(
      r + 1,
      [
        c(`B${r + 1}`, undefined, merged),
        c(`C${r + 1}`, undefined, merged),
        c(`D${r + 1}`, undefined, merged),
      ],
      {
        height: 18,
      }
    )
  );
  const mergeRef = `B${r}:D${r + 1}`;
  r += 3;

  section('Wrap, shrink, and rotation');
  rows.push(
    sheetRow(
      r,
      [
        c(`A${r}`, 'Wrap and shrink', label),
        c(
          `B${r}`,
          'Wrapped text breaks at spaces and fills the cell line by line.',
          box({ wrap: true, vertical: 'top' })
        ),
        c(`C${r}`, 'Shrink this long text to fit the cell', box({ shrink: true })),
      ],
      { height: 48 }
    )
  );
  r += 1;
  rows.push(
    sheetRow(
      r,
      [
        c(`A${r}`, 'Rotation', label),
        c(`B${r}`, 'Up 45 degrees', box({ rotation: 45 })),
        c(`C${r}`, 'Up 90 degrees', box({ rotation: 90 })),
        c(`D${r}`, 'Down 45 degrees', box({ rotation: 135 })),
        c(`E${r}`, 'ABC', box({ rotation: 255 })),
      ],
      { height: 84 }
    )
  );
  r += 2;

  section('Fills');
  const accents = [4, 5, 6, 7, 8, 9];
  rows.push(
    sheetRow(r, [
      c(`A${r}`, 'Theme colors', label),
      ...accents.map((theme, k) =>
        c(
          `${'BCDEFG'[k]}${r}`,
          `Accent ${k + 1}`,
          // Dark text on the light accents 3 and 4, white on the others.
          styles.xf({
            fill: { theme },
            font: { color: theme === 6 || theme === 7 ? COLORS.ink : COLORS.white },
          })
        )
      ),
    ])
  );
  r += 1;
  rows.push(
    sheetRow(r, [
      c(`A${r}`, 'With a tint', label),
      ...accents.map((theme, k) =>
        c(
          `${'BCDEFG'[k]}${r}`,
          `Tint 60%`,
          styles.xf({ fill: { theme, tint: 0.5999938962981048 } })
        )
      ),
    ])
  );
  return {
    name: 'Text',
    cols: columns([
      [1, 1, 18],
      [2, 2, 30],
      [3, 4, 20],
      [5, 7, 12],
    ]),
    rows: rows.join(''),
    merges: [...merges, mergeRef],
    // No gridlines: the sheet is about text, and its samples overflow their cells.
    gridLines: false,
    headings: true,
  };
}

const EMU_PER_PIXEL = 9525;

const anchorAt = (tag: string, col: number, row: number) =>
  `<xdr:${tag}><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:${tag}>`;

const shapeXml = (id: number, name: string, prst: string, fill: string) =>
  `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvSpPr/></xdr:nvSpPr>` +
  `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>` +
  `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="${COLORS.ink}"/></a:solidFill></a:ln></xdr:spPr></xdr:sp>`;

/** Layout: column widths, row heights, hidden rows and columns, borders, a picture, and shapes. */
export function buildLayoutSheet(styles: StyleTable, strings: StringTable): FixtureSheet {
  const c = cellWriter(strings);
  const group = styles.xf(groupStyle);
  const label = styles.xf({ font: { color: COLORS.muted } });
  const note = styles.xf({ font: { color: COLORS.muted }, fill: COLORS.white });
  const width = styles.xf({
    fill: COLORS.greenLight,
    border: THIN,
    align: { horizontal: 'center' },
  });
  const rows: string[] = [
    sheetRow(1, [c('A1', 'Column widths', group)], { height: 22 }),
    sheetRow(2, [
      c('A2', 'Width in characters', label),
      c('B2', '6', width),
      c('C2', '12', width),
      c('D2', 'Hidden', width),
      c('E2', '24', width),
    ]),
    sheetRow(3, [c('A3', 'Column D is hidden. Its letter is missing from the headings.', note)]),
    // The notes of rows 3, 10, and 11 sit in merged ranges with a white fill, so no gridline
    // crosses them.
    sheetRow(5, [c('A5', 'Row heights', group)], { height: 22 }),
    sheetRow(6, [c('A6', '12 points', label), c('B6', undefined, width)], { height: 12 }),
    sheetRow(7, [c('A7', '24 points', label), c('B7', undefined, width)], { height: 24 }),
    sheetRow(8, [c('A8', 'Hidden row', label)], { hidden: true }),
    sheetRow(9, [c('A9', '40 points', label), c('B9', undefined, width)], { height: 40 }),
    sheetRow(10, [c('A10', 'Row 8 is hidden. Its number is missing from the headings.', note)]),
    sheetRow(11, [c('F11', 'A picture and three shapes.', note)]),
    sheetRow(12, [c('A12', 'Borders', group)], { height: 22 }),
  ];
  const lines = ['thin', 'medium', 'thick', 'dashed', 'dotted', 'double'];
  lines.forEach((style, k) => {
    const r = 13 + k * 2;
    rows.push(
      sheetRow(r, [
        c(`A${r}`, style, label),
        c(
          `C${r}`,
          undefined,
          styles.xf({ border: { left: style, right: style, top: style, bottom: style } })
        ),
      ])
    );
  });
  const box = styles.xf({
    border: THIN,
    fill: COLORS.greenLight,
    align: { horizontal: 'center', vertical: 'center', wrap: true },
  });
  rows.push(
    sheetRow(26, [
      c('A26', 'Merged range', label),
      c('B26', 'B26:E28, centered', box),
      c('C26', undefined, box),
      c('E26', undefined, box),
    ]),
    sheetRow(27, [c('B27', undefined, box), c('C27', undefined, box), c('E27', undefined, box)]),
    sheetRow(28, [c('B28', undefined, box), c('C28', undefined, box), c('E28', undefined, box)])
  );
  const picture =
    '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Landscape picture"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>' +
    '<xdr:blipFill><a:blip r:embed="rIdImage1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>' +
    '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>';
  const ext = `<xdr:ext cx="${PICTURE_SIZE.width * EMU_PER_PIXEL}" cy="${PICTURE_SIZE.height * EMU_PER_PIXEL}"/>`;
  const anchors =
    `<xdr:oneCellAnchor>${anchorAt('from', 5, 1)}${ext}${picture}<xdr:clientData/></xdr:oneCellAnchor>` +
    `<xdr:twoCellAnchor>${anchorAt('from', 5, 13)}${anchorAt('to', 6, 18)}${shapeXml(3, 'Rounded rectangle', 'roundRect', COLORS.greenLight)}<xdr:clientData/></xdr:twoCellAnchor>` +
    `<xdr:twoCellAnchor>${anchorAt('from', 6, 13)}${anchorAt('to', 7, 18)}${shapeXml(4, 'Ellipse', 'ellipse', '7FBF9A')}<xdr:clientData/></xdr:twoCellAnchor>` +
    `<xdr:twoCellAnchor>${anchorAt('from', 7, 14)}${anchorAt('to', 8, 17)}${shapeXml(5, 'Arrow', 'rightArrow', COLORS.green)}<xdr:clientData/></xdr:twoCellAnchor>`;
  return {
    name: 'Layout',
    cols: columns([
      [1, 1, 22],
      [2, 2, 6],
      [3, 3, 12],
      [4, 4, 12, true],
      [5, 5, 24],
      [6, 9, 12],
    ]),
    rows: rows.join(''),
    merges: ['A3:E3', 'A10:E10', 'F11:K11', 'B26:E28'],
    gridLines: true,
    headings: true,
    drawing: {
      anchors,
      images: [
        {
          relId: 'rIdImage1',
          name: 'image1.png',
          bytes: landscapePng(),
          contentType: 'image/png',
        },
      ],
    },
  };
}

/** The tasks of each phase, in order. The fifth task of Design is the cancelled, hidden one (`CANCELLED`). */
const TASKS: readonly (readonly [string, readonly string[]])[] = [
  [
    'Plan',
    [
      'Set goals',
      'Map users',
      'List risks',
      'Draft budget',
      'Pick vendors',
      'Kick off',
      'Sign scope',
      'Book rooms',
    ],
  ],
  [
    'Design',
    [
      'Sketch flows',
      'Pick colors',
      'Wireframes',
      'User tests',
      'Print style',
      'Final mockups',
      'Icon set',
      'Design review',
    ],
  ],
  [
    'Build',
    [
      'Set up repo',
      'Data import',
      'Grid view',
      'Editing',
      'Save path',
      'Toolbar',
      'Shortcuts',
      'Code review',
    ],
  ],
  [
    'Test',
    [
      'Test plan',
      'Unit tests',
      'Load tests',
      'Bug bash',
      'Fix bugs',
      'Access audit',
      'Beta group',
      'Sign off',
    ],
  ],
  [
    'Release',
    [
      'Release notes',
      'Train support',
      'Price page',
      'Launch email',
      'Go live',
      'Watch metrics',
      'Retro',
      'Close out',
    ],
  ],
];
/** The index of the cancelled task, whose row is hidden. */
const CANCELLED = 12;

/** Task lengths in days, in turn. */
const LENGTHS = [3, 5, 2, 8, 4, 6, 3, 10, 2, 7, 5, 4, 9];

/** Schedule: dates, date formulas, percentages, a hidden row, and a list to scroll. */
export function buildScheduleSheet(styles: StyleTable, strings: StringTable): FixtureSheet {
  const c = cellWriter(strings);
  const head = styles.xf({
    font: { bold: true, color: COLORS.white },
    fill: COLORS.green,
    border: THIN,
    align: { vertical: 'center' },
  });
  const cell = styles.xf({ border: THIN });
  const date = styles.xf({ border: THIN, numFmt: 'ddd, mmm d, yyyy' });
  const days = styles.xf({ border: THIN, numFmt: '[=1]0" day";0" days"' });
  const note = styles.xf({ font: { color: COLORS.muted }, fill: COLORS.white });
  const pct = styles.xf({ border: THIN, numFmt: '0%' });
  const status = (color: string) => styles.xf({ border: THIN, font: { bold: true, color } });
  const STATUS = {
    done: ['Done', status(COLORS.green)],
    doing: ['In progress', status(COLORS.amber)],
    todo: ['Not started', status(COLORS.muted)],
  } as const;
  const rows: string[] = [
    sheetRow(
      1,
      ['Task', 'Team', 'Start', 'Length', 'End', 'Done', 'Status'].map((t, i) =>
        c(`${'ABCDEFG'[i]}1`, t, head)
      ),
      { height: 20 }
    ),
  ];
  // 0 is Sunday, 6 is Saturday, for a serial of the 1900 date system after February 1900.
  const weekday = (serial: number) => (serial + 6) % 7;
  const tasks = TASKS.flatMap(([phase, names]) => names.map((name) => `${phase}: ${name}`));
  let start = dateSerial(2026, 1, 5);
  for (let i = 0; i < tasks.length; i += 1) {
    const r = i + 2;
    // A task that would end on a Saturday or a Sunday takes one or two days less.
    const planned = LENGTHS[i % LENGTHS.length]!;
    const ends = weekday(start + planned);
    const length = planned - (ends === 6 ? 1 : ends === 0 ? 2 : 0);
    const done = i < 14 ? 1 : i < 20 ? ((i - 13) * 15) / 100 : 0;
    const [text, s] = done === 1 ? STATUS.done : done > 0 ? STATUS.doing : STATUS.todo;
    const cancelled = i === CANCELLED;
    rows.push(
      sheetRow(
        r,
        [
          c(`A${r}`, cancelled ? `${tasks[i]} (cancelled)` : tasks[i]!, cell),
          c(`B${r}`, `Team ${'ABCD'[i % 4]}`, cell),
          c(`C${r}`, start, date),
          c(`D${r}`, length, days),
          c(`E${r}`, { f: `C${r}+D${r}`, v: start + length }, date),
          c(`F${r}`, done, pct),
          c(`G${r}`, text, s),
          // The note sits in a merged range with a white fill, so no gridline crosses it.
          ...(r === 2
            ? [
                c('I2', `Row ${CANCELLED + 2} is hidden: a cancelled task.`, note),
                ...['J2', 'K2', 'L2'].map((ref) => c(ref, undefined, note)),
              ]
            : []),
        ],
        { hidden: cancelled }
      )
    );
    start += 1 + (i % 3);
    // Tasks start on weekdays: move a Saturday or Sunday start to the Monday after it.
    if (weekday(start) === 6) start += 2;
    else if (weekday(start) === 0) start += 1;
  }
  return {
    name: 'Schedule',
    cols: columns([
      [1, 1, 24],
      [2, 2, 10],
      [3, 3, 18],
      [4, 4, 9],
      [5, 5, 18],
      [6, 6, 8],
      [7, 7, 13],
    ]),
    rows: rows.join(''),
    merges: ['I2:L2'],
    gridLines: true,
    headings: true,
    viewChildren: frozenPane(1, 1),
  };
}
