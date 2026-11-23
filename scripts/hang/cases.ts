// The hostile workbooks of the browser hang test (`bun run test:hang:browser`). Each case builds
// its workbook at test time, so the repository holds no large binary. Every value here is one
// that the file's sender controls: AGENTS.md, "Security", names most of them.
//
// To add a case, add one entry to `FILE_CASES` with a unique `id`, the input that it `covers`, and
// the outcome that it must `expect`. Name the first sheet `HOSTILE_SHEET`: the `sheet` and `book`
// helpers do this. Set `firstCell` when the workbook could open and lose its data. Run it with
// `HANG_ONLY=<id> bun run test:hang:browser`.
//
// To add a network fault, add one entry to `NETWORK_CASES`, and apply a new `HangNetwork` value in
// `applyNetwork` of `run.ts`. To add a race, add one entry to `FILE_CASES` with a `race`, and run
// a new `HangRace` value in `runCase` of `run.ts`.
//
// Same role as pptx-editor/scripts/stress/hang-decks.ts at e22f3f3b (EigenPal), for workbooks.

import { FIRST_CELL } from './first-cells.ts';
import { zipSync } from 'fflate';
import { readZip, writeZip } from '../../packages/ooxml-core/src/zip.ts';
import { buildXlsx, c, row, type FixtureSheet, type FixtureWorkbook } from '../fixtures/builder.ts';

/** How the case must end: the workbook shows, or the editor refuses it with a reason. */
export type HangExpectation = 'open' | 'refuse';

/** The network fault of a case. Font faults apply from the first request of the page. */
export type HangNetwork =
  'fonts-stalled' | 'fonts-partly-stalled' | 'fonts-blocked' | 'offline-after-start';

export interface HangCase {
  readonly id: string;
  /** Turn off the demo's headings option, because that step rewrites the hostile value. */
  readonly headingsOff?: true;
  /** The text that cell A1 must show, for a case that could lose its data and still open. */
  readonly firstCell?: string;
  /** The hostile input or the fault. The report and each failure line show it. */
  readonly covers: string;
  readonly expect: HangExpectation;
  readonly network?: HangNetwork;
  /** The workbook to open. A case with none checks only the demo's own first open. */
  readonly bytes?: () => Uint8Array;
  /** The long work that the test starts on this workbook, to open another one during it. */
  readonly race?: HangRace;
}

/** The long work of the demo during which the test opens another workbook. */
export type HangRace = 'headings-switch' | 'headings-second-press' | 'save';

/** The name of the first sheet of every hostile workbook, which the test waits for. */
export const HOSTILE_SHEET = 'Hostile';
/** The name of the sheet of the workbook that the test opens after each case. */
export const RECOVERY_SHEET = 'Recovery';
/** The name of the sheet of the workbook that a race case opens during the long work. */
export const RACE_SHEET = 'Raced';

/** The uncompressed size of the shared string table of the `sst-50mb` case. */
export const LARGE_SST_BYTES = 50 * 1024 * 1024;

const enc = new TextEncoder();
const dec = new TextDecoder();

const sheet = (rows: string, extra: Partial<FixtureSheet> = {}): FixtureSheet => ({
  name: HOSTILE_SHEET,
  rows,
  ...extra,
});

const plainRows = row(1, [c('A1', 'first'), c('B1', 2), c('C1', 'third')]);

const book = (first: FixtureSheet, extra: Partial<FixtureWorkbook> = {}): Uint8Array =>
  buildXlsx({ sheets: [first], ...extra });

/** The same package with the text of one part changed. */
export function patchPart(
  bytes: Uint8Array,
  part: string,
  edit: (xml: string) => string
): Uint8Array {
  const read = readZip(bytes);
  if (!read.ok) throw new Error(`the built package does not read: ${read.reason}`);
  const data = read.entries.get(part);
  if (!data) throw new Error(`the built package has no part ${part}`);
  const entries = new Map(read.entries);
  entries.set(part, enc.encode(edit(dec.decode(data))));
  return writeZip(entries);
}

const STYLES = '/xl/styles.xml';
const STRINGS = '/xl/sharedStrings.xml';

/** A small seeded generator (mulberry32), so each build gives the same bytes. */
export function seeded(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const SHARED_STRING_SEED = 0x51a7;

/** The next 64 characters of `random`. */
function randomText(random: () => number): string {
  const alphabet = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 ';
  let text = '';
  for (let i = 0; i < 64; i += 1) text += alphabet[Math.floor(random() * alphabet.length)];
  return text;
}

/**
 * Shared strings of about `bytes` bytes of XML. The text is random, so the archive stays under
 * the decompression ratio limit, and the read must parse every byte.
 */
export function randomSharedStrings(bytes: number): string[] {
  const random = seeded(SHARED_STRING_SEED);
  const items: string[] = [];
  let total = 0;
  while (total < bytes) {
    const text = randomText(random);
    const item = `<si><t xml:space="preserve">${text}</t></si>`;
    items.push(item);
    total += item.length;
  }
  return items;
}

/** The text of the first item of `randomSharedStrings`. */
export function firstSharedString(): string {
  return randomText(seeded(SHARED_STRING_SEED));
}

/** The first `count` cells of a sheet that name the shared strings 0, 1, and so on. */
function stringRows(count: number): string {
  const rows: string[] = [];
  for (let r = 1; r <= count; r += 1) rows.push(row(r, [c(`A${r}`, r - 1, { t: 's' })]));
  return rows.join('');
}

/** `count` XML items from `item(index)`. */
const repeat = <T>(count: number, item: (i: number) => T): T[] =>
  Array.from({ length: count }, (_, i) => item(i));

/** A column name of the grid for the 0-based `index`. */
function columnName(index: number): string {
  let name = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26))
    name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
}

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** A relationships part with one external target. */
const externalRels = (id: string, type: string, target: string): Uint8Array =>
  enc.encode(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}" TargetMode="External"/>` +
      '</Relationships>'
  );

/** A workbook whose hyperlinks, picture, external link, and connection all name remote content. */
export function externalReferences(): Uint8Array {
  const built = book(
    sheet(plainRows, {
      afterMerges:
        '<hyperlinks><hyperlink ref="A1" r:id="rIdLink"/><hyperlink ref="B1" r:id="rIdScript"/></hyperlinks>',
      rels: [
        ['rIdLink', `${REL}/hyperlink`, 'https://hang-test.invalid/link', true],
        ['rIdScript', `${REL}/hyperlink`, 'javascript:void(document.title="hostile")', true],
      ],
      drawing: {
        anchors:
          '<xdr:oneCellAnchor><xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="952500" cy="952500"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="P"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:link="rIdRemote"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>',
      },
    }),
    {
      externalReferences: '<externalReference r:id="rIdX1"/>',
      extraParts: [
        [
          'xl/externalLinks/externalLink1.xml',
          'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml',
          enc.encode(
            `<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><externalBook xmlns:r="${REL}" r:id="rId1"/></externalLink>`
          ),
          `${REL}/externalLink`,
        ],
        [
          'xl/connections.xml',
          'application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml',
          enc.encode(
            '<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><connection id="1" name="c" type="4" refreshOnLoad="1"><webPr url="https://hang-test.invalid/data"/></connection></connections>'
          ),
          `${REL}/connections`,
        ],
      ],
    }
  );
  const read = readZip(built);
  if (!read.ok) throw new Error(`the built package does not read: ${read.reason}`);
  const entries = new Map(read.entries);
  entries.set(
    '/xl/drawings/_rels/drawing1.xml.rels',
    externalRels('rIdRemote', 'image', 'https://hang-test.invalid/remote.png')
  );
  entries.set(
    '/xl/externalLinks/_rels/externalLink1.xml.rels',
    externalRels('rId1', 'externalLinkPath', 'https://hang-test.invalid/book.xlsx')
  );
  return writeZip(entries);
}

/** A package with a deflated part of 100 MB of zero bytes, far past the ratio limit. */
export function zipRatioBomb(): Uint8Array {
  const read = readZip(book(sheet(plainRows)));
  if (!read.ok) throw new Error(`the built package does not read: ${read.reason}`);
  const files: Record<string, Uint8Array> = {};
  for (const [name, data] of read.entries) files[name.slice(1)] = data;
  files['xl/media/zeros.bin'] = new Uint8Array(100 * 1024 * 1024);
  return zipSync(files, { level: 9 });
}

/**
 * `length` cell style formats that follow the default format at index 0. Each one names the next
 * index in `xfId`, and the last one names index 1, so the chain never ends.
 */
export const styleChain = (length: number): string[] =>
  repeat(
    length,
    (i) =>
      `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="${i + 1 < length ? i + 2 : 1}"/>`
  );

/** A worksheet of 500000 numeric cells in 20 columns. */
export const denseWorkbook = (): Uint8Array =>
  book(
    sheet(
      repeat(25_000, (r) =>
        row(
          r + 1,
          repeat(20, (x) => `<c r="${columnName(x)}${r + 1}"><v>${r * x}</v></c>`)
        )
      ).join('')
    )
  );

const WORKBOOK = '/xl/workbook.xml';

/** A workbook whose relationships name targets outside the package. */
function relationshipTraversal(): Uint8Array {
  const up = '../../../../../etc/passwd';
  const built = book(
    sheet(plainRows, {
      afterMerges: '<hyperlinks><hyperlink ref="A1" r:id="rIdUp"/></hyperlinks>',
      rels: [
        ['rIdUp', `${REL}/hyperlink`, up],
        ['rIdPart', `${REL}/printerSettings`, up],
      ],
      drawing: {
        anchors:
          '<xdr:oneCellAnchor><xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="952500" cy="952500"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="P"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdMedia"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>',
      },
    })
  );
  return withPart(
    built,
    '/xl/drawings/_rels/drawing1.xml.rels',
    enc.encode(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMedia" Type="${REL}/image" Target="../media/../../../../secret.png"/></Relationships>`
    )
  );
}

/** A workbook with content that must stay inert: DDE, OLE, a control, and a query table. */
function activeContent(): Uint8Array {
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const built = book(
    sheet(plainRows, {
      rels: [
        ['rIdOle', `${REL}/oleObject`, '../embeddings/oleObject1.bin'],
        ['rIdControl', `${REL}/control`, '../controls/control1.xml'],
        ['rIdQuery', `${REL}/queryTable`, '../queryTables/queryTable1.xml'],
      ],
    }),
    {
      externalReferences: '<externalReference r:id="rIdDde"/>',
      extraParts: [
        [
          'xl/externalLinks/externalLink1.xml',
          'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml',
          enc.encode(
            `<externalLink xmlns="${NS}"><ddeLink ddeService="hostile" ddeTopic="topic"><ddeItems><ddeItem name="A1" advise="1"/></ddeItems></ddeLink></externalLink>`
          ),
          `${REL}/externalLink`,
        ],
        [
          'xl/embeddings/oleObject1.bin',
          'application/vnd.openxmlformats-officedocument.oleObject',
          enc.encode('hostile'),
        ],
        ['xl/controls/control1.xml', 'application/xml', enc.encode('<control/>')],
        [
          'xl/queryTables/queryTable1.xml',
          'application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml',
          enc.encode(`<queryTable xmlns="${NS}" name="q" connectionId="1" refreshOnLoad="1"/>`),
        ],
      ],
    }
  );
  return patchPart(built, '/xl/worksheets/sheet1.xml', (xml) =>
    xml.replace(
      '</worksheet>',
      '<oleObjects><oleObject progId="Hostile.1" shapeId="1025" r:id="rIdOle"/></oleObjects>' +
        '<controls><control shapeId="1026" r:id="rIdControl" name="Hostile"/></controls></worksheet>'
    )
  );
}

/** The same package with one more part. */
function withPart(bytes: Uint8Array, part: string, data: Uint8Array): Uint8Array {
  const read = readZip(bytes);
  if (!read.ok) throw new Error(`the built package does not read: ${read.reason}`);
  return writeZip(new Map(read.entries).set(part, data));
}

/** A workbook with `names` as its defined names. */
const withDefinedNames = (names: readonly string[]): Uint8Array =>
  patchPart(book(sheet(plainRows)), WORKBOOK, (xml) =>
    xml.replace(
      '</sheets>',
      `</sheets><definedNames>${names
        .map((name) => `<definedName name="${name}">${HOSTILE_SHEET}!$A$1</definedName>`)
        .join('')}</definedNames>`
    )
  );

/** Nested entities that expand to about 10 GB of text when a reader resolves them. */
const ENTITY_BOMB = `<!DOCTYPE sst [<!ENTITY a0 "${'x'.repeat(10)}">${repeat(
  9,
  (i) => `<!ENTITY a${i + 1} "${`&a${i};`.repeat(10)}">`
).join('')}]>`;

const FILE_CASES: readonly HangCase[] = [
  {
    id: 'xml-entity-expansion',
    covers: 'the shared string table declares a DTD with nested entity expansion, never expanded',
    expect: 'open',
    bytes: () =>
      patchPart(book(sheet(stringRows(1)), { strings: ['one'] }), STRINGS, (xml) =>
        xml.replace('<sst', `${ENTITY_BOMB}<sst`).replace('one', '&a9;')
      ),
  },
  {
    id: 'zip-path-traversal',
    covers: 'the package holds a part whose path climbs out of the package',
    expect: 'refuse',
    bytes: () => {
      const read = readZip(book(sheet(plainRows)));
      if (!read.ok) throw new Error(`the built package does not read: ${read.reason}`);
      const files = Object.fromEntries(
        [...read.entries].map(([name, data]) => [name.replace(/^\//, ''), data])
      );
      return zipSync({ ...files, '../../evil.xml': enc.encode('<evil/>') });
    },
  },
  {
    id: 'zip-leading-slash',
    // The reader keeps this entry as a part inside the package. The reader does not refuse it.
    covers: 'a zip entry name starts with a slash',
    expect: 'open',
    bytes: () => {
      const read = readZip(book(sheet(plainRows)));
      if (!read.ok) throw new Error(read.reason);
      const files = Object.fromEntries(
        [...read.entries].map(([name, data]) => [name.replace(/^\//, ''), data])
      );
      return zipSync({ ...files, '/evil.xml': enc.encode('<evil/>') });
    },
  },
  {
    id: 'relationship-target-traversal',
    covers: 'a hyperlink, a picture, and a part relationship name a target outside the package',
    expect: 'open',
    bytes: relationshipTraversal,
  },
  {
    id: 'sheet-names-object-keys',
    covers: 'worksheets are named __proto__, constructor, and prototype',
    expect: 'open',
    firstCell: 'first',
    bytes: () =>
      buildXlsx({
        sheets: [
          sheet(plainRows),
          ...['__proto__', 'constructor', 'prototype'].map((name) => ({ name, rows: plainRows })),
        ],
      }),
  },
  {
    id: 'active-content-inert',
    covers: 'a DDE link, an OLE object, an embedded control, and a query table stay inert',
    expect: 'open',
    firstCell: 'first',
    bytes: activeContent,
  },
  {
    id: 'names-object-keys',
    covers: 'defined names and an attribute are named __proto__, constructor, and prototype',
    expect: 'open',
    bytes: () =>
      patchPart(withDefinedNames(['__proto__', 'constructor', 'prototype']), WORKBOOK, (xml) =>
        xml.replace(`name="${HOSTILE_SHEET}"`, `name="${HOSTILE_SHEET}" __proto__="1"`)
      ),
  },
  {
    id: 'defined-names-many',
    covers: '100000 defined names',
    expect: 'open',
    bytes: () => withDefinedNames(repeat(100_000, (i) => `name_${i}`)),
  },
  {
    id: 'macro-part',
    covers: 'the package holds a macro part, which stays inert',
    expect: 'open',
    bytes: () => withPart(book(sheet(plainRows)), '/xl/vbaProject.bin', new Uint8Array(4096)),
  },
  {
    id: 'dimension-whole-sheet',
    covers: 'x:dimension/@ref names every cell of the grid',
    expect: 'open',
    bytes: () => book(sheet(plainRows, { dimension: 'A1:XFD1048576' })),
  },
  {
    id: 'dimension-garbage',
    covers: 'x:dimension/@ref is outside the grid and has a very long reference',
    expect: 'open',
    bytes: () => book(sheet(plainRows, { dimension: `A1:${'Z'.repeat(4096)}999999999999999` })),
  },
  {
    id: 'sst-count-huge',
    covers: 'x:sst/@count and @uniqueCount state 4294967295 items',
    expect: 'open',
    headingsOff: true,
    bytes: () =>
      patchPart(book(sheet(stringRows(3)), { strings: ['one', 'two', 'three'] }), STRINGS, (xml) =>
        xml.replace(/<sst\b[^>]*>/, (tag) =>
          tag
            .replace(/\s(count|uniqueCount)="[^"]*"/g, '')
            .replace('<sst', '<sst count="4294967295" uniqueCount="4294967295"')
        )
      ),
  },
  {
    id: 'sst-index-huge',
    covers: 'a cell names shared string 4294967295',
    expect: 'open',
    bytes: () =>
      book(sheet(row(1, [c('A1', 4294967295, { t: 's' }), c('B1', -1, { t: 's' })])), {
        strings: ['one'],
      }),
  },
  {
    id: 'row-number-hostile',
    covers: 'x:row/@r is out of range, negative, not a number, and out of order',
    expect: 'open',
    bytes: () =>
      book(
        sheet(
          [
            '<row r="1048576"><c r="A1048576"><v>1</v></c></row>',
            '<row r="4294967295"><c r="A4294967295"><v>2</v></c></row>',
            '<row r="-1"><c r="A1"><v>3</v></c></row>',
            '<row r="abc"><c r="B1"><v>4</v></c></row>',
            '<row r="2"><c r="A2"><v>5</v></c></row>',
          ].join('')
        )
      ),
  },
  {
    id: 'row-last-of-grid',
    covers: 'one cell in the last row and the last column of the grid',
    expect: 'open',
    bytes: () => book(sheet(row(1048576, [c('XFD1048576', 'last')]))),
  },
  {
    id: 'cell-reference-hostile',
    covers: 'x:c/@r is outside the grid, repeated, not a reference, and very long',
    expect: 'open',
    bytes: () =>
      book(
        sheet(
          row(1, [
            '<c r="ZZZZZZ1"><v>1</v></c>',
            '<c r="A0"><v>2</v></c>',
            '<c r="A99999999999"><v>3</v></c>',
            '<c r="A1"><v>4</v></c>',
            '<c r="A1"><v>5</v></c>',
            `<c r="${'A'.repeat(65536)}1"><v>6</v></c>`,
          ])
        )
      ),
  },
  {
    id: 'row-spans-huge',
    covers: 'x:row/@spans states ranges of up to 4294967295 columns',
    expect: 'open',
    bytes: () =>
      book(
        sheet(
          repeat(1000, (i) =>
            row(i + 1, [c(`A${i + 1}`, i)], 'spans="1:2147483647 5:4294967295 0:-1"')
          ).join('')
        )
      ),
  },
  {
    id: 'col-range-huge',
    covers: 'x:col/@min and @max are outside the grid or reversed',
    expect: 'open',
    bytes: () =>
      book(
        sheet(plainRows, {
          cols:
            '<cols><col min="1" max="2147483647" width="9" customWidth="1"/>' +
            '<col min="0" max="-1" width="9" customWidth="1"/>' +
            '<col min="16384" max="1" width="9" customWidth="1"/>' +
            '<col min="4294967295" max="4294967295" width="9" customWidth="1"/></cols>',
        })
      ),
  },
  {
    id: 'col-overlap-many',
    covers: '20000 x:col ranges that each cover the whole grid',
    expect: 'open',
    bytes: () =>
      book(
        sheet(plainRows, {
          cols: `<cols>${repeat(
            20_000,
            (i) => `<col min="1" max="16384" width="${8 + (i % 40)}" customWidth="1"/>`
          ).join('')}</cols>`,
        })
      ),
  },
  {
    id: 'everything-hidden',
    covers: 'every column and each stored row are hidden, and the default row has no height',
    expect: 'open',
    bytes: () =>
      book(
        sheet(repeat(50, (i) => row(i + 1, [c(`A${i + 1}`, 'hidden')], 'hidden="1"')).join(''), {
          cols: '<cols><col min="1" max="16384" width="9" hidden="1"/></cols>',
          formatPr: 'defaultRowHeight="15" zeroHeight="1"',
        })
      ),
  },
  {
    id: 'rows-alternate-hidden',
    covers: '100000 rows with their own heights, and 100000 hidden rows between them',
    expect: 'open',
    bytes: () =>
      book(
        sheet(
          repeat(200_000, (i) =>
            i % 2
              ? `<row r="${i + 1}" hidden="1"/>`
              : `<row r="${i + 1}" ht="20" customHeight="1"/>`
          ).join('')
        )
      ),
  },
  {
    id: 'axis-size-hostile',
    covers: 'row heights and column widths of 0, 1e308, negative, and not a number',
    expect: 'open',
    bytes: () =>
      book(
        sheet(
          [
            row(1, [c('A1', 'a')], 'ht="1e308" customHeight="1"'),
            row(2, [c('A2', 'b')], 'ht="-5" customHeight="1"'),
            row(3, [c('A3', 'c')], 'ht="NaN" customHeight="1"'),
            row(4, [c('A4', 'd')], 'ht="0" customHeight="1"'),
          ].join(''),
          {
            cols:
              '<cols><col min="1" max="1" width="1e308" customWidth="1"/>' +
              '<col min="2" max="2" width="-5" customWidth="1"/>' +
              '<col min="3" max="3" width="NaN" customWidth="1"/>' +
              '<col min="4" max="4" width="0" customWidth="1"/></cols>',
            formatPr: 'baseColWidth="2147483647" defaultColWidth="1e308" defaultRowHeight="1e308"',
          }
        )
      ),
  },
  {
    id: 'axis-default-zero',
    covers: 'the default row height and the default column width are 0',
    expect: 'open',
    bytes: () =>
      book(
        sheet(plainRows, { formatPr: 'baseColWidth="0" defaultColWidth="0" defaultRowHeight="0"' })
      ),
  },
  {
    id: 'merge-whole-sheet',
    covers: 'one merged range covers every cell of the grid',
    expect: 'open',
    bytes: () => book(sheet(plainRows, { merges: ['A1:XFD1048576'] })),
  },
  {
    id: 'merge-overlap-many',
    covers: '30000 merged ranges over the same rows, with repeated and duplicate ranges',
    expect: 'open',
    bytes: () =>
      book(
        sheet(plainRows, {
          merges: repeat(30_000, (i) => {
            const name = columnName(i % 16_384);
            return `${name}1:${name}1048576`;
          }),
        })
      ),
  },
  {
    id: 'merge-small-many',
    covers: '100000 small merged ranges that do not overlap',
    expect: 'open',
    bytes: () =>
      book(
        sheet(plainRows, {
          merges: repeat(100_000, (i) => `A${i * 2 + 2}:B${i * 2 + 2}`),
        })
      ),
  },
  {
    id: 'merge-hostile-reference',
    covers: 'merged ranges that are reversed, outside the grid, or not references',
    expect: 'open',
    bytes: () =>
      book(
        sheet(plainRows, {
          merges: ['XFD1048576:A1', 'A1:ZZZZZZ99999999', 'not-a-range', 'A1', ':'],
        })
      ),
  },
  {
    id: 'style-xml-deep',
    covers: 'the styles part nests 200000 elements',
    expect: 'open',
    bytes: () =>
      patchPart(book(sheet(plainRows)), STYLES, (xml) =>
        xml.replace(
          '</styleSheet>',
          `<extLst>${'<ext uri="a">'.repeat(200_000)}${'</ext>'.repeat(200_000)}</extLst></styleSheet>`
        )
      ),
  },
  {
    id: 'style-chain-cycle',
    covers:
      'a cell enters 60000 cell style formats whose x:xf/@xfId values form one long cycle, ' +
      'which a walk of the chain would never leave',
    expect: 'open',
    bytes: () =>
      book(sheet(row(1, [c('A1', 'chain', { s: 1 })])), {
        styles: {
          cellStyleXfs: styleChain(60_000),
          cellXfs: ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="1"/>'],
        },
      }),
  },
  {
    id: 'style-index-hostile',
    covers: 'x:c/@s, font, fill, border, and number format indexes are out of range',
    expect: 'open',
    bytes: () =>
      book(
        sheet(
          row(1, [
            '<c r="A1" s="4294967295"><v>1</v></c>',
            '<c r="B1" s="-1"><v>2</v></c>',
            c('C1', 3, { s: 1 }),
          ])
        ),
        {
          styles: {
            cellXfs: [
              '<xf numFmtId="4294967295" fontId="4294967295" fillId="4294967295" borderId="4294967295" xfId="4294967295"/>',
            ],
          },
        }
      ),
  },
  {
    id: 'style-many-formats',
    covers: '64000 fonts and cell formats, each used by one cell',
    expect: 'open',
    bytes: () =>
      book(sheet(repeat(64_000, (i) => row(i + 1, [c(`A${i + 1}`, i, { s: i + 1 })])).join('')), {
        styles: {
          fonts: repeat(
            64_000,
            (i) => `<font><sz val="${8 + (i % 20)}"/><name val="F${i}"/></font>`
          ),
          cellXfs: repeat(
            64_000,
            (i) =>
              `<xf numFmtId="0" fontId="${i + 1}" fillId="0" borderId="0" xfId="0" applyFont="1"/>`
          ),
        },
      }),
  },
  {
    id: 'font-size-hostile',
    covers: 'font sizes that are huge, zero, or negative, and huge indents and text rotations',
    expect: 'open',
    bytes: () =>
      book(
        sheet(
          row(1, [
            c('A1', 'huge', { s: 1 }),
            c('B1', 'zero', { s: 2 }),
            c('C1', 'negative', { s: 3 }),
            c('D1', 'indent', { s: 4 }),
            c('E1', 'rotation', { s: 5 }),
          ])
        ),
        {
          styles: {
            fonts: [
              '<font><sz val="1e308"/><name val="A"/></font>',
              '<font><sz val="0"/><name val="A"/></font>',
              '<font><sz val="-5"/><name val="A"/></font>',
            ],
            cellXfs: [
              '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
              '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
              '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
              '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left" indent="2147483647" wrapText="1"/></xf>',
              '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment textRotation="2147483647"/></xf>',
            ],
          },
        }
      ),
  },
  {
    id: 'number-format-hostile',
    covers: 'number format codes of 1 MB, with many sections, quotes, and fill characters',
    expect: 'open',
    bytes: () =>
      book(
        sheet(
          row(1, [
            c('A1', 1234.5, { s: 1 }),
            c('B1', 1234.5, { s: 2 }),
            c('C1', 1234.5, { s: 3 }),
            c('D1', 1234.5, { s: 4 }),
          ])
        ),
        {
          styles: {
            numFmts: [
              `<numFmt numFmtId="164" formatCode="${'#'.repeat(1_000_000)}"/>`,
              `<numFmt numFmtId="165" formatCode="${'0;'.repeat(100_000)}"/>`,
              `<numFmt numFmtId="166" formatCode="${'&quot;a'.repeat(100_000)}"/>`,
              `<numFmt numFmtId="167" formatCode="${'*x'.repeat(100_000)}0"/>`,
            ],
            cellXfs: [164, 165, 166, 167].map(
              (id) =>
                `<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`
            ),
          },
        }
      ),
  },
  {
    id: 'cell-values-hostile',
    covers: 'numeric cell values that overflow, and date serials far outside the calendar',
    expect: 'open',
    bytes: () =>
      book(
        sheet(
          row(1, [
            '<c r="A1"><v>1e308</v></c>',
            '<c r="B1"><v>1e400</v></c>',
            '<c r="C1"><v>NaN</v></c>',
            '<c r="D1" s="1"><v>1e308</v></c>',
            '<c r="E1" s="1"><v>-1e308</v></c>',
            `<c r="F1"><v>${'9'.repeat(100_000)}</v></c>`,
          ])
        ),
        {
          styles: {
            cellXfs: [
              '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
            ],
          },
        }
      ),
  },
  {
    id: 'text-long-cell',
    covers: 'one wrapped cell holds 10 MB of text with no space',
    expect: 'open',
    bytes: () => {
      const random = seeded(7);
      const text = Array.from({ length: 10_000_000 }, () =>
        String.fromCharCode(97 + Math.floor(random() * 26))
      ).join('');
      return book(sheet(row(1, [c('A1', text, { s: 1 })])), {
        styles: {
          cellXfs: [
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf>',
          ],
        },
      });
    },
  },
  {
    id: 'wrapped-rows-many',
    covers: '200000 rows without a height, each with 3000 characters of wrapped text',
    expect: 'open',
    firstCell: 'wrap wrap',
    bytes: () =>
      book(
        sheet(repeat(200_000, (r) => row(r + 1, [c(`A${r + 1}`, 0, { t: 's', s: 1 })])).join('')),
        {
          rawStrings: [`<si><t>${'wrap '.repeat(600)}</t></si>`],
          styles: {
            cellXfs: [
              '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf>',
            ],
          },
        }
      ),
  },
  {
    id: 'text-many-runs',
    covers: 'one shared string holds 200000 rich text runs',
    expect: 'open',
    firstCell: 'xxxxxxxxxx',
    bytes: () =>
      book(sheet(row(1, [c('A1', 0, { t: 's' })])), {
        rawStrings: [
          `<si>${repeat(200_000, (i) => `<r><rPr><b/><sz val="${8 + (i % 8)}"/></rPr><t>x</t></r>`).join('')}</si>`,
        ],
      }),
  },
  {
    id: 'view-hostile',
    covers: 'the sheet view, pane, and selection name cells and sizes outside the grid',
    expect: 'open',
    bytes: () =>
      book(
        sheet(plainRows, {
          viewAttrs: 'topLeftCell="XFD1048576" zoomScale="4294967295" zoomScaleNormal="0"',
          viewChildren:
            '<pane xSplit="4294967295" ySplit="4294967295" topLeftCell="ZZZZZZ99999999" state="frozen"/>' +
            `<selection activeCell="XFD1048576" sqref="${repeat(50_000, (i) => `A${i + 1}`).join(' ')}"/>`,
        })
      ),
  },
  {
    id: 'conditional-many',
    covers: '20000 conditional formatting rules that each cover the whole grid',
    expect: 'open',
    bytes: () =>
      book(
        sheet(plainRows, {
          afterMerges: repeat(
            20_000,
            (i) =>
              `<conditionalFormatting sqref="A1:XFD1048576"><cfRule type="cellIs" dxfId="0" priority="${i + 1}" operator="greaterThan"><formula>${i}</formula></cfRule></conditionalFormatting>`
          ).join(''),
        }),
        { styles: { dxfs: ['<dxf><font><b/></font></dxf>'] } }
      ),
  },
  {
    id: 'drawing-anchors-hostile',
    covers: '20000 drawing anchors with positions and sizes outside the grid',
    expect: 'open',
    bytes: () =>
      book(
        sheet(plainRows, {
          drawing: {
            anchors: repeat(
              20_000,
              (i) =>
                `<xdr:twoCellAnchor><xdr:from><xdr:col>${i * 1000}</xdr:col><xdr:colOff>9223372036854775807</xdr:colOff><xdr:row>4294967295</xdr:row><xdr:rowOff>-1</xdr:rowOff></xdr:from><xdr:to><xdr:col>2147483647</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2147483647</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="${i + 2}" name="S"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9223372036854775807" cy="9223372036854775807"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>`
            ).join(''),
          },
        })
      ),
  },
  {
    id: 'sheets-many',
    covers: '2000 worksheets, so 2000 sheet tabs',
    expect: 'open',
    bytes: () =>
      buildXlsx({
        sheets: [
          sheet(plainRows),
          ...repeat(1999, (i) => ({ name: `Sheet ${i + 2}`, rows: plainRows })),
        ],
      }),
  },
  {
    id: 'external-references',
    covers: 'hyperlinks, a remote picture, an external workbook, and a data connection',
    expect: 'open',
    bytes: externalReferences,
  },
  {
    id: 'zip-ratio-bomb',
    covers: 'a part of 100 MB of zero bytes that deflates far past the ratio limit',
    expect: 'refuse',
    bytes: zipRatioBomb,
  },
  {
    id: 'sheet-dense',
    covers: 'a worksheet of 500000 numeric cells in 20 columns',
    expect: 'open',
    bytes: denseWorkbook,
  },
  {
    id: 'race-open-during-switch',
    covers: 'another workbook opens while the headings switch of a large workbook runs',
    expect: 'open',
    race: 'headings-switch',
    bytes: denseWorkbook,
  },
  {
    id: 'race-second-press-during-switch',
    covers:
      'a second press of the headings option and a Save during the switch of a large workbook',
    expect: 'open',
    race: 'headings-second-press',
    bytes: denseWorkbook,
  },
  {
    id: 'race-open-during-save',
    covers: 'another workbook opens while the save of a large workbook runs',
    expect: 'open',
    race: 'save',
    bytes: denseWorkbook,
  },
  {
    id: 'sst-50mb',
    covers: 'a shared string table of 50 MB',
    expect: 'open',
    firstCell: firstSharedString(),
    bytes: () =>
      book(sheet(stringRows(1000)), {
        rawStrings: randomSharedStrings(LARGE_SST_BYTES),
      }),
  },
];

const NETWORK_CASES: readonly HangCase[] = [
  {
    id: 'fonts-stalled',
    covers: 'every font download stays open and sends no byte',
    expect: 'refuse',
    network: 'fonts-stalled',
  },
  {
    id: 'fonts-partly-stalled',
    covers: 'the bold font downloads stay open, and the other faces arrive',
    expect: 'open',
    network: 'fonts-partly-stalled',
  },
  {
    id: 'fonts-blocked',
    covers: 'the network refuses every font download',
    expect: 'refuse',
    network: 'fonts-blocked',
  },
  {
    id: 'network-blocked',
    covers: 'the network goes away after the demo starts, and a workbook names remote content',
    expect: 'open',
    network: 'offline-after-start',
    bytes: externalReferences,
  },
];

/** Every case, in the order of the run. */
export const HANG_CASES: readonly HangCase[] = [...FILE_CASES, ...NETWORK_CASES].map((c) => {
  const firstCell = c.firstCell ?? FIRST_CELL.get(c.id);
  return firstCell === undefined ? c : { ...c, firstCell };
});

/** The small workbook that the test opens after each case, to prove that the editor still works. */
export const recoveryWorkbook = (name: string = RECOVERY_SHEET): Uint8Array =>
  buildXlsx({ sheets: [{ name, rows: plainRows }] });
