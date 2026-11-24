// A minimal, valid SpreadsheetML package builder for synthetic fixtures and tests.
//
// Same role as pptx-editor/scripts/fixtures/builder.ts at 1943d2c9 (EigenPal), for workbooks.
// It writes the parts a consumer needs to open a workbook with no repair prompt: content types,
// package and workbook relationships, the workbook, one styles part, one theme, the shared string
// table, worksheets, and optional drawings with pictures.
//
// Every worksheet gets the page setup of design D8, so its printed page maps to the grid: one
// page, 100 % scale, fixed margins, and gridlines and headings set explicitly in
// `x:printOptions` to the same values as the sheet view. The paper is Letter (`paperSize="1"`):
// a fixture's content must fit one Letter page at 100 % inside 0.5 inch margins, or it sets
// `sheetPr` to fit the page and records the scale.
//
// The build is deterministic: the parts are written in a fixed order through the engine's own
// zip writer, whose entry time is fixed. That writer reads local time fields, so
// `scripts/fixtures/build.ts` refuses to write committed fixtures outside UTC.

import { writeZip } from '../../packages/ooxml-core/src/zip.ts';

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
export const NS_X = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const NS_PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'application/vnd.openxmlformats-officedocument';
/**
 * The macro-enabled form of the content type of the main part. ECMA-376 Part 1, 12.3.23 (Workbook
 * Part) lists only the workbook and the template types. A macro-enabled package uses this type
 * with the same root element and relationship (packages/core/src/store/package/main-part-type.ts).
 */
const MACRO_WORKBOOK = 'application/vnd.ms-excel.sheet.macroEnabled.main+xml';

/** Escape text for XML character data and attribute values. */
export const esc = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** One picture or shape of a sheet's drawing, as the XML of one `xdr:*Anchor`. */
export interface FixtureImage {
  /** The relationship id that the anchor XML uses in `r:embed`. */
  readonly relId: string;
  readonly name: string;
  readonly bytes: Uint8Array;
  readonly contentType: 'image/png' | 'image/jpeg';
}

export interface FixtureSheet {
  readonly name: string;
  /** `x:sheetPr` children, for example `<pageSetUpPr fitToPage="1"/>`. */
  readonly sheetPr?: string;
  /** `x:dimension/@ref`. No fixture writes one unless it sets this. */
  readonly dimension?: string;
  readonly state?: 'hidden' | 'veryHidden';
  /** `x:sheetData` children: the `x:row` elements. */
  readonly rows: string;
  /** The `x:cols` element, when the sheet has one. */
  readonly cols?: string;
  /** Merged ranges (`A1:B2`). */
  readonly merges?: readonly string[];
  /** `x:sheetFormatPr` attributes. `defaultRowHeight="15"` when absent. */
  readonly formatPr?: string;
  /** Extra attributes of `x:sheetView`, and its children (`x:pane`, `x:selection`). */
  readonly viewAttrs?: string;
  readonly viewChildren?: string;
  /** Show gridlines on screen and on the printed page. Off by default. */
  readonly gridLines?: boolean;
  /** Show row and column headings on screen and on the printed page. Off by default. */
  readonly headings?: boolean;
  /** `x:sheetCalcPr` attributes. */
  readonly calcPr?: string;
  /** Elements between `x:pageSetup` and `x:drawing` in schema order (for example `x:headerFooter`). */
  readonly afterPageSetup?: string;
  /** Elements after `x:sheetData` and before `x:mergeCells` (`x:sheetCalcPr` is separate). */
  readonly beforeMerges?: string;
  /** Elements between `x:mergeCells` and `x:printOptions` (conditional formatting, hyperlinks). */
  readonly afterMerges?: string;
  /** `x:extLst` children. */
  readonly ext?: string;
  /** The drawing: anchors XML and the pictures they embed. */
  readonly drawing?: { readonly anchors: string; readonly images?: readonly FixtureImage[] };
  /** Relationships of the sheet: `[id, type, target, external?]`. */
  readonly rels?: readonly (readonly [string, string, string, boolean?])[];
}

export interface FixtureWorkbook {
  readonly sheets: readonly FixtureSheet[];
  /** Shared strings, in order. A cell names one with `t="s"` and its index. */
  readonly strings?: readonly string[];
  /** Raw `x:si` XML items appended after `strings` (rich text runs, escapes). */
  readonly rawStrings?: readonly string[];
  /** Styles part content: each list holds the XML of its items after the defaults. */
  readonly styles?: {
    readonly numFmts?: readonly string[];
    readonly fonts?: readonly string[];
    readonly fills?: readonly string[];
    readonly borders?: readonly string[];
    /** Cell style formats after the Normal style format (`xfId` 0). */
    readonly cellStyleXfs?: readonly string[];
    /** The Normal style format (`xfId` 0). It replaces the default when present. */
    readonly normalStyleXf?: string;
    readonly cellXfs?: readonly string[];
    /** `x:cellStyle` items after the Normal style. */
    readonly cellStyles?: readonly string[];
    readonly dxfs?: readonly string[];
    readonly indexedColors?: readonly string[];
  };
  /** The default font: `x:font` children. Carlito 11 with the minor theme font when absent. */
  readonly defaultFont?: string;
  readonly date1904?: boolean;
  /** `x:workbookProtection` attributes; no element when absent. */
  readonly workbookProtection?: string;
  /** `x:calcPr` attributes. */
  readonly calcPr?: string;
  /** `x:externalReference` elements, written in `x:externalReferences` before the names. */
  readonly externalReferences?: string;
  /** `x:definedName` elements. */
  readonly definedNames?: string;
  readonly activeTab?: number;
  /** The `x:bookViews` content in place of a view with `activeTab`; an empty string writes none. */
  readonly bookViews?: string;
  /** Theme fonts and colors. */
  readonly theme?: { readonly minor?: string; readonly major?: string; readonly accent1?: string };
  /**
   * Write `docProps/app.xml` with the sheet names in `TitlesOfParts`, as the extended properties
   * of a saved workbook hold them. A sheet op leaves these titles as they were (design D1 of the
   * change `workbook-structure-editing`).
   */
  readonly appProps?: boolean;
  /** Workbook elements after `x:calcPr` in schema order (for example `x:customWorkbookViews`). */
  readonly afterCalcPr?: string;
  /**
   * Extra parts: `[partName, contentType, bytes, workbook relationship type?]`. An empty content
   * type writes no override. The workbook relationship of the extra part at index `i` has the
   * id `rIdX{i + 1}`.
   */
  readonly extraParts?: readonly (readonly [string, string, Uint8Array, string?])[];
  /** Raw `x:sheet` elements after the sheets, for example one that names an extra part. */
  readonly extraSheets?: string;
  /**
   * Give the main part the macro-enabled content type. The builder writes it in the `Override`
   * element of the main part (ECMA-376 Part 2, 7.2.3.2.5, `CT_Override`). The builder writes no
   * macro part. A script cannot make a valid macro project, and a consumer can report damage in
   * an invalid one. So the package is a macro-enabled workbook with no code. Its catalog entry
   * sets the extension `xlsm` (`scripts/fixtures/catalog.ts`).
   */
  readonly macroEnabled?: boolean;
}

const DEFAULT_FONT =
  '<sz val="11"/><color theme="1"/><name val="Carlito"/><family val="2"/><scheme val="minor"/>';

function theme(t: FixtureWorkbook['theme'] = {}): string {
  const clr = (name: string, value: string): string =>
    `<a:${name}><a:srgbClr val="${value}"/></a:${name}>`;
  const fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
  const line = '<a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>';
  return (
    `${XML}<a:theme xmlns:a="${NS_A}" name="Fixture Theme"><a:themeElements>` +
    '<a:clrScheme name="Fixture">' +
    '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>' +
    '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
    clr('dk2', '44546A') +
    clr('lt2', 'E7E6E6') +
    clr('accent1', t.accent1 ?? '4472C4') +
    clr('accent2', 'ED7D31') +
    clr('accent3', 'A5A5A5') +
    clr('accent4', 'FFC000') +
    clr('accent5', '5B9BD5') +
    clr('accent6', '70AD47') +
    clr('hlink', '0563C1') +
    clr('folHlink', '954F72') +
    '</a:clrScheme>' +
    '<a:fontScheme name="Fixture">' +
    `<a:majorFont><a:latin typeface="${esc(t.major ?? 'Carlito')}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>` +
    `<a:minorFont><a:latin typeface="${esc(t.minor ?? 'Carlito')}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>` +
    '</a:fontScheme>' +
    '<a:fmtScheme name="Fixture">' +
    `<a:fillStyleLst>${fill}${fill}${fill}</a:fillStyleLst>` +
    `<a:lnStyleLst>${line}${line}${line}</a:lnStyleLst>` +
    '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>' +
    `<a:bgFillStyleLst>${fill}${fill}${fill}</a:bgFillStyleLst>` +
    '</a:fmtScheme></a:themeElements></a:theme>'
  );
}

function styles(wb: FixtureWorkbook): string {
  const s = wb.styles ?? {};
  const list = (name: string, defaults: readonly string[], extra: readonly string[] = []) => {
    const items = [...defaults, ...extra];
    return `<${name} count="${items.length}">${items.join('')}</${name}>`;
  };
  const numFmts = s.numFmts?.length
    ? `<numFmts count="${s.numFmts.length}">${s.numFmts.join('')}</numFmts>`
    : '';
  const dxfs = `<dxfs count="${s.dxfs?.length ?? 0}">${(s.dxfs ?? []).join('')}</dxfs>`;
  const colors = s.indexedColors
    ? `<colors><indexedColors>${s.indexedColors.map((c) => `<rgbColor rgb="${c}"/>`).join('')}</indexedColors></colors>`
    : '';
  return (
    `${XML}<styleSheet xmlns="${NS_X}">` +
    numFmts +
    list('fonts', [`<font>${wb.defaultFont ?? DEFAULT_FONT}</font>`], s.fonts) +
    list(
      'fills',
      [
        '<fill><patternFill patternType="none"/></fill>',
        '<fill><patternFill patternType="gray125"/></fill>',
      ],
      s.fills
    ) +
    list('borders', ['<border><left/><right/><top/><bottom/><diagonal/></border>'], s.borders) +
    list(
      'cellStyleXfs',
      [s.normalStyleXf ?? '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'],
      s.cellStyleXfs
    ) +
    list('cellXfs', ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'], s.cellXfs) +
    list('cellStyles', ['<cellStyle name="Normal" xfId="0" builtinId="0"/>'], s.cellStyles) +
    dxfs +
    '<tableStyles count="0"/>' +
    colors +
    '</styleSheet>'
  );
}

function worksheet(sheet: FixtureSheet, active: boolean): string {
  const grid = sheet.gridLines === true ? '1' : '0';
  const headings = sheet.headings === true ? '1' : '0';
  const merges = sheet.merges?.length
    ? `<mergeCells count="${sheet.merges.length}">${sheet.merges.map((m) => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>`
    : '';
  return (
    `${XML}<worksheet xmlns="${NS_X}" xmlns:r="${NS_R}">` +
    (sheet.sheetPr ? `<sheetPr>${sheet.sheetPr}</sheetPr>` : '') +
    (sheet.dimension ? `<dimension ref="${esc(sheet.dimension)}"/>` : '') +
    `<sheetViews><sheetView showGridLines="${grid}" showRowColHeaders="${headings}"${active ? ' tabSelected="1"' : ''} workbookViewId="0"${sheet.viewAttrs ? ` ${sheet.viewAttrs}` : ''}>` +
    `${sheet.viewChildren ?? ''}</sheetView></sheetViews>` +
    `<sheetFormatPr ${sheet.formatPr ?? 'defaultRowHeight="15"'}/>` +
    (sheet.cols ?? '') +
    `<sheetData>${sheet.rows}</sheetData>` +
    (sheet.calcPr ? `<sheetCalcPr ${sheet.calcPr}/>` : '') +
    (sheet.beforeMerges ?? '') +
    merges +
    (sheet.afterMerges ?? '') +
    `<printOptions gridLines="${grid}" headings="${headings}"/>` +
    '<pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5" header="0" footer="0"/>' +
    '<pageSetup paperSize="1" scale="100" orientation="portrait"/>' +
    (sheet.afterPageSetup ?? '') +
    (sheet.drawing ? '<drawing r:id="rIdDrawing"/>' : '') +
    (sheet.ext ? `<extLst>${sheet.ext}</extLst>` : '') +
    '</worksheet>'
  );
}

function rels(list: readonly (readonly [string, string, string, boolean?])[]): string {
  return (
    `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    list
      .map(
        ([id, type, target, external]) =>
          `<Relationship Id="${id}" Type="${type}" Target="${esc(target)}"${external ? ' TargetMode="External"' : ''}/>`
      )
      .join('') +
    '</Relationships>'
  );
}

/** Build the bytes of one workbook. */
export function buildXlsx(wb: FixtureWorkbook): Uint8Array {
  const enc = new TextEncoder();
  const parts = new Map<string, Uint8Array>();
  const overrides: [string, string][] = [];
  const put = (name: string, text: string, type?: string) => {
    parts.set(name, enc.encode(text));
    if (type) overrides.push([name, type]);
  };
  const wbRels: [string, string, string][] = [];
  const sheetEls: string[] = [];
  wb.sheets.forEach((sheet, i) => {
    const n = i + 1;
    put(
      `xl/worksheets/sheet${n}.xml`,
      worksheet(sheet, i === (wb.activeTab ?? 0)),
      `${CT}.spreadsheetml.worksheet+xml`
    );
    wbRels.push([`rId${n}`, `${R}/worksheet`, `worksheets/sheet${n}.xml`]);
    sheetEls.push(
      `<sheet name="${esc(sheet.name)}" sheetId="${n}"${sheet.state ? ` state="${sheet.state}"` : ''} r:id="rId${n}"/>`
    );
    const sheetRels: (readonly [string, string, string, boolean?])[] = [...(sheet.rels ?? [])];
    if (sheet.drawing) {
      sheetRels.push(['rIdDrawing', `${R}/drawing`, `../drawings/drawing${n}.xml`]);
      put(
        `xl/drawings/drawing${n}.xml`,
        `${XML}<xdr:wsDr xmlns:xdr="${NS_XDR}" xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:pic="${NS_PIC}">${sheet.drawing.anchors}</xdr:wsDr>`,
        `${CT}.drawing+xml`
      );
      const images = sheet.drawing.images ?? [];
      images.forEach((image) => {
        parts.set(`xl/media/${image.name}`, image.bytes);
      });
      if (images.length)
        put(
          `xl/drawings/_rels/drawing${n}.xml.rels`,
          rels(
            images.map((image) => [image.relId, `${R}/image`, `../media/${image.name}`] as const)
          )
        );
    }
    if (sheetRels.length) put(`xl/worksheets/_rels/sheet${n}.xml.rels`, rels(sheetRels));
  });
  const k = wb.sheets.length;
  wbRels.push([`rId${k + 1}`, `${R}/styles`, 'styles.xml']);
  wbRels.push([`rId${k + 2}`, `${R}/theme`, 'theme/theme1.xml']);
  const strings = [
    ...(wb.strings ?? []).map((s) => `<si><t xml:space="preserve">${esc(s)}</t></si>`),
    ...(wb.rawStrings ?? []),
  ];
  if (strings.length) {
    wbRels.push([`rId${k + 3}`, `${R}/sharedStrings`, 'sharedStrings.xml']);
    put(
      'xl/sharedStrings.xml',
      `${XML}<sst xmlns="${NS_X}" count="${strings.length}" uniqueCount="${strings.length}">${strings.join('')}</sst>`,
      `${CT}.spreadsheetml.sharedStrings+xml`
    );
  }
  (wb.extraParts ?? []).forEach(([name, type, bytes, relType], i) => {
    parts.set(name.replace(/^\//, ''), bytes);
    if (type) overrides.push([name.replace(/^\//, ''), type]);
    if (relType) wbRels.push([`rIdX${i + 1}`, relType, name.replace(/^\/?xl\//, '')]);
  });
  put(
    'xl/workbook.xml',
    `${XML}<workbook xmlns="${NS_X}" xmlns:r="${NS_R}">` +
      `<workbookPr${wb.date1904 ? ' date1904="1"' : ''}/>` +
      (wb.workbookProtection !== undefined
        ? `<workbookProtection ${wb.workbookProtection}/>`
        : '') +
      (wb.bookViews === undefined
        ? `<bookViews><workbookView activeTab="${wb.activeTab ?? 0}"/></bookViews>`
        : wb.bookViews && `<bookViews>${wb.bookViews}</bookViews>`) +
      `<sheets>${sheetEls.join('')}${wb.extraSheets ?? ''}</sheets>` +
      (wb.externalReferences
        ? `<externalReferences>${wb.externalReferences}</externalReferences>`
        : '') +
      (wb.definedNames ? `<definedNames>${wb.definedNames}</definedNames>` : '') +
      `<calcPr ${wb.calcPr ?? 'calcId="191029"'}/>` +
      (wb.afterCalcPr ?? '') +
      '</workbook>',
    wb.macroEnabled ? MACRO_WORKBOOK : `${CT}.spreadsheetml.sheet.main+xml`
  );
  put('xl/_rels/workbook.xml.rels', rels(wbRels));
  put('xl/styles.xml', styles(wb), `${CT}.spreadsheetml.styles+xml`);
  put('xl/theme/theme1.xml', theme(wb.theme), `${CT}.theme+xml`);
  const packageRels: [string, string, string][] = [
    ['rId1', `${R}/officeDocument`, 'xl/workbook.xml'],
  ];
  if (wb.appProps) {
    const names = wb.sheets.map((sheet) => `<vt:lpstr>${esc(sheet.name)}</vt:lpstr>`);
    put(
      'docProps/app.xml',
      `${XML}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
        '<DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>' +
        '<HeadingPairs><vt:vector size="2" baseType="variant">' +
        `<vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>${names.length}</vt:i4></vt:variant>` +
        '</vt:vector></HeadingPairs>' +
        `<TitlesOfParts><vt:vector size="${names.length}" baseType="lpstr">${names.join('')}</vt:vector></TitlesOfParts>` +
        '<LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>' +
        '<HyperlinksChanged>false</HyperlinksChanged></Properties>',
      `${CT}.extended-properties+xml`
    );
    packageRels.push([
      'rId2',
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
      'docProps/app.xml',
    ]);
  }
  put('_rels/.rels', rels(packageRels));
  const defaults =
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Default Extension="png" ContentType="image/png"/>' +
    '<Default Extension="jpeg" ContentType="image/jpeg"/>';
  const ct =
    `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${defaults}` +
    overrides
      .map(([name, type]) => `<Override PartName="/${name}" ContentType="${type}"/>`)
      .join('') +
    '</Types>';
  const ordered = new Map<string, Uint8Array>([['/[Content_Types].xml', enc.encode(ct)]]);
  for (const name of [...parts.keys()].sort()) ordered.set(`/${name}`, parts.get(name)!);
  return writeZip(ordered);
}

// ---------------------------------------------------------------- cell helpers

/** One `x:c`. `v` is the stored value; `f` the formula text; `s` the cell format index. */
export function c(
  ref: string,
  v?: string | number | boolean,
  opts: { t?: string; f?: string; s?: number; is?: string } = {}
): string {
  const t =
    opts.t ??
    (typeof v === 'boolean' ? 'b' : typeof v === 'string' && !opts.f ? 'inlineStr' : undefined);
  const attrs = `r="${ref}"${opts.s ? ` s="${opts.s}"` : ''}${t && t !== 'n' ? ` t="${t}"` : ''}`;
  const f = opts.f !== undefined ? `<f>${esc(opts.f)}</f>` : '';
  if (t === 'inlineStr')
    return `<c ${attrs}><is><t xml:space="preserve">${esc(opts.is ?? String(v))}</t></is></c>`;
  const value =
    v === undefined ? '' : `<v>${esc(typeof v === 'boolean' ? (v ? '1' : '0') : String(v))}</v>`;
  return `<c ${attrs}>${f}${value}</c>`;
}

/** One `x:row` with its cells. */
export function row(r: number, cells: readonly string[], attrs = ''): string {
  return `<row r="${r}"${attrs ? ` ${attrs}` : ''}>${cells.join('')}</row>`;
}

/** Re-export so fixture modules import one place. */
export { NS_A, NS_XDR, R as REL_BASE };
