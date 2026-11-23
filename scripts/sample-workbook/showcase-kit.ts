// The building blocks of the demo's sample workbook (`showcase.ts`): a style table that hands
// out cell format indexes, a shared string table, cell and row helpers, and date serials.
//
// Same role as pptx-editor/scripts/sample-deck/showcase-kit.ts at 79a5d22d (EigenPal), for a
// workbook. It writes through the fixture builder (`scripts/fixtures/builder.ts`).

import { columnLetters } from '../../packages/core/src/store/index.ts';
import { esc } from '../fixtures/builder.ts';

/** The colors of the sample, as RGB hex. Green is the `.XLSX` label of the brand lockup. */
export const COLORS = {
  ink: '201F1E',
  muted: '6B6A68',
  green: '107C41',
  greenLight: 'DFF0E5',
  band: 'F3F2F1',
  white: 'FFFFFF',
  red: 'C00000',
  amber: 'B85C00',
  blue: '1F5FA8',
} as const;

/** A thin line on every side of a cell. */
export const THIN = { left: 'thin', right: 'thin', top: 'thin', bottom: 'thin' } as const;

/** `x:family` (ECMA-376 Part 1, 18.8.18) of the font families that the sample uses. */
const FONT_FAMILY: Readonly<Record<string, number>> = {
  Calibri: 2,
  Arial: 2,
  Cambria: 1,
  'Times New Roman': 1,
  'Courier New': 3,
};

/** One cell format, as the parts of `x:font`, `x:fill`, `x:border`, `x:numFmt`, and alignment. */
export interface CellStyle {
  readonly font?: {
    readonly name?: string;
    readonly size?: number;
    readonly bold?: boolean;
    readonly italic?: boolean;
    /** `x:u/@val`. */
    readonly underline?: 'single' | 'double';
    readonly strike?: boolean;
    /** RGB hex. The theme text color when absent. */
    readonly color?: string;
  };
  /** A solid fill: RGB hex, or a theme color index with a tint. */
  readonly fill?: string | { readonly theme: number; readonly tint?: number };
  /** The line style of each side that has one. */
  readonly border?: {
    readonly left?: string;
    readonly right?: string;
    readonly top?: string;
    readonly bottom?: string;
  };
  /** A built-in number format id, or a format code. */
  readonly numFmt?: number | string;
  readonly align?: {
    readonly horizontal?: string;
    readonly vertical?: string;
    readonly wrap?: boolean;
    readonly shrink?: boolean;
    readonly indent?: number;
    readonly rotation?: number;
  };
}

/** Collects fonts, fills, borders, number formats, and cell formats, each once. */
export class StyleTable {
  private readonly lists = {
    fonts: [] as string[],
    fills: [] as string[],
    borders: [] as string[],
  };
  private readonly formats: string[] = [];
  private readonly formatIds = new Map<string, number>();
  private readonly xfs: string[] = [];
  private readonly xfIndex = new Map<string, number>();

  /** The index of `xml` in a list that starts after `base` default items. */
  private put(list: 'fonts' | 'fills' | 'borders', xml: string, base: number): number {
    const items = this.lists[list];
    const at = items.indexOf(xml);
    if (at >= 0) return base + at;
    items.push(xml);
    return base + items.length - 1;
  }

  private font(f: NonNullable<CellStyle['font']>): number {
    const name = f.name ?? 'Calibri';
    const xml =
      '<font>' +
      (f.bold ? '<b/>' : '') +
      (f.italic ? '<i/>' : '') +
      (f.strike ? '<strike/>' : '') +
      (f.underline ? `<u${f.underline === 'double' ? ' val="double"' : ''}/>` : '') +
      `<sz val="${f.size ?? 11}"/>` +
      (f.color ? `<color rgb="FF${f.color}"/>` : '<color theme="1"/>') +
      `<name val="${esc(name)}"/><family val="${FONT_FAMILY[name] ?? 0}"/>` +
      (name === 'Calibri' ? '<scheme val="minor"/>' : '') +
      '</font>';
    return this.put('fonts', xml, 1);
  }

  private fill(fill: NonNullable<CellStyle['fill']>): number {
    const color =
      typeof fill === 'string'
        ? `rgb="FF${fill}"`
        : `theme="${fill.theme}"${fill.tint ? ` tint="${fill.tint}"` : ''}`;
    return this.put(
      'fills',
      `<fill><patternFill patternType="solid"><fgColor ${color}/><bgColor indexed="64"/></patternFill></fill>`,
      2
    );
  }

  private border(b: NonNullable<CellStyle['border']>): number {
    const side = (name: string, style?: string) =>
      style ? `<${name} style="${style}"><color rgb="FF${COLORS.ink}"/></${name}>` : `<${name}/>`;
    const xml = `<border>${side('left', b.left)}${side('right', b.right)}${side('top', b.top)}${side('bottom', b.bottom)}<diagonal/></border>`;
    return this.put('borders', xml, 1);
  }

  private numFmt(code: number | string): number {
    if (typeof code === 'number') return code;
    let id = this.formatIds.get(code);
    if (id === undefined) {
      id = 164 + this.formats.length;
      this.formatIds.set(code, id);
      this.formats.push(`<numFmt numFmtId="${id}" formatCode="${esc(code)}"/>`);
    }
    return id;
  }

  /** The `s` index of a cell with this style. The empty style is 0. */
  xf(style: CellStyle = {}): number {
    const fontId = style.font ? this.font(style.font) : 0;
    const fillId = style.fill ? this.fill(style.fill) : 0;
    const borderId = style.border ? this.border(style.border) : 0;
    const numFmtId = style.numFmt !== undefined ? this.numFmt(style.numFmt) : 0;
    const a = style.align;
    const alignment = a
      ? '<alignment' +
        (a.horizontal ? ` horizontal="${a.horizontal}"` : '') +
        (a.vertical ? ` vertical="${a.vertical}"` : '') +
        (a.rotation !== undefined ? ` textRotation="${a.rotation}"` : '') +
        (a.wrap ? ' wrapText="1"' : '') +
        (a.indent ? ` indent="${a.indent}"` : '') +
        (a.shrink ? ' shrinkToFit="1"' : '') +
        '/>'
      : '';
    const applied =
      (numFmtId ? ' applyNumberFormat="1"' : '') +
      (fontId ? ' applyFont="1"' : '') +
      (fillId ? ' applyFill="1"' : '') +
      (borderId ? ' applyBorder="1"' : '') +
      (alignment ? ' applyAlignment="1"' : '');
    const xml =
      `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"${applied}` +
      (alignment ? `>${alignment}</xf>` : '/>');
    if (xml === '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>') return 0;
    let index = this.xfIndex.get(xml);
    if (index === undefined) {
      this.xfs.push(xml);
      index = this.xfs.length;
      this.xfIndex.set(xml, index);
    }
    return index;
  }

  /** The styles part lists, for `FixtureWorkbook.styles`. */
  styles() {
    return {
      numFmts: this.formats,
      fonts: this.lists.fonts,
      fills: this.lists.fills,
      borders: this.lists.borders,
      cellXfs: this.xfs,
    };
  }
}

/** The shared string table: each text once, plain or as rich text runs. */
export class StringTable {
  readonly items: string[] = [];
  private readonly index = new Map<string, number>();

  private add(xml: string): number {
    let i = this.index.get(xml);
    if (i === undefined) {
      i = this.items.length;
      this.items.push(xml);
      this.index.set(xml, i);
    }
    return i;
  }

  text(value: string): number {
    return this.add(`<si><t xml:space="preserve">${esc(value)}</t></si>`);
  }

  /** Rich text: each run is its text and its `x:rPr` children (empty for the cell font). */
  rich(runs: readonly (readonly [string, string])[]): number {
    return this.add(
      `<si>${runs.map(([text, rPr]) => `<r>${rPr ? `<rPr>${rPr}</rPr>` : ''}<t xml:space="preserve">${esc(text)}</t></r>`).join('')}</si>`
    );
  }
}

/** A cell value: text, a number, a boolean, an error, or a formula with its cached result. */
export type CellValue =
  | string
  | number
  | boolean
  | { readonly error: string }
  | { readonly rich: number }
  | { readonly f: string; readonly v: number | string };

/** Writes `x:c` elements against one string table. */
export function cellWriter(strings: StringTable) {
  return (ref: string, value: CellValue | undefined, s = 0): string => {
    const style = s ? ` s="${s}"` : '';
    if (value === undefined) return `<c r="${ref}"${style}/>`;
    if (typeof value === 'string')
      return `<c r="${ref}"${style} t="s"><v>${strings.text(value)}</v></c>`;
    if (typeof value === 'number') return `<c r="${ref}"${style}><v>${value}</v></c>`;
    if (typeof value === 'boolean')
      return `<c r="${ref}"${style} t="b"><v>${value ? 1 : 0}</v></c>`;
    if ('error' in value) return `<c r="${ref}"${style} t="e"><v>${esc(value.error)}</v></c>`;
    if ('rich' in value) return `<c r="${ref}"${style} t="s"><v>${value.rich}</v></c>`;
    const t = typeof value.v === 'string' ? ' t="str"' : '';
    return `<c r="${ref}"${style}${t}><f>${esc(value.f)}</f><v>${esc(String(value.v))}</v></c>`;
  };
}

/** One `x:row`. `height` sets a custom height in points. */
export function sheetRow(
  r: number,
  cells: readonly string[],
  opts: { readonly height?: number; readonly hidden?: boolean } = {}
): string {
  const attrs =
    (opts.height !== undefined ? ` ht="${opts.height}" customHeight="1"` : '') +
    (opts.hidden ? ' hidden="1"' : '');
  return `<row r="${r}"${attrs}>${cells.join('')}</row>`;
}

/** `x:cols` from `[first, last, width, hidden?]` entries (1-based columns, width in characters). */
export function columns(list: readonly (readonly [number, number, number, boolean?])[]): string {
  return `<cols>${list
    .map(
      ([min, max, width, hidden]) =>
        `<col min="${min}" max="${max}" width="${width}" customWidth="1"${hidden ? ' hidden="1"' : ''}/>`
    )
    .join('')}</cols>`;
}

/** The 1900 date system serial of a calendar date (after 1 March 1900). */
export function dateSerial(year: number, month: number, day: number): number {
  return (Date.UTC(year, month - 1, day) - Date.UTC(1899, 11, 30)) / 86_400_000;
}

/**
 * A frozen pane of `rows` top rows and `cols` left columns. The selection is `active`, or the
 * first cell under the pane.
 */
export function frozenPane(rows: number, cols: number, active?: string): string {
  const top = `${columnLetters(cols + 1)}${rows + 1}`;
  const pane = rows && cols ? 'bottomRight' : rows ? 'bottomLeft' : cols ? 'topRight' : 'topLeft';
  return (
    `<pane${cols ? ` xSplit="${cols}"` : ''}${rows ? ` ySplit="${rows}"` : ''} topLeftCell="${top}" activePane="${pane}" state="frozen"/>` +
    `<selection pane="${pane}" activeCell="${active ?? top}" sqref="${active ?? top}"/>`
  );
}
