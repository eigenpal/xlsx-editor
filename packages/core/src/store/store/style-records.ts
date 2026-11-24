// The style records of a format edit: find or append the `x:numFmt`, `x:font`, `x:fill`,
// `x:border`, and `x:xf` records that a format patch needs (design D2, "Ops carry intent").
//
// The rules, each from a defect that DOCX Editor and PPTX Editor fixed on save:
// - Never change an existing record in place: other cells use it. A patch builds a new record
//   from a cell's record and the patch, reuses an existing record with the same canonical form
//   (`canonicalForm`: names, attributes, and children, without ids or prefixes), and appends
//   one only when none exists.
// - Write children in schema order (sml.xsd): `x:xf` holds `alignment`, `protection`, `extLst`
//   (`CT_Xf`); `x:border` holds `start`, `end`, `left`, `right`, `top`, `bottom`, `diagonal`,
//   `vertical`, `horizontal` (`CT_Border`); `x:patternFill` holds `fgColor`, `bgColor`. A font
//   (`CT_Font`, a choice) is written in one fixed order, so that the same font has one form.
// - Merge, never duplicate: records go into the first container of their list. A missing
//   container is created in its place in the `CT_Stylesheet` sequence.
// - Allocate unique ids: a new custom number format takes the lowest id from 164 up that no
//   `x:numFmt` of the part uses.
// - Write `count` from the content, as the save rules do (`package-write/save.ts`).
//
// Every list is read from the elements that exist, never from a `count` that the file states.
// A list that would pass `MAX_STYLE_RECORDS` refuses the edit.

import type { Attr, Element, Node, XmlPart } from '../package/tree.ts';
import { NS } from '../package/namespaces.ts';
import { attr, child, children, intAttr } from '../package/query.ts';
import { canonicalForm } from '../package-write/serialize.ts';
import { IdSource, plainAttr, smlElement, withAttr, withNamespacesOf } from './tree-edit.ts';
import { FIRST_CUSTOM_NUMFMT_ID, type FormatPatch, type PatchGroups } from './format-patch.ts';

/** The most records that one style list holds after an edit. */
export const MAX_STYLE_RECORDS = 64_000;

/** The children of `x:styleSheet`, in schema order (`CT_Stylesheet`). */
const STYLESHEET_ORDER = [
  'numFmts',
  'fonts',
  'fills',
  'borders',
  'cellStyleXfs',
  'cellXfs',
  'cellStyles',
  'dxfs',
  'tableStyles',
  'colors',
  'extLst',
] as const;

/** The one order in which a new font record writes its children. */
const FONT_ORDER = [
  'b',
  'i',
  'strike',
  'condense',
  'extend',
  'outline',
  'shadow',
  'u',
  'vertAlign',
  'sz',
  'color',
  'name',
  'family',
  'charset',
  'scheme',
] as const;

/** `CT_Border`, in schema order. */
const BORDER_ORDER = [
  'start',
  'end',
  'left',
  'right',
  'top',
  'bottom',
  'diagonal',
  'vertical',
  'horizontal',
] as const;

/** `CT_Xf`, in schema order. */
const XF_ORDER = ['alignment', 'protection', 'extLst'] as const;

/** Edge bits of a cell's border mask. */
export const EDGE_LEFT = 1;
export const EDGE_RIGHT = 2;
export const EDGE_TOP = 4;
export const EDGE_BOTTOM = 8;
export const EDGE_ALL = EDGE_LEFT | EDGE_RIGHT | EDGE_TOP | EDGE_BOTTOM;

type ListName = 'numFmts' | 'fonts' | 'fills' | 'borders' | 'cellXfs';
const ITEM: Readonly<Record<ListName, string>> = {
  numFmts: 'numFmt',
  fonts: 'font',
  fills: 'fill',
  borders: 'border',
  cellXfs: 'xf',
};

/**
 * The `numFmtId` of an `x:numFmt` record, read as `styleLists` reads it (`intAttr`): white
 * space collapsed and leading zeros allowed, so the store and the cascade see the same ids.
 */
function numFmtIdOf(el: Element): number | undefined {
  const id = intAttr(el, 'numFmtId');
  return id !== undefined && id >= 0 ? id : undefined;
}

/**
 * The `numFmtId` that a cell format or a differential format names. It is trimmed, so a padded
 * id stays in use: a reader that takes fewer ids only keeps more ids from a new code.
 */
const namedNumFmtId = (el: Element): number | undefined => index9(attr(el, 'numFmtId')?.trim());

/** A non-negative decimal integer of up to 9 digits, or undefined. */
function index9(raw: string | undefined): number | undefined {
  if (raw === undefined || raw.length === 0 || raw.length > 9) return undefined;
  let n = 0;
  for (let i = 0; i < raw.length; i += 1) {
    const c = raw.charCodeAt(i);
    if (c < 48 || c > 57) return undefined;
    n = n * 10 + (c - 48);
  }
  return n;
}

/** A deep copy of a node with new ids, so that no two nodes of a part share an id. */
function cloneNode(ids: IdSource, node: Node): Node {
  if (node.type === 'text') return { ...node, id: ids.id() };
  return { ...node, id: ids.id(), children: node.children.map((c) => cloneNode(ids, c)) };
}
const clone = (ids: IdSource, el: Element): Element => cloneNode(ids, el) as Element;

/** One list of the styles part: the records that exist, and the ones this edit appends. */
class RecordList {
  readonly added: Element[] = [];
  #keys: Map<string, number> | null = null;
  constructor(readonly existing: readonly Element[]) {}

  get length(): number {
    return this.existing.length + this.added.length;
  }

  at(i: number | undefined): Element | undefined {
    if (i === undefined) return undefined;
    return i < this.existing.length ? this.existing[i] : this.added[i - this.existing.length];
  }

  /** The index of a record with the same canonical form, appending `el` when none exists. */
  intern(el: Element): number | undefined {
    if (!this.#keys) {
      this.#keys = new Map();
      for (let i = 0; i < this.existing.length; i += 1) {
        const key = canonicalForm(this.existing[i]!);
        if (!this.#keys.has(key)) this.#keys.set(key, i);
      }
    }
    const key = canonicalForm(el);
    const hit = this.#keys.get(key);
    if (hit !== undefined) return hit;
    if (this.length >= MAX_STYLE_RECORDS) return undefined;
    this.added.push(el);
    this.#keys.set(key, this.length - 1);
    return this.length - 1;
  }

  /** Append a record that is known to be new. */
  push(el: Element): number | undefined {
    if (this.length >= MAX_STYLE_RECORDS) return undefined;
    this.added.push(el);
    return this.length - 1;
  }
}

/** Why a style edit stopped: a list is full. */
export class StyleLimit extends Error {}

/** The records of one styles part, and the records a format edit appends to it. */
export class StyleRecords {
  readonly #part: XmlPart;
  readonly #ids: IdSource;
  readonly #lists: Record<ListName, RecordList>;
  /** Custom number format code to id. The first definition of a code wins. */
  readonly #codes = new Map<string, number>();
  readonly #usedIds = new Set<number>();
  /** The cell style formats (`x:cellStyleXfs`) that `xfId` names. */
  #styleXfs: readonly Element[];
  /** The `x:cellStyleXfs` and `x:cellStyles` that a bootstrapped part gets, when it has none. */
  #styleSeed: Element[] = [];
  #nextCustom = FIRST_CUSTOM_NUMFMT_ID;

  constructor(part: XmlPart) {
    this.#part = part;
    this.#ids = new IdSource(part);
    const root = part.root;
    const list = (name: ListName) =>
      new RecordList(children(child(root, `x:${name}`), `x:${ITEM[name]}`));
    this.#lists = {
      numFmts: list('numFmts'),
      fonts: list('fonts'),
      fills: list('fills'),
      borders: list('borders'),
      cellXfs: list('cellXfs'),
    };
    this.#styleXfs = children(child(root, 'x:cellStyleXfs'), 'x:xf');
    this.#bootstrap();
    // The ids as the cascade reads them (`styleLists`): the first record of an id is the one
    // that cells show. Only a custom id gives a code to reuse: a record that redefines a
    // built-in id (ECMA-376 Part 1, 18.8.30) is not read the same way by every consumer. A later
    // record with the same id gives no code to reuse.
    for (const el of this.#lists.numFmts.existing) {
      const id = numFmtIdOf(el);
      const code = attr(el, 'formatCode');
      if (id === undefined || this.#usedIds.has(id)) continue;
      this.#usedIds.add(id);
      if (code !== undefined && id >= FIRST_CUSTOM_NUMFMT_ID && !this.#codes.has(code))
        this.#codes.set(code, id);
    }
    // An id that a cell format, a cell style format, or a differential format names is in use
    // also when no `x:numFmt` record defines it: a new code never takes it, or the cells that
    // name it would change with the edit. The lists are the records that the part holds.
    const named = [
      ...this.#lists.cellXfs.existing,
      ...this.#styleXfs,
      ...children(child(root, 'x:dxfs'), 'x:dxf').flatMap((dxf) => children(dxf, 'x:numFmt')),
    ];
    for (const el of named) {
      const id = namedNumFmtId(el);
      if (id !== undefined && id >= FIRST_CUSTOM_NUMFMT_ID) this.#usedIds.add(id);
    }
  }

  /**
   * A styles part with no cell format gets the default records first: one font, the two reserved
   * fills, one empty border, and the default cell format at index 0. Without them, a new
   * format at index 0 would become the format of every cell that has no `@s`.
   */
  #bootstrap(): void {
    const l = this.#lists;
    if (l.cellXfs.length > 0) return;
    const add = (list: ListName, el: Element) => l[list].push(el);
    if (l.fonts.length === 0)
      add(
        'fonts',
        this.#el(
          'font',
          'fonts',
          [],
          [
            this.#el('sz', 'generic', [plainAttr('val', '11')]),
            this.#el('name', 'generic', [plainAttr('val', 'Calibri')]),
          ]
        )
      );
    for (const type of ['none', 'gray125'].slice(l.fills.length))
      add(
        'fills',
        this.#el(
          'fill',
          'fills',
          [],
          [this.#el('patternFill', 'generic', [plainAttr('patternType', type)])]
        )
      );
    if (l.borders.length === 0)
      add(
        'borders',
        this.#el(
          'border',
          'borders',
          [],
          ['left', 'right', 'top', 'bottom', 'diagonal'].map((e) => this.#el(e, 'generic'))
        )
      );
    // The default cell format names the default cell style (`@xfId`), which the part gets with
    // its `Normal` built-in style when it has none (ECMA-376 Part 1, 18.8.9 and 18.8.8).
    const zero = ['numFmtId', 'fontId', 'fillId', 'borderId'].map((a) => plainAttr(a, '0'));
    if (this.#styleXfs.length === 0 && !child(this.#part.root, 'x:cellStyles')) {
      const styleXf = this.#el('xf', 'cellStyleXfs', zero);
      this.#styleXfs = [styleXf];
      this.#styleSeed = [
        this.#el('cellStyleXfs', 'styleSheet', [plainAttr('count', '1')], [styleXf]),
        this.#el(
          'cellStyles',
          'styleSheet',
          [plainAttr('count', '1')],
          [
            this.#el('cellStyle', 'generic', [
              plainAttr('name', 'Normal'),
              plainAttr('xfId', '0'),
              plainAttr('builtinId', '0'),
            ]),
          ]
        ),
      ];
    }
    const xfId = this.#styleXfs.length > 0 ? [plainAttr('xfId', '0')] : [];
    add('cellXfs', this.#el('xf', 'cellXfs', [...zero, ...xfId]));
  }

  /** The number of cell formats (`x:cellXfs`), with the ones that this edit appended. */
  get cellFormats(): number {
    return this.#lists.cellXfs.length;
  }

  /** Whether the edit appended a record. */
  get changed(): boolean {
    return Object.values(this.#lists).some((l) => l.added.length > 0);
  }

  /**
   * The cell format index for a cell whose format is `base`, with `patch` applied. `edges` is
   * the border mask of the cell (`EDGE_*`), read only when the patch draws borders.
   * Throws `StyleLimit` when a list is full.
   */
  xfFor(base: number, patch: FormatPatch, groups: PatchGroups, edges: number): number {
    const xfs = this.#lists.cellXfs;
    const baseXf = xfs.at(base < xfs.length ? base : 0);
    const attrs: Attr[] = [...(baseXf?.attrs ?? [])];
    const set = (name: string, value: string | undefined) => {
      const at = attrs.findIndex((a) => a.ns === '' && a.local === name);
      if (value === undefined) {
        if (at >= 0) attrs.splice(at, 1);
      } else if (at < 0) attrs.push(plainAttr(name, value));
      else attrs[at] = { ...attrs[at]!, value };
    };
    const ref = (name: string) => index9(attr(baseXf, name));
    // `apply*` says that the cell format sets a record of its own instead of its cell style's
    // (Part 1, 18.8.45). It is written when the record differs from the style's, and removed
    // when it is the style's again, so that removing a property finds the original format.
    const style = this.#styleXfs[ref('xfId') ?? 0];
    const own = (id: string, apply: string, value: number) => {
      set(id, String(value));
      set(apply, value === (index9(attr(style, id)) ?? 0) ? undefined : '1');
    };
    if (groups.font) {
      const font = this.#font(
        this.#lists.fonts.at(ref('fontId')) ?? this.#lists.fonts.at(0),
        patch
      );
      own('fontId', 'applyFont', this.#intern('fonts', font));
    }
    if (groups.fill && patch.fill)
      own('fillId', 'applyFill', this.#intern('fills', this.#fill(patch.fill)));
    if (groups.border && patch.borders && (patch.borders.edges === 'none' || edges !== 0)) {
      const base = this.#lists.borders.at(ref('borderId')) ?? this.#lists.borders.at(0);
      own(
        'borderId',
        'applyBorder',
        this.#intern('borders', this.#border(base, patch.borders, edges))
      );
    }
    if (groups.numFmt && patch.numberFormat)
      own('numFmtId', 'applyNumberFormat', this.#numFmt(patch.numberFormat));
    const kids = new Map<string, Element>();
    const other: Element[] = [];
    for (const c of baseXf?.children ?? []) {
      if (c.type !== 'element') continue;
      if (c.ns === NS.x && (XF_ORDER as readonly string[]).includes(c.local) && !kids.has(c.local))
        kids.set(c.local, clone(this.#ids, c));
      else other.push(clone(this.#ids, c));
    }
    if (groups.alignment) {
      const alignment = this.#alignment(kids.get('alignment'), patch);
      if (alignment) kids.set('alignment', alignment);
      else kids.delete('alignment');
      set('applyAlignment', alignment ? '1' : undefined);
    }
    const ordered = XF_ORDER.flatMap((name) => (kids.has(name) ? [kids.get(name)!] : []));
    const xf = smlElement(this.#ids, 'xf', 'cellXfs', attrs, [...ordered, ...other]);
    return this.#intern('cellXfs', withNamespacesOf(baseXf, xf));
  }

  /**
   * The cell format index of the default cell format (index 0) with the given protection
   * (`x:protection`, ECMA-376 Part 1, 18.8.33): a clear of formats on a protected sheet uses it,
   * so the clear keeps the protection of each cell. Throws `StyleLimit` when the list is full.
   */
  defaultWithProtection(locked: boolean, hidden: boolean): number {
    const zero = this.#lists.cellXfs.at(0);
    const attrs: Attr[] = (zero?.attrs ?? []).filter(
      (a) => !(a.ns === '' && a.local === 'applyProtection')
    );
    attrs.push(plainAttr('applyProtection', '1'));
    // `@locked` is true by default, and `@hidden` false (`CT_CellProtection`).
    const protection = this.#el('protection', 'generic', [
      ...(locked ? [] : [plainAttr('locked', '0')]),
      ...(hidden ? [plainAttr('hidden', '1')] : []),
    ]);
    const kids: Element[] = [];
    let placed = false;
    for (const c of zero?.children ?? []) {
      if (c.type !== 'element') continue;
      if (c.ns === NS.x && c.local === 'protection') continue;
      // `x:protection` follows `x:alignment` and precedes `x:extLst` (`CT_Xf`).
      if (!placed && c.ns === NS.x && c.local === 'extLst') {
        kids.push(protection);
        placed = true;
      }
      kids.push(clone(this.#ids, c));
    }
    if (!placed) kids.push(protection);
    const xf = smlElement(this.#ids, 'xf', 'cellXfs', attrs, kids);
    return this.#intern('cellXfs', withNamespacesOf(zero, xf));
  }

  #intern(list: ListName, el: Element): number {
    const i = this.#lists[list].intern(el);
    if (i === undefined)
      throw new StyleLimit(`the ${list} list holds ${MAX_STYLE_RECORDS} records`);
    return i;
  }

  #el(
    local: string,
    parent: Parameters<typeof smlElement>[2],
    attrs: Attr[] = [],
    kids: Node[] = []
  ) {
    return smlElement(this.#ids, local, parent, attrs, kids);
  }

  /**
   * A `CT_Color` element (ECMA-376 Part 1, 18.8.19): `@rgb` for `RRGGBB`, `@auto` for `auto`,
   * and `indexed="64"`, the system foreground, without a color.
   */
  #color(local: string, rgb: string | undefined): Element {
    return this.#el(local, 'generic', [
      rgb === undefined
        ? plainAttr('indexed', '64')
        : rgb === 'auto'
          ? plainAttr('auto', '1')
          : plainAttr('rgb', `FF${rgb.toUpperCase()}`),
    ]);
  }

  #font(base: Element | undefined, patch: FormatPatch): Element {
    const kept = new Map<string, Element>();
    const other: Element[] = [];
    for (const c of base?.children ?? []) {
      if (c.type !== 'element') continue;
      if (c.ns === NS.x && (FONT_ORDER as readonly string[]).includes(c.local)) {
        if (!kept.has(c.local)) kept.set(c.local, clone(this.#ids, c));
      } else other.push(clone(this.#ids, c));
    }
    const flag = (local: string, on: boolean | undefined) => {
      if (on === undefined) return;
      if (on) kept.set(local, this.#el(local, 'generic'));
      else kept.delete(local);
    };
    flag('b', patch.bold);
    flag('i', patch.italic);
    flag('u', patch.underline);
    if (patch.fontSize !== undefined)
      kept.set('sz', this.#el('sz', 'generic', [plainAttr('val', String(patch.fontSize))]));
    if (patch.fontColor !== undefined) kept.set('color', this.#color('color', patch.fontColor));
    if (patch.fontName !== undefined) {
      kept.set('name', this.#el('name', 'generic', [plainAttr('val', patch.fontName)]));
      // A theme scheme overrides the name (Part 1, 18.8.35), so a named font drops it.
      kept.delete('scheme');
    }
    const ordered = FONT_ORDER.flatMap((name) => (kept.has(name) ? [kept.get(name)!] : []));
    const font = this.#el('font', 'fonts', [...(base?.attrs ?? [])], [...ordered, ...other]);
    return withNamespacesOf(base, font);
  }

  #fill(fill: NonNullable<FormatPatch['fill']>): Element {
    const pattern =
      fill.type === 'none'
        ? this.#el('patternFill', 'generic', [plainAttr('patternType', 'none')])
        : this.#el(
            'patternFill',
            'generic',
            [plainAttr('patternType', 'solid')],
            [this.#color('fgColor', fill.color), this.#color('bgColor', undefined)]
          );
    return this.#el('fill', 'fills', [], [pattern]);
  }

  #border(
    base: Element | undefined,
    borders: NonNullable<FormatPatch['borders']>,
    edges: number
  ): Element {
    const kept = new Map<string, Element>();
    const other: Element[] = [];
    for (const c of base?.children ?? []) {
      if (c.type !== 'element') continue;
      if (c.ns === NS.x && (BORDER_ORDER as readonly string[]).includes(c.local)) {
        if (!kept.has(c.local)) kept.set(c.local, clone(this.#ids, c));
      } else other.push(clone(this.#ids, c));
    }
    let attrs = [...(base?.attrs ?? [])];
    const line = (local: string) =>
      this.#el(
        local,
        'generic',
        [plainAttr('style', borders.style ?? 'thin')],
        [this.#color('color', borders.color)]
      );
    if (borders.edges === 'none') {
      kept.clear();
      attrs = attrs.filter(
        (a) => !(a.ns === '' && (a.local === 'diagonalUp' || a.local === 'diagonalDown'))
      );
    } else {
      const sides: [number, string, string][] = [
        [EDGE_LEFT, 'left', 'start'],
        [EDGE_RIGHT, 'right', 'end'],
        [EDGE_TOP, 'top', ''],
        [EDGE_BOTTOM, 'bottom', ''],
      ];
      for (const [bit, local, logical] of sides) {
        if (!(edges & bit)) continue;
        kept.set(local, line(local));
        // `start` and `end` stand for left and right only when those are absent.
        if (logical) kept.delete(logical);
      }
    }
    for (const local of ['left', 'right', 'top', 'bottom', 'diagonal'])
      if (
        !kept.has(local) &&
        !(local === 'left' && kept.has('start')) &&
        !(local === 'right' && kept.has('end'))
      )
        kept.set(local, this.#el(local, 'generic'));
    const ordered = BORDER_ORDER.flatMap((name) => (kept.has(name) ? [kept.get(name)!] : []));
    return withNamespacesOf(base, this.#el('border', 'borders', attrs, [...ordered, ...other]));
  }

  #numFmt(format: NonNullable<FormatPatch['numberFormat']>): number {
    if (format.id !== undefined) return format.id;
    const hit = this.#codes.get(format.code);
    if (hit !== undefined) return hit;
    while (this.#usedIds.has(this.#nextCustom)) this.#nextCustom += 1;
    const id = this.#nextCustom;
    const el = this.#el('numFmt', 'numFmts', [
      plainAttr('numFmtId', String(id)),
      plainAttr('formatCode', format.code),
    ]);
    if (this.#lists.numFmts.push(el) === undefined)
      throw new StyleLimit(`the numFmts list holds ${MAX_STYLE_RECORDS} records`);
    this.#usedIds.add(id);
    this.#codes.set(format.code, id);
    return id;
  }

  #alignment(base: Element | undefined, patch: FormatPatch): Element | undefined {
    let el = base ?? this.#el('alignment', 'generic');
    if (patch.horizontal !== undefined)
      el = withAttr(
        el,
        'horizontal',
        patch.horizontal === 'general' ? undefined : patch.horizontal
      );
    if (patch.vertical !== undefined)
      el = withAttr(el, 'vertical', patch.vertical === 'bottom' ? undefined : patch.vertical);
    if (patch.wrapText !== undefined)
      el = withAttr(el, 'wrapText', patch.wrapText ? '1' : undefined);
    return el.attrs.length === 0 && el.children.length === 0 ? undefined : el;
  }

  /** The styles part with the appended records, or the same part when nothing was appended. */
  part(): XmlPart {
    if (!this.changed) return this.#part;
    let root = this.#part.root;
    for (const seed of this.#styleSeed)
      root = { ...root, children: insertInOrder(root.children, seed) };
    for (const name of ['numFmts', 'fonts', 'fills', 'borders', 'cellXfs'] as const) {
      const added = this.#lists[name].added;
      if (added.length === 0) continue;
      const container = child(root, `x:${name}`);
      if (container) {
        const kids = [...container.children, ...added];
        const updated = withAttr({ ...container, children: kids }, 'count', String(countOf(kids)));
        root = { ...root, children: root.children.map((c) => (c === container ? updated : c)) };
      } else {
        const created = this.#el(
          name,
          'styleSheet',
          [plainAttr('count', String(added.length))],
          added
        );
        root = { ...root, children: insertInOrder(root.children, created) };
      }
    }
    return { partName: this.#part.partName, root, nextId: this.#ids.next };
  }
}

/** The element children that a `count` counts: all but `x:extLst`. */
function countOf(kids: readonly Node[]): number {
  let n = 0;
  for (const c of kids) if (c.type === 'element' && c.local !== 'extLst') n += 1;
  return n;
}

const rank = (node: Node): number =>
  node.type === 'element' && node.ns === NS.x
    ? (STYLESHEET_ORDER as readonly string[]).indexOf(node.local)
    : -1;

/** `kids` with a new container before the first known child that follows it in the schema. */
function insertInOrder(kids: readonly Node[], el: Element): Node[] {
  const own = rank(el);
  const at = kids.findIndex((n) => rank(n) > own);
  const out = [...kids];
  out.splice(at < 0 ? out.length : at, 0, el);
  return out;
}
