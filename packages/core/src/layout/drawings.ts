// Drawings over the grid (design D6, ECMA-376 Part 1, 20.5): the anchors of a worksheet's
// drawing part, placed on the column and row offsets of the sheet, and the objects they hold.
//
// The anchor types are in `dml-spreadsheetDrawing.xsd`:
// - `xdr:twoCellAnchor` (`CT_TwoCellAnchor`): from `xdr:from` to `xdr:to`, each a `CT_Marker`:
//   a 0-based cell (`xdr:col`, `xdr:row`) and an offset in EMU (`xdr:colOff`, `xdr:rowOff`).
//   The engine places the anchor from its markers and does not read `@editAs` (`ST_EditAs`).
// - `xdr:oneCellAnchor` (`CT_OneCellAnchor`): from `xdr:from`, with the size of `xdr:ext`
//   (`a:CT_PositiveSize2D`).
// - `xdr:absoluteAnchor` (`CT_AbsoluteAnchor`): at `xdr:pos` (`a:CT_Point2D`), with the size of
//   `xdr:ext`.
//
// Each anchor holds one object of `EG_ObjectChoices`: `xdr:sp`, `xdr:grpSp`, `xdr:graphicFrame`,
// `xdr:cxnSp`, `xdr:pic`, or `xdr:contentPart`.
//
// An offset larger than its cell stays inside the cell: a hidden column holds no offset.
//
// Objects: a picture (`xdr:pic`) names its image with `a:blip/@r:embed`, an internal relationship
// of the drawing part. An external image (`r:link`, or an embed relationship with
// `TargetMode="External"`) is never loaded (AGENTS.md, "Security"); it lays out as an unsupported
// object. So does a picture whose image part is not in the package. A shape (`xdr:sp`, `xdr:cxnSp`)
// evaluates `a:prstGeom` or `a:custGeom` through `@eigenpal/ooxml-core`, with the solid fill and
// line of `xdr:spPr`, or the colors of its `xdr:style` references. Groups, chart frames, shape
// text, and gradient or pattern fills are not laid out yet (the approximation "Drawing objects").
//
// The anchors of a part are read once per part revision; the rectangles depend on the sheet
// geometry. Nothing loops over a number that the file states: at most `MAX_ANCHORS` are read.

import {
  customGeometryDefinition,
  evaluateGeometry,
  evaluatePresetGeometry,
  parseHex,
  resolveDrawingColor,
  type DrawingGeometryElement,
  type EvaluatedGeometry,
  type Rgba,
} from '@eigenpal/ooxml-core';
import type { Element } from '../store/package/tree.ts';
import { attr, child, elements, textOf } from '../store/package/query.ts';
import { NS, REL } from '../store/package/namespaces.ts';
import {
  contentTypeOf,
  partOf,
  relationshipById,
  relTarget,
  xmlPartOf,
  type XlsxPackage,
} from '../store/package/package.ts';
import { MAX_COLS, MAX_ROWS } from '../store/cell-ref.ts';
import type { WorkbookTheme } from '../cascade/index.ts';
import type { SheetGeometry } from './grid-geometry.ts';
import type { Rect } from './cell-text.ts';
import { EMU_PER_POINT } from './units.ts';

/** At most this many anchors are read from one drawing part. */
export const MAX_ANCHORS = 10_000;
/**
 * The largest coordinate of `ST_Coordinate` (`ST_CoordinateUnqualified`, dml-main.xsd), in EMU.
 * The schema minimum is -27273042329600; the engine uses the negative of the maximum.
 */
const MAX_COORDINATE = 27_273_042_316_900;

interface AnchorPoint {
  readonly col: number;
  readonly colOff: number;
  readonly row: number;
  readonly rowOff: number;
}

type Placement =
  | { readonly kind: 'twoCell'; readonly from: AnchorPoint; readonly to: AnchorPoint }
  | {
      readonly kind: 'oneCell';
      readonly from: AnchorPoint;
      readonly cx: number;
      readonly cy: number;
    }
  | {
      readonly kind: 'absolute';
      readonly x: number;
      readonly y: number;
      readonly cx: number;
      readonly cy: number;
    };

interface DrawingCommon {
  /** The position of the anchor in the drawing part. */
  readonly index: number;
  readonly name: string;
  /** In points, from the top-left corner of the sheet. */
  readonly rect: Rect;
  /** Clockwise degrees (`a:xfrm/@rot`), in [0, 360). */
  readonly rotation: number;
  readonly flipH: boolean;
  readonly flipV: boolean;
}

/** A picture: the painter reads the bytes of `partName`, a part of the package. */
export interface PictureLayout extends DrawingCommon {
  readonly kind: 'picture';
  readonly partName: string;
  readonly contentType: string | undefined;
}

/** A shape: paths in points relative to the top-left corner of `rect`. */
export interface ShapeLayout extends DrawingCommon {
  readonly kind: 'shape';
  readonly geometry: EvaluatedGeometry;
  readonly fill: Rgba | undefined;
  readonly line: { readonly color: Rgba; readonly width: number } | undefined;
}

/**
 * An object that layout places but does not draw: a group, a chart frame, an external image, or
 * a picture without its image part.
 */
export interface UnsupportedDrawing extends DrawingCommon {
  readonly kind: 'unsupported';
  /**
   * The local name of the object element; `externalImage` for a picture that links an external
   * image, and `missingImage` for a picture whose embed names no image part of the package.
   */
  readonly element: string;
}

export type DrawingLayout = PictureLayout | ShapeLayout | UnsupportedDrawing;

interface ParsedAnchor {
  readonly index: number;
  readonly placement: Placement;
  readonly object: Element | undefined;
}

const numberText = (el: Element | undefined, min: number, max: number): number | undefined => {
  if (!el) return undefined;
  const raw = textOf(el).trim();
  if (raw.length === 0 || raw.length > 20) return undefined;
  const n = Number(raw);
  return Number.isInteger(n) && n >= min && n <= max ? n : undefined;
};

const numberAttr = (el: Element | undefined, name: string, min: number, max: number) => {
  const raw = attr(el, name);
  if (raw === undefined || raw.length === 0 || raw.length > 20) return undefined;
  const n = Number(raw);
  return Number.isFinite(n) && n >= min && n <= max ? n : undefined;
};

function point(el: Element | undefined): AnchorPoint | undefined {
  const col = numberText(child(el, 'xdr:col'), 0, MAX_COLS - 1);
  const row = numberText(child(el, 'xdr:row'), 0, MAX_ROWS - 1);
  if (col === undefined || row === undefined) return undefined;
  return {
    col,
    row,
    colOff: numberText(child(el, 'xdr:colOff'), -MAX_COORDINATE, MAX_COORDINATE) ?? 0,
    rowOff: numberText(child(el, 'xdr:rowOff'), -MAX_COORDINATE, MAX_COORDINATE) ?? 0,
  };
}

function extent(el: Element | undefined): { cx: number; cy: number } | undefined {
  const cx = numberAttr(el, 'cx', 0, MAX_COORDINATE);
  const cy = numberAttr(el, 'cy', 0, MAX_COORDINATE);
  return cx === undefined || cy === undefined ? undefined : { cx, cy };
}

function placementOf(anchor: Element): Placement | undefined {
  if (anchor.local === 'twoCellAnchor') {
    const from = point(child(anchor, 'xdr:from'));
    const to = point(child(anchor, 'xdr:to'));
    return from && to ? { kind: 'twoCell', from, to } : undefined;
  }
  const ext = extent(child(anchor, 'xdr:ext'));
  if (!ext) return undefined;
  if (anchor.local === 'oneCellAnchor') {
    const from = point(child(anchor, 'xdr:from'));
    return from ? { kind: 'oneCell', from, ...ext } : undefined;
  }
  const pos = child(anchor, 'xdr:pos');
  const x = numberAttr(pos, 'x', -MAX_COORDINATE, MAX_COORDINATE);
  const y = numberAttr(pos, 'y', -MAX_COORDINATE, MAX_COORDINATE);
  return x === undefined || y === undefined ? undefined : { kind: 'absolute', x, y, ...ext };
}

const OBJECTS = new Set(['sp', 'pic', 'cxnSp', 'grpSp', 'graphicFrame', 'contentPart']);
const ANCHORS = new Set(['twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor']);

/** The anchors of a drawing part root, with `mc:AlternateContent` read through its fallback. */
function anchorsOf(root: Element): Element[] {
  const out: Element[] = [];
  for (const el of elements(root)) {
    if (out.length >= MAX_ANCHORS) break;
    if (el.ns === NS.xdr && ANCHORS.has(el.local)) out.push(el);
    else if (el.ns === NS.mc && el.local === 'AlternateContent') {
      const fallback = child(el, 'mc:Fallback');
      for (const inner of elements(fallback))
        if (inner.ns === NS.xdr && ANCHORS.has(inner.local) && out.length < MAX_ANCHORS)
          out.push(inner);
    }
  }
  return out;
}

const parsedCache = new WeakMap<Element, readonly ParsedAnchor[]>();

function parsedAnchors(root: Element): readonly ParsedAnchor[] {
  const hit = parsedCache.get(root);
  if (hit) return hit;
  const out: ParsedAnchor[] = [];
  anchorsOf(root).forEach((anchor, index) => {
    const placement = placementOf(anchor);
    if (!placement) return;
    const object = elements(anchor).find((el) => el.ns === NS.xdr && OBJECTS.has(el.local));
    out.push({ index, placement, object });
  });
  parsedCache.set(root, out);
  return out;
}

/** The drawing part of a worksheet (`x:drawing/@r:id`), or undefined. */
export function drawingPartOf(
  pkg: XlsxPackage,
  sheetPart: string,
  worksheet: Element
): string | undefined {
  const id = attr(child(worksheet, 'x:drawing'), 'r:id');
  if (!id) return undefined;
  const rel = relationshipById(pkg, sheetPart, id);
  return rel && !rel.external && rel.type === REL.drawing ? rel.target : undefined;
}

const pt = (emu: number) => emu / EMU_PER_POINT;

/** Degrees in [0, 360). */
const normalizedDegrees = (deg: number): number => {
  const d = deg % 360;
  return d < 0 ? d + 360 : d === 0 ? 0 : d;
};

/**
 * The area that a drawing can paint, in the points of its `rect`: the bounding box of the
 * rotated rectangle, grown by half the line width of a shape. A straight connector has a zero
 * width or height; its bounds still have the width of its line.
 */
export function drawingBounds(d: DrawingLayout): Rect {
  const rad = (d.rotation * Math.PI) / 180;
  const cos = Math.abs(Math.cos(rad));
  const sin = Math.abs(Math.sin(rad));
  const w = d.rect.width * cos + d.rect.height * sin;
  const h = d.rect.width * sin + d.rect.height * cos;
  const grow = d.kind === 'shape' && d.line ? d.line.width / 2 : 0;
  const cx = d.rect.x + d.rect.width / 2;
  const cy = d.rect.y + d.rect.height / 2;
  return { x: cx - w / 2 - grow, y: cy - h / 2 - grow, width: w + 2 * grow, height: h + 2 * grow };
}

function cellPoint(geometry: SheetGeometry, p: AnchorPoint): { x: number; y: number } {
  const clamp = (offset: number, size: number) => Math.min(Math.max(0, offset), size);
  return {
    x: geometry.columns.offsetOf(p.col) + clamp(pt(p.colOff), geometry.columns.sizeOf(p.col)),
    y: geometry.rows.offsetOf(p.row) + clamp(pt(p.rowOff), geometry.rows.sizeOf(p.row)),
  };
}

/** The rectangle of a placement on the grid, in points. */
function rectOf(geometry: SheetGeometry, placement: Placement): Rect {
  if (placement.kind === 'absolute')
    return {
      x: pt(placement.x),
      y: pt(placement.y),
      width: pt(placement.cx),
      height: pt(placement.cy),
    };
  const from = cellPoint(geometry, placement.from);
  if (placement.kind === 'oneCell')
    return { x: from.x, y: from.y, width: pt(placement.cx), height: pt(placement.cy) };
  const to = cellPoint(geometry, placement.to);
  return {
    x: Math.min(from.x, to.x),
    y: Math.min(from.y, to.y),
    width: Math.abs(to.x - from.x),
    height: Math.abs(to.y - from.y),
  };
}

/**
 * The palette index of a scheme color name (`a:schemeClr/@val`, `ST_SchemeColorVal`), in theme
 * index order (quick reference). `phClr` has no entry.
 */
const SCHEME_INDEX: Readonly<Record<string, number>> = Object.freeze({
  lt1: 0,
  bg1: 0,
  dk1: 1,
  tx1: 1,
  lt2: 2,
  bg2: 2,
  dk2: 3,
  tx2: 3,
  accent1: 4,
  accent2: 5,
  accent3: 6,
  accent4: 7,
  accent5: 8,
  accent6: 9,
  hlink: 10,
  folHlink: 11,
});

/** The color of the first DrawingML color model child of `el` (`EG_ColorChoice`, dml-main.xsd). */
function colorOf(el: Element | undefined, theme: WorkbookTheme): Rgba | undefined {
  const model = elements(el).find((c) => c.ns === NS.a);
  if (!model) return undefined;
  return resolveDrawingColor(
    {
      model: model.local,
      readAttribute: (name) => attr(model, name),
      transforms: elements(model)
        .filter((t) => t.ns === NS.a)
        .map((t) => {
          const val = attr(t, 'val');
          return val === undefined ? { name: t.local } : { name: t.local, val };
        }),
    },
    (name) => {
      const i = Object.hasOwn(SCHEME_INDEX, name) ? SCHEME_INDEX[name] : undefined;
      return i === undefined ? undefined : parseHex(theme.palette[i]);
    }
  );
}

/** An adapter from the canonical tree to the geometry reader of `@eigenpal/ooxml-core`. */
function geometryElement(el: Element): DrawingGeometryElement {
  return {
    namespaceUri: el.ns,
    localName: el.local,
    attribute: (name) => attr(el, name),
    childElements: () => elements(el).map(geometryElement),
  };
}

/** The `a:gd` (`CT_GeomGuide`) entries of `a:prstGeom/a:avLst` whose `@fmla` is `val n`. */
function adjustValues(prst: Element): Map<string, number> {
  const out = new Map<string, number>();
  for (const gd of elements(child(prst, 'a:avLst'))) {
    const name = attr(gd, 'name');
    const fmla = attr(gd, 'fmla') ?? '';
    if (!name || !fmla.startsWith('val ') || fmla.length > 32) continue;
    const v = Number(fmla.slice(4));
    if (Number.isFinite(v)) out.set(name, v);
  }
  return out;
}

/**
 * A shape (`xdr:sp`, `CT_Shape`, or `xdr:cxnSp`, `CT_Connector`): the geometry of `a:prstGeom`
 * (`CT_PresetGeometry2D`, `@prst` of `ST_ShapeType`) or `a:custGeom`, and the solid fill and line
 * (`a:ln`, `CT_LineProperties`) of `xdr:spPr` (`CT_ShapeProperties`), else the `a:fillRef` and
 * `a:lnRef` colors of `xdr:style` (`CT_ShapeStyle`).
 */
function shapeOf(
  object: Element,
  common: DrawingCommon,
  theme: WorkbookTheme
): ShapeLayout | UnsupportedDrawing {
  const spPr = child(object, 'xdr:spPr');
  const w = common.rect.width * EMU_PER_POINT;
  const h = common.rect.height * EMU_PER_POINT;
  const prst = child(spPr, 'a:prstGeom');
  const cust = child(spPr, 'a:custGeom');
  let geometry: EvaluatedGeometry;
  if (prst)
    geometry = evaluatePresetGeometry(attr(prst, 'prst') ?? 'rect', w, h, adjustValues(prst));
  else if (cust) geometry = evaluateGeometry(customGeometryDefinition(geometryElement(cust)), w, h);
  else return { ...common, kind: 'unsupported', element: object.local };
  const toPoints = (v: number) => v / EMU_PER_POINT;
  const scaled: EvaluatedGeometry = {
    ...geometry,
    paths: geometry.paths.map((p) => ({
      ...p,
      commands: p.commands.map((c) => {
        if (c.op === 'Z') return c;
        if (c.op === 'C')
          return {
            op: 'C' as const,
            x1: toPoints(c.x1),
            y1: toPoints(c.y1),
            x2: toPoints(c.x2),
            y2: toPoints(c.y2),
            x: toPoints(c.x),
            y: toPoints(c.y),
          };
        if (c.op === 'Q')
          return {
            op: 'Q' as const,
            x1: toPoints(c.x1),
            y1: toPoints(c.y1),
            x: toPoints(c.x),
            y: toPoints(c.y),
          };
        return { op: c.op, x: toPoints(c.x), y: toPoints(c.y) };
      }),
    })),
    textRect: {
      l: toPoints(geometry.textRect.l),
      t: toPoints(geometry.textRect.t),
      r: toPoints(geometry.textRect.r),
      b: toPoints(geometry.textRect.b),
    },
  };
  const style = child(object, 'xdr:style');
  const fill = child(spPr, 'a:noFill')
    ? undefined
    : (colorOf(child(spPr, 'a:solidFill'), theme) ??
      (object.local === 'sp' ? colorOf(child(style, 'a:fillRef'), theme) : undefined));
  const ln = child(spPr, 'a:ln');
  const lineColor = child(ln, 'a:noFill')
    ? undefined
    : (colorOf(child(ln, 'a:solidFill'), theme) ?? colorOf(child(style, 'a:lnRef'), theme));
  // `a:ln/@w` is `ST_LineWidth` (0 to 20116800 EMU) and has no schema default. 9525 EMU (0.75 pt)
  // is the engine's width for a line with no `@w`.
  const lineWidth = numberAttr(ln, 'w', 0, 20_116_800) ?? 9525;
  return {
    ...common,
    kind: 'shape',
    geometry: scaled,
    fill,
    line: lineColor ? { color: lineColor, width: lineWidth / EMU_PER_POINT } : undefined,
  };
}

function objectLayout(
  pkg: XlsxPackage,
  drawingPart: string,
  anchor: ParsedAnchor,
  rect: Rect,
  theme: WorkbookTheme
): DrawingLayout | undefined {
  const object = anchor.object;
  if (!object) return undefined;
  const nv = elements(object).find((el) => el.local.startsWith('nv'));
  const cNvPr = elements(nv).find((el) => el.local === 'cNvPr');
  if (attr(cNvPr, 'hidden') === '1' || attr(cNvPr, 'hidden') === 'true') return undefined;
  const xfrm = child(child(object, 'xdr:spPr'), 'a:xfrm');
  const common: DrawingCommon = {
    index: anchor.index,
    name: (attr(cNvPr, 'name') ?? '').slice(0, 255),
    rect,
    rotation: normalizedDegrees(
      (numberAttr(xfrm, 'rot', -21_600_000 * 100, 21_600_000 * 100) ?? 0) / 60_000
    ),
    flipH: attr(xfrm, 'flipH') === '1' || attr(xfrm, 'flipH') === 'true',
    flipV: attr(xfrm, 'flipV') === '1' || attr(xfrm, 'flipV') === 'true',
  };
  if (object.local === 'pic') {
    const blip = child(child(object, 'xdr:blipFill'), 'a:blip');
    const embed = attr(blip, 'r:embed');
    const rel = embed ? relationshipById(pkg, drawingPart, embed) : undefined;
    if (rel?.external || (!embed && attr(blip, 'r:link') !== undefined))
      return { ...common, kind: 'unsupported', element: 'externalImage' };
    const target = embed ? relTarget(pkg, drawingPart, embed) : undefined;
    if (!target || rel?.type !== REL.image || !partOf(pkg, target))
      return { ...common, kind: 'unsupported', element: 'missingImage' };
    return {
      ...common,
      kind: 'picture',
      partName: target,
      contentType: contentTypeOf(pkg, target),
    };
  }
  if (object.local === 'sp' || object.local === 'cxnSp') return shapeOf(object, common, theme);
  return { ...common, kind: 'unsupported', element: object.local };
}

const layoutCache = new WeakMap<
  Element,
  WeakMap<SheetGeometry, { theme: WorkbookTheme; layouts: readonly DrawingLayout[] }>
>();

/**
 * Every drawing of a worksheet, placed on its geometry, in document order (the order paint
 * stacks them). Cached per drawing part revision, geometry, and theme.
 */
export function sheetDrawings(
  pkg: XlsxPackage,
  drawingPart: string | undefined,
  geometry: SheetGeometry,
  theme: WorkbookTheme
): readonly DrawingLayout[] {
  const part = xmlPartOf(pkg, drawingPart);
  if (!drawingPart || !part || part.root.ns !== NS.xdr || part.root.local !== 'wsDr') return [];
  let byGeometry = layoutCache.get(part.root);
  if (!byGeometry) {
    byGeometry = new WeakMap();
    layoutCache.set(part.root, byGeometry);
  }
  const hit = byGeometry.get(geometry);
  if (hit && hit.theme === theme) return hit.layouts;
  const layouts: DrawingLayout[] = [];
  for (const anchor of parsedAnchors(part.root)) {
    const layout = objectLayout(
      pkg,
      drawingPart,
      anchor,
      rectOf(geometry, anchor.placement),
      theme
    );
    if (layout) layouts.push(layout);
  }
  byGeometry.set(geometry, { theme, layouts });
  return layouts;
}
