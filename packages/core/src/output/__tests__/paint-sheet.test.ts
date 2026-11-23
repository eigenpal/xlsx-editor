// The canvas paint of a sheet window and its accessible rows, in a happy-dom window of its own
// (the engine tests run with no global DOM). happy-dom has no 2D context, so a recording context
// stands in: text draws glyph by glyph at the measured positions with the measured face, fills
// and borders draw on the grid pixels, shapes draw their geometry, and pictures draw from
// package bytes only.

import { beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { Window } from 'happy-dom';
import { unzipSync, zipSync } from 'fflate';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { FIXTURES } from '../../../../../scripts/fixtures/catalog.ts';
import { readXlsx, workbookIndex, type XlsxPackage } from '../../store/index.ts';
import {
  layoutSheet,
  PRINT_PROFILE,
  type SheetLayout,
  type TextMeasurer,
} from '../../layout/index.ts';
import { substituteMeasurer } from '../../layout/__tests__/test-fonts.ts';
import {
  accessibleGrid,
  borderLine,
  createPictureStore,
  fillPaint,
  paintSheet,
  type PaintOptions,
} from '../index.ts';
import { fontOf } from '../paint-sheet.ts';

let measurer: TextMeasurer;
beforeAll(async () => {
  measurer = await substituteMeasurer();
});

const open = (bytes: Uint8Array): XlsxPackage => {
  const r = readXlsx(bytes);
  if (!r.ok) throw new Error('read');
  return r.pkg;
};
const fixture = (name: string) =>
  open(new Uint8Array(readFileSync(`${import.meta.dir}/../../../../../fixtures/${name}.xlsx`)));

const VIEW = { width: 540, height: 720, scrollX: 0, scrollY: 0 };
const layout = (pkg: XlsxPackage, sheet: string): SheetLayout => {
  const out = layoutSheet(pkg, sheet, VIEW, { measurer, profile: PRINT_PROFILE });
  if (!out) throw new Error('no sheet');
  return out;
};
const doc = () => new Window().document as unknown as Document;

interface Call {
  readonly op: string;
  readonly args: readonly unknown[];
  readonly font?: string;
  readonly fill?: unknown;
}

/**
 * A canvas whose 2D context records each call and the state it ran with. `save` and `restore`
 * keep a stack of the state, as a real 2D context does, so a property set inside a save goes
 * back to its earlier value on restore.
 */
function recordingCanvas(patterns?: { matrix?: DOMMatrix2DInit }[]): {
  canvas: HTMLCanvasElement;
  calls: Call[];
} {
  const calls: Call[] = [];
  // The defaults of a 2D context for the text and line state that a save and restore keep.
  const state: Record<string, unknown> = {
    font: '10px sans-serif',
    fillStyle: '#000000',
    strokeStyle: '#000000',
    lineWidth: 1,
    globalAlpha: 1,
    textAlign: 'start',
    textBaseline: 'alphabetic',
    fontKerning: 'auto',
  };
  const saved: Record<string, unknown>[] = [];
  const gradient = { addColorStop: () => {} };
  const ctx = new Proxy(state, {
    get(target, key: string) {
      if (key in target) return target[key];
      if (key === 'createPattern' && patterns)
        return () => {
          const p: { matrix?: DOMMatrix2DInit; setTransform: (m: DOMMatrix2DInit) => void } = {
            setTransform: (m) => (p.matrix = m),
          };
          patterns.push(p);
          return p;
        };
      return (...args: unknown[]) => {
        calls.push({ op: key, args, font: target.font as string, fill: target.fillStyle });
        if (key === 'save') saved.push({ ...target });
        if (key === 'restore') {
          const back = saved.pop();
          if (back) {
            for (const k of Object.keys(target)) delete target[k];
            Object.assign(target, back);
          }
        }
        return key.startsWith('create') ? gradient : undefined;
      };
    },
    set(target, key: string, value) {
      target[key] = value;
      return true;
    },
  });
  const owner = doc();
  const canvas = owner.createElement('canvas') as HTMLCanvasElement;
  (canvas as unknown as { getContext: () => unknown }).getContext = () => ctx;
  state.canvas = canvas;
  if (patterns) {
    // Tile bitmaps come from the document of this canvas: give them a context that draws nothing.
    const make = owner.createElement.bind(owner);
    owner.createElement = ((tag: string) => {
      const el = make(tag) as unknown as { getContext: () => unknown };
      el.getContext = () => ({ fillStyle: '', fillRect: () => {} });
      return el;
    }) as typeof owner.createElement;
  }
  return { canvas, calls };
}

const OPTIONS: PaintOptions = { scale: 1, width: 540, height: 720, gridlines: false };

/** Each `fillText` call as its glyph and the canvas font it ran with. */
const drawnGlyphs = (calls: readonly Call[]) =>
  calls.filter((c) => c.op === 'fillText').map((c) => ({ glyph: String(c.args[0]), font: c.font }));

/**
 * The glyphs that a layout paints, in paint order, with the canvas font of each span: panes,
 * then text boxes, lines, and spans. The paint skips spaces and draws a code point at a time.
 * The font comes from `fontOf`, as in the paint, so these tests check the cache, not `fontOf`.
 */
function expectedGlyphs(l: SheetLayout) {
  const out: { glyph: string; font: string }[] = [];
  for (const pane of l.panes)
    for (const box of pane.texts)
      for (const line of box.text.lines)
        for (const span of line.spans) {
          const s = span.style;
          const font = fontOf(s.faceId, s.sizePt, s.synthBold, s.synthItalic);
          let i = 0;
          for (const glyph of span.text) {
            const at = i;
            i += glyph.length;
            const end = span.positions[i] ?? span.x + span.width;
            // A glyph whose advance lies a full font size outside the clip cannot show.
            const clip = box.text.rotation ? null : box.clip;
            const hidden =
              clip !== null &&
              (end < clip.x - s.sizePt || span.positions[at]! > clip.x + clip.width + s.sizePt);
            if (glyph !== ' ' && !hidden) out.push({ glyph, font });
          }
        }
  return out;
}

describe('paintSheet', () => {
  test('cell text draws glyph by glyph at the positions that layout measured', () => {
    const l = layout(fixture('basic-values'), 'Values');
    const { canvas, calls } = recordingCanvas();
    const stats = paintSheet(canvas, l, OPTIONS);
    expect(stats.texts).toBeGreaterThan(5);
    const span = l.panes.flatMap((p) => p.texts).find((t) => t.row === 0 && t.col === 0)!.text
      .lines[0]!.spans[0]!;
    const glyphs = calls.filter((c) => c.op === 'fillText').slice(0, span.text.length - 1);
    expect(glyphs.map((c) => c.args[0]).join('')).toBe('Sharedtext');
    expect(glyphs[0]!.args[1]).toBe(span.positions[0]);
    expect(glyphs[0]!.font).toContain(`${span.style.faceId}, sans-serif`);
  });

  test('a run that cannot reach the clip of its cell draws no glyph', () => {
    const runs = Array.from({ length: 200 }, (_, i) => `<r><t>r${i % 10}</t></r>`).join('');
    const pkg = open(
      buildXlsx({
        sheets: [{ name: 'S', rows: row(1, [c('A1', 0, { t: 's' })]) }],
        rawStrings: [`<si>${runs}</si>`],
      })
    );
    const l = layout(pkg, 'S');
    const box = l.panes.flatMap((p) => p.texts)[0]!;
    const spans = box.text.lines[0]!.spans;
    const right = box.clip.x + box.clip.width;
    const reach = spans.filter((s) => s.x - s.style.sizePt <= right);
    expect(reach.length).toBeLessThan(spans.length);
    const { canvas, calls } = recordingCanvas();
    paintSheet(canvas, l, OPTIONS);
    const glyphs = drawnGlyphs(calls)
      .map((g) => g.glyph)
      .join('');
    expect(glyphs).toBe(reach.map((s) => s.text).join(''));
  });

  test('a very long unwrapped text draws only the glyphs that can reach its clip', () => {
    const pkg = open(
      buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 'x'.repeat(200_000))]) }] })
    );
    const l = layout(pkg, 'S');
    const box = l.panes.flatMap((p) => p.texts)[0]!;
    const span = box.text.lines[0]!.spans[0]!;
    const right = box.clip.x + box.clip.width;
    const inClip = span.positions.filter((x) => x >= box.clip.x && x <= right).length;
    expect(inClip).toBeGreaterThan(0);
    const { canvas, calls } = recordingCanvas();
    paintSheet(canvas, l, OPTIONS);
    const drawn = drawnGlyphs(calls).length;
    expect(drawn).toBeGreaterThanOrEqual(inClip);
    expect(drawn).toBeLessThan(inClip + 10);
  });

  test('each cell of a row draws with its own font after the restore of the cell before', () => {
    // Every cell text paints inside its own save and restore. A cache of the last font must not
    // outlive the restore, or the second bold cell of a row draws with the default font that the
    // restore put back.
    const bold = '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>';
    const pkg = open(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            rows: row(1, [
              c('A1', 'Plain'),
              c('B1', 'One', { s: 1 }),
              c('C1', 'Two', { s: 1 }),
              c('D1', 'Three', { s: 1 }),
              c('E1', 'Plain'),
              c('F1', 'Again'),
            ]),
          },
        ],
        styles: {
          fonts: [
            '<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>',
          ],
          cellXfs: [bold],
        },
      })
    );
    const l = layout(pkg, 'S');
    const bolds = l.panes.flatMap((p) => p.texts).map((t) => t.text.lines[0]!.spans[0]!.style.bold);
    expect(bolds).toEqual([false, true, true, true, false, false]);
    const { canvas, calls } = recordingCanvas();
    paintSheet(canvas, l, OPTIONS);
    const drawn = drawnGlyphs(calls);
    expect(drawn).toEqual(expectedGlyphs(l));
    // The bold cells draw in a font of their own, not in the font of the plain cells.
    const fontOfFirstGlyph = (first: string) => drawn.find((g) => g.glyph === first)!.font;
    expect(fontOfFirstGlyph('O')).not.toBe(fontOfFirstGlyph('P'));
    expect(new Set(drawn.map((g) => g.font)).size).toBe(2);
  });

  test('a run of spaces in another font leaves the next run with its own font', () => {
    // A run of only spaces draws no glyph and no line, so its paint returns early. It still set
    // its font, and the next run must not take that font for its own.
    const pkg = open(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            rows:
              '<row r="1"><c r="A1" t="inlineStr"><is><r><t xml:space="preserve">Lead </t></r>' +
              '<r><rPr><b/><sz val="11"/><rFont val="Calibri"/></rPr><t xml:space="preserve">   </t></r>' +
              '<r><t>Tail</t></r></is></c></row>',
          },
        ],
      })
    );
    const l = layout(pkg, 'S');
    const spans = l.panes.flatMap((p) => p.texts).flatMap((t) => t.text.lines[0]!.spans);
    expect(spans.map((sp) => [sp.text, sp.style.bold])).toEqual([
      ['Lead ', false],
      ['   ', true],
      ['Tail', false],
    ]);
    const { canvas, calls } = recordingCanvas();
    paintSheet(canvas, l, OPTIONS);
    expect(drawnGlyphs(calls)).toEqual(expectedGlyphs(l));
  });

  test.each(FIXTURES.map((f) => f.name))(
    'every glyph of %s draws with the font of its span',
    (name) => {
      const pkg = fixture(name);
      for (const sheet of workbookIndex(pkg)!.sheets) {
        if (sheet.kind !== 'worksheet') continue;
        const l = layout(pkg, sheet.name);
        const { canvas, calls } = recordingCanvas();
        paintSheet(canvas, l, OPTIONS);
        // The sheet name goes into the comparison, so a failure names it.
        expect({ sheet: sheet.name, glyphs: drawnGlyphs(calls) }).toEqual({
          sheet: sheet.name,
          glyphs: expectedGlyphs(l),
        });
      }
    }
  );

  test('scale and the device pixel ratio size the canvas; layout stays in points', () => {
    const l = layout(fixture('basic-values'), 'Values');
    const { canvas, calls } = recordingCanvas();
    paintSheet(canvas, l, {
      ...OPTIONS,
      scale: 2,
      width: 100,
      height: 50,
      gridlines: true,
      devicePixelRatio: 2,
    });
    expect([canvas.width, canvas.height]).toEqual([400, 200]);
    expect([canvas.style.width, canvas.style.height]).toEqual(['200px', '100px']);
    expect(calls.find((c) => c.op === 'setTransform')!.args).toEqual([4, 0, 0, 4, 0, 0]);
    // Gridlines are one grid pixel wide rectangles.
    expect(calls.some((c) => c.op === 'fillRect' && c.fill === '#D4D4D4')).toBe(true);
  });

  test('fills, borders, merged ranges, and shapes paint', () => {
    const pkg = fixture('grid-structure');
    const l = layout(pkg, l0(pkg));
    const { canvas, calls } = recordingCanvas();
    const stats = paintSheet(canvas, l, OPTIONS);
    expect(stats.cells).toBeGreaterThan(0);
    expect(stats.drawings).toBeGreaterThan(0);
    expect(calls.some((c) => c.op === 'bezierCurveTo' || c.op === 'lineTo')).toBe(true);
  });

  test('the tile grid starts at the pane sheet origin, or at patternOrigin when it is set', () => {
    const pkg = fixture('fills-patterns');
    const l = layoutSheet(
      pkg,
      'Fills',
      { ...VIEW, scrollX: 13, scrollY: 7 },
      {
        measurer,
        profile: PRINT_PROFILE,
      }
    )!;
    const origin = l.panes[0]!.sheetOrigin;
    // At 1 device pixel for each point, the bitmap is one tile of 8 pixels: the phase is mod 8.
    const phase = (v: number) => ((Math.round(v) % 8) + 8) % 8;
    expect([phase(origin.x), phase(origin.y)]).not.toEqual([4, 4]);
    const screen: { matrix?: DOMMatrix2DInit }[] = [];
    paintSheet(recordingCanvas(screen).canvas, l, OPTIONS);
    expect(screen.length).toBeGreaterThan(0);
    for (const p of screen)
      expect(p.matrix).toEqual({
        a: 1,
        b: 0,
        c: 0,
        d: 1,
        e: phase(origin.x),
        f: phase(origin.y),
      });
    const page: { matrix?: DOMMatrix2DInit }[] = [];
    paintSheet(recordingCanvas(page).canvas, l, { ...OPTIONS, patternOrigin: { x: -36, y: 756 } });
    for (const p of page)
      expect(p.matrix).toEqual({ a: 1, b: 0, c: 0, d: 1, e: phase(-36), f: phase(756) });
  });

  test('a picture without a decoder paints a placeholder, never a URL', () => {
    const pkg = fixture('grid-structure');
    const l = layout(pkg, l0(pkg));
    const pictures = createPictureStore(pkg, undefined);
    const { canvas, calls } = recordingCanvas();
    paintSheet(canvas, l, { ...OPTIONS, pictures });
    expect(calls.some((c) => c.op === 'strokeRect')).toBe(true);
    expect(calls.some((c) => c.op === 'drawImage')).toBe(false);
  });

  test('a picture whose header size cannot be read is refused before decoding', () => {
    const files = unzipSync(
      readFileSync(`${import.meta.dir}/../../../../../fixtures/grid-structure.xlsx`)
    );
    // A RIFF WebP header: `imageSize` reads no size from it, so the pixel cap cannot apply.
    files['xl/media/image1.png'] = new TextEncoder().encode('RIFF\0\0\0\0WEBPVP8 ');
    const pkg = open(zipSync(files));
    let decodes = 0;
    const win = {
      Blob,
      createImageBitmap: () => {
        decodes += 1;
        return Promise.reject(new Error('not reached'));
      },
    };
    const store = createPictureStore(pkg, win);
    const state = store.get('/xl/media/image1.png', 'image/png');
    expect(state).toEqual({ kind: 'unsupported', reason: 'unknown-size' });
    expect(decodes).toBe(0);
  });

  test('without a 2D context the paint draws nothing and says so', () => {
    const l = layout(fixture('basic-values'), 'Values');
    const canvas = doc().createElement('canvas') as HTMLCanvasElement;
    (canvas as unknown as { getContext: () => null }).getContext = () => null;
    expect(paintSheet(canvas, l, OPTIONS).texts).toBe(0);
  });
});

describe('accessibleGrid', () => {
  const LABELS = {
    staleCell: (text: string) => `${text}, stale`,
    staleNoResult: 'no result, stale',
  };

  test('a stale cell box with no text reads as having no result, with no bare stale words', () => {
    const l = layout(fixture('basic-values'), 'Values');
    const box = l.panes[0]!.texts[0]!;
    const empty = { ...box, row: 40, col: 3, stale: true, text: { ...box.text, lines: [] } };
    const synthetic = { ...l, panes: [{ ...l.panes[0]!, texts: [empty] }] };
    const cell = accessibleGrid(doc(), synthetic, LABELS).querySelector('[data-cell="D41"]')!;
    expect(cell.textContent).toBe('no result, stale');
    expect(cell.getAttribute('data-stale')).toBe('1');
  });

  test('a stale formula cell reads its shown text with the stale words; another cell does not', () => {
    const pkg = open(
      buildXlsx({
        calcPr: 'fullCalcOnLoad="1"',
        sheets: [{ name: 'S', rows: row(1, [c('A1', 2), c('B1', 4, { f: 'A1*2' })]) }],
      })
    );
    const rows = accessibleGrid(doc(), layout(pkg, 'S'), LABELS);
    const b1 = rows.querySelector('[data-cell="B1"]')!;
    expect(b1.textContent).toBe('4, stale');
    expect(b1.getAttribute('data-stale')).toBe('1');
    const a1 = rows.querySelector('[data-cell="A1"]')!;
    expect(a1.textContent).toBe('2');
    expect(a1.hasAttribute('data-stale')).toBe(false);
  });

  test('each cell text is a grid cell of its row, in reading order', () => {
    const l = layout(fixture('basic-values'), 'Values');
    const rows = accessibleGrid(doc(), l, LABELS);
    expect(rows.getAttribute('role')).toBe('rowgroup');
    const a1 = rows.querySelector('[data-cell="A1"]')!;
    expect(a1.getAttribute('role')).toBe('gridcell');
    expect(a1.textContent).toBe('Shared text');
    expect(a1.parentElement!.getAttribute('aria-rowindex')).toBe('1');
    expect(a1.getAttribute('aria-colindex')).toBe('1');
  });

  test('cell text cannot become markup', () => {
    const pkg = open(
      buildXlsx({
        sheets: [{ name: 'S', rows: row(1, [c('A1', '<script>x</script>')]) }],
      })
    );
    const rows = accessibleGrid(doc(), layout(pkg, 'S'), LABELS);
    expect(rows.querySelector('script')).toBeNull();
    expect(rows.querySelector('[data-cell="A1"]')!.textContent).toBe('<script>x</script>');
  });

  test('a frozen pane draws its edge lines only when the options ask for them', () => {
    const rows = Array.from({ length: 30 }, (_, i) =>
      row(i + 1, [c(`A${i + 1}`, `r${i + 1}`), c(`B${i + 1}`, 'b')])
    ).join('');
    const pkg = open(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            rows,
            viewChildren:
              '<pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/>',
          },
        ],
      })
    );
    const l = layout(pkg, 'S');
    expect(l.frozenEdges.x).toBeGreaterThan(0);
    expect(l.frozenEdges.y).toBe(28);
    const edges = (frozenLines: boolean) => {
      const { canvas, calls } = recordingCanvas();
      paintSheet(canvas, l, { ...OPTIONS, frozenLines });
      return calls.filter((call) => call.op === 'fillRect' && call.fill === '#808080');
    };
    expect(edges(false)).toEqual([]);
    expect(edges(true).map((call) => call.args.slice(0, 2))).toEqual([
      [l.frozenEdges.x, 0],
      [0, l.frozenEdges.y],
    ]);
  });
});

describe('paint styles', () => {
  test('border styles map to grid-pixel lines', () => {
    expect(borderLine('thin')).toEqual({ width: 1, dash: [], double: false });
    expect(borderLine('double')!.double).toBe(true);
    expect(borderLine('none')).toBeUndefined();
    expect(borderLine('__proto__')).toBeUndefined();
  });
  test('a pattern fill paints its tile, with the mix of its colors as the fallback', () => {
    expect(
      fillPaint({
        kind: 'pattern',
        pattern: 'mediumGray',
        foreground: '000000',
        background: undefined,
      })
    ).toEqual({
      kind: 'pattern',
      pattern: 'mediumGray',
      foreground: '#000000',
      background: '#FFFFFF',
      mix: '#808080',
    });
    // A pattern with no tile paints its foreground share as 0: the background.
    expect(
      fillPaint({ kind: 'pattern', pattern: 'unknown', foreground: '000000', background: 'FF0000' })
    ).toEqual({ kind: 'color', color: '#FF0000' });
    expect(fillPaint({ kind: 'solid', color: 'bad"' })).toEqual({
      kind: 'color',
      color: '#FFFFFF',
    });
  });
});

function l0(pkg: XlsxPackage): string {
  return workbookIndex(pkg)!.sheets[0]!.name;
}
