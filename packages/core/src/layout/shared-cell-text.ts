// The text layout of cells with very many runs. One rich shared string can fill every visible
// row, and a layout of its runs for each cell makes each paint cost the runs times the rows. So
// the cells of one pane that show the same runs in the same format and size share one layout, and
// each cell keeps only the spans that can reach its clip. The text for assistive technology stays
// complete in `shownText`.

import type { ResolvedCellFormat } from '../cascade/index.ts';
import type { CellTextLayout, Rect, TextLine, TextSpan } from './cell-text.ts';

/** A cell with at least this many runs shares its layout. */
export const MANY_RUNS = 256;

/** The largest font size of each line of a layout. */
const lineSizes = (text: CellTextLayout): readonly number[] =>
  text.lines.map((line) => line.spans.reduce((size, s) => Math.max(size, s.style.sizePt), 0));

/** The first index of the ascending `spans` for which `reaches` holds, or their count. */
function firstReaching(spans: readonly TextSpan[], reaches: (s: TextSpan) => boolean): number {
  let low = 0;
  let high = spans.length;
  while (low < high) {
    const mid = (low + high) >>> 1;
    if (reaches(spans[mid]!)) high = mid;
    else low = mid + 1;
  }
  return low;
}

/** The layout of one cell, placed from the layout that its pane shares. */
export class PlacedCellText {
  /** The layout in the coordinates of the sheet. Its lines are those of the shared layout. */
  readonly text: CellTextLayout;

  constructor(
    private readonly shared: SharedLayout,
    private readonly dx: number,
    private readonly dy: number
  ) {
    const { text } = shared;
    this.text = {
      ...text,
      left: text.left + dx,
      right: text.right + dx,
      areaLeft: text.areaLeft + dx,
      areaRight: text.areaRight + dx,
    };
  }

  /** The complete text of the cell, with its lines joined by a space. */
  get shownText(): string {
    return this.shared.shownText;
  }

  /** The layout moved by the pane offset, with the spans that can reach `clip`. */
  inView(clip: Rect, paneDx: number, paneDy: number): CellTextLayout {
    const x = this.dx + paneDx;
    const y = this.dy + paneDy;
    const lines = this.shared.text.lines.map((line, i): TextLine => {
      // A glyph can overhang its advance, so a span must lie a full font size away.
      const size = this.shared.sizes[i]!;
      const baseline = line.baseline + y;
      const near = baseline + size * 2 >= clip.y && baseline - size * 2 <= clip.y + clip.height;
      const right = clip.x + clip.width;
      const from = firstReaching(line.spans, (s) => s.x + x + s.width + size >= clip.x);
      const to = firstReaching(line.spans, (s) => s.x + x - size > right);
      const spans = near ? line.spans.slice(from, Math.max(from, to)) : [];
      return {
        ...line,
        x: line.x + x,
        baseline,
        spans: spans.map((s) => ({
          ...s,
          x: s.x + x,
          positions: s.positions.map((p) => p + x),
        })),
      };
    });
    return {
      ...this.shared.text,
      lines,
      left: this.shared.text.left + x,
      right: this.shared.text.right + x,
      areaLeft: this.shared.text.areaLeft + x,
      areaRight: this.shared.text.areaRight + x,
    };
  }
}

interface SharedLayout {
  readonly format: ResolvedCellFormat;
  readonly box: Rect;
  readonly text: CellTextLayout;
  readonly sizes: readonly number[];
  readonly shownText: string;
}

/** The shared layouts of one pane. */
export class SharedCellTexts {
  private readonly known = new Map<object, SharedLayout[]>();

  /** The layout of `runs` in `format` for `box`, when a cell of the pane already made it. */
  find(runs: object, format: ResolvedCellFormat, box: Rect): PlacedCellText | undefined {
    const shared = this.known
      .get(runs)
      ?.find(
        (s) =>
          s.format.font === format.font &&
          s.format.alignment === format.alignment &&
          s.box.width === box.width &&
          s.box.height === box.height
      );
    return shared && new PlacedCellText(shared, box.x - shared.box.x, box.y - shared.box.y);
  }

  /** Keeps the layout `text` of `runs` that was made for `box`. */
  add(runs: object, format: ResolvedCellFormat, box: Rect, text: CellTextLayout): PlacedCellText {
    const shownText = text.lines.map((l) => l.spans.map((s) => s.text).join('')).join(' ');
    const shared = { format, box, text, sizes: lineSizes(text), shownText };
    const list = this.known.get(runs);
    if (list) list.push(shared);
    else this.known.set(runs, [shared]);
    return new PlacedCellText(shared, 0, 0);
  }
}
