// The offset index of one grid axis (design D6): the size of every row or column, and the
// offset of any of them, without a walk over the axis.
//
// The axis is a default size plus sorted segments of indexes with their own size (an `x:col`
// range from `@min` to `@max`, `CT_Col`; a row with `x:row/@ht`, `CT_Row`; a hidden row). Lookups are binary searches over the segments, so the
// cost depends on the number of segments, which is in proportion to the content of the file,
// never on the sheet size: the offset of row 1048576 is one search.

/** A run of indexes with one size. Indexes are 0-based and inclusive. */
export interface AxisSegment {
  readonly first: number;
  readonly last: number;
  /** Size in points. 0 for a hidden row or column. */
  readonly size: number;
}

export class AxisIndex {
  /** The number of indexes on the axis: 1048576 rows or 16384 columns. */
  readonly count: number;
  readonly defaultSize: number;
  private readonly segments: readonly AxisSegment[];
  /** The offset of each segment's first index. */
  private readonly starts: readonly number[];

  /**
   * `segments` may arrive in any order. Where two overlap, the one that starts first keeps
   * the shared indexes (and, for equal starts, the one given first); the other is clipped.
   */
  constructor(count: number, defaultSize: number, segments: readonly AxisSegment[]) {
    this.count = count;
    this.defaultSize = defaultSize;
    const sorted = normalize(count, segments);
    const starts: number[] = [];
    let offset = 0;
    let next = 0;
    for (const s of sorted) {
      offset += (s.first - next) * defaultSize;
      starts.push(offset);
      offset += (s.last - s.first + 1) * s.size;
      next = s.last + 1;
    }
    this.segments = sorted;
    this.starts = starts;
  }

  /** The index of the last segment whose first index is at most `i`, or -1. */
  private segmentAt(i: number): number {
    let lo = 0;
    let hi = this.segments.length - 1;
    let found = -1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      if (this.segments[mid]!.first <= i) {
        found = mid;
        lo = mid + 1;
      } else hi = mid - 1;
    }
    return found;
  }

  /** The size of index `i`, in points. */
  sizeOf(i: number): number {
    const k = this.segmentAt(i);
    const s = k < 0 ? undefined : this.segments[k]!;
    return s && i <= s.last ? s.size : this.defaultSize;
  }

  /** The offset of the start of index `i` from the start of the axis, in points. */
  offsetOf(i: number): number {
    const k = this.segmentAt(i);
    if (k < 0) return i * this.defaultSize;
    const s = this.segments[k]!;
    const start = this.starts[k]!;
    if (i <= s.last) return start + (i - s.first) * s.size;
    return start + (s.last - s.first + 1) * s.size + (i - s.last - 1) * this.defaultSize;
  }

  /** The length of the whole axis, in points. */
  get total(): number {
    return this.offsetOf(this.count);
  }

  /**
   * The index that holds `offset`: the last index whose start is at most `offset`, skipping
   * indexes of size 0. Offsets before the axis give 0, and offsets past it give the last index.
   */
  indexAt(offset: number): number {
    if (offset <= 0) return this.firstVisibleFrom(0);
    // Find the segment boundary at or before the offset, then step within the run.
    let lo = 0;
    let hi = this.segments.length - 1;
    let k = -1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      if (this.starts[mid]! <= offset) {
        k = mid;
        lo = mid + 1;
      } else hi = mid - 1;
    }
    let base = 0;
    let baseOffset = 0;
    if (k >= 0) {
      const s = this.segments[k]!;
      const start = this.starts[k]!;
      const end = start + (s.last - s.first + 1) * s.size;
      if (offset < end && s.size > 0)
        return Math.min(s.first + Math.floor((offset - start) / s.size), s.last);
      base = s.last + 1;
      baseOffset = end;
    }
    if (this.defaultSize <= 0) return this.firstVisibleFrom(base);
    const i = base + Math.floor((offset - baseOffset) / this.defaultSize);
    return Math.min(i, this.count - 1);
  }

  /** The first index at or after `i` with a size above 0, or `i` when there is none. */
  private firstVisibleFrom(i: number): number {
    let at = i;
    for (;;) {
      const k = this.segmentAt(at);
      const s = k < 0 ? undefined : this.segments[k]!;
      if (!s || at > s.last || s.size > 0) return at;
      if (s.last + 1 >= this.count) return i;
      at = s.last + 1;
    }
  }
}

/** Sort, clip to the axis, and drop the overlapping parts of later segments. */
function normalize(count: number, segments: readonly AxisSegment[]): AxisSegment[] {
  const valid = segments
    .map((s, order) => ({ ...s, order }))
    .filter(
      (s) =>
        Number.isInteger(s.first) &&
        Number.isInteger(s.last) &&
        s.first <= s.last &&
        s.last >= 0 &&
        s.first < count &&
        Number.isFinite(s.size) &&
        s.size >= 0
    )
    .map((s) => ({ ...s, first: Math.max(0, s.first), last: Math.min(count - 1, s.last) }))
    .sort((a, b) => a.first - b.first || a.order - b.order);
  const out: AxisSegment[] = [];
  let next = 0;
  for (const s of valid) {
    const first = Math.max(s.first, next);
    if (first > s.last) continue;
    const prev = out[out.length - 1];
    if (prev && prev.last + 1 === first && prev.size === s.size)
      out[out.length - 1] = { first: prev.first, last: s.last, size: s.size };
    else out.push({ first, last: s.last, size: s.size });
    next = s.last + 1;
  }
  return out;
}
