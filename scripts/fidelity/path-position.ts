// The position that the shapes of a path gradient paint at a point: the last shape over the
// point, projected on its gradient line. Shared by the tests that compare paint with the
// reference PDF.

import type { PathShape } from '../../packages/core/src/output/paint-fills.ts';

function inside(points: PathShape['points'], px: number, py: number): boolean {
  let hit = false;
  for (let i = 0, j = points.length - 1; i < points.length; j = i++) {
    const [xi, yi] = points[i]!;
    const [xj, yj] = points[j]!;
    if (yi > py !== yj > py && px < ((xj - xi) * (py - yi)) / (yj - yi) + xi) hit = !hit;
  }
  return hit;
}

/** The position from 0 to 1 that `shapes` paint at a point, or NaN outside every shape. */
export function paintedPosition(shapes: readonly PathShape[], px: number, py: number): number {
  for (let i = shapes.length - 1; i >= 0; i--) {
    const s = shapes[i]!;
    if (!inside(s.points, px, py)) continue;
    const [fx, fy] = s.from;
    const [tx, ty] = s.to;
    const len = (tx - fx) ** 2 + (ty - fy) ** 2;
    if (len === 0) return 0;
    return Math.min(1, Math.max(0, ((px - fx) * (tx - fx) + (py - fy) * (ty - fy)) / len));
  }
  return Number.NaN;
}
