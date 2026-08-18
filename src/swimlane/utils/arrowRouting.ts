export interface IRect {
  left: number;
  top: number;
  right: number;
  bottom: number;
  cx: number;
  cy: number;
}

export interface IPathResult {
  d: string;
  labelX: number;
  labelY: number;
}

export function rectFromDomRect(domRect: DOMRect, containerRect: DOMRect, scrollLeft: number, scrollTop: number): IRect {
  const left = domRect.left - containerRect.left + scrollLeft;
  const top = domRect.top - containerRect.top + scrollTop;
  return {
    left,
    top,
    right: left + domRect.width,
    bottom: top + domRect.height,
    cx: left + domRect.width / 2,
    cy: top + domRect.height / 2
  };
}

// Horizontal-exit/entry connector: leaves and enters sideways, bends once
// vertically. Ported from the earlier proven version. Label point sits on
// the bend's own vertical segment, not some independently-guessed spot -
// it's the only point guaranteed to actually be on the drawn line.
function hBendPath(x1: number, y1: number, x2: number, y2: number): IPathResult {
  const dx = x2 - x1;
  const dy = y2 - y1;
  const r = 12;

  if (Math.abs(dy) < 4 || Math.abs(dx) < 2 * r) {
    return { d: `M ${x1} ${y1} L ${x2} ${y2}`, labelX: (x1 + x2) / 2, labelY: (y1 + y2) / 2 };
  }

  const midX = x1 + dx / 2;
  const vDir = dy > 0 ? 1 : -1;
  const rr = Math.min(r, Math.abs(dy) / 2);
  const y1c = y1 + vDir * rr;
  const y2c = y2 - vDir * rr;
  const hDir1 = midX >= x1 ? 1 : -1;
  const hDir2 = x2 >= midX ? 1 : -1;
  const midX1 = midX - hDir1 * rr;
  const midX2 = midX + hDir2 * rr;

  return {
    d: `M ${x1} ${y1} L ${midX1} ${y1} Q ${midX} ${y1} ${midX} ${y1c} L ${midX} ${y2c} Q ${midX} ${y2} ${midX2} ${y2} L ${x2} ${y2}`,
    labelX: midX,
    labelY: (y1 + y2) / 2
  };
}

// Vertical-exit/entry connector: leaves and enters top/bottom, bends once
// horizontally. Ported from the earlier proven version. Label point sits
// on the bend's own horizontal segment for the same reason as hBendPath.
function vBendPath(x1: number, y1: number, x2: number, y2: number): IPathResult {
  const dx = x2 - x1;
  const dy = y2 - y1;
  const r = 12;

  if (Math.abs(dx) < 4 || Math.abs(dy) < 2 * r) {
    return { d: `M ${x1} ${y1} L ${x2} ${y2}`, labelX: (x1 + x2) / 2, labelY: (y1 + y2) / 2 };
  }

  const midY = y1 + dy / 2;
  const hDir = dx > 0 ? 1 : -1;
  const rr = Math.min(r, Math.abs(dx) / 2);
  const x1c = x1 + hDir * rr;
  const x2c = x2 - hDir * rr;
  const vDir1 = midY >= y1 ? 1 : -1;
  const vDir2 = y2 >= midY ? 1 : -1;
  const midY1 = midY - vDir1 * rr;
  const midY2 = midY + vDir2 * rr;

  return {
    d: `M ${x1} ${y1} L ${x1} ${midY1} Q ${x1} ${midY} ${x1c} ${midY} L ${x2c} ${midY} Q ${x2} ${midY} ${x2} ${midY2} L ${x2} ${y2}`,
    labelX: (x1 + x2) / 2,
    labelY: midY
  };
}

export type Side = 'top' | 'bottom' | 'left' | 'right';

/**
 * Which side of each box a connector attaches to, based on the dominant
 * direction between them - exported separately from connectorPath so
 * SwimlaneCanvas can figure out, ahead of drawing anything, which other
 * edges land on the exact same side of the exact same box (multiple
 * dependencies converging on one node, or one decision's several
 * branches leaving from the same node) and spread their attachment
 * points apart instead of every one of them puncturing the box at the
 * identical pixel.
 */
export function pickSides(a: IRect, b: IRect): { fromSide: Side; toSide: Side } {
  const dx = b.cx - a.cx;
  const dy = b.cy - a.cy;
  if (Math.abs(dy) < 6) {
    return dx >= 0 ? { fromSide: 'right', toSide: 'left' } : { fromSide: 'left', toSide: 'right' };
  }
  if (Math.abs(dx) < 6) {
    return dy >= 0 ? { fromSide: 'bottom', toSide: 'top' } : { fromSide: 'top', toSide: 'bottom' };
  }
  if (Math.abs(dx) >= Math.abs(dy)) {
    return dx >= 0 ? { fromSide: 'right', toSide: 'left' } : { fromSide: 'left', toSide: 'right' };
  }
  return dy >= 0 ? { fromSide: 'bottom', toSide: 'top' } : { fromSide: 'top', toSide: 'bottom' };
}

/**
 * Picks the anchor side on each box based on the dominant direction
 * between them, so a box directly below its source gets a straight
 * vertical connector instead of always exiting right / entering left -
 * this is what keeps lines from cutting diagonally through unrelated
 * boxes between them (they exit/enter at whichever side actually faces
 * the other node, and bend through the gap between rows/columns).
 *
 * aOffset/bOffset nudge the attachment point away from dead-center on
 * whichever side gets picked - top/bottom sides shift along X, left/right
 * sides shift along Y - so that when SwimlaneCanvas assigns different
 * offsets to edges sharing a side, they visibly land at different points
 * along the box's edge instead of every one entering/exiting through the
 * same hole.
 *
 * Returns a label point that's actually ON the drawn path - a plain
 * midpoint between the two boxes' centers can land nowhere near the real
 * line once it bends, which is exactly what made a "Yes" label appear to
 * sit next to the wrong branch's box.
 */
export function connectorPath(a: IRect, b: IRect, aOffset = 0, bOffset = 0): IPathResult {
  const dx = b.cx - a.cx;
  const dy = b.cy - a.cy;
  const { fromSide } = pickSides(a, b);
  const onXAxis = fromSide === 'top' || fromSide === 'bottom'; // offset shifts X, not Y

  if (Math.abs(dy) < 6) {
    const goRight = dx >= 0;
    const x1 = goRight ? a.right : a.left;
    const x2 = goRight ? b.left : b.right;
    const y1 = a.cy + aOffset;
    const y2 = b.cy + bOffset;
    // Same row doesn't mean same Y once per-edge spread offsets are
    // applied (see the attachGroups comment above) - two boxes dead level
    // with each other still end up connected by a visibly slanted line if
    // one end's offset differs from the other's. Only draw the plain
    // straight line when both ends actually land on the same Y; otherwise
    // bend through the gap like any other offset pair does, instead of
    // cutting diagonally across the row.
    if (aOffset === bOffset) {
      return { d: `M ${x1} ${y1} L ${x2} ${y2}`, labelX: (x1 + x2) / 2, labelY: (y1 + y2) / 2 };
    }
    return hBendPath(x1, y1, x2, y2);
  }
  if (Math.abs(dx) < 6) {
    const goDown = dy >= 0;
    const y1 = goDown ? a.bottom : a.top;
    const y2 = goDown ? b.top : b.bottom;
    const x1 = a.cx + aOffset;
    const x2 = b.cx + bOffset;
    // Same reasoning as the same-row case above, mirrored for a same-
    // column pair whose X ends up different once offsets are applied.
    if (aOffset === bOffset) {
      return { d: `M ${x1} ${y1} L ${x2} ${y2}`, labelX: (x1 + x2) / 2, labelY: (y1 + y2) / 2 };
    }
    return vBendPath(x1, y1, x2, y2);
  }
  if (onXAxis) {
    const goDown = dy >= 0;
    const y1 = goDown ? a.bottom : a.top;
    const y2 = goDown ? b.top : b.bottom;
    return vBendPath(a.cx + aOffset, y1, b.cx + bOffset, y2);
  }
  const goRight = dx >= 0;
  const x1 = goRight ? a.right : a.left;
  const x2 = goRight ? b.left : b.right;
  return hBendPath(x1, a.cy + aOffset, x2, b.cy + bOffset);
}

/**
 * Every step gets its own column now, so the vertical channel directly
 * above and below a box, within its own column, is ALWAYS empty - no
 * other step ever shares that column. That makes "straight up out of the
 * box, across a dedicated empty strip above every lane, straight down
 * into the target box" a route that's safe by construction for ANY two
 * boxes, no matter how far apart or which lanes they're in - unlike
 * routing sideways through the source's own lane row first (the old
 * gutterPath approach), which cuts through every other box sitting in
 * that same row between the source and the margin whenever the source
 * isn't already near the edge. Label sits on the highway's own
 * horizontal run, not the source/target's straight-line midpoint.
 *
 * aOffset/bOffset shift the up-tick/down-tick away from dead-center of
 * the source/target column, same reasoning as connectorPath - every
 * highway/local-hop edge attaches from the top, so without this, any two
 * of them sharing a source or target would puncture that box at the
 * exact same point.
 */
export function highwayPath(a: IRect, b: IRect, highwayY: number, aOffset = 0, bOffset = 0): IPathResult {
  const r = 10;
  const x1 = a.cx + aOffset;
  const x2 = b.cx + bOffset;
  const hDir = x2 >= x1 ? 1 : -1;
  return {
    d: `M ${x1} ${a.top} L ${x1} ${highwayY + r} Q ${x1} ${highwayY} ${x1 + hDir * r} ${highwayY} L ${x2 - hDir * r} ${highwayY} Q ${x2} ${highwayY} ${x2} ${highwayY + r} L ${x2} ${b.top}`,
    labelX: (x1 + x2) / 2,
    labelY: highwayY
  };
}
