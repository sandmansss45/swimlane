export interface IRect {
  left: number;
  top: number;
  right: number;
  bottom: number;
  cx: number;
  cy: number;
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
// vertically. Ported from the earlier proven version.
function hBendPath(x1: number, y1: number, x2: number, y2: number): string {
  const dx = x2 - x1;
  const dy = y2 - y1;
  const r = 12;

  if (Math.abs(dy) < 4 || Math.abs(dx) < 2 * r) {
    return `M ${x1} ${y1} L ${x2} ${y2}`;
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

  return `M ${x1} ${y1} L ${midX1} ${y1} Q ${midX} ${y1} ${midX} ${y1c} L ${midX} ${y2c} Q ${midX} ${y2} ${midX2} ${y2} L ${x2} ${y2}`;
}

// Vertical-exit/entry connector: leaves and enters top/bottom, bends once
// horizontally. Ported from the earlier proven version.
function vBendPath(x1: number, y1: number, x2: number, y2: number): string {
  const dx = x2 - x1;
  const dy = y2 - y1;
  const r = 12;

  if (Math.abs(dx) < 4 || Math.abs(dy) < 2 * r) {
    return `M ${x1} ${y1} L ${x2} ${y2}`;
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

  return `M ${x1} ${y1} L ${x1} ${midY1} Q ${x1} ${midY} ${x1c} ${midY} L ${x2c} ${midY} Q ${x2} ${midY} ${x2} ${midY2} L ${x2} ${y2}`;
}

/**
 * Picks the anchor side on each box based on the dominant direction
 * between them, so a box directly below its source gets a straight
 * vertical connector instead of always exiting right / entering left -
 * this is what keeps lines from cutting diagonally through unrelated
 * boxes between them (they exit/enter at whichever side actually faces
 * the other node, and bend through the gap between rows/columns).
 */
export function connectorPath(a: IRect, b: IRect): string {
  const dx = b.cx - a.cx;
  const dy = b.cy - a.cy;

  if (Math.abs(dy) < 6) {
    const goRight = dx >= 0;
    const x1 = goRight ? a.right : a.left;
    const x2 = goRight ? b.left : b.right;
    return `M ${x1} ${a.cy} L ${x2} ${b.cy}`;
  }
  if (Math.abs(dx) < 6) {
    const goDown = dy >= 0;
    const y1 = goDown ? a.bottom : a.top;
    const y2 = goDown ? b.top : b.bottom;
    return `M ${a.cx} ${y1} L ${b.cx} ${y2}`;
  }
  if (Math.abs(dx) >= Math.abs(dy)) {
    const goRight = dx >= 0;
    const x1 = goRight ? a.right : a.left;
    const x2 = goRight ? b.left : b.right;
    return hBendPath(x1, a.cy, x2, b.cy);
  }
  const goDown = dy >= 0;
  const y1 = goDown ? a.bottom : a.top;
  const y2 = goDown ? b.top : b.bottom;
  return vBendPath(a.cx, y1, b.cx, y2);
}

export function pathMidpoint(a: IRect, b: IRect): { x: number; y: number } {
  return { x: (a.cx + b.cx) / 2, y: (a.cy + b.cy) / 2 };
}

/**
 * Multiple lanes stack vertically within the same column, so a "same
 * column" connection between two different lanes isn't actually adjacent -
 * there's usually at least one other lane's box sitting physically between
 * source and target. A straight connectorPath line there cuts right
 * through whatever's in between. This routes through the lane-label
 * gutter on the left instead (dedicated whitespace, never has task
 * content in it) - exit left, travel down the gutter, re-enter left -
 * which guarantees the line never crosses another box's content no matter
 * how many lanes it spans.
 */
export function gutterPath(a: IRect, b: IRect, gutterX: number): string {
  const r = 10;
  const y1 = a.cy;
  const y2 = b.cy;
  const vDir = y2 >= y1 ? 1 : -1;
  return `M ${a.left} ${y1} L ${gutterX + r} ${y1} Q ${gutterX} ${y1} ${gutterX} ${y1 + vDir * r} L ${gutterX} ${y2 - vDir * r} Q ${gutterX} ${y2} ${gutterX + r} ${y2} L ${b.left} ${y2}`;
}
