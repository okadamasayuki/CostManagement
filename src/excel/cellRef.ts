/** A1 形式のセル参照 <-> 行列番号 (いずれも 1 始まり) の相互変換 */

export interface CellAddr {
  row: number;
  col: number;
}

/** 選択範囲。row/col はいずれも 1 始まり、両端を含む。 */
export interface RangeRect {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

/** 1 -> 'A', 27 -> 'AA' */
export function colToLetter(col: number): string {
  let n = col;
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** 'A' -> 1, 'AA' -> 27 */
export function letterToCol(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n;
}

export function addrToA1(addr: CellAddr): string {
  return `${colToLetter(addr.col)}${addr.row}`;
}

export function a1ToAddr(a1: string): CellAddr | null {
  const m = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(a1.trim());
  if (!m) return null;
  return { col: letterToCol(m[1]), row: parseInt(m[2], 10) };
}

export function rectToA1(rect: RangeRect): string {
  const tl = addrToA1({ row: rect.top, col: rect.left });
  if (rect.top === rect.bottom && rect.left === rect.right) return tl;
  return `${tl}:${addrToA1({ row: rect.bottom, col: rect.right })}`;
}

/** 'A1' / 'A1:C5' / 'A:C' (列全体) / '1:5' (行全体) を解釈する */
export function parseA1Range(text: string, maxRow = 1048576, maxCol = 16384): RangeRect | null {
  const t = text.trim().replace(/\$/g, '');
  if (!t) return null;

  const colOnly = /^([A-Za-z]+):([A-Za-z]+)$/.exec(t);
  if (colOnly) {
    const a = letterToCol(colOnly[1]);
    const b = letterToCol(colOnly[2]);
    return { top: 1, bottom: maxRow, left: Math.min(a, b), right: Math.max(a, b) };
  }

  const rowOnly = /^(\d+):(\d+)$/.exec(t);
  if (rowOnly) {
    const a = parseInt(rowOnly[1], 10);
    const b = parseInt(rowOnly[2], 10);
    return { left: 1, right: maxCol, top: Math.min(a, b), bottom: Math.max(a, b) };
  }

  const parts = t.split(':');
  const start = a1ToAddr(parts[0]);
  if (!start) return null;
  if (parts.length === 1) {
    return { top: start.row, left: start.col, bottom: start.row, right: start.col };
  }
  const end = a1ToAddr(parts[1]);
  if (!end) return null;
  return {
    top: Math.min(start.row, end.row),
    bottom: Math.max(start.row, end.row),
    left: Math.min(start.col, end.col),
    right: Math.max(start.col, end.col),
  };
}

export function normalizeRect(a: CellAddr, b: CellAddr): RangeRect {
  return {
    top: Math.min(a.row, b.row),
    bottom: Math.max(a.row, b.row),
    left: Math.min(a.col, b.col),
    right: Math.max(a.col, b.col),
  };
}

export function rectContains(rect: RangeRect, row: number, col: number): boolean {
  return row >= rect.top && row <= rect.bottom && col >= rect.left && col <= rect.right;
}

export function rectCellCount(rect: RangeRect): number {
  return (rect.bottom - rect.top + 1) * (rect.right - rect.left + 1);
}
