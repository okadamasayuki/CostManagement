/**
 * 行 / 列の位置とサイズの計算。
 *
 * 以前は「先頭から N 行 / N 列ぶんの配列」を持っていたため、
 * その上限 (200 列 = GR 列) より先が扱えなかった。
 * かといって Excel の上限 (1,048,576 行 × 16,384 列) ぶんの配列を
 * 持つのは非現実的なので、
 *
 *   ・実際にサイズ指定がある範囲 … 個別の値を積み上げた配列で持つ
 *   ・それより先                 … 既定サイズ × 個数 で計算する
 *
 * という二段構えにする。これで配列の大きさは実データ量に比例したまま、
 * Excel と同じ範囲までスクロールできる。
 */

/** Excel のシートの上限 */
export const EXCEL_MAX_ROWS = 1_048_576;
export const EXCEL_MAX_COLS = 16_384;

export interface AxisMetrics {
  /** 移動できる最大の行/列番号 */
  count: number;
  /** 個別のサイズを持っている範囲 (1..measured) */
  measured: number;
  /** offsets[i] = i 番目の開始位置 (px)。長さは measured + 2。 */
  offsets: number[];
  /** measured を超えた分に使う既定サイズ */
  defaultSize: number;
}

/**
 * @param sizes 1 始まりのサイズ配列 (未指定の要素は defaultSize を使う)
 * @param measured 個別サイズを持つ範囲の上限
 * @param count 移動できる最大番号
 */
export function buildAxis(
  sizes: Array<number | undefined>,
  measured: number,
  count: number,
  defaultSize: number,
): AxisMetrics {
  const m = Math.max(0, Math.min(measured, count));
  const offsets = new Array<number>(m + 2);
  offsets[0] = 0;
  offsets[1] = 0;
  for (let i = 1; i <= m; i++) {
    offsets[i + 1] = offsets[i] + (sizes[i] ?? defaultSize);
  }
  return { count, measured: m, offsets, defaultSize };
}

/** i 番目の開始位置 (px) */
export function axisOffset(a: AxisMetrics, i: number): number {
  if (i <= a.measured + 1) return a.offsets[i];
  return a.offsets[a.measured + 1] + (i - a.measured - 1) * a.defaultSize;
}

/** i 番目のサイズ (px) */
export function axisSize(a: AxisMetrics, i: number): number {
  if (i <= a.measured) return a.offsets[i + 1] - a.offsets[i];
  return a.defaultSize;
}

/** 全体のサイズ (px) */
export function axisTotal(a: AxisMetrics): number {
  return axisOffset(a, a.count + 1);
}

/** px 位置を含む番号 (1 始まり) */
export function axisIndexAt(a: AxisMetrics, pos: number): number {
  const clamped = Math.max(0, pos);
  const measuredEnd = a.offsets[a.measured + 1] ?? 0;

  // 個別サイズの範囲より先は割り算で求まる
  if (clamped >= measuredEnd) {
    const i = a.measured + 1 + Math.floor((clamped - measuredEnd) / a.defaultSize);
    return Math.min(Math.max(i, 1), a.count);
  }

  let lo = 1;
  let hi = a.measured;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (a.offsets[mid] <= clamped) lo = mid;
    else hi = mid - 1;
  }
  return Math.min(Math.max(lo, 1), a.count);
}
