/**
 * 年度の一括置換。
 *
 * ■ 最重要ポイント: 置換は必ず「1 パスの同時置換」で行う
 *
 * 2023->2024, 2024->2025, 2025->2026 を順番に実行すると、
 * 最初の置換で 2024 になった文字列が次の置換で 2025 になり、
 * 元の 2023 が 2025 まで飛んでしまう (連鎖事故)。
 * そのため本実装では文字列を 1 回だけ走査し、見つかった年を
 * その場で確定させる。走査済みの位置は二度と評価しない。
 */

export type YearMapper = (year: number) => number | null;

export function shiftMapper(delta: number, minYear: number, maxYear: number): YearMapper {
  return (y) => (y >= minYear && y <= maxYear ? y + delta : null);
}

export function pairMapper(pairs: Array<{ from: number; to: number }>): YearMapper {
  const map = new Map<number, number>();
  for (const p of pairs) map.set(p.from, p.to);
  return (y) => (map.has(y) ? (map.get(y) as number) : null);
}

/**
 * 文字列中の 4 桁の年を 1 パスで同時置換する。
 * @param wholeNumberOnly true なら前後に数字がある場合を対象外にする
 *        (例: '20240401' の中の '2024' は置換しない)
 * @returns 変更後の文字列。変更がなければ null
 */
export function replaceYearsInString(
  s: string,
  mapper: YearMapper,
  wholeNumberOnly: boolean,
): string | null {
  const re = /\d{4}/g;
  let out = '';
  let last = 0;
  let changed = false;
  let m: RegExpExecArray | null;

  while ((m = re.exec(s)) !== null) {
    const start = m.index;
    const end = start + 4;
    if (wholeNumberOnly) {
      const prev = start > 0 ? s[start - 1] : '';
      const next = end < s.length ? s[end] : '';
      if (/\d/.test(prev) || /\d/.test(next)) continue;
    }
    const to = mapper(parseInt(m[0], 10));
    if (to === null) continue;
    out += s.slice(last, start) + String(to);
    last = end;
    changed = true;
    // re.lastIndex は既に end を指しているため、置換後の文字列を
    // 再走査することはない = 連鎖しない
  }
  if (!changed) return null;
  return out + s.slice(last);
}

const ERA_RE = /(令和|平成|昭和|大正)\s*(元|\d{1,2})\s*年/g;

/**
 * 和暦 (令和6年 など) の年をずらす。元号はまたがない。
 * 「令和6年度」→「令和7年度」のような表記のために用意している。
 */
export function replaceEraYears(s: string, delta: number): string | null {
  let changed = false;
  const out = s.replace(ERA_RE, (whole, era: string, num: string) => {
    const n = num === '元' ? 1 : parseInt(num, 10);
    const next = n + delta;
    if (next < 1 || next > 99) return whole;
    changed = true;
    return `${era}${next}年`;
  });
  return changed ? out : null;
}

/** 数値セル用: 値そのものが対象年と一致するときだけ置換する */
export function mapNumericYear(value: number, mapper: YearMapper): number | null {
  if (!Number.isInteger(value)) return null;
  return mapper(value);
}

/** 置換対象になりうる年をプレビュー用に列挙する */
export function collectYears(s: string, wholeNumberOnly: boolean): number[] {
  const found: number[] = [];
  const re = /\d{4}/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(s)) !== null) {
    const start = m.index;
    const end = start + 4;
    if (wholeNumberOnly) {
      const prev = start > 0 ? s[start - 1] : '';
      const next = end < s.length ? s[end] : '';
      if (/\d/.test(prev) || /\d/.test(next)) continue;
    }
    found.push(parseInt(m[0], 10));
  }
  return found;
}
