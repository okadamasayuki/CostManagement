import { resolveColor } from './color';

/**
 * Excel の表示形式 (numFmt) の簡易実装。
 * 画面表示を「Excel で開いたときの見た目」に近づけるためだけのもので、
 * ファイルに書き出す値そのものには一切影響しない。
 * 完全な互換ではなく、実務でよく使う書式に絞って対応している。
 */

const DATE_TOKEN = /(yy|mm|dd|hh|ss|aaa|ge?)/i;

function pad(n: number, len = 2): string {
  return String(n).padStart(len, '0');
}

function formatDate(d: Date, fmt: string): string {
  const map: Array<[RegExp, () => string]> = [
    [/yyyy/gi, () => String(d.getFullYear())],
    [/yy/gi, () => pad(d.getFullYear() % 100)],
    [/mmmm/g, () => `${d.getMonth() + 1}月`],
    [/mm/g, () => pad(d.getMonth() + 1)],
    [/m(?![a-z])/g, () => String(d.getMonth() + 1)],
    [/dd/gi, () => pad(d.getDate())],
    [/d(?![a-z])/gi, () => String(d.getDate())],
    [/hh/gi, () => pad(d.getHours())],
    [/ss/gi, () => pad(d.getSeconds())],
  ];
  // 'mm' は月と分の両方に使われるが、実務上は月が圧倒的に多いので月として扱う
  let out = fmt.replace(/\[[^\]]*\]/g, '').replace(/"/g, '');
  for (const [re, fn] of map) out = out.replace(re, fn());
  return out.trim() || d.toLocaleDateString('ja-JP');
}

function formatNumeric(n: number, fmt: string): string {
  const isPercent = fmt.includes('%');
  const v = isPercent ? n * 100 : n;

  const decMatch = /\.(0+)/.exec(fmt);
  const decimals = decMatch ? decMatch[1].length : Number.isInteger(v) ? 0 : undefined;
  const useGrouping = fmt.includes('#,##') || fmt.includes(',#');

  let s =
    decimals === undefined
      ? String(v)
      : v.toLocaleString('ja-JP', {
          minimumFractionDigits: decimals,
          maximumFractionDigits: decimals,
          useGrouping,
        });
  if (decimals !== undefined && !useGrouping) {
    s = v.toFixed(decimals);
  } else if (decimals === undefined && useGrouping) {
    s = v.toLocaleString('ja-JP');
  }

  if (isPercent) s += '%';
  if (fmt.includes('¥') || fmt.includes('\\¥')) s = `¥${s}`;
  if (fmt.includes('$')) s = `$${s}`;
  return s;
}

/** 表示形式を適用した文字列を返す */
export function applyNumFmt(value: number | Date, numFmt: string | undefined): string {
  if (value instanceof Date) {
    return numFmt && DATE_TOKEN.test(numFmt)
      ? formatDate(value, numFmt)
      : value.toLocaleDateString('ja-JP');
  }
  if (!numFmt || numFmt === 'General' || numFmt === '@') {
    return Number.isInteger(value) ? String(value) : String(Number(value.toFixed(10)));
  }
  // 正/負/ゼロ のセクション分けがある場合は該当セクションを使う
  const sections = numFmt.split(';');
  const fmt = value < 0 && sections.length > 1 ? sections[1] : sections[0];
  if (DATE_TOKEN.test(fmt)) {
    return formatDate(new Date(value), fmt);
  }
  return formatNumeric(value, fmt);
}

/** ExcelJS の width (文字数) を px に変換する近似 */
export function charWidthToPx(width: number | undefined): number {
  if (!width || width <= 0) return 72;
  return Math.round(width * 7 + 5);
}

/** ポイント -> px */
export function pointsToPx(pt: number | undefined): number {
  if (!pt || pt <= 0) return 20;
  return Math.round(pt * (96 / 72));
}

/**
 * ExcelJS の色オブジェクトから 'FFRRGGBB' を取り出す。
 * テーマ色 (theme + tint) や古い色番号 (indexed) も実際の RGB に解決する。
 */
export function colorToArgb(color: unknown): string | undefined {
  return resolveColor(color)?.argb;
}

/** 'FFRRGGBB' -> '#RRGGBB' */
export function argbToCss(argb: string | undefined): string | undefined {
  if (!argb) return undefined;
  const hex = argb.length === 8 ? argb.slice(2) : argb;
  return `#${hex}`;
}

/** '#RRGGBB' -> 'FFRRGGBB' */
export function cssToArgb(css: string): string {
  const hex = css.replace('#', '').toUpperCase();
  return hex.length === 6 ? `FF${hex}` : hex;
}

/** 背景色に対して読みやすい文字色 (黒 or 白) を選ぶ */
export function readableTextColor(argb: string | undefined): string {
  if (!argb) return '#111827';
  const hex = argb.length === 8 ? argb.slice(2) : argb;
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
  return luminance > 0.6 ? '#111827' : '#ffffff';
}
