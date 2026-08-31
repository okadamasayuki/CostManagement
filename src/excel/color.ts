/**
 * 塗りつぶし色の解決。
 *
 * Excel の色指定には 3 通りある:
 *   1. argb   … 'FFFFFF00' のような直接指定
 *   2. theme  … テーマ色の番号 + 明度調整 (tint)。
 *               「塗りつぶしの色」から標準パレットを選ぶとこちらになる。
 *   3. indexed… 古い形式の色番号
 *
 * このうち 2 と 3 は数値でしか入っておらず、そのままでは
 * 「同じ色かどうか」の判定も画面表示もできない。ここで実際の RGB に
 * 解決する。
 *
 * 判定 (どのセルが同じ色か) には元の指定をそのまま表す key を使う。
 * RGB への変換は標準テーマに基づく近似なので、独自テーマのブックでは
 * 表示色がずれることがあるが、key で比較する限り判定は常に正しい。
 */

/** Office 標準テーマの色。xlsx の theme 属性の並び順 (lt1/dk1 が入れ替わる点に注意)。 */
const THEME_COLORS = [
  'FFFFFF', // 0 lt1  背景 1
  '000000', // 1 dk1  テキスト 1
  'E7E6E6', // 2 lt2  背景 2
  '44546A', // 3 dk2  テキスト 2
  '4472C4', // 4 accent1
  'ED7D31', // 5 accent2
  'A5A5A5', // 6 accent3
  'FFC000', // 7 accent4
  '5B9BD5', // 8 accent5
  '70AD47', // 9 accent6
  '0563C1', // 10 ハイパーリンク
  '954F72', // 11 表示済みハイパーリンク
];

/** 古い形式の色番号 (BIFF8 標準パレット) */
const INDEXED_COLORS = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
];

export interface FillRef {
  /** 同じ色かどうかの判定に使う。元の指定をそのまま表す。 */
  key: string;
  /** 表示用の 'FFRRGGBB' */
  argb: string;
  /** テーマ色などから推定した色か (独自テーマだとずれる可能性がある) */
  isApprox: boolean;
}

// --- HSL 経由の明度調整 (tint) -------------------------------------------

function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
  const rr = r / 255;
  const gg = g / 255;
  const bb = b / 255;
  const max = Math.max(rr, gg, bb);
  const min = Math.min(rr, gg, bb);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === rr) h = ((gg - bb) / d + (gg < bb ? 6 : 0)) / 6;
  else if (max === gg) h = ((bb - rr) / d + 2) / 6;
  else h = ((rr - gg) / d + 4) / 6;
  return [h, s, l];
}

function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  if (s === 0) {
    const v = Math.round(l * 255);
    return [v, v, v];
  }
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const hue = (t: number): number => {
    let tt = t;
    if (tt < 0) tt += 1;
    if (tt > 1) tt -= 1;
    if (tt < 1 / 6) return p + (q - p) * 6 * tt;
    if (tt < 1 / 2) return q;
    if (tt < 2 / 3) return p + (q - p) * (2 / 3 - tt) * 6;
    return p;
  };
  return [
    Math.round(hue(h + 1 / 3) * 255),
    Math.round(hue(h) * 255),
    Math.round(hue(h - 1 / 3) * 255),
  ];
}

/**
 * tint を適用する (ECMA-376 の定義)。
 *   tint < 0 … 暗くする   L' = L * (1 + tint)
 *   tint > 0 … 明るくする L' = L * (1 - tint) + tint
 */
export function applyTint(hex6: string, tint: number): string {
  if (!tint) return hex6.toUpperCase();
  const r = parseInt(hex6.slice(0, 2), 16);
  const g = parseInt(hex6.slice(2, 4), 16);
  const b = parseInt(hex6.slice(4, 6), 16);
  const [h, s, l] = rgbToHsl(r, g, b);
  const nl = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const [nr, ng, nb] = hslToRgb(h, s, Math.min(1, Math.max(0, nl)));
  return [nr, ng, nb].map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase();
}

/** ExcelJS の色オブジェクトを解決する。色指定が無ければ null。 */
export function resolveColor(color: unknown): FillRef | null {
  if (!color || typeof color !== 'object') return null;
  const c = color as { argb?: string; theme?: number; tint?: number; indexed?: number };

  if (typeof c.argb === 'string' && /^[0-9A-Fa-f]{6,8}$/.test(c.argb)) {
    const argb = (c.argb.length === 6 ? `FF${c.argb}` : c.argb).toUpperCase();
    // argb にも tint が併記されることがある
    if (c.tint) {
      const tinted = applyTint(argb.slice(2), c.tint);
      return { key: `argb:${argb}+${c.tint}`, argb: `FF${tinted}`, isApprox: false };
    }
    return { key: `argb:${argb}`, argb, isApprox: false };
  }

  if (typeof c.theme === 'number') {
    const base = THEME_COLORS[c.theme] ?? '808080';
    const tint = c.tint ?? 0;
    return {
      key: `theme:${c.theme}${tint ? `+${tint}` : ''}`,
      argb: `FF${applyTint(base, tint)}`,
      isApprox: true,
    };
  }

  if (typeof c.indexed === 'number') {
    const base = INDEXED_COLORS[c.indexed];
    // 64/65 は「自動」を意味し、実際の色ではない
    if (!base) return null;
    return { key: `indexed:${c.indexed}`, argb: `FF${base}`, isApprox: true };
  }

  return null;
}

/** UI のパレットで選んだ色 (argb 直接指定) の key */
export function argbKey(argb: string): string {
  return `argb:${argb.toUpperCase()}`;
}
