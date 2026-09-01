import type ExcelJS from 'exceljs';
import type { LoadedWorkbook } from './types';
import { colToLetter } from './cellRef';
import { asRecord, forEachExistingCell } from './exceljsCompat';
import { extractCellValue } from './view';

/**
 * 2 年分のファイルを見比べて「毎年書き換えられている欄」を見つける。
 *
 * 考え方:
 *   同じ様式のファイルを 2 年分並べると、
 *     ・値が変わっているセル … 担当部署が毎年記入している欄
 *     ・値が同じセル         … 様式 (見出し、費目名、固定値)
 *   という区別ができる。これを使って「記入欄だけ入力できるようにし、
 *   様式はロックする」を自動で組み立てる。
 *
 * 注意している点:
 *   ・年度の数字だけの違い (2023年度 → 2024年度) は様式とみなす。
 *     これを入れないと、見出しまで記入欄と判定されてしまう。
 *   ・数式は「式そのもの」で比べる。合計欄は結果が毎年変わるが
 *     記入欄ではないため。
 */

/** 年らしい 4 桁の数字を伏せて、年違いを無視した形にする */
export function normalizeYears(text: string, minYear = 1990, maxYear = 2100): string {
  return text.replace(/\d{4}/g, (m) => {
    const n = parseInt(m, 10);
    return n >= minYear && n <= maxYear ? '####' : m;
  });
}

/**
 * 行やセルを「作らずに」読む。
 *
 * ExcelJS の getRow() / getCell() は存在しない行・セルをその場で作るため、
 * 見比べただけのつもりが古い方のファイルの dimension を書き換えてしまう。
 * 比較は読むだけなので、必ずこちらを使う。
 */
function findCell(ws: ExcelJS.Worksheet, row: number, col: number): ExcelJS.Cell | null {
  const r = ws.findRow(row);
  if (!r) return null;
  return r.findCell(col) ?? null;
}

/** 例として見せるときの表示。人が Excel で見るのと同じ形にする。 */
function cellDisplay(cell: ExcelJS.Cell | null): string {
  const v = cell?.value;
  if (!cell || v === null || v === undefined || v === '') return '(空欄)';
  const ex = extractCellValue(v, cell.numFmt);
  // 数式で計算結果が保存されていないときは、式そのものを見せる
  return ex.text !== '' ? ex.text : ex.raw !== '' ? ex.raw : '(空欄)';
}

/** 比較のためのセルの中身。数式は式そのものを見る。 */
function cellSignature(cell: ExcelJS.Cell, compareFormulaText: boolean): string | null {
  const v = cell.value;
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'object' && v !== null) {
    const o = asRecord(v);
    if ('formula' in o) {
      return compareFormulaText ? `=${String(o.formula)}` : String(o.result ?? '');
    }
    if ('sharedFormula' in o) {
      return compareFormulaText ? `=shared:${String(o.sharedFormula)}` : String(o.result ?? '');
    }
  }
  return extractCellValue(v, cell.numFmt).raw;
}

export interface DetectOptions {
  /** 年度の数字だけの違いは様式とみなす */
  ignoreYearOnly: boolean;
  /** 数式は式そのもので比べる (合計欄などを記入欄と誤判定しないため) */
  compareFormulaText: boolean;
}

export interface SheetHits {
  /** 変化していたセル (A1 形式) */
  changed: Set<string>;
  /** 変化していなかったセル */
  unchanged: Set<string>;
}

export interface DetectSample {
  book: string;
  sheet: string;
  addr: string;
  before: string;
  after: string;
}

export interface DetectResult {
  /** 比較できたファイルの組 */
  pairCount: number;
  /** 比較できたシート数 */
  sheetCount: number;
  changedCount: number;
  unchangedCount: number;
  /** 年の違いだけだったセル数 (様式として扱った) */
  yearOnlyCount: number;
  /** `${bookId}::${シート名}` -> 判定結果。新しい方のブックに対して持つ。 */
  hits: Map<string, SheetHits>;
  samples: DetectSample[];
  /** 比較相手が見つからなかったファイル */
  unpaired: string[];
}

/** ファイル名やパスから年を取り出す (一番大きい 4 桁を採る) */
function yearOf(path: string): number {
  let best = 0;
  for (const m of path.matchAll(/\d{4}/g)) {
    const n = parseInt(m[0], 10);
    if (n >= 1990 && n <= 2100) best = Math.max(best, n);
  }
  return best;
}

/**
 * 年違いのファイル同士を組にする。
 * パスの年を伏せた形が同じものを「同じ様式のファイル」とみなす。
 */
export function pairByYear(books: LoadedWorkbook[]): {
  pairs: Array<{ older: LoadedWorkbook; newer: LoadedWorkbook }>;
  unpaired: LoadedWorkbook[];
} {
  const groups = new Map<string, LoadedWorkbook[]>();
  for (const b of books) {
    if (b.loadError) continue;
    const key = normalizeYears(b.relPath);
    const list = groups.get(key);
    if (list) list.push(b);
    else groups.set(key, [b]);
  }

  const pairs: Array<{ older: LoadedWorkbook; newer: LoadedWorkbook }> = [];
  const unpaired: LoadedWorkbook[] = [];
  for (const list of groups.values()) {
    if (list.length < 2) {
      unpaired.push(...list);
      continue;
    }
    const sorted = [...list].sort((a, b) => yearOf(a.relPath) - yearOf(b.relPath));
    // 3 年分あるときは隣り合う組すべてを見る
    for (let i = 0; i + 1 < sorted.length; i++) {
      pairs.push({ older: sorted[i], newer: sorted[i + 1] });
    }
  }
  return { pairs, unpaired };
}

/** 年違いを無視してシート名を対応づける */
function pairSheets(
  older: ExcelJS.Workbook,
  newer: ExcelJS.Workbook,
): Array<{ a: ExcelJS.Worksheet; b: ExcelJS.Worksheet }> {
  const byKey = new Map<string, ExcelJS.Worksheet>();
  older.eachSheet((ws) => byKey.set(normalizeYears(ws.name), ws));
  const out: Array<{ a: ExcelJS.Worksheet; b: ExcelJS.Worksheet }> = [];
  newer.eachSheet((ws) => {
    const a = byKey.get(normalizeYears(ws.name));
    if (a) out.push({ a, b: ws });
  });
  return out;
}

const MAX_SAMPLES = 30;

/** 2 年分を比べて、記入されている欄を見つける */
export function detectInputCells(books: LoadedWorkbook[], opts: DetectOptions): DetectResult {
  const { pairs, unpaired } = pairByYear(books);
  const hits = new Map<string, SheetHits>();
  const samples: DetectSample[] = [];
  let sheetCount = 0;
  let changedCount = 0;
  let unchangedCount = 0;
  let yearOnlyCount = 0;

  for (const { older, newer } of pairs) {
    for (const { a, b } of pairSheets(older.wb, newer.wb)) {
      sheetCount++;
      const key = `${newer.id}::${b.name}`;
      const sheet: SheetHits = hits.get(key) ?? { changed: new Set(), unchanged: new Set() };
      hits.set(key, sheet);

      // 新しい方に存在するセルを基準に見る
      forEachExistingCell(b, 1_048_576, 16_384, (cellB, row, col) => {
        const addr = `${colToLetter(col)}${row}`;
        const sigB = cellSignature(cellB, opts.compareFormulaText);
        const cellA = findCell(a, row, col);
        const sigA = cellA ? cellSignature(cellA, opts.compareFormulaText) : null;

        if (sigA === null && sigB === null) return; // どちらも空欄

        if (sigA === sigB) {
          sheet.unchanged.add(addr);
          unchangedCount++;
          return;
        }

        // 年の数字だけの違いなら様式とみなす
        if (
          opts.ignoreYearOnly &&
          sigA !== null &&
          sigB !== null &&
          normalizeYears(sigA) === normalizeYears(sigB)
        ) {
          sheet.unchanged.add(addr);
          yearOnlyCount++;
          unchangedCount++;
          return;
        }

        sheet.changed.add(addr);
        changedCount++;
        if (samples.length < MAX_SAMPLES) {
          samples.push({
            book: newer.relPath,
            sheet: b.name,
            addr,
            before: cellDisplay(cellA),
            after: cellDisplay(cellB),
          });
        }
      });
    }
  }

  return {
    pairCount: pairs.length,
    sheetCount,
    changedCount,
    unchangedCount,
    yearOnlyCount,
    hits,
    samples,
    unpaired: unpaired.map((b) => b.relPath),
  };
}
