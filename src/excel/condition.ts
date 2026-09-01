import type ExcelJS from 'exceljs';
import type { CellCondition } from '../recipe/types';
import { extractCellValue } from './view';

/**
 * セルが条件に合うかを判定する。
 *
 * 「数値が入っているセル」「500,000 を超えるセル」「"合計" を含むセル」
 * のように、中身を見て対象を絞り込むために使う。
 *
 * 数式のセルは、判定には計算結果を使う (画面に見えている値と一致させるため)。
 * ただし kind='formula' は数式であること自体を条件にする。
 */
export function matchesCondition(cell: ExcelJS.Cell, cond: CellCondition): boolean {
  const v = extractCellValue(cell.value, cell.numFmt);
  const isFormula = v.kind === 'formula';

  // 種類の判定
  switch (cond.kind) {
    case 'blank':
      return v.kind === 'blank';
    case 'formula':
      if (!isFormula) return false;
      break;
    case 'number':
      if (numericValue(cell) === null) return false;
      break;
    case 'text':
      // 数式でも結果が文字なら文字として扱う
      if (v.kind === 'blank' || numericValue(cell) !== null) return false;
      break;
    case 'any':
      break;
  }

  if (cond.number) {
    const n = numericValue(cell);
    if (n === null) return false;
    if (!compareNumber(n, cond.number)) return false;
  }

  if (cond.text) {
    if (!compareText(v.text, cond.text)) return false;
  }

  return true;
}

/** セルの数値。数式なら計算結果を見る。数値でなければ null。 */
function numericValue(cell: ExcelJS.Cell): number | null {
  const raw = cell.value;
  if (typeof raw === 'number') return raw;
  if (raw && typeof raw === 'object') {
    const o = raw as { result?: unknown };
    if (typeof o.result === 'number') return o.result;
  }
  return null;
}

function compareNumber(n: number, c: NonNullable<CellCondition['number']>): boolean {
  switch (c.op) {
    case 'gt':
      return n > c.a;
    case 'ge':
      return n >= c.a;
    case 'lt':
      return n < c.a;
    case 'le':
      return n <= c.a;
    case 'eq':
      return n === c.a;
    case 'ne':
      return n !== c.a;
    case 'between': {
      const lo = Math.min(c.a, c.b ?? c.a);
      const hi = Math.max(c.a, c.b ?? c.a);
      return n >= lo && n <= hi;
    }
  }
}

function compareText(text: string, c: NonNullable<CellCondition['text']>): boolean {
  const a = c.matchCase ? text : text.toLowerCase();
  const b = c.matchCase ? c.value : c.value.toLowerCase();
  if (!b) return true;
  switch (c.op) {
    case 'contains':
      return a.includes(b);
    case 'startsWith':
      return a.startsWith(b);
    case 'endsWith':
      return a.endsWith(b);
    case 'equals':
      return a === b;
  }
}
