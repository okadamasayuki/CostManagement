import type ExcelJS from 'exceljs';
import type { CellKind, CellView, SheetView } from './types';
import { applyNumFmt, charWidthToPx, colorToArgb, pointsToPx } from './format';
import { parseA1Range } from './cellRef';
import { asRecord, getMerges, getSheetProtection } from './exceljsCompat';

/** 空のシートでも最低これだけの行/列は描画して、Excel らしい見た目を保つ */
const MIN_ROWS = 40;
const MIN_COLS = 20;
/** 描画が重くなりすぎないための上限 (超過分はスクロールしても表示しない) */
const MAX_RENDER_ROWS = 5000;
const MAX_RENDER_COLS = 200;

interface Extracted {
  text: string;
  raw: string;
  kind: CellKind;
}

/** ExcelJS の多様な CellValue から表示用テキストと数式バー用の文字列を作る */
export function extractCellValue(value: ExcelJS.CellValue, numFmt?: string): Extracted {
  if (value === null || value === undefined || value === '') {
    return { text: '', raw: '', kind: 'blank' };
  }
  if (typeof value === 'number') {
    return { text: applyNumFmt(value, numFmt), raw: String(value), kind: 'number' };
  }
  if (typeof value === 'boolean') {
    return { text: value ? 'TRUE' : 'FALSE', raw: String(value), kind: 'other' };
  }
  if (typeof value === 'string') {
    return { text: value, raw: value, kind: 'text' };
  }
  if (value instanceof Date) {
    return { text: applyNumFmt(value, numFmt), raw: value.toISOString(), kind: 'date' };
  }
  if (typeof value === 'object') {
    const v = asRecord(value);

    if ('formula' in v || 'sharedFormula' in v) {
      const formula = (v.formula ?? v.sharedFormula) as string;
      const result = v.result;
      const shown =
        result === null || result === undefined
          ? ''
          : extractCellValue(result as ExcelJS.CellValue, numFmt).text;
      return { text: shown, raw: `=${formula}`, kind: 'formula' };
    }
    if ('richText' in v && Array.isArray(v.richText)) {
      const text = (v.richText as Array<{ text?: string }>).map((r) => r.text ?? '').join('');
      return { text, raw: text, kind: 'text' };
    }
    if ('text' in v) {
      const text = String(v.text ?? '');
      return { text, raw: text, kind: 'text' };
    }
    if ('error' in v) {
      return { text: String(v.error), raw: String(v.error), kind: 'error' };
    }
  }
  const s = String(value);
  return { text: s, raw: s, kind: 'other' };
}

/** ExcelJS のセルから「ロックされているか」を読む。Excel の既定は locked=true。 */
export function isCellLocked(cell: ExcelJS.Cell): boolean {
  const p = cell.protection as { locked?: boolean } | undefined;
  if (p && typeof p.locked === 'boolean') return p.locked;
  return true;
}

function alignOf(cell: ExcelJS.Cell, kind: CellKind): 'left' | 'center' | 'right' | undefined {
  const h = cell.alignment?.horizontal;
  if (h === 'center' || h === 'centerContinuous') return 'center';
  if (h === 'right') return 'right';
  if (h === 'left') return 'left';
  // 未指定なら Excel の既定 (数値は右寄せ)
  if (kind === 'number' || kind === 'date') return 'right';
  return undefined;
}

/**
 * ワークシートを描画用のスナップショットへ変換する。
 * 操作のたびに作り直す前提なので、副作用は持たない。
 */
export function buildSheetView(ws: ExcelJS.Worksheet, index: number): SheetView {
  const cells = new Map<string, CellView>();
  let maxRow = 1;
  let maxCol = 1;
  let lockedCount = 0;
  let unlockedCount = 0;
  let usedCellCount = 0;

  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber > MAX_RENDER_ROWS) return;
    maxRow = Math.max(maxRow, rowNumber);
    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      if (colNumber > MAX_RENDER_COLS) return;
      maxCol = Math.max(maxCol, colNumber);

      const numFmt = cell.numFmt;
      const extracted = extractCellValue(cell.value, numFmt);
      const locked = isCellLocked(cell);
      if (locked) lockedCount++;
      else unlockedCount++;
      if (extracted.kind !== 'blank') usedCellCount++;

      const fill = cell.fill as { type?: string; fgColor?: unknown } | undefined;
      const fillArgb =
        fill && fill.type === 'pattern' ? colorToArgb(fill.fgColor) : undefined;

      cells.set(`${rowNumber}:${colNumber}`, {
        text: extracted.text,
        raw: extracted.raw,
        kind: extracted.kind,
        locked,
        fillArgb,
        fontArgb: colorToArgb(cell.font?.color),
        bold: cell.font?.bold ?? undefined,
        italic: cell.font?.italic ?? undefined,
        align: alignOf(cell, extracted.kind),
        numFmt,
      });
    });
  });

  // 結合セルを反映する
  const merges = getMerges(ws);
  for (const range of merges) {
    const rect = parseA1Range(range);
    if (!rect) continue;
    for (let r = rect.top; r <= rect.bottom && r <= MAX_RENDER_ROWS; r++) {
      for (let c = rect.left; c <= rect.right && c <= MAX_RENDER_COLS; c++) {
        const key = `${r}:${c}`;
        const existing = cells.get(key);
        if (r === rect.top && c === rect.left) {
          const base = existing ?? emptyCellView();
          cells.set(key, {
            ...base,
            mergeSpan: { rows: rect.bottom - rect.top + 1, cols: rect.right - rect.left + 1 },
          });
        } else {
          cells.set(key, { ...(existing ?? emptyCellView()), mergedHidden: true });
        }
      }
    }
    maxRow = Math.max(maxRow, Math.min(rect.bottom, MAX_RENDER_ROWS));
    maxCol = Math.max(maxCol, Math.min(rect.right, MAX_RENDER_COLS));
  }

  const rowCount = Math.min(Math.max(maxRow + 8, MIN_ROWS), MAX_RENDER_ROWS);
  const colCount = Math.min(Math.max(maxCol + 3, MIN_COLS), MAX_RENDER_COLS);

  const colWidths: number[] = [0];
  for (let c = 1; c <= colCount; c++) {
    colWidths.push(charWidthToPx(ws.getColumn(c)?.width));
  }
  const rowHeights: number[] = [0];
  for (let r = 1; r <= rowCount; r++) {
    rowHeights.push(pointsToPx(ws.getRow(r)?.height));
  }

  const sp = getSheetProtection(ws);

  return {
    name: ws.name,
    index,
    rowCount,
    colCount,
    cells,
    colWidths,
    rowHeights,
    state: (ws.state as SheetView['state']) ?? 'visible',
    isProtected: Boolean(sp && sp.sheet !== false),
    hasPassword: Boolean(sp && (sp.hashValue || sp.password)),
    lockedCount,
    unlockedCount,
    usedCellCount,
  };
}

function emptyCellView(): CellView {
  return { text: '', raw: '', kind: 'blank', locked: true };
}
