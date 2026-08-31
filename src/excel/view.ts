import type ExcelJS from 'exceljs';
import type { CellKind, CellView, SheetView } from './types';
import { applyNumFmt, charWidthToPx, colorToArgb, pointsToPx } from './format';
import { parseA1Range } from './cellRef';
import {
  asRecord,
  definedExtent,
  forEachExistingCell,
  getMerges,
  getSheetProtection,
  hasAnyStyle,
  readColWidth,
  readRowHeight,
} from './exceljsCompat';
import { EXCEL_MAX_COLS, EXCEL_MAX_ROWS, buildAxis } from './axis';

/**
 * 画面用に保持するセル数の上限。
 *
 * 行・列の「番号」に上限は設けない (以前は 200 列 = GR 列で切っていた)。
 * ただし極端に大きなシートで画面用データを作りすぎるとブラウザーが
 * 固まるため、セルの「数」だけを上限とする。
 * 上限に達した場合は画面表示が一部欠けるが、ロックや置換などの処理は
 * 画面用データを使わないので、ファイル全体に対して正しく適用される。
 */
const MAX_VIEW_CELLS = 400_000;

/** 既定のサイズ (Excel の標準に合わせる) */
const DEFAULT_ROW_PX = 20;
const DEFAULT_COL_PX = 72;

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

  // 値の無いセルでも、書式やロック設定を持っていれば画面に出す。
  // (「ロックを外したが空欄」のセルが、ロック済みに見えてしまうのを防ぐ)
  let truncated = false;
  forEachExistingCell(ws, EXCEL_MAX_ROWS, EXCEL_MAX_COLS, (cell, rowNumber, colNumber) => {
    if (cells.size >= MAX_VIEW_CELLS) {
      truncated = true;
      return;
    }
    const numFmt = cell.numFmt;
    const extracted = extractCellValue(cell.value, numFmt);
    const styled = hasAnyStyle(cell);
    // 値も書式も無いセルは描画しない (数が膨大になるため)
    if (extracted.kind === 'blank' && !styled) return;

    maxRow = Math.max(maxRow, rowNumber);
    maxCol = Math.max(maxCol, colNumber);

    const locked = isCellLocked(cell);
    if (locked) lockedCount++;
    else unlockedCount++;
    if (extracted.kind !== 'blank') usedCellCount++;

    const fill = cell.fill as { type?: string; fgColor?: unknown } | undefined;
    const fillArgb = fill && fill.type === 'pattern' ? colorToArgb(fill.fgColor) : undefined;

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

  // 結合セルを反映する
  const merges = getMerges(ws);
  for (const range of merges) {
    const rect = parseA1Range(range);
    if (!rect) continue;
    for (let r = rect.top; r <= rect.bottom; r++) {
      for (let c = rect.left; c <= rect.right; c++) {
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
    maxRow = Math.max(maxRow, rect.bottom);
    maxCol = Math.max(maxCol, rect.right);
  }

  const contentBottom = maxRow;
  const contentRight = maxCol;

  // 個別のサイズを持つ範囲。中身のある範囲に加え、ファイル側で
  // 幅や高さだけ指定されている行・列も拾う。
  const extent = definedExtent(ws);
  const measuredRows = Math.min(Math.max(maxRow, extent.rows), EXCEL_MAX_ROWS);
  const measuredCols = Math.min(Math.max(maxCol, extent.cols), EXCEL_MAX_COLS);

  const colSizes: Array<number | undefined> = [undefined];
  for (let c = 1; c <= measuredCols; c++) {
    const w = readColWidth(ws, c);
    colSizes.push(w === undefined ? undefined : charWidthToPx(w));
  }
  const rowSizes: Array<number | undefined> = [undefined];
  for (let r = 1; r <= measuredRows; r++) {
    const h = readRowHeight(ws, r);
    rowSizes.push(h === undefined ? undefined : pointsToPx(h));
  }

  const rows = buildAxis(rowSizes, measuredRows, EXCEL_MAX_ROWS, DEFAULT_ROW_PX);
  const cols = buildAxis(colSizes, measuredCols, EXCEL_MAX_COLS, DEFAULT_COL_PX);

  const sp = getSheetProtection(ws);

  return {
    name: ws.name,
    index,
    rowCount: EXCEL_MAX_ROWS,
    colCount: EXCEL_MAX_COLS,
    cells,
    rows,
    cols,
    truncated,
    state: (ws.state as SheetView['state']) ?? 'visible',
    isProtected: Boolean(sp && sp.sheet !== false),
    hasPassword: Boolean(sp && (sp.hashValue || sp.password)),
    lockedCount,
    unlockedCount,
    usedCellCount,
    contentBottom,
    contentRight,
  };
}

function emptyCellView(): CellView {
  return { text: '', raw: '', kind: 'blank', locked: true };
}
