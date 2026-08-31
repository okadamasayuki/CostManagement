import type ExcelJS from 'exceljs';

/**
 * ExcelJS の型定義に含まれていないが実行時には存在するプロパティへの
 * アクセスをここに集約する。キャストを各所に散らかさないため。
 */

export interface SheetProtectionModel {
  sheet?: boolean;
  password?: string;
  hashValue?: string;
  algorithmName?: string;
  saltValue?: string;
  spinCount?: number;
}

/** ws.sheetProtection (未保護なら null) */
export function getSheetProtection(ws: ExcelJS.Worksheet): SheetProtectionModel | null {
  return (ws as unknown as { sheetProtection: SheetProtectionModel | null }).sheetProtection ?? null;
}

/** ws.model.merges ('A1:B2' の配列) */
export function getMerges(ws: ExcelJS.Worksheet): string[] {
  return (ws.model as unknown as { merges?: string[] }).merges ?? [];
}

/** CellValue をオブジェクトとして覗くためのキャスト */
export function asRecord(value: unknown): Record<string, unknown> {
  return value as Record<string, unknown>;
}

/** 部分的なスタイルを ExcelJS.Style として渡すためのキャスト */
export function asStyle(style: Record<string, unknown>): ExcelJS.Style {
  return style as unknown as ExcelJS.Style;
}

/**
 * 既に存在するセルだけを走査する。
 *
 * ExcelJS の eachRow/eachCell は、
 *   - includeEmpty: false … 値の無いセルを飛ばす
 *     (値は無いが書式やロック設定を持つセルまで見えなくなる)
 *   - includeEmpty: true  … getRow/getCell で行やセルを新たに作ってしまう
 *     (保存するファイルに空の行が増える)
 * のどちらも都合が悪いため、内部の配列を読み取り専用で走査する。
 *
 * 「値は無いがロックを外したセル」「値は無いが色を付けたセル」を
 * 画面に正しく出すために必要。
 */
export function forEachExistingCell(
  ws: ExcelJS.Worksheet,
  maxRow: number,
  maxCol: number,
  fn: (cell: ExcelJS.Cell, row: number, col: number) => void,
): void {
  const rows = (ws as unknown as { _rows?: Array<{ number: number; _cells?: unknown[] } | undefined> })
    ._rows;
  if (!rows) return;
  for (const row of rows) {
    if (!row || row.number > maxRow) continue;
    const cells = row._cells;
    if (!cells) continue;
    for (let i = 0; i < cells.length; i++) {
      const cell = cells[i];
      if (!cell) continue;
      const col = i + 1;
      if (col > maxCol) continue;
      fn(cell as ExcelJS.Cell, row.number, col);
    }
  }
}

/** 書式が 1 つでも設定されているか */
export function hasAnyStyle(cell: ExcelJS.Cell): boolean {
  const style = cell.style as Record<string, unknown> | undefined;
  return Boolean(style) && Object.keys(style as object).length > 0;
}
