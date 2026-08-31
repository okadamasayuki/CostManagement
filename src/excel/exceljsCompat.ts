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
