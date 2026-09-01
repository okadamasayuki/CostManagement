import type ExcelJS from 'exceljs';
import type { RangeRect } from './cellRef';
import type { FileSystemDirectoryHandleLike } from './fsTypes';
import type { AxisMetrics } from './axis';

/**
 * 読み込んだブック 1 冊。
 * `wb` は ExcelJS のライブなモデルで、各種操作はこれを直接書き換える。
 */
export interface LoadedWorkbook {
  id: string;
  /** 選んだフォルダからの相対パス (単一ファイル選択時はファイル名のみ) */
  relPath: string;
  fileName: string;
  wb: ExcelJS.Workbook;
  /** File System Access API 経由で開いた場合のみ。元ファイルへ上書き保存できる。 */
  handle?: FileSystemFileHandle;
  /**
   * 元ファイルが置かれているフォルダー。ファイル名を変えて保存する際に、
   * 同じフォルダーへ新しい名前で作成するために使う。
   */
  parentDir?: FileSystemDirectoryHandleLike;
  /**
   * 読み込み元の File。File はディスク上のファイルへの参照でメモリを
   * ほとんど使わないため保持しておき、「元に戻す」で再読み込みできる。
   */
  sourceFile?: File;
  dirty: boolean;
  sizeBytes: number;
  /** 読み込みに失敗した場合の理由 */
  loadError?: string;
}

/** 読み込み時にスキップしたファイルの記録 */
export interface SkippedFile {
  relPath: string;
  reason: string;
}

export type CellKind = 'blank' | 'number' | 'text' | 'date' | 'formula' | 'error' | 'other';

/** 描画用に切り出した 1 セルの情報 */
export interface CellView {
  /** 画面に表示する文字列 */
  text: string;
  /** 数式バーに出す元の内容 ('=SUM(A1:A3)' など) */
  raw: string;
  kind: CellKind;
  /** Excel の「セルのロック」属性。未指定は true (Excel の既定) */
  locked: boolean;
  /** 塗りつぶし色 'FFFFFF00' 形式。無地は undefined */
  fillArgb?: string;
  fontArgb?: string;
  bold?: boolean;
  italic?: boolean;
  align?: 'left' | 'center' | 'right';
  numFmt?: string;
  /** 結合セルの左上以外 (描画をスキップする) */
  mergedHidden?: boolean;
  /** 結合セルの左上の場合の広がり */
  mergeSpan?: { rows: number; cols: number };
}

export interface SheetView {
  name: string;
  index: number;
  rowCount: number;
  colCount: number;
  /** キーは `${row}:${col}` */
  cells: Map<string, CellView>;
  /** 行の位置とサイズ */
  rows: AxisMetrics;
  /** 列の位置とサイズ */
  cols: AxisMetrics;
  /** 画面用データが上限に達して一部省かれたか (処理自体には影響しない) */
  truncated: boolean;
  state: 'visible' | 'hidden' | 'veryHidden';
  /** シート保護が有効か */
  isProtected: boolean;
  /** シート保護にパスワードが掛かっているか */
  hasPassword: boolean;
  lockedCount: number;
  unlockedCount: number;
  /** 値または数式が入っているセル数 */
  usedCellCount: number;
  /**
   * シートに実際に中身 (値または書式) がある範囲の右下。
   * 画面には見栄えのため余分な行・列も描くが、そこは
   * シートの内容ではないのでロック状態の網掛けを出さない。
   */
  contentBottom: number;
  contentRight: number;
}

/** 操作の適用範囲 */
export interface OpScope {
  books: 'current' | 'all' | 'glob' | 'folder' | 'selected';
  /** books==='glob' のときのファイル名パターン (例: `*原価*.xlsx`) */
  bookGlob?: string;
  /**
   * books==='folder' のときの対象フォルダー (読み込み元からの相対パス)。
   * そのフォルダーの配下 (孫フォルダー以下も含む) が全て対象になる。
   * 空文字は最上位 = 全ブック。
   */
  bookFolder?: string;
  sheets: 'current' | 'all' | 'glob';
  /** sheets==='glob' のときのシート名パターン */
  sheetGlob?: string;
}

export const DEFAULT_SCOPE: OpScope = { books: 'current', sheets: 'current' };

/** 1 回の操作の実行結果 */
export interface OpResult {
  /** 変更されたセル数 / シート数など、人間向けの要約 */
  summary: string;
  changedCells: number;
  changedSheets: number;
  changedBooks: number;
  details: OpDetail[];
}

export interface OpDetail {
  book: string;
  sheet: string;
  message: string;
  count: number;
}

export interface TargetRange {
  /** 明示的な範囲。未指定なら「使用範囲全体」 */
  rect?: RangeRect;
}
