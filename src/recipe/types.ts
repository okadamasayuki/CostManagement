import type { OpScope } from '../excel/types';

/**
 * 手順書 (レシピ) のデータ構造。
 *
 * 毎年ほぼ同じ作業を繰り返す運用を想定しているため、
 * 「対象ファイルを固定名で持たない」ことを重視している。
 * ブック/シートはファイル名パターン (glob) で指定でき、年度は
 * 「+1 年ずらす」という相対指定ができる。これにより同じ手順書を
 * 翌年もそのまま読み込んで実行できる。
 */

export const RECIPE_FORMAT = 'excel-lock-manager/recipe';
export const RECIPE_VERSION = 1;

/** 操作対象のセル範囲の指定方法 */
export type RangeSpec =
  /** 'A1:C5' のような固定アドレス */
  | { kind: 'a1'; a1: string }
  /** データが入っている範囲全体 */
  | { kind: 'used' }
  /** シート全体 */
  | { kind: 'sheet' };

export interface YearTargets {
  /** セルの値 (文字列・数値) を対象にする */
  values: boolean;
  /** 数式の中身も置換する */
  formulas: boolean;
  /** シート名も置換する */
  sheetNames: boolean;
  /** 保存時のファイル名も置換する */
  fileNames: boolean;
  /** 「令和6年」のような和暦も一緒にずらす */
  japaneseEra: boolean;
}

export const DEFAULT_YEAR_TARGETS: YearTargets = {
  values: true,
  formulas: true,
  sheetNames: true,
  fileNames: false,
  japaneseEra: false,
};

export interface SheetProtectOptions {
  /** ロックされたセルを選択できるようにする */
  selectLockedCells: boolean;
  /** 書式変更を許可する */
  formatCells: boolean;
  /** 行/列の挿入・削除を許可する */
  insertRows: boolean;
  deleteRows: boolean;
  /** オートフィルタの操作を許可する */
  autoFilter: boolean;
  /** 並べ替えを許可する */
  sort: boolean;
}

export const DEFAULT_PROTECT_OPTIONS: SheetProtectOptions = {
  selectLockedCells: true,
  formatCells: false,
  insertRows: false,
  deleteRows: false,
  autoFilter: false,
  sort: false,
};

export type StepBody =
  /** 指定範囲のロックを ON/OFF する */
  | { op: 'setLock'; range: RangeSpec; locked: boolean }
  /** 指定範囲「以外」をロックする (=指定範囲だけ入力可能にする) */
  | { op: 'lockAllExcept'; range: RangeSpec; alsoUnlockTarget: boolean }
  /** シート保護を有効化 */
  | { op: 'protectSheet'; password?: string; options: SheetProtectOptions }
  /** シート保護を解除 */
  | { op: 'unprotectSheet' }
  /** 指定範囲を塗る / 塗りを消す (colorArgb=null で消去) */
  | { op: 'fillRange'; range: RangeSpec; colorArgb: string | null }
  /**
   * 塗りつぶしの色を手がかりにロックを切り替える。
   * 「黄色が入力欄」のような既存の色分け運用があるファイルに対して、
   * 色からロック設定を起こせる。
   */
  | {
      op: 'setLockByFill';
      /** 対象にする色の key (excel/color.ts の FillRef.key) */
      colorKeys: string[];
      /** 人が読むための色名。実行時の判定には使わない。 */
      colorLabels?: string[];
      /** match='in' … 指定色のセル / 'out' … 指定色以外のセル */
      match: 'in' | 'out';
      locked: boolean;
      /**
       * match='out' のとき、指定色のセルを逆の状態にするか。
       *
       * Excel のセルは既定で全てロック済みのため、「指定色以外をロック」
       * だけでは何も変わらず、肝心の指定色のセルもロックされたままになる。
       * 「この色の欄だけ入力させたい」という使い方をそのまま実現するために、
       * 既定で有効にする。
       */
      alsoSetMatched?: boolean;
      /** 'out' のとき、色の付いていないセルも対象に含めるか */
      includeUnfilled: boolean;
      range: RangeSpec;
    }
  /** ロック状態に応じて一括で塗る */
  | {
      op: 'fillByLockState';
      target: 'locked' | 'unlocked';
      colorArgb: string | null;
      onlyUsedRange: boolean;
    }
  /** 年を相対的にずらす (2024 -> 2025 など)。同時置換なので連鎖しない。 */
  | {
      op: 'shiftYears';
      delta: number;
      minYear: number;
      maxYear: number;
      wholeNumberOnly: boolean;
      targets: YearTargets;
      range: RangeSpec;
    }
  /** 明示的な対応表で年を置換する */
  | {
      op: 'mapYears';
      pairs: Array<{ from: number; to: number }>;
      wholeNumberOnly: boolean;
      targets: YearTargets;
      range: RangeSpec;
    }
  /** 汎用の文字列置換 */
  | {
      op: 'replaceText';
      find: string;
      replace: string;
      matchCase: boolean;
      wholeCell: boolean;
      targets: Pick<YearTargets, 'values' | 'formulas' | 'sheetNames' | 'fileNames'>;
      range: RangeSpec;
    };

export interface RecipeStep {
  id: string;
  /** 手順書に出る見出し。自動生成されるが編集できる。 */
  label: string;
  /** 担当者向けの補足メモ */
  note?: string;
  /** この手順を実行するかどうか (手順書上で一時的に外せる) */
  enabled: boolean;
  scope: OpScope;
  body: StepBody;
}

export interface Recipe {
  format: typeof RECIPE_FORMAT;
  version: number;
  title: string;
  description: string;
  createdAt: string;
  /** 作成時に読み込んでいたファイルの一覧 (参考情報。実行時の対象は scope で決まる) */
  sourceHint: string[];
  steps: RecipeStep[];
}

export function emptyRecipe(): Recipe {
  return {
    format: RECIPE_FORMAT,
    version: RECIPE_VERSION,
    title: '年次更新手順',
    description: '',
    createdAt: new Date().toISOString(),
    sourceHint: [],
    steps: [],
  };
}
