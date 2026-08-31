import type ExcelJS from 'exceljs';
import type { LoadedWorkbook, OpDetail, OpResult, OpScope } from './types';
import type { RangeSpec, RecipeStep, SheetProtectOptions, StepBody } from '../recipe/types';
import type { RangeRect } from './cellRef';
import { colToLetter, parseA1Range, rectToA1 } from './cellRef';
import { asRecord, asStyle, getSheetProtection } from './exceljsCompat';
import { resolveColor, type FillRef } from './color';
import { isUnderFolder } from './folders';
import {
  mapNumericYear,
  pairMapper,
  replaceEraYears,
  replaceYearsInString,
  shiftMapper,
  type YearMapper,
} from './yearShift';

/** 1 シートあたりの走査上限。巨大シートでブラウザが固まるのを防ぐ。 */
const MAX_CELLS_PER_SHEET = 400_000;

export interface OpContext {
  books: LoadedWorkbook[];
  currentBookId: string | null;
  currentSheetName: string | null;
  /** グリッド上の選択範囲 */
  selection: RangeRect | null;
}

export interface ApplyOptions {
  /** true のとき実際には書き換えず、件数の集計だけ行う */
  dryRun?: boolean;
}

interface SheetTarget {
  book: LoadedWorkbook;
  ws: ExcelJS.Worksheet;
}

// ---------------------------------------------------------------------------
// 対象の解決
// ---------------------------------------------------------------------------

function globToRegExp(glob: string): RegExp {
  const escaped = glob
    .replace(/[.+^${}()|[\]\\]/g, '\\$&')
    .replace(/\*/g, '.*')
    .replace(/\?/g, '.');
  return new RegExp(`^${escaped}$`, 'i');
}

export function matchesGlob(name: string, glob: string | undefined): boolean {
  if (!glob || !glob.trim()) return true;
  // カンマ区切りで複数パターンを許可する
  return glob
    .split(',')
    .map((g) => g.trim())
    .filter(Boolean)
    .some((g) => globToRegExp(g).test(name));
}

export function resolveTargets(ctx: OpContext, scope: OpScope): SheetTarget[] {
  let books: LoadedWorkbook[];
  switch (scope.books) {
    case 'current':
      books = ctx.books.filter((b) => b.id === ctx.currentBookId);
      break;
    case 'glob':
      books = ctx.books.filter(
        (b) => matchesGlob(b.fileName, scope.bookGlob) || matchesGlob(b.relPath, scope.bookGlob),
      );
      break;
    case 'folder':
      books = ctx.books.filter((b) => isUnderFolder(b.relPath, scope.bookFolder ?? ''));
      break;
    default:
      books = [...ctx.books];
  }
  books = books.filter((b) => !b.loadError);

  const targets: SheetTarget[] = [];
  for (const book of books) {
    book.wb.eachSheet((ws) => {
      if (scope.sheets === 'current') {
        // 「現在のシート」は現在のブックでのみ意味を持つ。
        // 複数ブックを対象にしている場合は同名シートを対象にする。
        if (ws.name !== ctx.currentSheetName) return;
      } else if (scope.sheets === 'glob') {
        if (!matchesGlob(ws.name, scope.sheetGlob)) return;
      }
      targets.push({ book, ws });
    });
  }
  return targets;
}

/** 実データが入っている範囲。空シートなら null */
export function usedRect(ws: ExcelJS.Worksheet): RangeRect | null {
  const d = ws.dimensions as
    | { top: number; left: number; bottom: number; right: number }
    | null
    | undefined;
  if (!d || !d.bottom || !d.right) return null;
  return {
    top: Math.max(1, d.top),
    left: Math.max(1, d.left),
    bottom: d.bottom,
    right: d.right,
  };
}

export function resolveRange(ws: ExcelJS.Worksheet, spec: RangeSpec): RangeRect | null {
  if (spec.kind === 'a1') {
    const rect = parseA1Range(spec.a1);
    if (!rect) return null;
    // 列全体/行全体の指定はデータ範囲でクリップして走査量を抑える
    const used = usedRect(ws);
    if (!used) return rect.bottom > 10000 || rect.right > 1000 ? null : rect;
    return {
      top: Math.max(rect.top, 1),
      left: Math.max(rect.left, 1),
      bottom: Math.min(rect.bottom, Math.max(used.bottom, rect.top)),
      right: Math.min(rect.right, Math.max(used.right, rect.left)),
    };
  }
  // 'used' も 'sheet' もデータ範囲を対象にする。
  // Excel ではセルの既定が「ロック済み」なので、データのない領域は
  // 触らなくても「ロックされている」状態になる。
  return usedRect(ws);
}

function rectCells(rect: RangeRect): number {
  return (rect.bottom - rect.top + 1) * (rect.right - rect.left + 1);
}

// ---------------------------------------------------------------------------
// スタイル書き換えのための安全なヘルパー
//
// ExcelJS は同じ styleId を持つセル間で style オブジェクトを共有している
// (StylesXform.getStyleModel が id 単位でキャッシュを返すため)。
// その場で書き換えると無関係なセルまで巻き添えになるので、
// 必ずコピーしてから差し替える。
// ---------------------------------------------------------------------------

type Style = Record<string, unknown>;

function setStylePart(cell: ExcelJS.Cell, part: string, value: unknown): void {
  const next: Style = { ...((cell.style as Style) ?? {}) };
  if (value === undefined) delete next[part];
  else next[part] = value;
  cell.style = asStyle(next);
}

function getLocked(cell: ExcelJS.Cell): boolean {
  const p = cell.protection as { locked?: boolean } | undefined;
  return p && typeof p.locked === 'boolean' ? p.locked : true;
}

function setLocked(cell: ExcelJS.Cell, locked: boolean): boolean {
  if (getLocked(cell) === locked) return false;
  const prev = (cell.protection as Record<string, unknown>) ?? {};
  setStylePart(cell, 'protection', { ...prev, locked });
  return true;
}

function getFillArgb(cell: ExcelJS.Cell): string | undefined {
  return getFillRef(cell)?.argb;
}

/**
 * セルの塗りつぶし色を解決して返す。塗りが無ければ null。
 * テーマ色・古い色番号も実際の色に解決されるため、
 * 「同じ色のセル」を色の指定方法によらず拾える。
 */
export function getFillRef(cell: ExcelJS.Cell): FillRef | null {
  const fill = cell.fill as { type?: string; pattern?: string; fgColor?: unknown } | undefined;
  if (!fill || fill.type !== 'pattern') return null;
  if (fill.pattern === 'none') return null;
  return resolveColor(fill.fgColor);
}

function setFill(cell: ExcelJS.Cell, argb: string | null): boolean {
  const current = getFillArgb(cell);
  if (argb === null) {
    if (!cell.fill) return false;
    setStylePart(cell, 'fill', undefined);
    return true;
  }
  if (current === argb) return false;
  setStylePart(cell, 'fill', {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb },
    bgColor: { argb },
  });
  return true;
}

/** rect 内の全セルを走査する。上限を超えた場合は打ち切って false を返す。 */
function forEachCell(
  ws: ExcelJS.Worksheet,
  rect: RangeRect,
  fn: (cell: ExcelJS.Cell, row: number, col: number) => void,
): boolean {
  if (rectCells(rect) > MAX_CELLS_PER_SHEET) return false;
  for (let r = rect.top; r <= rect.bottom; r++) {
    const row = ws.getRow(r);
    for (let c = rect.left; c <= rect.right; c++) {
      fn(row.getCell(c), r, c);
    }
  }
  return true;
}

// ---------------------------------------------------------------------------
// 各操作の実装
// ---------------------------------------------------------------------------

function opSetLock(
  t: SheetTarget,
  body: Extract<StepBody, { op: 'setLock' }>,
  dryRun: boolean,
): OpDetail | null {
  const rect = resolveRange(t.ws, body.range);
  if (!rect) return null;
  let count = 0;
  const ok = forEachCell(t.ws, rect, (cell) => {
    if (getLocked(cell) === body.locked) return;
    count++;
    if (!dryRun) setLocked(cell, body.locked);
  });
  if (!ok) {
    return detail(t, `範囲が大きすぎるため処理を中止しました (${rectToA1(rect)})`, 0);
  }
  if (count === 0) return null;
  return detail(
    t,
    `${rectToA1(rect)} の ${count} セルを${body.locked ? 'ロック' : 'ロック解除'}`,
    count,
  );
}

function opLockAllExcept(
  t: SheetTarget,
  body: Extract<StepBody, { op: 'lockAllExcept' }>,
  dryRun: boolean,
): OpDetail | null {
  const keep = resolveRange(t.ws, body.range);
  const all = usedRect(t.ws);
  if (!all) return null;
  // 例外範囲がデータ範囲の外にある場合も考慮して和を取る
  const scan: RangeRect = keep
    ? {
        top: Math.min(all.top, keep.top),
        left: Math.min(all.left, keep.left),
        bottom: Math.max(all.bottom, keep.bottom),
        right: Math.max(all.right, keep.right),
      }
    : all;

  let locked = 0;
  let unlocked = 0;
  const ok = forEachCell(t.ws, scan, (cell, r, c) => {
    const inKeep =
      keep !== null && r >= keep.top && r <= keep.bottom && c >= keep.left && c <= keep.right;
    if (inKeep) {
      if (!body.alsoUnlockTarget) return;
      if (getLocked(cell)) {
        unlocked++;
        if (!dryRun) setLocked(cell, false);
      }
    } else if (!getLocked(cell)) {
      locked++;
      if (!dryRun) setLocked(cell, true);
    }
  });
  if (!ok) return detail(t, '範囲が大きすぎるため処理を中止しました', 0);
  if (locked === 0 && unlocked === 0) return null;
  const parts: string[] = [];
  if (locked) parts.push(`${locked} セルをロック`);
  if (unlocked) parts.push(`${keep ? rectToA1(keep) : ''} の ${unlocked} セルを入力可能に`);
  return detail(t, parts.join(' / '), locked + unlocked);
}

async function opProtectSheet(
  t: SheetTarget,
  body: Extract<StepBody, { op: 'protectSheet' }>,
  dryRun: boolean,
): Promise<OpDetail | null> {
  if (dryRun) {
    return detail(t, `シート保護を有効化${body.password ? ' (パスワードあり)' : ''}`, 1);
  }
  await t.ws.protect(body.password ?? '', toExcelJsProtectOptions(body.options));
  return detail(t, `シート保護を有効化${body.password ? ' (パスワードあり)' : ''}`, 1);
}

function toExcelJsProtectOptions(o: SheetProtectOptions): Record<string, boolean> {
  // ExcelJS のオプションは「許可するか」を表す
  return {
    selectLockedCells: o.selectLockedCells,
    selectUnlockedCells: true,
    formatCells: o.formatCells,
    formatColumns: o.formatCells,
    formatRows: o.formatCells,
    insertRows: o.insertRows,
    insertColumns: o.insertRows,
    deleteRows: o.deleteRows,
    deleteColumns: o.deleteRows,
    autoFilter: o.autoFilter,
    sort: o.sort,
    insertHyperlinks: false,
    pivotTables: false,
    objects: false,
    scenarios: false,
  };
}

function opUnprotectSheet(t: SheetTarget, dryRun: boolean): OpDetail | null {
  const sp = getSheetProtection(t.ws);
  if (!sp) return null;
  if (!dryRun) t.ws.unprotect();
  return detail(t, 'シート保護を解除', 1);
}

function opFillRange(
  t: SheetTarget,
  body: Extract<StepBody, { op: 'fillRange' }>,
  dryRun: boolean,
): OpDetail | null {
  const rect = resolveRange(t.ws, body.range);
  if (!rect) return null;
  let count = 0;
  const ok = forEachCell(t.ws, rect, (cell) => {
    if (body.colorArgb === null) {
      if (!cell.fill) return;
      count++;
      if (!dryRun) setFill(cell, null);
      return;
    }
    if (getFillArgb(cell) === body.colorArgb) return;
    count++;
    if (!dryRun) setFill(cell, body.colorArgb);
  });
  if (!ok) return detail(t, '範囲が大きすぎるため処理を中止しました', 0);
  if (count === 0) return null;
  return detail(
    t,
    `${rectToA1(rect)} の ${count} セルの塗りを${body.colorArgb === null ? '解除' : '変更'}`,
    count,
  );
}

/**
 * 塗りつぶしの色でロックを切り替える。
 * 「この色のセルだけ入力させたい」という既存の色分け運用を、
 * そのままロック設定に変換できる。
 */
function opSetLockByFill(
  t: SheetTarget,
  body: Extract<StepBody, { op: 'setLockByFill' }>,
  dryRun: boolean,
): OpDetail | null {
  const rect = resolveRange(t.ws, body.range);
  if (!rect) return null;
  const wanted = new Set(body.colorKeys);
  if (wanted.size === 0) return null;

  let count = 0;
  const ok = forEachCell(t.ws, rect, (cell) => {
    const ref = getFillRef(cell);
    const isMatch =
      body.match === 'in'
        ? ref !== null && wanted.has(ref.key)
        : ref === null
          ? body.includeUnfilled
          : !wanted.has(ref.key);
    if (!isMatch) return;
    if (getLocked(cell) === body.locked) return;
    count++;
    if (!dryRun) setLocked(cell, body.locked);
  });
  if (!ok) return detail(t, '範囲が大きすぎるため処理を中止しました', 0);
  if (count === 0) return null;
  return detail(
    t,
    `色が${body.match === 'in' ? '一致' : '不一致'}の ${count} セルを${
      body.locked ? 'ロック' : 'ロック解除'
    }`,
    count,
  );
}

function opFillByLockState(
  t: SheetTarget,
  body: Extract<StepBody, { op: 'fillByLockState' }>,
  dryRun: boolean,
): OpDetail | null {
  const rect = usedRect(t.ws);
  if (!rect) return null;
  const wantLocked = body.target === 'locked';
  let count = 0;
  const ok = forEachCell(t.ws, rect, (cell) => {
    if (getLocked(cell) !== wantLocked) return;
    if (body.onlyUsedRange && cell.value === null) return;
    if (body.colorArgb === null) {
      if (!cell.fill) return;
      count++;
      if (!dryRun) setFill(cell, null);
      return;
    }
    if (getFillArgb(cell) === body.colorArgb) return;
    count++;
    if (!dryRun) setFill(cell, body.colorArgb);
  });
  if (!ok) return detail(t, '範囲が大きすぎるため処理を中止しました', 0);
  if (count === 0) return null;
  return detail(
    t,
    `${wantLocked ? 'ロック済み' : 'ロック解除済み'}の ${count} セルを塗り${
      body.colorArgb === null ? '解除' : '変更'
    }`,
    count,
  );
}

// --- 年度置換 -------------------------------------------------------------

interface TextRewriter {
  (s: string): string | null;
}

function makeRewriter(mapper: YearMapper, wholeNumberOnly: boolean, era: number | null): TextRewriter {
  return (s: string) => {
    let out = replaceYearsInString(s, mapper, wholeNumberOnly);
    if (era !== null) {
      const eraOut = replaceEraYears(out ?? s, era);
      if (eraOut !== null) out = eraOut;
    }
    return out;
  };
}

/**
 * セルの値を書き換える。ExcelJS の CellValue は形が多いため、
 * 型ごとに元の構造を保ったまま文字列部分だけ差し替える。
 */
function rewriteCellValue(
  cell: ExcelJS.Cell,
  rewrite: TextRewriter,
  mapper: YearMapper,
  includeFormulas: boolean,
  dryRun: boolean,
): boolean {
  const v = cell.value;
  if (v === null || v === undefined || v === '') return false;

  if (typeof v === 'string') {
    const next = rewrite(v);
    if (next === null) return false;
    if (!dryRun) cell.value = next;
    return true;
  }

  if (typeof v === 'number') {
    const next = mapNumericYear(v, mapper);
    if (next === null || next === v) return false;
    if (!dryRun) cell.value = next;
    return true;
  }

  if (typeof v === 'object') {
    const obj = asRecord(v);

    // 共有数式の従属セルはマスター側を書き換えれば追随するので触らない
    if ('sharedFormula' in obj && !('formula' in obj)) return false;

    if ('formula' in obj) {
      if (!includeFormulas) return false;
      const next = rewrite(String(obj.formula));
      if (next === null) return false;
      if (!dryRun) {
        // 数式を書き換えたらキャッシュ結果は捨て、Excel に再計算させる
        cell.value = { formula: next, date1904: false } as ExcelJS.CellFormulaValue;
      }
      return true;
    }

    if ('richText' in obj && Array.isArray(obj.richText)) {
      const runs = obj.richText as Array<{ text?: string }>;
      let changed = false;
      const nextRuns = runs.map((run) => {
        const next = run.text ? rewrite(run.text) : null;
        if (next === null) return run;
        changed = true;
        return { ...run, text: next };
      });
      if (!changed) return false;
      if (!dryRun) cell.value = { richText: nextRuns } as ExcelJS.CellRichTextValue;
      return true;
    }

    if ('text' in obj && 'hyperlink' in obj) {
      const next = rewrite(String(obj.text));
      if (next === null) return false;
      if (!dryRun) cell.value = { ...obj, text: next } as ExcelJS.CellHyperlinkValue;
      return true;
    }
  }
  return false;
}

function opReplaceYears(
  t: SheetTarget,
  body: Extract<StepBody, { op: 'shiftYears' | 'mapYears' }>,
  dryRun: boolean,
  renameLog: Map<string, string>,
): OpDetail | null {
  const mapper =
    body.op === 'shiftYears'
      ? shiftMapper(body.delta, body.minYear, body.maxYear)
      : pairMapper(body.pairs);
  const eraDelta =
    body.targets.japaneseEra && body.op === 'shiftYears' ? body.delta : null;
  const rewrite = makeRewriter(mapper, body.wholeNumberOnly, eraDelta);

  let cellCount = 0;
  const messages: string[] = [];

  if (body.targets.values || body.targets.formulas) {
    const rect = resolveRange(t.ws, body.range);
    if (rect) {
      const ok = forEachCell(t.ws, rect, (cell) => {
        const isFormula =
          cell.value !== null &&
          typeof cell.value === 'object' &&
          'formula' in (cell.value as object);
        if (isFormula && !body.targets.formulas) return;
        if (!isFormula && !body.targets.values) return;
        if (rewriteCellValue(cell, rewrite, mapper, body.targets.formulas, dryRun)) cellCount++;
      });
      if (!ok) return detail(t, '範囲が大きすぎるため処理を中止しました', 0);
    }
  }
  if (cellCount > 0) messages.push(`${cellCount} セルの年を変更`);

  if (body.targets.sheetNames) {
    const next = rewrite(t.ws.name);
    if (next !== null && next !== t.ws.name) {
      const conflict = sheetNameExists(t.book, next, t.ws.id);
      if (conflict) {
        messages.push(`シート名「${t.ws.name}」→「${next}」は同名シートがあるため見送り`);
      } else {
        messages.push(`シート名「${t.ws.name}」→「${next}」`);
        renameLog.set(`${t.book.id}::${t.ws.name}`, next);
        if (!dryRun) t.ws.name = next;
        cellCount++;
      }
    }
  }

  if (messages.length === 0) return null;
  return detail(t, messages.join(' / '), cellCount);
}

function sheetNameExists(book: LoadedWorkbook, name: string, exceptId: number): boolean {
  let found = false;
  book.wb.eachSheet((ws) => {
    if (ws.id !== exceptId && ws.name === name) found = true;
  });
  return found;
}

function opReplaceText(
  t: SheetTarget,
  body: Extract<StepBody, { op: 'replaceText' }>,
  dryRun: boolean,
): OpDetail | null {
  if (!body.find) return null;
  const flags = body.matchCase ? 'g' : 'gi';
  const escaped = body.find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp(escaped, flags);

  const rewrite: TextRewriter = (s) => {
    if (body.wholeCell) {
      const same = body.matchCase ? s === body.find : s.toLowerCase() === body.find.toLowerCase();
      return same ? body.replace : null;
    }
    re.lastIndex = 0;
    if (!re.test(s)) return null;
    re.lastIndex = 0;
    return s.replace(re, body.replace);
  };
  const noopMapper: YearMapper = () => null;

  let count = 0;
  const messages: string[] = [];

  if (body.targets.values || body.targets.formulas) {
    const rect = resolveRange(t.ws, body.range);
    if (rect) {
      const ok = forEachCell(t.ws, rect, (cell) => {
        const isFormula =
          cell.value !== null &&
          typeof cell.value === 'object' &&
          'formula' in (cell.value as object);
        if (isFormula && !body.targets.formulas) return;
        if (!isFormula && !body.targets.values) return;
        if (rewriteCellValue(cell, rewrite, noopMapper, body.targets.formulas, dryRun)) count++;
      });
      if (!ok) return detail(t, '範囲が大きすぎるため処理を中止しました', 0);
    }
  }
  if (count) messages.push(`${count} セルを置換`);

  if (body.targets.sheetNames) {
    const next = rewrite(t.ws.name);
    if (next !== null && next !== t.ws.name && !sheetNameExists(t.book, next, t.ws.id)) {
      messages.push(`シート名「${t.ws.name}」→「${next}」`);
      if (!dryRun) t.ws.name = next;
      count++;
    }
  }

  if (!messages.length) return null;
  return detail(t, messages.join(' / '), count);
}

function detail(t: SheetTarget, message: string, count: number): OpDetail {
  return { book: t.book.relPath, sheet: t.ws.name, message, count };
}

// ---------------------------------------------------------------------------
// エントリポイント
// ---------------------------------------------------------------------------

/** ファイル名の年度置換の結果 (保存時のファイル名に反映される) */
export interface FileRename {
  bookId: string;
  from: string;
  to: string;
}

export interface StepOutcome extends OpResult {
  fileRenames: FileRename[];
  /**
   * 適用範囲に一致したシートの数。
   * 0 のときは対象が 1 つも無かったということで、
   * 手順書のフォルダー指定が今のファイル構成に合っていない可能性がある。
   */
  targetSheets: number;
}

export async function applyStep(
  step: RecipeStep,
  ctx: OpContext,
  opts: ApplyOptions = {},
): Promise<StepOutcome> {
  const dryRun = opts.dryRun ?? false;
  const targets = resolveTargets(ctx, step.scope);
  const details: OpDetail[] = [];
  const renameLog = new Map<string, string>();
  const fileRenames: FileRename[] = [];
  const body = step.body;

  for (const t of targets) {
    let d: OpDetail | null = null;
    switch (body.op) {
      case 'setLock':
        d = opSetLock(t, body, dryRun);
        break;
      case 'lockAllExcept':
        d = opLockAllExcept(t, body, dryRun);
        break;
      case 'protectSheet':
        d = await opProtectSheet(t, body, dryRun);
        break;
      case 'unprotectSheet':
        d = opUnprotectSheet(t, dryRun);
        break;
      case 'fillRange':
        d = opFillRange(t, body, dryRun);
        break;
      case 'fillByLockState':
        d = opFillByLockState(t, body, dryRun);
        break;
      case 'setLockByFill':
        d = opSetLockByFill(t, body, dryRun);
        break;
      case 'shiftYears':
      case 'mapYears':
        d = opReplaceYears(t, body, dryRun, renameLog);
        break;
      case 'replaceText':
        d = opReplaceText(t, body, dryRun);
        break;
    }
    if (d) details.push(d);
  }

  // ファイル名の置換はブック単位で 1 回だけ行う
  if (
    (body.op === 'shiftYears' || body.op === 'mapYears' || body.op === 'replaceText') &&
    body.targets.fileNames
  ) {
    const seen = new Set<string>();
    for (const t of targets) {
      if (seen.has(t.book.id)) continue;
      seen.add(t.book.id);
      const rewrite =
        body.op === 'replaceText'
          ? makeTextRewriter(body)
          : makeRewriter(
              body.op === 'shiftYears'
                ? shiftMapper(body.delta, body.minYear, body.maxYear)
                : pairMapper(body.pairs),
              body.wholeNumberOnly,
              body.targets.japaneseEra && body.op === 'shiftYears' ? body.delta : null,
            );
      const next = rewrite(t.book.fileName);
      if (next !== null && next !== t.book.fileName) {
        fileRenames.push({ bookId: t.book.id, from: t.book.fileName, to: next });
        details.push({
          book: t.book.relPath,
          sheet: '(ファイル名)',
          message: `「${t.book.fileName}」→「${next}」`,
          count: 1,
        });
      }
    }
  }

  const changedBooks = new Set(details.filter((d) => d.count > 0).map((d) => d.book)).size;
  const changedSheets = details.filter((d) => d.count > 0 && d.sheet !== '(ファイル名)').length;
  const changedCells = details.reduce((sum, d) => sum + d.count, 0);

  if (!dryRun && changedCells > 0) {
    const touched = new Set(details.map((d) => d.book));
    for (const b of ctx.books) if (touched.has(b.relPath)) b.dirty = true;
  }

  return {
    summary: buildSummary(body, changedBooks, changedSheets, changedCells, targets.length),
    changedCells,
    changedSheets,
    changedBooks,
    details,
    fileRenames,
    targetSheets: targets.length,
  };
}

function makeTextRewriter(body: Extract<StepBody, { op: 'replaceText' }>): TextRewriter {
  const flags = body.matchCase ? 'g' : 'gi';
  const escaped = body.find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp(escaped, flags);
  return (s) => {
    re.lastIndex = 0;
    if (!re.test(s)) return null;
    re.lastIndex = 0;
    return s.replace(re, body.replace);
  };
}

const LOCK_OPS = new Set(['setLock', 'lockAllExcept', 'setLockByFill']);

function buildSummary(
  body: StepBody,
  books: number,
  sheets: number,
  cells: number,
  targetCount: number,
): string {
  if (targetCount === 0) return '対象のシートがありませんでした';
  if (cells === 0) {
    // Excel のセルは既定でロック済みなので、「ロックする」操作は
    // 何も変わらないことが多い。失敗と誤解されないよう理由を添える。
    if (LOCK_OPS.has(body.op)) {
      return `対象 ${targetCount} シートは、すでに指定どおりのロック状態でした (変更なし)`;
    }
    return `対象 ${targetCount} シートを確認しましたが、変更はありませんでした`;
  }
  const unit =
    body.op === 'protectSheet' || body.op === 'unprotectSheet' ? `${sheets} シート` : `${cells} 箇所`;
  return `${books} ブック / ${sheets} シート / ${unit} を変更しました`;
}


// ---------------------------------------------------------------------------
// 使われている塗りつぶし色の列挙
// ---------------------------------------------------------------------------

export interface UsedColor extends FillRef {
  /** その色で塗られているセルの数 */
  count: number;
  /** 見つかった場所の例 (先頭 1 件) */
  sample: string;
}

/**
 * 対象のシート群で実際に使われている塗りつぶし色を、多い順に列挙する。
 *
 * 「どの色が入力欄なのか」はファイルを見ないと分からないため、
 * 画面上で選ばせるための一覧を作る。
 */
export function collectUsedColors(
  ctx: OpContext,
  scope: OpScope,
  limit = 40,
): UsedColor[] {
  const found = new Map<string, UsedColor>();
  for (const t of resolveTargets(ctx, scope)) {
    const rect = usedRect(t.ws);
    if (!rect) continue;
    forEachCell(t.ws, rect, (cell, r, c) => {
      const ref = getFillRef(cell);
      if (!ref) return;
      const hit = found.get(ref.key);
      if (hit) {
        hit.count++;
      } else {
        found.set(ref.key, {
          ...ref,
          count: 1,
          sample: `${t.ws.name}!${colToLetter(c)}${r}`,
        });
      }
    });
  }
  return [...found.values()].sort((a, b) => b.count - a.count).slice(0, limit);
}
