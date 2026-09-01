import ExcelJS from 'exceljs';
import { useSyncExternalStore } from 'react';
import type { LoadedWorkbook, OpScope, SkippedFile } from '../excel/types';
import type { RangeRect } from '../excel/cellRef';
import type { OpContext, StepOutcome } from '../excel/ops';
import { applyStep } from '../excel/ops';
import type { Recipe, RecipeStep, StepBody } from '../recipe/types';
import { emptyRecipe } from '../recipe/types';
import { autoLabel } from '../recipe/describe';
import type { RunReport } from '../recipe/runner';

/**
 * アプリ全体の状態。
 *
 * ExcelJS の Workbook は巨大な可変オブジェクトなので、React の state に
 * 入れて毎回コピーするのは現実的でない。ここでは状態をモジュール内に
 * 保持し、変更時に version を進めて再描画を促す方式にしている。
 */

export type RibbonTab = 'file' | 'lock' | 'format' | 'year' | 'recipe';

export interface Toast {
  id: number;
  kind: 'info' | 'success' | 'warn' | 'error';
  message: string;
  detail?: string;
}

export interface BusyState {
  active: boolean;
  label: string;
  done: number;
  total: number;
}

export interface HistoryEntry {
  id: number;
  at: Date;
  label: string;
  summary: string;
  outcome: StepOutcome;
}

export interface AppState {
  books: LoadedWorkbook[];
  skipped: SkippedFile[];
  currentBookId: string | null;
  /** 一覧で選択しているブック (Shift / Ctrl クリックで複数選べる) */
  selectedBookIds: string[];
  currentSheetName: string | null;
  selection: RangeRect | null;
  anchor: { row: number; col: number } | null;
  activeTab: RibbonTab;
  scope: OpScope;
  recipe: Recipe;
  history: HistoryEntry[];
  toasts: Toast[];
  busy: BusyState;
  /** 保存時に差し替えるファイル名 (年度置換の結果) */
  renames: Record<string, string>;
  /** ロック状態の可視化を有効にするか */
  showLockOverlay: boolean;
  lastRunReport: RunReport | null;
  /** 直前の操作を手順として記録するか */
  recording: boolean;
  /**
   * ファイルを読み込んだ直後にロック状態をそろえるかどうか。
   * Excel のセルは既定で全てロック済みのため、
   * 「まっさらな状態から始めたい」場合に unlock を選ぶ。
   */
  initialLockMode: 'keep' | 'unlock' | 'lock';
  /** 試算 (dry run) の結果。実行前の確認に使う。 */
  preview: { label: string; outcome: StepOutcome } | null;
  /**
   * ブックの中身が書き換わるたびに増える。
   * 画面用スナップショットの作り直しをこれだけに絞ることで、
   * 入力欄のタイプなど無関係な状態変化で再構築が走らないようにする。
   */
  docVersion: number;
}

let state: AppState = {
  books: [],
  skipped: [],
  currentBookId: null,
  selectedBookIds: [],
  currentSheetName: null,
  selection: null,
  anchor: null,
  activeTab: 'file',
  scope: { books: 'current', sheets: 'current' },
  recipe: emptyRecipe(),
  history: [],
  toasts: [],
  busy: { active: false, label: '', done: 0, total: 0 },
  renames: {},
  showLockOverlay: true,
  lastRunReport: null,
  recording: true,
  initialLockMode: 'keep',
  preview: null,
  docVersion: 0,
};

let version = 0;
const listeners = new Set<() => void>();

function emit(): void {
  version++;
  listeners.forEach((l) => l());
}

export function subscribe(fn: () => void): () => void {
  listeners.add(fn);
  return () => listeners.delete(fn);
}

export function getState(): AppState {
  return state;
}

export function setState(patch: Partial<AppState>): void {
  state = { ...state, ...patch };
  emit();
}

/** ExcelJS のモデルを直接書き換えたあとに呼ぶ */
export function touch(): void {
  state = { ...state, docVersion: state.docVersion + 1 };
  emit();
}

export function useStore(): AppState {
  useSyncExternalStore(
    subscribe,
    () => version,
    () => version,
  );
  return state;
}

// ---------------------------------------------------------------------------
// 派生値
// ---------------------------------------------------------------------------

export function currentBook(): LoadedWorkbook | null {
  return state.books.find((b) => b.id === state.currentBookId) ?? null;
}

export function currentSheet(): ExcelJS.Worksheet | null {
  const book = currentBook();
  if (!book || book.loadError) return null;
  if (state.currentSheetName) {
    const ws = book.wb.getWorksheet(state.currentSheetName);
    if (ws) return ws;
  }
  let first: ExcelJS.Worksheet | null = null;
  book.wb.eachSheet((ws) => {
    if (!first) first = ws;
  });
  return first;
}

export function opContext(): OpContext {
  return {
    books: state.books,
    currentBookId: state.currentBookId,
    currentSheetName: state.currentSheetName,
    selection: state.selection,
    selectedBookIds: state.selectedBookIds,
  };
}

// ---------------------------------------------------------------------------
// 通知
// ---------------------------------------------------------------------------

let toastSeq = 0;
export function toast(kind: Toast['kind'], message: string, detail?: string): void {
  const t: Toast = { id: ++toastSeq, kind, message, detail };
  setState({ toasts: [...state.toasts, t] });
  const ttl = kind === 'error' ? 12000 : 6000;
  setTimeout(() => dismissToast(t.id), ttl);
}

export function dismissToast(id: number): void {
  setState({ toasts: state.toasts.filter((t) => t.id !== id) });
}

export function setBusy(label: string, done = 0, total = 0): void {
  setState({ busy: { active: true, label, done, total } });
}
export function clearBusy(): void {
  setState({ busy: { active: false, label: '', done: 0, total: 0 } });
}

// ---------------------------------------------------------------------------
// ブックの管理
// ---------------------------------------------------------------------------

export function addBooks(books: LoadedWorkbook[], skipped: SkippedFile[]): void {
  const merged = [...state.books, ...books];
  const first = books.find((b) => !b.loadError);
  let sheetName = state.currentSheetName;
  if (first) {
    let name: string | null = null;
    first.wb.eachSheet((ws) => {
      if (!name) name = ws.name;
    });
    sheetName = name;
  }
  setState({
    books: merged,
    docVersion: state.docVersion + 1,
    skipped: [...state.skipped, ...skipped],
    currentBookId: state.currentBookId ?? first?.id ?? null,
    currentSheetName: state.currentBookId ? state.currentSheetName : sheetName,
    selection: null,
    anchor: null,
  });
}

/** 一覧での複数選択を差し替える */
export function setSelectedBooks(ids: string[]): void {
  setState({ selectedBookIds: ids });
}

export function selectBook(id: string): void {
  const book = state.books.find((b) => b.id === id);
  if (!book) return;
  let name: string | null = null;
  if (!book.loadError) {
    book.wb.eachSheet((ws) => {
      if (!name) name = ws.name;
    });
  }
  setState({
    currentBookId: id,
    selectedBookIds: [id],
    currentSheetName: name,
    selection: null,
    anchor: null,
  });
}

export function selectSheet(name: string): void {
  setState({ currentSheetName: name, selection: null, anchor: null });
}

export function closeAll(): void {
  setState({
    books: [],
    skipped: [],
    currentBookId: null,
    selectedBookIds: [],
    currentSheetName: null,
    selection: null,
    anchor: null,
    renames: {},
    lastRunReport: null,
  });
}

export function closeBook(id: string): void {
  const rest = state.books.filter((b) => b.id !== id);
  const nextId = state.currentBookId === id ? (rest[0]?.id ?? null) : state.currentBookId;
  let name: string | null = null;
  const nb = rest.find((b) => b.id === nextId);
  if (nb && !nb.loadError) {
    nb.wb.eachSheet((ws) => {
      if (!name) name = ws.name;
    });
  }
  const renames = { ...state.renames };
  delete renames[id];
  setState({
    books: rest,
    selectedBookIds: state.selectedBookIds.filter((x) => x !== id),
    currentBookId: nextId,
    currentSheetName: nextId === state.currentBookId ? state.currentSheetName : name,
    renames,
  });
}

/** 変更を破棄して元のファイルから読み直す */
export async function revertBook(id: string): Promise<void> {
  const book = state.books.find((b) => b.id === id);
  if (!book) return;
  let file = book.sourceFile;
  if (!file && book.handle) {
    file = await (book.handle as unknown as { getFile(): Promise<File> }).getFile();
  }
  if (!file) {
    toast('error', '元のファイルを読み直せませんでした', 'もう一度ファイルを開き直してください。');
    return;
  }
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await file.arrayBuffer());
  book.wb = wb;
  book.dirty = false;
  book.loadError = undefined;
  const renames = { ...state.renames };
  delete renames[id];
  setState({ renames, docVersion: state.docVersion + 1 });
  toast('success', `「${book.fileName}」を読み込み直しました`);
}

export function dirtyBooks(): LoadedWorkbook[] {
  return state.books.filter((b) => b.dirty && !b.loadError);
}

// ---------------------------------------------------------------------------
// 操作の実行と記録
// ---------------------------------------------------------------------------

let historySeq = 0;
let stepSeq = 0;

export function nextStepId(): string {
  return `step-${++stepSeq}-${Date.now().toString(36)}`;
}

/**
 * 操作を実行し、履歴と手順書に記録する。
 * 「毎年同じ作業をする」という運用に合わせ、実際に行った操作が
 * そのまま手順書になるようにしている。
 */
export async function runOperation(
  body: StepBody,
  opts: { scope?: OpScope; label?: string; record?: boolean } = {},
): Promise<StepOutcome> {
  const scope = opts.scope ?? state.scope;
  const step: RecipeStep = {
    id: nextStepId(),
    label: opts.label ?? autoLabel(body),
    enabled: true,
    scope,
    body,
  };
  const outcome = await applyStep(step, opContext(), { dryRun: false });

  /**
   * 「一覧で選択したブック」は画面上の一時的な状態なので、
   * 手順書にはファイル名の一覧として残す。
   * こうしておけば、後日 読み込み直しても同じ対象に当てられる。
   */
  const recorded: RecipeStep =
    scope.books === 'selected'
      ? {
          ...step,
          scope: {
            ...scope,
            books: 'glob',
            bookGlob: state.books
              .filter((b) => state.selectedBookIds.includes(b.id))
              .map((b) => b.fileName)
              .join(', '),
          },
        }
      : step;

  const renames = { ...state.renames };
  for (const r of outcome.fileRenames) renames[r.bookId] = r.to;

  const entry: HistoryEntry = {
    id: ++historySeq,
    at: new Date(),
    label: step.label,
    summary: outcome.summary,
    outcome,
  };

  const record = opts.record ?? state.recording;
  setState({
    history: [entry, ...state.history].slice(0, 200),
    docVersion: state.docVersion + 1,
    renames,
    recipe: record ? { ...state.recipe, steps: [...state.recipe.steps, recorded] } : state.recipe,
  });
  return outcome;
}

/** 変更を加えずに件数だけ数える */
export async function previewOperation(body: StepBody, scope?: OpScope): Promise<StepOutcome> {
  const step: RecipeStep = {
    id: 'preview',
    label: autoLabel(body),
    enabled: true,
    scope: scope ?? state.scope,
    body,
  };
  return applyStep(step, opContext(), { dryRun: true });
}

export function setRecipe(recipe: Recipe): void {
  setState({ recipe });
}

export function updateRecipe(patch: Partial<Recipe>): void {
  setState({ recipe: { ...state.recipe, ...patch } });
}

export function removeStep(id: string): void {
  setState({
    recipe: { ...state.recipe, steps: state.recipe.steps.filter((s) => s.id !== id) },
  });
}

export function toggleStep(id: string): void {
  setState({
    recipe: {
      ...state.recipe,
      steps: state.recipe.steps.map((s) => (s.id === id ? { ...s, enabled: !s.enabled } : s)),
    },
  });
}

export function moveStep(id: string, dir: -1 | 1): void {
  const steps = [...state.recipe.steps];
  const i = steps.findIndex((s) => s.id === id);
  const j = i + dir;
  if (i < 0 || j < 0 || j >= steps.length) return;
  [steps[i], steps[j]] = [steps[j], steps[i]];
  setState({ recipe: { ...state.recipe, steps } });
}

export function patchStep(id: string, patch: Partial<RecipeStep>): void {
  setState({
    recipe: {
      ...state.recipe,
      steps: state.recipe.steps.map((s) => (s.id === id ? { ...s, ...patch } : s)),
    },
  });
}

export function clearRecipeSteps(): void {
  setState({ recipe: { ...state.recipe, steps: [] } });
}
