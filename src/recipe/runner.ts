import type { OpContext, FileRename, StepOutcome } from '../excel/ops';
import { applyStep } from '../excel/ops';
import {
  DEFAULT_PROTECT_OPTIONS,
  DEFAULT_YEAR_TARGETS,
  RECIPE_FORMAT,
  type Recipe,
  type RecipeStep,
} from './types';

export interface StepReport {
  step: RecipeStep;
  outcome: StepOutcome;
  error?: string;
}

export interface RunReport {
  startedAt: Date;
  finishedAt: Date;
  steps: StepReport[];
  fileRenames: FileRename[];
  totalChangedCells: number;
  dryRun: boolean;
}

/**
 * 手順書を順番に実行する。
 * 途中の手順が失敗しても止めず、最後まで実行して結果をまとめて返す
 * (どこで何が起きたかを一覧で確認できるようにするため)。
 */
export async function runRecipe(
  recipe: Recipe,
  ctx: OpContext,
  opts: { dryRun?: boolean; onProgress?: (i: number, total: number, label: string) => void } = {},
): Promise<RunReport> {
  const startedAt = new Date();
  const steps: StepReport[] = [];
  const fileRenames: FileRename[] = [];
  const active = recipe.steps.filter((s) => s.enabled);

  for (let i = 0; i < active.length; i++) {
    const step = active[i];
    opts.onProgress?.(i, active.length, step.label);
    try {
      const outcome = await applyStep(step, ctx, { dryRun: opts.dryRun });
      steps.push({ step, outcome });
      fileRenames.push(...outcome.fileRenames);
    } catch (e) {
      steps.push({
        step,
        outcome: {
          summary: '実行時にエラーが発生しました',
          changedCells: 0,
          changedSheets: 0,
          changedBooks: 0,
          details: [],
          fileRenames: [],
        },
        error: e instanceof Error ? e.message : String(e),
      });
    }
    await new Promise((r) => setTimeout(r, 0));
  }
  opts.onProgress?.(active.length, active.length, '完了');

  return {
    startedAt,
    finishedAt: new Date(),
    steps,
    fileRenames,
    totalChangedCells: steps.reduce((s, r) => s + r.outcome.changedCells, 0),
    dryRun: opts.dryRun ?? false,
  };
}

export class RecipeParseError extends Error {}

/**
 * 読み込んだ JSON を検証して Recipe にする。
 * 手で編集された手順書や旧バージョンでも壊れないよう、
 * 足りないフィールドは既定値で補う。
 */
export function parseRecipe(json: string): Recipe {
  let data: unknown;
  try {
    data = JSON.parse(json);
  } catch {
    throw new RecipeParseError('JSON として読み取れませんでした。ファイルが壊れていないか確認してください。');
  }
  if (!data || typeof data !== 'object') {
    throw new RecipeParseError('手順書の形式ではありません。');
  }
  const r = data as Record<string, unknown>;
  if (r.format !== RECIPE_FORMAT) {
    throw new RecipeParseError(
      `この JSON は本ツールの手順書ではないようです (format: ${String(r.format ?? 'なし')})。`,
    );
  }
  if (!Array.isArray(r.steps)) {
    throw new RecipeParseError('手順 (steps) が含まれていません。');
  }

  const steps: RecipeStep[] = [];
  (r.steps as unknown[]).forEach((raw, i) => {
    const step = normalizeStep(raw, i);
    if (step) steps.push(step);
  });
  if (!steps.length) {
    throw new RecipeParseError('有効な手順が 1 件もありませんでした。');
  }

  return {
    format: RECIPE_FORMAT,
    version: typeof r.version === 'number' ? r.version : 1,
    title: typeof r.title === 'string' ? r.title : '読み込んだ手順',
    description: typeof r.description === 'string' ? r.description : '',
    createdAt: typeof r.createdAt === 'string' ? r.createdAt : new Date().toISOString(),
    sourceHint: Array.isArray(r.sourceHint) ? (r.sourceHint as string[]).map(String) : [],
    steps,
  };
}

const KNOWN_OPS = new Set([
  'setLock',
  'lockAllExcept',
  'protectSheet',
  'unprotectSheet',
  'fillRange',
  'fillByLockState',
  'shiftYears',
  'mapYears',
  'replaceText',
]);

function normalizeStep(raw: unknown, index: number): RecipeStep | null {
  if (!raw || typeof raw !== 'object') return null;
  const s = raw as Record<string, unknown>;
  const body = s.body as Record<string, unknown> | undefined;
  if (!body || typeof body.op !== 'string' || !KNOWN_OPS.has(body.op)) return null;

  // 型の欠落を既定値で補う
  if ('targets' in body) {
    body.targets = { ...DEFAULT_YEAR_TARGETS, ...(body.targets as object) };
  }
  if (body.op === 'protectSheet') {
    body.options = { ...DEFAULT_PROTECT_OPTIONS, ...((body.options as object) ?? {}) };
  }
  if ('range' in body) {
    const range = body.range as Record<string, unknown> | undefined;
    if (!range || typeof range.kind !== 'string') body.range = { kind: 'used' };
  }

  const scope = (s.scope as Record<string, unknown>) ?? {};
  return {
    id: typeof s.id === 'string' ? s.id : `step-${index + 1}`,
    label: typeof s.label === 'string' ? s.label : `手順 ${index + 1}`,
    note: typeof s.note === 'string' ? s.note : undefined,
    enabled: s.enabled !== false,
    scope: {
      books: (['current', 'all', 'glob'] as const).includes(scope.books as 'all')
        ? (scope.books as 'all')
        : 'all',
      bookGlob: typeof scope.bookGlob === 'string' ? scope.bookGlob : undefined,
      sheets: (['current', 'all', 'glob'] as const).includes(scope.sheets as 'all')
        ? (scope.sheets as 'all')
        : 'all',
      sheetGlob: typeof scope.sheetGlob === 'string' ? scope.sheetGlob : undefined,
    },
    body: body as unknown as RecipeStep['body'],
  };
}
