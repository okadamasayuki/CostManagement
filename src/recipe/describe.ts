import type { OpScope } from '../excel/types';
import type { RangeSpec, RecipeStep, StepBody } from './types';
import { argbToCss } from '../excel/format';

/** 操作内容を日本語の文章にする。手順書とプレビューの両方で使う。 */

export function describeRange(range: RangeSpec): string {
  switch (range.kind) {
    case 'a1':
      return `セル範囲 ${range.a1}`;
    case 'used':
      return 'データが入っている範囲全体';
    case 'sheet':
      return 'シート全体 (データ範囲)';
  }
}

export function describeScope(scope: OpScope): string {
  const book =
    scope.books === 'current'
      ? '選択中のブック'
      : scope.books === 'all'
        ? '読み込んだすべてのブック'
        : `ファイル名が「${scope.bookGlob || '*'}」に一致するブック`;
  const sheet =
    scope.sheets === 'current'
      ? '選択中のシート'
      : scope.sheets === 'all'
        ? 'すべてのシート'
        : `シート名が「${scope.sheetGlob || '*'}」に一致するシート`;
  return `${book} の ${sheet}`;
}

export function describeBody(body: StepBody): string {
  switch (body.op) {
    case 'setLock':
      return `${describeRange(body.range)}を${body.locked ? 'ロックする' : 'ロック解除する (入力可能にする)'}`;
    case 'lockAllExcept':
      return (
        `${describeRange(body.range)}「以外」をすべてロックする` +
        (body.alsoUnlockTarget ? `。指定範囲は入力可能にする` : '')
      );
    case 'protectSheet':
      return (
        'シートの保護を有効にする' +
        (body.password ? ' (パスワードあり)' : ' (パスワードなし)') +
        `。許可する操作: ${describeProtectOptions(body)}`
      );
    case 'unprotectSheet':
      return 'シートの保護を解除する';
    case 'fillRange':
      return body.colorArgb === null
        ? `${describeRange(body.range)}の塗りつぶしを解除する`
        : `${describeRange(body.range)}を ${argbToCss(body.colorArgb)} で塗りつぶす`;
    case 'setLockByFill': {
      const colors = body.colorLabels?.length
        ? body.colorLabels.join('・')
        : `${body.colorKeys.length} 色`;
      const target =
        body.match === 'in'
          ? `${colors} で塗られているセル`
          : `${colors} 以外で塗られているセル${body.includeUnfilled ? 'と、塗りのないセル' : ''}`;
      return `${describeRange(body.range)}のうち、${target}を${
        body.locked ? 'ロックする' : 'ロック解除する (入力可能にする)'
      }`;
    }
    case 'fillByLockState':
      return body.colorArgb === null
        ? `${body.target === 'locked' ? 'ロックされている' : 'ロックされていない'}セルの塗りつぶしを解除する`
        : `${body.target === 'locked' ? 'ロックされている' : 'ロックされていない'}セルを ${argbToCss(
            body.colorArgb,
          )} で塗りつぶす`;
    case 'shiftYears': {
      const dir = body.delta >= 0 ? `${body.delta} 年進める` : `${Math.abs(body.delta)} 年戻す`;
      const sample = `${body.minYear}→${body.minYear + body.delta}, ${body.minYear + 1}→${
        body.minYear + 1 + body.delta
      } …`;
      return (
        `${body.minYear}〜${body.maxYear} 年の西暦をすべて${dir} (${sample})。` +
        `対象: ${describeYearTargets(body.targets)}。` +
        (body.wholeNumberOnly ? '前後に数字が続く場合 (例: 20240401) は対象外。' : '数字の途中でも置換する。')
      );
    }
    case 'mapYears':
      return (
        `西暦を対応表どおりに置換する (${body.pairs.map((p) => `${p.from}→${p.to}`).join(', ')})。` +
        `対象: ${describeYearTargets(body.targets)}`
      );
    case 'replaceText':
      return (
        `「${body.find}」を「${body.replace}」に置換する` +
        `${body.wholeCell ? ' (セル全体が一致する場合のみ)' : ''}` +
        `${body.matchCase ? ' (大文字小文字を区別)' : ''}。` +
        `対象: ${describeYearTargets({ ...body.targets, japaneseEra: false })}`
      );
  }
}

function describeYearTargets(t: {
  values: boolean;
  formulas: boolean;
  sheetNames: boolean;
  fileNames: boolean;
  japaneseEra: boolean;
}): string {
  const parts: string[] = [];
  if (t.values) parts.push('セルの値');
  if (t.formulas) parts.push('数式');
  if (t.sheetNames) parts.push('シート名');
  if (t.fileNames) parts.push('ファイル名');
  if (t.japaneseEra) parts.push('和暦');
  return parts.length ? parts.join('・') : '(なし)';
}

function describeProtectOptions(body: Extract<StepBody, { op: 'protectSheet' }>): string {
  const o = body.options;
  const allowed: string[] = [];
  if (o.selectLockedCells) allowed.push('ロックセルの選択');
  if (o.formatCells) allowed.push('書式変更');
  if (o.insertRows) allowed.push('行/列の挿入');
  if (o.deleteRows) allowed.push('行/列の削除');
  if (o.autoFilter) allowed.push('オートフィルタ');
  if (o.sort) allowed.push('並べ替え');
  return allowed.length ? allowed.join('・') : 'なし (入力のみ)';
}

/** 手順書の 1 行分の見出しを自動生成する */
export function autoLabel(body: StepBody): string {
  switch (body.op) {
    case 'setLock':
      return body.locked ? '範囲をロック' : '範囲のロックを解除';
    case 'lockAllExcept':
      return '指定範囲以外をロック';
    case 'protectSheet':
      return 'シート保護を有効化';
    case 'unprotectSheet':
      return 'シート保護を解除';
    case 'fillRange':
      return body.colorArgb === null ? '塗りつぶしを解除' : '範囲を塗りつぶし';
    case 'setLockByFill':
      return `色で${body.locked ? 'ロック' : 'ロック解除'} (${
        body.match === 'in' ? '指定色' : '指定色以外'
      })`;
    case 'fillByLockState':
      return `${body.target === 'locked' ? 'ロック済み' : 'ロック解除'}セルを色分け`;
    case 'shiftYears':
      return `年度を ${body.delta >= 0 ? '+' : ''}${body.delta} 年ずらす`;
    case 'mapYears':
      return '年度を対応表で置換';
    case 'replaceText':
      return `文字列置換 (${body.find} → ${body.replace})`;
  }
}

export function describeStep(step: RecipeStep): string {
  return `${describeScope(step.scope)} に対して、${describeBody(step.body)}`;
}
