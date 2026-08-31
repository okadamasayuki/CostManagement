import type { Recipe } from './types';
import { describeBody, describeScope } from './describe';

/**
 * 手順書の出力。
 * - JSON: 本ツールに読み込ませて自動実行するための形式
 * - HTML: 人が読む/印刷する/回覧する手順書 (単体で開ける自己完結ファイル)
 * - Markdown: 社内 Wiki へ貼り付ける用
 * いずれもローカルでファイルを生成するだけで、送信は行わない。
 */

export function recipeToJson(recipe: Recipe): string {
  return JSON.stringify(recipe, null, 2);
}

export function recipeToMarkdown(recipe: Recipe): string {
  const lines: string[] = [];
  lines.push(`# ${recipe.title}`);
  lines.push('');
  if (recipe.description) {
    lines.push(recipe.description);
    lines.push('');
  }
  lines.push(`作成日: ${formatDate(recipe.createdAt)}`);
  lines.push('');
  if (recipe.sourceHint.length) {
    lines.push('## 作成時の対象ファイル');
    lines.push('');
    for (const s of recipe.sourceHint) lines.push(`- ${s}`);
    lines.push('');
  }
  lines.push('## 作業手順');
  lines.push('');
  const active = recipe.steps.filter((s) => s.enabled);
  active.forEach((step, i) => {
    lines.push(`### 手順 ${i + 1}: ${step.label}`);
    lines.push('');
    lines.push(`- **対象**: ${describeScope(step.scope)}`);
    lines.push(`- **操作**: ${describeBody(step.body)}`);
    if (step.note) lines.push(`- **メモ**: ${step.note}`);
    lines.push('');
  });
  lines.push('## 自動実行する場合');
  lines.push('');
  lines.push('1. 本ツールを開く');
  lines.push('2. 「ファイル」タブ → 「フォルダーを開く」で対象フォルダーを選ぶ');
  lines.push('3. 「手順書」タブ → 「手順書を読み込む」で同梱の JSON ファイルを選ぶ');
  lines.push('4. 「すべて実行」を押す');
  lines.push('5. 「ファイル」タブ → 「すべて保存」で書き出す');
  lines.push('');
  return lines.join('\n');
}

export function recipeToHtml(recipe: Recipe): string {
  const active = recipe.steps.filter((s) => s.enabled);
  const rows = active
    .map(
      (step, i) => `
      <tr>
        <td class="num">${i + 1}</td>
        <td>
          <div class="label">${esc(step.label)}</div>
          <div class="scope">対象: ${esc(describeScope(step.scope))}</div>
          <div class="body">${esc(describeBody(step.body))}</div>
          ${step.note ? `<div class="note">メモ: ${esc(step.note)}</div>` : ''}
        </td>
      </tr>`,
    )
    .join('');

  const sources = recipe.sourceHint.length
    ? `<h2>作成時の対象ファイル</h2><ul>${recipe.sourceHint
        .map((s) => `<li>${esc(s)}</li>`)
        .join('')}</ul>`
    : '';

  return `<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; connect-src 'none';">
<title>${esc(recipe.title)}</title>
<style>
  body { font-family: "Yu Gothic UI", "Meiryo", sans-serif; margin: 32px auto; max-width: 900px;
         color: #1f2937; line-height: 1.7; }
  h1 { font-size: 22px; border-bottom: 3px solid #217346; padding-bottom: 8px; }
  h2 { font-size: 16px; margin-top: 28px; color: #217346; }
  .meta { color: #6b7280; font-size: 13px; }
  .desc { background: #f3f4f6; padding: 12px 16px; border-radius: 6px; white-space: pre-wrap; }
  table { border-collapse: collapse; width: 100%; margin-top: 12px; }
  th, td { border: 1px solid #d1d5db; padding: 10px 12px; vertical-align: top; text-align: left; }
  th { background: #217346; color: #fff; font-size: 13px; }
  td.num { width: 44px; text-align: center; font-weight: 700; background: #f9fafb; }
  .label { font-weight: 700; font-size: 15px; }
  .scope { color: #6b7280; font-size: 12px; margin-top: 4px; }
  .body { margin-top: 6px; font-size: 14px; }
  .note { margin-top: 6px; font-size: 13px; color: #92400e; background: #fffbeb;
          padding: 6px 10px; border-radius: 4px; }
  ol { padding-left: 20px; }
  @media print { body { margin: 0; } h1 { page-break-after: avoid; } tr { page-break-inside: avoid; } }
</style>
</head>
<body>
  <h1>${esc(recipe.title)}</h1>
  <p class="meta">作成日: ${esc(formatDate(recipe.createdAt))} ／ 手順数: ${active.length}</p>
  ${recipe.description ? `<div class="desc">${esc(recipe.description)}</div>` : ''}
  ${sources}
  <h2>作業手順</h2>
  <table>
    <thead><tr><th>#</th><th>内容</th></tr></thead>
    <tbody>${rows}</tbody>
  </table>
  <h2>本ツールで自動実行する場合</h2>
  <ol>
    <li>ツール (単一 HTML ファイル) をブラウザーで開く</li>
    <li>「ファイル」タブ →「フォルダーを開く」で対象フォルダーを選ぶ</li>
    <li>「手順書」タブ →「手順書を読み込む」で対になる JSON ファイルを選ぶ</li>
    <li>「すべて実行」を押す</li>
    <li>「ファイル」タブ →「すべて保存」で書き出す</li>
  </ol>
  <h2>注意事項</h2>
  <ul>
    <li>実行前に必ず対象フォルダーのバックアップを取ってください。</li>
    <li>年度の置換は「同時置換」で行われるため、2024→2025 と 2025→2026 を同時に指定しても値が二重に進むことはありません。</li>
    <li>本ツールは外部通信を一切行いません。ファイルはブラウザー内でのみ処理されます。</li>
  </ul>
</body>
</html>`;
}

function esc(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function formatDate(iso: string): string {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return iso;
  return `${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日`;
}
