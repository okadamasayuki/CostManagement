import { useRef, useState } from 'react';
import { BigButton, Btn, Check, Field, Modal, NoteBox, RCol, RGroup } from '../ui';
import { downloadText } from '../../excel/saver';
import { recipeToHtml, recipeToJson, recipeToMarkdown } from '../../recipe/document';
import { describeScope } from '../../recipe/describe';
import { RecipeParseError, parseRecipe, runRecipe, type RunReport } from '../../recipe/runner';
import { emptyRecipe } from '../../recipe/types';
import {
  clearBusy,
  clearRecipeSteps,
  getState,
  opContext,
  setBusy,
  setRecipe,
  setState,
  toast,
  updateRecipe,
  useStore,
} from '../../state/store';

/**
 * 手順書 (レシピ) の作成・書き出し・読み込み・自動実行。
 *
 * 「毎年同じ作業をしている」という運用に向けて、
 * 画面で行った操作がそのまま手順として溜まっていくようにしている。
 */
export function RecipePanel() {
  const s = useStore();
  const importInput = useRef<HTMLInputElement>(null);
  const [report, setReport] = useState<RunReport | null>(null);
  const [showMeta, setShowMeta] = useState(false);

  const stepCount = s.recipe.steps.filter((st) => st.enabled).length;
  const ready = s.books.length > 0;
  const fileStem = s.recipe.title.replace(/[\\/:*?"<>|]/g, '_') || '手順書';

  function exportJson() {
    downloadText(recipeToJson(s.recipe), `${fileStem}.json`, 'application/json');
    toast('success', '手順書 (JSON) を保存しました', '来年この JSON を読み込めば同じ操作を再現できます。');
  }
  function exportHtml() {
    downloadText(recipeToHtml(s.recipe), `${fileStem}.html`, 'text/html');
    toast('success', '手順書 (HTML) を保存しました', 'そのまま印刷・回覧できます。');
  }
  function exportMarkdown() {
    downloadText(recipeToMarkdown(s.recipe), `${fileStem}.md`, 'text/markdown');
    toast('success', '手順書 (Markdown) を保存しました');
  }

  async function importJson(file: File) {
    try {
      const recipe = parseRecipe(await file.text());
      setRecipe(recipe);
      toast(
        'success',
        `手順書「${recipe.title}」を読み込みました`,
        `${recipe.steps.length} 手順。「すべて実行」で適用できます。`,
      );
    } catch (e) {
      toast(
        'error',
        '手順書を読み込めませんでした',
        e instanceof RecipeParseError ? e.message : e instanceof Error ? e.message : String(e),
      );
    }
  }

  async function run() {
    if (!stepCount) {
      toast('warn', '実行できる手順がありません');
      return;
    }
    setBusy('手順を実行しています…', 0, stepCount);
    try {
      const r = await runRecipe(getState().recipe, opContext(), {
        dryRun: false,
        onProgress: (i, total, label) => setBusy(`実行中: ${label} (${i + 1}/${total})`, i, total),
      });
      setReport(r);
      const renames = { ...getState().renames };
      for (const fr of r.fileRenames) renames[fr.bookId] = fr.to;
      setState({ renames, lastRunReport: r });
      const failed = r.steps.filter((x) => x.error).length;
      // 対象が 1 つも無かった手順は、フォルダー指定が今の構成に
      // 合っていない可能性が高いので目立たせる
      const empty = r.steps.filter((x) => !x.error && x.outcome.targetSheets === 0);
      const notes: string[] = [];
      if (failed) notes.push(`${failed} 手順でエラーが発生しました。`);
      if (empty.length) {
        notes.push(
          `${empty.length} 手順は対象が 1 つもありませんでした` +
            `（${empty.map((x) => x.step.label).slice(0, 3).join('、')}）。` +
            `手順書の「適用先」が今のファイル構成と合っているか確認してください。`,
        );
      }
      if (!notes.length) notes.push(`${r.steps.length} 手順を処理しました。`);
      toast(
        failed || empty.length ? 'warn' : 'success',
        `実行完了: ${r.totalChangedCells} 箇所`,
        notes.join('\n'),
      );
    } catch (e) {
      toast('error', '手順の実行に失敗しました', e instanceof Error ? e.message : String(e));
    } finally {
      clearBusy();
    }
  }

  return (
    <>
      <input
        ref={importInput}
        type="file"
        accept=".json,application/json"
        className="hidden-input"
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) void importJson(f);
          e.target.value = '';
        }}
      />

      <RGroup title="記録">
        <RCol>
          <Check
            label="操作を手順として記録する"
            checked={s.recording}
            onChange={(v) => setState({ recording: v })}
            title="ON の間、ロックや年度更新などの操作が手順書に追加されていきます"
          />
          <div style={{ fontSize: 10.5, color: 'var(--text-dim)', width: 200, lineHeight: 1.6 }}>
            現在 <b>{s.recipe.steps.length}</b> 手順
            {stepCount !== s.recipe.steps.length && `（有効 ${stepCount}）`}
            <br />
            右側のパネルで並べ替え・削除ができます。
          </div>
          <Btn onClick={() => setShowMeta(true)}>✎ 手順書の名前とメモ</Btn>
          <Btn
            kind="danger"
            disabled={!s.recipe.steps.length}
            onClick={() => {
              if (confirm('記録した手順をすべて削除しますか?')) {
                clearRecipeSteps();
                toast('info', '手順をすべて削除しました');
              }
            }}
          >
            🗑 手順をすべて削除
          </Btn>
        </RCol>
      </RGroup>

      <RGroup title="書き出し">
        <BigButton
          icon="📋"
          label={<>手順書を<br />書き出す (HTML)</>}
          disabled={!s.recipe.steps.length}
          title="人が読む・印刷する用の手順書"
          onClick={exportHtml}
        />
        <BigButton
          icon="⚙️"
          label={<>手順書を<br />書き出す (JSON)</>}
          primary
          disabled={!s.recipe.steps.length}
          title="来年このツールに読み込ませて自動実行するための形式"
          onClick={exportJson}
        />
        <RCol>
          <Btn onClick={exportMarkdown} disabled={!s.recipe.steps.length}>
            📝 Markdown で書き出す
          </Btn>
          <div style={{ fontSize: 10.5, color: 'var(--text-dim)', width: 170, lineHeight: 1.6 }}>
            HTML は回覧用、JSON は自動実行用です。両方を保存しておくと便利です。
          </div>
        </RCol>
      </RGroup>

      <RGroup title="読み込みと実行">
        <BigButton icon="📂" label={<>手順書を<br />読み込む</>} onClick={() => importInput.current?.click()} />
        <BigButton
          icon="▶️"
          label={<>すべて<br />実行</>}
          primary
          disabled={!ready || !stepCount}
          title="手順書に書かれた適用先で実行します"
          onClick={() => void run()}
        />
      </RGroup>

      <RGroup title="来年の使い方">
        <div style={{ width: 260 }}>
          <NoteBox>
            <b>1年目:</b> 画面で作業 → JSON と HTML を書き出す
            <br />
            <b>2年目以降:</b> フォルダーを開く → JSON を読み込む →「すべて実行」→ 保存
            <br />
            <span style={{ color: 'var(--warn)' }}>
              ※ 年度は「+1 年ずらす」で記録しておくと、翌年もそのまま使えます。
            </span>
          </NoteBox>
        </div>
      </RGroup>

      {showMeta && (
        <Modal
          title="手順書の情報"
          onClose={() => setShowMeta(false)}
          footer={<Btn kind="accent" onClick={() => setShowMeta(false)}>閉じる</Btn>}
        >
          <Field label="タイトル">
            <input
              type="text"
              style={{ width: '100%' }}
              value={s.recipe.title}
              onChange={(e) => updateRecipe({ title: e.target.value })}
            />
          </Field>
          <div style={{ marginTop: 12 }}>
            <div style={{ marginBottom: 4, color: 'var(--text-dim)' }}>説明 / 引き継ぎメモ</div>
            <textarea
              rows={6}
              style={{
                width: '100%',
                fontSize: 12,
                padding: 6,
                border: '1px solid var(--border)',
                borderRadius: 3,
              }}
              value={s.recipe.description}
              placeholder="例: 毎年 4 月の年度更新作業。対象は 原価管理\\2025年度 フォルダー配下すべて。"
              onChange={(e) => updateRecipe({ description: e.target.value })}
            />
          </div>
          <div style={{ marginTop: 12, display: 'flex', gap: 8 }}>
            <Btn
              onClick={() => {
                updateRecipe({ sourceHint: getState().books.map((b) => b.relPath) });
                toast('info', '現在読み込んでいるファイル一覧を手順書に記録しました');
              }}
            >
              現在のファイル一覧を記録
            </Btn>
            <Btn
              kind="danger"
              onClick={() => {
                setRecipe(emptyRecipe());
                setShowMeta(false);
                toast('info', '手順書を新規作成しました');
              }}
            >
              手順書を新規作成
            </Btn>
          </div>
        </Modal>
      )}

      {report && (
        <Modal
          title="実行結果"

          wide
          onClose={() => setReport(null)}
          footer={<Btn kind="accent" onClick={() => setReport(null)}>閉じる</Btn>}
        >
          <p>
            {report.steps.length} 手順 / 合計 <b>{report.totalChangedCells}</b> 箇所
          </p>
          <table className="plain">
            <thead>
              <tr>
                <th style={{ width: 30 }}>#</th>
                <th>手順</th>
                <th>結果</th>
              </tr>
            </thead>
            <tbody>
              {report.steps.map((r, i) => (
                <tr key={r.step.id + i}>
                  <td className="num">{i + 1}</td>
                  <td>{r.step.label}</td>
                  <td
                    style={
                      r.error
                        ? { color: 'var(--error)' }
                        : r.outcome.targetSheets === 0
                          ? { color: 'var(--warn)' }
                          : undefined
                    }
                  >
                    {r.error ?? r.outcome.summary}
                    {!r.error && r.outcome.targetSheets === 0 && (
                      <div style={{ fontSize: 10.5 }}>
                        適用先: {describeScope(r.step.scope)}
                      </div>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          {report.fileRenames.length > 0 && (
            <>
              <h4 style={{ margin: '14px 0 6px' }}>保存時に変わるファイル名</h4>
              <table className="plain">
                <tbody>
                  {report.fileRenames.map((f) => (
                    <tr key={f.bookId}>
                      <td>{f.from}</td>
                      <td>→ {f.to}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </>
          )}
          <NoteBox kind="warn">
            変更はまだメモリー上だけです。「ファイル」タブから<b>保存</b>してください。
          </NoteBox>
        </Modal>
      )}
    </>
  );
}
