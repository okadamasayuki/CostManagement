import type { SheetView } from '../excel/types';
import type { StepOutcome } from '../excel/ops';
import { describeBody, describeScope } from '../recipe/describe';
import {
  moveStep,
  patchStep,
  removeStep,
  setState,
  toggleStep,
  useStore,
  currentBook,
} from '../state/store';

/** 画面右側の情報パネル。リボンのタブに応じて内容が変わる。 */
export function RightPanel(props: { view: SheetView | null }) {
  const s = useStore();
  if (s.activeTab === 'recipe') return <RecipeSteps />;
  return <ContextInfo view={props.view} />;
}

function ContextInfo({ view }: { view: SheetView | null }) {
  const s = useStore();
  const book = currentBook();
  const last = s.history[0];
  const preview = s.preview;

  return (
    <div className="rightpanel">
      {preview && (
        <div className="rp-section">
          <div className="rp-title">
            <span>試算結果 — {preview.label}</span>
            <span className="spacer" />
            <button className="icon-btn" title="閉じる" onClick={() => setState({ preview: null })}>
              ✕
            </button>
          </div>
          <div className="rp-body">
            <div className="note-box warn" style={{ marginBottom: 8 }}>
              まだ変更されていません。内容を確認してから実行してください。
            </div>
            <OutcomeView outcome={preview.outcome} testId="preview-details" />
          </div>
        </div>
      )}

      {view && (
        <div className="rp-section">
          <div className="rp-title">シートの状況</div>
          <div className="rp-body">
            <dl className="kv">
              <dt>シート名</dt>
              <dd>{view.name}</dd>
              <dt>データ範囲</dt>
              <dd>
                {view.contentBottom.toLocaleString()} 行 × {view.contentRight.toLocaleString()} 列
              </dd>
              <dt>値のあるセル</dt>
              <dd>{view.usedCellCount.toLocaleString()}</dd>
              <dt>🔒 ロック済み</dt>
              <dd>{view.lockedCount.toLocaleString()}</dd>
              <dt>🔓 ロック解除</dt>
              <dd>{view.unlockedCount.toLocaleString()}</dd>
              <dt>シート保護</dt>
              <dd>{view.isProtected ? (view.hasPassword ? '有効 (PW)' : '有効') : '無効'}</dd>
            </dl>
            {view.truncated && (
              <div className="note-box warn" style={{ marginTop: 8 }}>
                <b>シートが大きいため、画面表示の一部を省いています。</b>
                ロック・色・置換などの処理は画面表示を使わないので、
                ファイル全体に対して正しく適用されます。
              </div>
            )}
            {!view.isProtected && view.unlockedCount > 0 && (
              <div className="note-box warn" style={{ marginTop: 8 }}>
                <b>シート保護が無効です。</b>
                このままではロック設定は効きません。「ロック」タブの
                <b>「シート保護を有効化」</b>を実行してください。
              </div>
            )}
            {view.isProtected && (
              <div className="note-box ok" style={{ marginTop: 8 }}>
                シート保護が有効です。ロックしたセルは Excel 上で編集できません。
              </div>
            )}
          </div>
        </div>
      )}

      {book?.loadError && (
        <div className="rp-section">
          <div className="rp-title">読み込みエラー</div>
          <div className="rp-body">
            <div className="note-box err">
              <b>{book.fileName}</b> を読み込めませんでした。
              <br />
              {book.loadError}
              <br />
              <br />
              パスワード付きのブックや、Excel 以外のツールで作られたファイルの可能性があります。
            </div>
          </div>
        </div>
      )}

      <div className="rp-section" style={{ flex: 1, minHeight: 0, display: 'flex', flexDirection: 'column' }}>
        <div className="rp-title">
          <span>操作の履歴</span>
          <span className="spacer" />
          <span style={{ fontWeight: 400 }}>{s.history.length} 件</span>
        </div>
        <div className="rp-body" style={{ overflow: 'auto' }}>
          {!last && (
            <div style={{ color: 'var(--text-dim)' }}>
              まだ操作していません。リボンから操作すると、ここに結果が表示され、
              同時に「手順書」へ記録されます。
            </div>
          )}
          {s.history.slice(0, 12).map((h) => (
            <div key={h.id} style={{ marginBottom: 10 }}>
              <div style={{ fontWeight: 600 }}>{h.label}</div>
              <div style={{ color: 'var(--text-dim)', fontSize: 10.5 }}>
                {h.at.toLocaleTimeString('ja-JP')} — {h.summary}
              </div>
              {h.id === last?.id && <OutcomeView outcome={h.outcome} />}
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

function OutcomeView({ outcome, testId }: { outcome: StepOutcome; testId?: string }) {
  if (!outcome.details.length) return null;
  return (
    <ul className="detail-list" data-testid={testId}>
      {outcome.details.slice(0, 60).map((d, i) => (
        <li key={i}>
          {d.message}
          <span className="where">
            {d.book} / {d.sheet}
          </span>
        </li>
      ))}
      {outcome.details.length > 60 && (
        <li style={{ color: 'var(--text-dim)' }}>ほか {outcome.details.length - 60} 件</li>
      )}
    </ul>
  );
}

function RecipeSteps() {
  const s = useStore();
  const steps = s.recipe.steps;

  return (
    <div className="rightpanel">
      <div className="rp-title">
        <span>手順書: {s.recipe.title}</span>
        <span className="spacer" />
        <span style={{ fontWeight: 400 }}>{steps.length} 手順</span>
      </div>
      <div className="rp-body" style={{ overflow: 'auto', flex: 1 }}>
        {!steps.length && (
          <div className="note-box">
            手順がまだありません。
            <br />
            <br />
            「ロック」「書式」「年度更新」タブで操作すると、その内容が自動的に手順として記録されます。
            一通り作業したら「手順書を書き出す」で保存してください。
          </div>
        )}
        <ol className="step-list">
          {steps.map((step, i) => (
            <li key={step.id} className={`step-item${step.enabled ? '' : ' disabled'}`}>
              <div className="step-head">
                <span className="step-num">{i + 1}</span>
                <input
                  className="step-label"
                  value={step.label}
                  onChange={(e) => patchStep(step.id, { label: e.target.value })}
                  style={{ border: 'none', background: 'transparent', outline: 'none', minWidth: 0 }}
                />
                <button
                  className="icon-btn"
                  title={step.enabled ? 'この手順を無効にする' : 'この手順を有効にする'}
                  onClick={() => toggleStep(step.id)}
                >
                  {step.enabled ? '☑' : '☐'}
                </button>
                <button
                  className="icon-btn"
                  title="上へ"
                  disabled={i === 0}
                  onClick={() => moveStep(step.id, -1)}
                >
                  ▲
                </button>
                <button
                  className="icon-btn"
                  title="下へ"
                  disabled={i === steps.length - 1}
                  onClick={() => moveStep(step.id, 1)}
                >
                  ▼
                </button>
                <button className="icon-btn" title="削除" onClick={() => removeStep(step.id)}>
                  ✕
                </button>
              </div>
              <div className="step-body">
                <div className="step-scope">対象: {describeScope(step.scope)}</div>
                {describeBody(step.body)}
              </div>
            </li>
          ))}
        </ol>
      </div>
    </div>
  );
}
