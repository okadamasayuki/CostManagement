import { useEffect, useState } from 'react';
import { Btn, Check, ColorSwatches, Modal, NoteBox } from '../ui';
import type { DetectReport } from '../../excel/ops';
import type { StepBody } from '../../recipe/types';
import { describeScope } from '../../recipe/describe';
import { getState, previewOperation, runOperation, setState, toast, useStore } from '../../state/store';

/**
 * 2 年分のファイルを見比べて、記入欄を自動で見つける画面。
 *
 * 200 ファイルの様式を人が見て「ここは主管部が毎年書く欄」と
 * 判断していくのは現実的でない。同じ様式が 2 年分あるなら、
 *   ・毎年 値が書き換わっているセル → 記入欄
 *   ・毎年 同じ値のセル             → 様式 (見出し・費目名など)
 * という当たりが機械的に付けられる。それを一気に色分け / ロックする。
 */
export function DetectInputDialog(props: { onClose(): void }) {
  const s = useStore();
  const [ignoreYearOnly, setIgnoreYearOnly] = useState(true);
  const [compareFormulaText, setCompareFormulaText] = useState(true);
  const [doFill, setDoFill] = useState(true);
  const [color, setColor] = useState<string | null>('FFFFF2CC');
  const [unlockChanged, setUnlockChanged] = useState(true);
  const [lockUnchanged, setLockUnchanged] = useState(true);
  const scopeNow = getState().scope;
  const [preview, setPreview] = useState<{ summary: string; detect?: DetectReport } | null>(null);
  const [busy, setBusy] = useState(false);

  /**
   * 開いたときは読み込んだ全ブック・全シートを対象にする。
   *
   * この操作は「年違いの 2 冊」を組にして見比べるものなので、
   * 既定の「選択中のブックのみ」のままだと組が作れず必ず 0 件になる。
   * シートも年違いで対応づけるため、全シートを見る。
   */
  useEffect(() => {
    const sc = getState().scope;
    if (sc.books === 'current' || sc.sheets === 'current') {
      setState({
        scope: {
          ...sc,
          books: sc.books === 'current' ? 'all' : sc.books,
          sheets: 'all',
        },
      });
    }
  }, []);

  /** 設定を変えたら試算結果は捨てる (古い数字を見せない) */
  const touch = () => setPreview(null);

  function buildBody(): StepBody {
    return {
      op: 'detectInputCells',
      ignoreYearOnly,
      compareFormulaText,
      fillChanged: doFill ? color : null,
      unlockChanged,
      lockUnchanged,
    };
  }

  const nothingToDo = !doFill && !unlockChanged && !lockUnchanged;

  async function doPreview() {
    setBusy(true);
    try {
      const outcome = await previewOperation(buildBody());
      setPreview({ summary: outcome.summary, detect: outcome.detect });
    } finally {
      setBusy(false);
    }
  }

  async function doApply() {
    if (nothingToDo) {
      toast('warn', '何をするかを 1 つ以上選んでください');
      return;
    }
    setBusy(true);
    try {
      const outcome = await runOperation(buildBody());
      toast(
        outcome.changedCells ? 'success' : 'info',
        outcome.summary,
        lockUnchanged || unlockChanged
          ? '仕上げに「ロック」タブでシート保護を有効にしてください。'
          : undefined,
      );
      props.onClose();
    } finally {
      setBusy(false);
    }
  }

  const d = preview?.detect;

  return (
    <Modal
      title="2 年分を見比べて記入欄を判定する"
      wide
      onClose={props.onClose}
      footer={
        <>
          <Btn onClick={props.onClose}>キャンセル</Btn>
          <Btn onClick={() => void doPreview()} disabled={busy}>
            {busy ? '判定中…' : '判定してみる (試算)'}
          </Btn>
          <Btn kind="accent" onClick={() => void doApply()} disabled={busy || nothingToDo}>
            実行する
          </Btn>
        </>
      }
    >
      <NoteBox>
        同じ様式のファイルが<b>年違いで 2 つ以上</b>読み込まれていると、
        中身を突き合わせて
        <b>「毎年 値が書き換わっているセル = 主管部の記入欄」</b>
        を割り出せます。逆に毎年同じ値のセルは様式とみなします。
        <br />
        判定は<b>新しい方のファイル</b>に反映します (古い方は変更しません)。
      </NoteBox>

      <div
        style={{
          display: 'flex',
          gap: 16,
          alignItems: 'flex-start',
          background: 'var(--panel-bg)',
          border: '1px solid var(--border)',
          borderRadius: 4,
          padding: '8px 10px',
          marginTop: 10,
        }}
      >
        <div style={{ fontSize: 12, whiteSpace: 'nowrap' }}>
          🎯 適用先:{' '}
          <b style={{ color: 'var(--excel-green)' }} data-testid="dialog-scope">
            {describeScope(scopeNow)}
          </b>
        </div>
        <div style={{ fontSize: 11, color: 'var(--text-dim)', lineHeight: 1.6 }}>
          対象: <b style={{ color: 'var(--text)' }}>{describeScope(s.scope)}</b>
          <br />
          2 年分のフォルダーが<b>両方入るように</b>指定してください。
          <br />
          ファイル名やフォルダー名の<b>年の数字だけが違うもの</b>を組にします。
          <br />
          （例: <code>2024/予算表.xlsx</code> と <code>2025/予算表.xlsx</code>）
          <br />
          シートも同じように年違いで対応づけるので、全シートが対象です。
        </div>
      </div>

      <h4 style={{ margin: '14px 0 6px' }}>1. 比べ方</h4>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
        <Check
          label="年度の数字だけの違いは「変わっていない」とみなす"
          checked={ignoreYearOnly}
          onChange={(v) => {
            setIgnoreYearOnly(v);
            touch();
          }}
          title="「2024年度予算」→「2025年度予算」のような見出しを記入欄と誤判定しないためのものです"
        />
        <Check
          label="数式は計算結果ではなく、式そのもので比べる"
          checked={compareFormulaText}
          onChange={(v) => {
            setCompareFormulaText(v);
            touch();
          }}
          title="合計欄は結果が毎年変わりますが、記入欄ではありません"
        />
      </div>

      <h4 style={{ margin: '16px 0 6px' }}>2. 判定したらどうするか</h4>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
        <Check
          label="記入欄と判定したセルを塗る"
          checked={doFill}
          onChange={(v) => {
            setDoFill(v);
            touch();
          }}
        />
        {doFill && (
          <div style={{ paddingLeft: 20, maxWidth: 200 }}>
            <ColorSwatches
              value={color}
              onChange={(c) => {
                setColor(c);
                touch();
              }}
            />
          </div>
        )}
        <Check
          label="🔓 記入欄と判定したセルのロックを外す（入力できるようにする）"
          checked={unlockChanged}
          onChange={(v) => {
            setUnlockChanged(v);
            touch();
          }}
        />
        <Check
          label="🔒 様式と判定したセルをロックする"
          checked={lockUnchanged}
          onChange={(v) => {
            setLockUnchanged(v);
            touch();
          }}
        />
      </div>
      {nothingToDo && (
        <NoteBox kind="warn">何をするかを 1 つ以上選んでください。</NoteBox>
      )}

      <div style={{ marginTop: 12, fontSize: 11, color: 'var(--text-dim)' }}>
        下の<b>「判定してみる (試算)」</b>を押すと、変更しないまま件数と判定の例を確認できます。
      </div>

      {preview && (
        <div data-testid="detect-preview">
          <NoteBox kind={d && d.pairCount > 0 ? 'info' : 'warn'}>
            <b>試算:</b> {preview.summary}
            {d && d.yearOnlyCount > 0 && (
              <>
                <br />
                うち {d.yearOnlyCount.toLocaleString()} セルは年の数字だけの違いだったので、
                様式として扱いました。
              </>
            )}
          </NoteBox>

          {d && d.pairCount === 0 && (
            <NoteBox kind="warn">
              年違いで対になるファイルが 1 組も見つかりませんでした。
              <b>2 年分のファイルを一緒に読み込んで</b>から、
              上の「適用先」で両方が入る範囲を選んでください。
            </NoteBox>
          )}

          {d && d.samples.length > 0 && (
            <>
              <h4 style={{ margin: '14px 0 6px' }}>記入欄と判定したセルの例</h4>
              <ul className="detail-list" data-testid="detect-samples">
                {d.samples.map((x, i) => (
                  <li key={i}>
                    <b>{x.addr}</b>: {x.before} → {x.after}
                    <span className="where">
                      {x.book} / {x.sheet}
                    </span>
                  </li>
                ))}
              </ul>
              <div style={{ fontSize: 11, color: 'var(--text-dim)' }}>
                （先頭 {d.samples.length} 件だけ表示しています）
              </div>
            </>
          )}

          {d && d.unpaired.length > 0 && (
            <NoteBox kind="warn" testId="detect-unpaired">
              <b>対になる年が見つからなかったファイル ({d.unpaired.length} 件):</b>
              <br />
              {d.unpaired.slice(0, 12).join(' / ')}
              {d.unpaired.length > 12 && ` ほか ${d.unpaired.length - 12} 件`}
              <br />
              これらは判定できないので、そのままにしています。
            </NoteBox>
          )}
        </div>
      )}

      <NoteBox>
        あくまで<b>2 年分の値の違いからの推定</b>です。実行後は画面の色分けで
        「記入欄になっているか」を必ず確認してください。
        判定が粗いときは「書式・色」タブの手動の塗り分けで直せます。
        <br />
        ロックは<b>シート保護を有効にして初めて効きます</b>。
        仕上げに「ロック」タブの「シート保護を有効化」を実行してください。
      </NoteBox>
    </Modal>
  );
}
