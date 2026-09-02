import { useEffect, useState } from 'react';
import { Btn, Check, Modal, NoteBox } from '../ui';
import { collectUsedColors, type UsedColor } from '../../excel/ops';
import { argbToCss, readableTextColor } from '../../excel/format';
import type { StepBody } from '../../recipe/types';
import { describeScope } from '../../recipe/describe';
import { getState, opContext, previewOperation, runOperation, toast, useStore } from '../../state/store';

/**
 * 塗りつぶしの色からロックを設定する画面。
 *
 * 「黄色が入力欄」のような色分け運用が既にあるファイルでは、
 * どの色が使われているかを人が把握していないことが多い。
 * そこで実際に使われている色を数えて一覧にし、そこから選ばせる。
 */
export function ColorLockDialog(props: { onClose(): void }) {
  const store = useStore();
  const [colors, setColors] = useState<UsedColor[] | null>(null);
  const [selected, setSelected] = useState<Set<string>>(new Set());
  /**
   * よく使う 3 通りを、そのままの言葉で選べるようにする。
   * Excel のセルは既定で全てロック済みなので、
   * 「この色だけ入力できるようにする」= 色以外をロック + 色は解除、になる。
   * 細かく指定したい場合だけ、従来の指定に切り替えられる。
   */
  const [mode, setMode] = useState<'only' | 'unlock' | 'lock'>('only');
  const [detailed, setDetailed] = useState(false);
  const [match, setMatch] = useState<'in' | 'out'>('out');
  const [locked, setLocked] = useState(true);
  const [includeUnfilled, setIncludeUnfilled] = useState(true);
  const [alsoSetMatched, setAlsoSetMatched] = useState(true);
  const scopeNow = getState().scope;
  const [preview, setPreview] = useState<string | null>(null);

  const scope = store.scope;

  useEffect(() => {
    // 対象範囲の色を数える。件数が多いと少し時間がかかるので描画後に行う。
    // 適用先を変えたら数え直す。
    setColors(null);
    setSelected(new Set());
    setPreview(null);
    const id = setTimeout(() => setColors(collectUsedColors(opContext(), getState().scope)), 0);
    return () => clearTimeout(id);
  }, [scope.books, scope.sheets, scope.bookGlob, scope.sheetGlob]);

  function buildBody(): StepBody | null {
    if (!selected.size) {
      toast('warn', '色を 1 つ以上選んでください');
      return null;
    }
    const labels = (colors ?? [])
      .filter((c) => selected.has(c.key))
      .map((c) => argbToCss(c.argb) ?? c.key);
    // 分かりやすい 3 通りを、実際の指定へ翻訳する
    const spec = detailed
      ? { match, locked, includeUnfilled, alsoSetMatched }
      : mode === 'only'
        ? { match: 'out' as const, locked: true, includeUnfilled: true, alsoSetMatched: true }
        : mode === 'unlock'
          ? { match: 'in' as const, locked: false, includeUnfilled: true, alsoSetMatched: false }
          : { match: 'in' as const, locked: true, includeUnfilled: true, alsoSetMatched: false };
    return {
      op: 'setLockByFill',
      colorKeys: [...selected],
      colorLabels: labels,
      ...spec,
      range: { kind: 'used' },
    };
  }

  async function doPreview() {
    const body = buildBody();
    if (!body) return;
    const outcome = await previewOperation(body);
    setPreview(outcome.summary);
  }

  async function doApply() {
    const body = buildBody();
    if (!body) return;
    const outcome = await runOperation(body);
    toast(
      outcome.changedCells ? 'success' : 'info',
      outcome.summary,
      '仕上げに「ロック」タブでシート保護を有効にしてください。',
    );
    props.onClose();
  }

  const hasApprox = (colors ?? []).some((c) => c.isApprox && selected.has(c.key));

  return (
    <Modal
      title="色からロックを設定する"
      wide
      onClose={props.onClose}
      footer={
        <>
          <Btn onClick={props.onClose}>キャンセル</Btn>
          <Btn onClick={() => void doPreview()} disabled={!selected.size}>
            試算
          </Btn>
          <Btn kind="accent" onClick={() => void doApply()} disabled={!selected.size}>
            実行する
          </Btn>
        </>
      }
    >
      <div
        style={{
          display: 'flex',
          gap: 16,
          alignItems: 'flex-start',
          background: 'var(--panel-bg)',
          border: '1px solid var(--border)',
          borderRadius: 4,
          padding: '8px 10px',
        }}
      >
        <div style={{ fontSize: 12, whiteSpace: 'nowrap' }}>
          🎯 適用先:{' '}
          <b style={{ color: 'var(--excel-green)' }} data-testid="dialog-scope">
            {describeScope(scopeNow)}
          </b>
        </div>
        <div style={{ fontSize: 11, color: 'var(--text-dim)', lineHeight: 1.6 }}>
          対象: <b style={{ color: 'var(--text)' }}>{describeScope(scope)}</b>
          <br />
          範囲は各シートのデータが入っている範囲全体です。
          <br />
          適用先を変えると、色の一覧も数え直します。
        </div>
      </div>

      <h4 style={{ margin: '14px 0 6px' }}>
        1. 対象の色を選ぶ
        {colors && <span style={{ fontWeight: 400, color: 'var(--text-dim)' }}>（{colors.length} 色）</span>}
      </h4>

      {!colors && <div style={{ color: 'var(--text-dim)' }}>使われている色を調べています…</div>}

      {colors && colors.length === 0 && (
        <NoteBox kind="warn">
          対象の範囲に塗りつぶされたセルが見つかりませんでした。
          リボンの「適用先」で対象のブック / シートを広げてみてください。
        </NoteBox>
      )}

      {colors && colors.length > 0 && (
        <>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, margin: '8px 0' }}>
            {colors.map((c) => {
              const on = selected.has(c.key);
              const css = argbToCss(c.argb);
              return (
                <button
                  key={c.key}
                  type="button"
                  data-color={c.key}
                  onClick={() => {
                    const next = new Set(selected);
                    if (on) next.delete(c.key);
                    else next.add(c.key);
                    setSelected(next);
                    setPreview(null);
                  }}
                  title={`${c.count.toLocaleString()} セル / 例: ${c.sample}${
                    c.isApprox ? '\nテーマ色のため画面上の色は近似です' : ''
                  }`}
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    gap: 6,
                    padding: '4px 9px 4px 5px',
                    border: on ? '2px solid var(--excel-green)' : '1px solid var(--border)',
                    borderRadius: 4,
                    background: on ? '#e5f0ea' : '#fff',
                    cursor: 'pointer',
                    fontSize: 11.5,
                  }}
                >
                  <span
                    style={{
                      width: 20,
                      height: 20,
                      borderRadius: 3,
                      background: css,
                      border: '1px solid rgba(0,0,0,0.3)',
                      color: readableTextColor(c.argb),
                      fontSize: 12,
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                    }}
                  >
                    {on ? '✓' : ''}
                  </span>
                  <span style={{ fontVariantNumeric: 'tabular-nums' }}>
                    {c.count.toLocaleString()} セル
                  </span>
                  {c.isApprox && <span style={{ color: 'var(--text-dim)' }}>※</span>}
                </button>
              );
            })}
          </div>
          <div style={{ fontSize: 11, color: 'var(--text-dim)' }}>
            マウスを重ねると、そのセル数と見つかった場所の例が出ます。複数選べます。
          </div>
        </>
      )}

      <h4 style={{ margin: '16px 0 6px' }}>2. どうするか決める</h4>
      {!detailed ? (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 7 }}>
          {(
            [
              [
                'only',
                'この色のセル「だけ」入力できるようにする',
                'それ以外はすべてロックします。配って記入してもらう様式は、たいていこれです。',
              ],
              [
                'unlock',
                'この色のセルのロックを外す',
                'それ以外のロック状態はそのままにします。',
              ],
              [
                'lock',
                'この色のセルをロックする',
                'それ以外のロック状態はそのままにします。',
              ],
            ] as const
          ).map(([v, title, desc]) => (
            <label className="check" key={v} style={{ alignItems: 'flex-start' }}>
              <input
                type="radio"
                data-testid={`mode-${v}`}
                style={{ marginTop: 3 }}
                checked={mode === v}
                onChange={() => {
                  setMode(v);
                  setPreview(null);
                }}
              />
              <span>
                <b>{title}</b>
                {v === 'only' && (
                  <span style={{ color: 'var(--excel-green)', fontWeight: 700 }}>　← おすすめ</span>
                )}
                <br />
                <span style={{ color: 'var(--text-dim)', fontSize: 11.5 }}>{desc}</span>
              </span>
            </label>
          ))}
          <button
            type="button"
            onClick={() => setDetailed(true)}
            style={{
              alignSelf: 'flex-start',
              background: 'none',
              border: 'none',
              color: 'var(--text-dim)',
              textDecoration: 'underline',
              cursor: 'pointer',
              fontSize: 11.5,
              padding: '2px 0',
            }}
          >
            細かく指定する…
          </button>
        </div>
      ) : (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
          <label className="check">
            <input
              type="radio"
              checked={match === 'in'}
              onChange={() => {
                setMatch('in');
                setPreview(null);
              }}
            />
            選んだ色の<b>セルを</b>
          </label>
          <label className="check">
            <input
              type="radio"
              checked={match === 'out'}
              onChange={() => {
                setMatch('out');
                setPreview(null);
              }}
            />
            選んだ色<b>以外の</b>セルを
          </label>
          {match === 'out' && (
            <div style={{ paddingLeft: 20 }}>
              <Check
                label="塗りつぶしのないセルも対象に含める"
                checked={includeUnfilled}
                onChange={(v) => {
                  setIncludeUnfilled(v);
                  setPreview(null);
                }}
              />
            </div>
          )}
          {match === 'out' && (
            <div style={{ paddingLeft: 20 }}>
              <Check
                label={`選んだ色のセルは逆に${locked ? '入力できるようにする' : 'ロックする'}`}
                checked={alsoSetMatched}
                onChange={(v) => {
                  setAlsoSetMatched(v);
                  setPreview(null);
                }}
                title="Excel のセルは既定で全てロック済みなので、これを外すと選んだ色のセルもロックされたままになります"
              />
            </div>
          )}
          <div style={{ display: 'flex', gap: 14, marginTop: 4 }}>
            <label className="check">
              <input
                type="radio"
                checked={!locked}
                onChange={() => {
                  setLocked(false);
                  setPreview(null);
                }}
              />
              🔓 ロック解除する（入力できるようにする）
            </label>
            <label className="check">
              <input
                type="radio"
                checked={locked}
                onChange={() => {
                  setLocked(true);
                  setPreview(null);
                }}
              />
              🔒 ロックする
            </label>
          </div>
        </div>
      )}

      {preview && (
        <NoteBox kind="info">
          <b>試算:</b> {preview}
          <br />
          まだ変更していません。よければ「実行する」を押してください。
        </NoteBox>
      )}

      {hasApprox && (
        <NoteBox kind="warn">
          ※ が付いた色は<b>テーマ色</b>です。画面に出している色は標準テーマからの推定なので、
          実際の Excel と少し違って見えることがあります。
          <b>どのセルを選ぶかの判定は色の指定そのもので行う</b>ため、
          判定結果は正確です。
        </NoteBox>
      )}

      <NoteBox>
        ロックの設定だけでは Excel 上の動作は変わりません。
        仕上げに「ロック」タブの<b>「シート保護を有効化」</b>を実行してください。
        {(detailed ? match === 'out' && alsoSetMatched : mode === 'only') && (
          <>
            <br />
            この設定では、<b>選んだ色のセルだけが入力できる</b>状態になります。
          </>
        )}
      </NoteBox>
    </Modal>
  );
}
