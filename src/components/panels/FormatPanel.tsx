import { useState } from 'react';
import {
  BigButton,
  Btn,
  Check,
  ColorSwatches,
  NoteBox,
  RCol,
  RGroup,
  ScopeSelector,
} from '../ui';
import { rectToA1 } from '../../excel/cellRef';
import { argbToCss } from '../../excel/format';
import { ColorLockDialog } from './ColorLockDialog';
import { runOperation, setState, toast, useStore } from '../../state/store';

/**
 * 塗りつぶし操作。
 * 「ロックしていないセル (= 入力してもらう欄) だけ色を付ける」という
 * 実務でよく使う一括処理をワンボタンで行えるようにしている。
 */
export function FormatPanel() {
  const s = useStore();
  const [color, setColor] = useState<string | null>('FFFFF2CC');
  const [onlyWithValue, setOnlyWithValue] = useState(false);
  const [showColorLock, setShowColorLock] = useState(false);

  const selectionA1 = s.selection ? rectToA1(s.selection) : null;
  const ready = s.books.length > 0;

  /**
   * 色をクリックしたら、選んでいるセルにその場で反映する。
   * (Excel の「塗りつぶしの色」と同じ感覚で使えるように)
   *
   * セルを選んでいないときは色を選ぶだけにする。
   * こうしておくと「ロック状態で一括」で使う色を、
   * うっかり塗ってしまわずに変更できる。
   */
  async function pickColor(argb: string | null) {
    setColor(argb);
    if (!selectionA1) {
      toast(
        'info',
        argb === null ? '「塗りなし」を選びました' : '色を選びました',
        'セルを選んでから色を押すと、その場で塗れます。',
      );
      return;
    }
    const r = await runOperation({
      op: 'fillRange',
      range: { kind: 'a1', a1: selectionA1 },
      colorArgb: argb,
    });
    toast(r.changedCells ? 'success' : 'info', r.summary);
  }

  async function fillByLock(target: 'locked' | 'unlocked', clear: boolean) {
    if (!clear && color === null) {
      toast('warn', '色が選ばれていません', '左の「塗りつぶしの色」で色を選んでください。');
      return;
    }
    const r = await runOperation({
      op: 'fillByLockState',
      target,
      colorArgb: clear ? null : color,
      onlyUsedRange: onlyWithValue,
    });
    toast(r.changedCells ? 'success' : 'info', r.summary);
  }

  return (
    <>
      <RGroup title="塗りつぶしの色">
        <RCol>
          <ColorSwatches value={color} onChange={(c) => void pickColor(c)} />
          <div style={{ fontSize: 10.5, color: 'var(--text-dim)', width: 150, lineHeight: 1.6 }}>
            {selectionA1 ? (
              <>
                <b>{selectionA1}</b> にその場で反映します。
              </>
            ) : (
              'セルを選んでから色を押すと、その場で塗れます。'
            )}
            <br />
            斜線は<b>塗りなし</b>、右端の四角で任意の色を選べます。
          </div>
        </RCol>
      </RGroup>

      <RGroup title="ロック状態で一括">
        <BigButton
          icon="🟡"
          label={<>ロック解除セル<br />を色分け</>}
          primary
          disabled={!ready}
          title="入力してもらう欄 (ロックしていないセル) をまとめて塗ります"
          onClick={() => void fillByLock('unlocked', false)}
        />
        <BigButton
          icon="⬜"
          label={<>ロック済みセル<br />を色分け</>}
          disabled={!ready}
          onClick={() => void fillByLock('locked', false)}
        />
        <RCol>
          <div
            style={{ display: 'flex', alignItems: 'center', gap: 5, fontSize: 10.5 }}
            title="左の「塗りつぶしの色」で選んだ色を使います"
          >
            <span
              style={{
                width: 16,
                height: 16,
                borderRadius: 3,
                border: '1px solid var(--border-strong)',
                background: color ? argbToCss(color) : undefined,
                backgroundImage: color
                  ? undefined
                  : 'repeating-linear-gradient(45deg,#fff,#fff 3px,#e11 3px,#e11 4px)',
                flexShrink: 0,
              }}
            />
            <span style={{ color: 'var(--text-dim)' }}>左で選んだ色を使います</span>
          </div>
          <Btn onClick={() => void fillByLock('unlocked', true)} disabled={!ready}>
            ロック解除セルの塗りを解除
          </Btn>
          <Btn onClick={() => void fillByLock('locked', true)} disabled={!ready}>
            ロック済みセルの塗りを解除
          </Btn>
          <Check
            label="値が入っているセルのみ"
            checked={onlyWithValue}
            onChange={setOnlyWithValue}
            title="空欄には色を付けません"
          />
        </RCol>
      </RGroup>

      <RGroup title="色からロック">
        <BigButton
          icon="🔎"
          label={<>色から<br />ロックを設定</>}
          primary
          disabled={!ready}
          title="実際に使われている色を一覧から選び、その色のセル (またはそれ以外) のロックを切り替えます"
          onClick={() => setShowColorLock(true)}
        />
        <RCol>
          <div style={{ width: 216 }}>
            <NoteBox>
              「<b>黄色が入力欄</b>」のような色分けが既にあるファイルなら、
              その色を選ぶだけでロック設定を起こせます。
              ファイル内で実際に使われている色を数えて一覧にします。
            </NoteBox>
          </div>
        </RCol>
      </RGroup>

      <RGroup title="適用先">
        <ScopeSelector scope={s.scope} onChange={(scope) => setState({ scope })} />
      </RGroup>

      <RGroup title="画面表示">
        <RCol>
          <Check
            label="ロック状態を色分け表示"
            checked={s.showLockOverlay}
            onChange={(v) => setState({ showLockOverlay: v })}
          />
          <div style={{ width: 224 }}>
            <NoteBox>
              画面上の斜線 / 黄色は<b>この画面だけの表示</b>で、ファイルには保存されません。
              実際のファイルに色を付けたい場合は上のボタンを使ってください。
            </NoteBox>
          </div>
        </RCol>
      </RGroup>

      {showColorLock && <ColorLockDialog onClose={() => setShowColorLock(false)} />}
    </>
  );
}
