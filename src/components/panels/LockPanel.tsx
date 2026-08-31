import { useState } from 'react';
import {
  BigButton,
  Btn,
  Check,
  Field,
  NoteBox,
  RangeSelector,
  RCol,
  RGroup,
  ScopeSelector,
  toRangeSpec,
  type RangeMode,
} from '../ui';
import { DEFAULT_PROTECT_OPTIONS, type SheetProtectOptions } from '../../recipe/types';
import { rectToA1 } from '../../excel/cellRef';
import { runOperation, setState, toast, useStore } from '../../state/store';

/**
 * ロック (セルの保護) 操作。
 *
 * Excel の仕様上の注意:
 *   ・セルの「ロック」属性は既定で ON。
 *   ・「シートの保護」を有効にしないとロックは一切効かない。
 * この 2 点を UI 上で常に見えるようにしている。
 */
export function LockPanel() {
  const s = useStore();
  const [rangeMode, setRangeMode] = useState<RangeMode>('selection');
  const [a1, setA1] = useState('');
  const [password, setPassword] = useState('');
  const [options, setOptions] = useState<SheetProtectOptions>(DEFAULT_PROTECT_OPTIONS);
  const [alsoUnlock, setAlsoUnlock] = useState(true);

  const selectionA1 = s.selection ? rectToA1(s.selection) : null;
  const ready = s.books.length > 0;

  function resolveRange() {
    const spec = toRangeSpec(rangeMode, a1, selectionA1);
    if (!spec) {
      toast('warn', '対象の範囲が決まっていません', 'グリッド上でセルを選ぶか、アドレスを入力してください。');
      return null;
    }
    return spec;
  }

  async function setLock(locked: boolean) {
    const range = resolveRange();
    if (!range) return;
    const r = await runOperation({ op: 'setLock', range, locked });
    toast(r.changedCells ? 'success' : 'info', r.summary);
  }

  async function lockAllExcept() {
    const range = resolveRange();
    if (!range) return;
    const r = await runOperation({ op: 'lockAllExcept', range, alsoUnlockTarget: alsoUnlock });
    toast(r.changedCells ? 'success' : 'info', r.summary, '仕上げに「シート保護を有効化」を実行してください。');
  }

  async function lockWholeSheet(locked: boolean) {
    const r = await runOperation({ op: 'setLock', range: { kind: 'used' }, locked });
    toast(r.changedCells ? 'success' : 'info', r.summary);
  }

  async function protectSheet() {
    const r = await runOperation({
      op: 'protectSheet',
      password: password || undefined,
      options,
    });
    toast('success', r.summary, password ? 'パスワードは手順書には保存されません。' : undefined);
  }

  async function unprotectSheet() {
    const r = await runOperation({ op: 'unprotectSheet' });
    toast(r.changedSheets ? 'success' : 'info', r.summary);
  }

  return (
    <>
      <RGroup title="セルのロック">
        <BigButton
          icon="🔒"
          label={<>選択範囲を<br />ロック</>}
          disabled={!ready}
          onClick={() => void setLock(true)}
        />
        <BigButton
          icon="🔓"
          label={<>選択範囲の<br />ロック解除</>}
          disabled={!ready}
          onClick={() => void setLock(false)}
        />
        <BigButton
          icon="🎯"
          label={<>選択範囲以外を<br />ロック</>}
          primary
          disabled={!ready}
          title="指定した範囲だけを入力可能にし、それ以外をすべてロックします"
          onClick={() => void lockAllExcept()}
        />
      </RGroup>

      <RGroup title="対象">
        <RangeSelector
          mode={rangeMode}
          a1={a1}
          selectionA1={selectionA1}
          onModeChange={setRangeMode}
          onA1Change={setA1}
        />
        <ScopeSelector scope={s.scope} onChange={(scope) => setState({ scope })} />
      </RGroup>

      <RGroup title="シート全体">
        <RCol>
          <Btn onClick={() => void lockWholeSheet(true)} disabled={!ready}>
            🔒 すべてロック
          </Btn>
          <Btn onClick={() => void lockWholeSheet(false)} disabled={!ready}>
            🔓 すべてロック解除
          </Btn>
          <Check
            label="「以外をロック」で対象範囲も解除する"
            checked={alsoUnlock}
            onChange={setAlsoUnlock}
            title="オフにすると、対象範囲のロック状態はそのままで、周囲だけをロックします"
          />
        </RCol>
      </RGroup>

      <RGroup title="シートの保護">
        <BigButton
          icon="🛡️"
          label={<>シート保護を<br />有効化</>}
          primary
          disabled={!ready}
          onClick={() => void protectSheet()}
        />
        <BigButton
          icon="🔧"
          label={<>シート保護を<br />解除</>}
          disabled={!ready}
          onClick={() => void unprotectSheet()}
        />
        <RCol>
          <Field label="パスワード">
            <input
              type="password"
              style={{ width: 120 }}
              value={password}
              placeholder="任意"
              onChange={(e) => setPassword(e.target.value)}
            />
          </Field>
          <Check
            label="ロックセルの選択を許可"
            checked={options.selectLockedCells}
            onChange={(v) => setOptions({ ...options, selectLockedCells: v })}
          />
          <Check
            label="書式変更を許可"
            checked={options.formatCells}
            onChange={(v) => setOptions({ ...options, formatCells: v })}
          />
          <Check
            label="オートフィルタを許可"
            checked={options.autoFilter}
            onChange={(v) => setOptions({ ...options, autoFilter: v })}
          />
          <Check
            label="並べ替えを許可"
            checked={options.sort}
            onChange={(v) => setOptions({ ...options, sort: v })}
          />
        </RCol>
      </RGroup>

      <RGroup title="ヒント">
        <div style={{ width: 246 }}>
          <NoteBox kind="warn">
            Excel では<b>セルのロックは「シートの保護」を有効にして初めて効きます</b>。
            <br />
            典型的な流れ:
            <br />① 入力させたい範囲を選ぶ → ②「選択範囲以外をロック」 → ③「シート保護を有効化」
          </NoteBox>
        </div>
      </RGroup>
    </>
  );
}
