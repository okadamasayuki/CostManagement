import { useState } from 'react';
import { Btn, Check, ColorSwatches, Field, Modal, NoteBox, ScopeSelector } from '../ui';
import type { CellCondition, ConditionAction, StepBody } from '../../recipe/types';
import { DEFAULT_CONDITION } from '../../recipe/types';
import { describeCondition, describeScope } from '../../recipe/describe';
import { previewOperation, runOperation, setState, toast, useStore } from '../../state/store';

/**
 * 条件を指定してセルを塗る / ロックする画面。
 *
 * 「数値が入っているセルに色を付ける」「500,000 を超えるセルを目立たせる」
 * のように、1 つ 1 つ選ばなくても中身から対象を決められるようにする。
 * 200 ファイルのような数を手で塗るのは現実的でないため。
 */
export function ConditionDialog(props: { onClose(): void }) {
  const s = useStore();
  const [condition, setCondition] = useState<CellCondition>(DEFAULT_CONDITION);
  const [useNumber, setUseNumber] = useState(false);
  const [useText, setUseText] = useState(false);
  const [action, setAction] = useState<'fill' | 'lock' | 'unlock'>('fill');
  const [color, setColor] = useState<string | null>('FFFFF2CC');
  const [preview, setPreview] = useState<string | null>(null);

  const patch = (p: Partial<CellCondition>) => {
    setCondition({ ...condition, ...p });
    setPreview(null);
  };

  function buildCondition(): CellCondition {
    return {
      kind: condition.kind,
      number: useNumber ? (condition.number ?? { op: 'gt', a: 0 }) : undefined,
      text: useText ? (condition.text ?? { op: 'contains', value: '', matchCase: false }) : undefined,
    };
  }

  function buildBody(): StepBody {
    const act: ConditionAction =
      action === 'fill' ? { kind: 'fill', colorArgb: color } : { kind: 'lock', locked: action === 'lock' };
    return {
      op: 'applyByCondition',
      condition: buildCondition(),
      action: act,
      range: { kind: 'used' },
    };
  }

  async function doPreview() {
    const outcome = await previewOperation(buildBody());
    setPreview(outcome.summary);
  }

  async function doApply() {
    const outcome = await runOperation(buildBody());
    toast(outcome.changedCells ? 'success' : 'info', outcome.summary);
    props.onClose();
  }

  const num = condition.number ?? { op: 'gt' as const, a: 0 };
  const txt = condition.text ?? { op: 'contains' as const, value: '', matchCase: false };

  return (
    <Modal
      title="条件を指定して塗る / ロックする"
      wide
      onClose={props.onClose}
      footer={
        <>
          <Btn onClick={props.onClose}>キャンセル</Btn>
          <Btn onClick={() => void doPreview()}>試算</Btn>
          <Btn kind="accent" onClick={() => void doApply()}>
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
        <ScopeSelector scope={s.scope} onChange={(next) => setState({ scope: next })} />
        <div style={{ fontSize: 11, color: 'var(--text-dim)', lineHeight: 1.6 }}>
          対象: <b style={{ color: 'var(--text)' }}>{describeScope(s.scope)}</b>
          <br />
          範囲は各シートのデータが入っている範囲全体です。
        </div>
      </div>

      <h4 style={{ margin: '14px 0 6px' }}>1. どのセルを対象にするか</h4>
      <Field label="セルの種類">
        <select
          data-testid="cond-kind"
          value={condition.kind}
          onChange={(e) => patch({ kind: e.target.value as CellCondition['kind'] })}
        >
          <option value="number">数値が入っているセル</option>
          <option value="text">文字が入っているセル</option>
          <option value="formula">数式が入っているセル</option>
          <option value="blank">空のセル</option>
          <option value="any">すべてのセル</option>
        </select>
      </Field>

      <div style={{ marginTop: 8 }}>
        <Check
          label="値の大きさで絞り込む"
          checked={useNumber}
          onChange={(v) => {
            setUseNumber(v);
            setPreview(null);
          }}
        />
        {useNumber && (
          <div style={{ display: 'flex', gap: 6, alignItems: 'center', paddingLeft: 20, marginTop: 4 }}>
            <input
              type="number"
              data-testid="cond-num-a"
              style={{ width: 120 }}
              value={num.a}
              onChange={(e) => patch({ number: { ...num, a: Number(e.target.value) } })}
            />
            <select
              data-testid="cond-num-op"
              value={num.op}
              onChange={(e) => patch({ number: { ...num, op: e.target.value as typeof num.op } })}
            >
              <option value="gt">より大きい</option>
              <option value="ge">以上</option>
              <option value="lt">より小さい</option>
              <option value="le">以下</option>
              <option value="eq">と等しい</option>
              <option value="ne">以外</option>
              <option value="between">〜の範囲</option>
            </select>
            {num.op === 'between' && (
              <input
                type="number"
                style={{ width: 120 }}
                value={num.b ?? num.a}
                onChange={(e) => patch({ number: { ...num, b: Number(e.target.value) } })}
              />
            )}
          </div>
        )}
      </div>

      <div style={{ marginTop: 8 }}>
        <Check
          label="文字で絞り込む"
          checked={useText}
          onChange={(v) => {
            setUseText(v);
            setPreview(null);
          }}
        />
        {useText && (
          <div style={{ display: 'flex', gap: 6, alignItems: 'center', paddingLeft: 20, marginTop: 4 }}>
            <input
              type="text"
              data-testid="cond-text"
              placeholder="例: 合計"
              style={{ width: 160 }}
              value={txt.value}
              onChange={(e) => patch({ text: { ...txt, value: e.target.value } })}
            />
            <select
              value={txt.op}
              onChange={(e) => patch({ text: { ...txt, op: e.target.value as typeof txt.op } })}
            >
              <option value="contains">を含む</option>
              <option value="startsWith">で始まる</option>
              <option value="endsWith">で終わる</option>
              <option value="equals">と一致する</option>
            </select>
            <Check
              label="大文字小文字を区別"
              checked={txt.matchCase}
              onChange={(v) => patch({ text: { ...txt, matchCase: v } })}
            />
          </div>
        )}
      </div>

      <h4 style={{ margin: '16px 0 6px' }}>2. そのセルをどうするか</h4>
      <div style={{ display: 'flex', gap: 14, flexWrap: 'wrap' }}>
        {(
          [
            ['fill', '🎨 塗りつぶす'],
            ['lock', '🔒 ロックする'],
            ['unlock', '🔓 入力できるようにする'],
          ] as const
        ).map(([v, label]) => (
          <label className="check" key={v}>
            <input
              type="radio"
              checked={action === v}
              onChange={() => {
                setAction(v);
                setPreview(null);
              }}
            />
            {label}
          </label>
        ))}
      </div>
      {action === 'fill' && (
        <div style={{ marginTop: 8, maxWidth: 200 }}>
          <ColorSwatches
            value={color}
            onChange={(c) => {
              setColor(c);
              setPreview(null);
            }}
          />
        </div>
      )}

      <NoteBox>
        <b>指定内容:</b> {describeCondition(buildCondition())} を
        {action === 'fill'
          ? color === null
            ? '塗りつぶし解除'
            : '塗りつぶす'
          : action === 'lock'
            ? 'ロックする'
            : '入力できるようにする'}
      </NoteBox>

      {preview && (
        <NoteBox kind="info">
          <b>試算:</b> {preview}
          <br />
          まだ変更していません。よければ「実行する」を押してください。
        </NoteBox>
      )}
    </Modal>
  );
}
