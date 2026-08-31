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
  ScopeBadge,
  ScopeSelector,
  toRangeSpec,
  type RangeMode,
} from '../ui';
import { DEFAULT_YEAR_TARGETS, type StepBody, type YearTargets } from '../../recipe/types';
import { rectToA1 } from '../../excel/cellRef';
import { previewOperation, runOperation, setState, toast, useStore } from '../../state/store';

/**
 * 年度の一括更新。
 *
 * 既定は「+1 年ずらす」。2023→2024, 2024→2025, 2025→2026 … を
 * 1 回の走査で同時に置換するため、値が二重に進むことはない。
 */
export function YearPanel() {
  const s = useStore();
  const thisYear = new Date().getFullYear();
  const [mode, setMode] = useState<'shift' | 'map'>('shift');
  const [delta, setDelta] = useState(1);
  const [minYear, setMinYear] = useState(thisYear - 10);
  const [maxYear, setMaxYear] = useState(thisYear + 10);
  const [wholeNumberOnly, setWholeNumberOnly] = useState(true);
  const [targets, setTargets] = useState<YearTargets>(DEFAULT_YEAR_TARGETS);
  const [pairsText, setPairsText] = useState('2024→2025\n2025→2026');
  const [rangeMode, setRangeMode] = useState<RangeMode>('used');
  const [a1, setA1] = useState('');
  const [replaceFind, setReplaceFind] = useState('');
  const [replaceTo, setReplaceTo] = useState('');
  const [wholeCell, setWholeCell] = useState(false);

  const selectionA1 = s.selection ? rectToA1(s.selection) : null;
  const ready = s.books.length > 0;

  function parsePairs(): Array<{ from: number; to: number }> {
    return pairsText
      .split(/[\n,]/)
      .map((line) => line.trim())
      .filter(Boolean)
      .map((line) => {
        const m = /^(\d{4})\s*(?:→|->|=>|:|、|,)\s*(\d{4})$/.exec(line);
        return m ? { from: parseInt(m[1], 10), to: parseInt(m[2], 10) } : null;
      })
      .filter((p): p is { from: number; to: number } => p !== null);
  }

  function buildBody(): StepBody | null {
    const range = toRangeSpec(rangeMode, a1, selectionA1) ?? { kind: 'used' as const };
    if (mode === 'shift') {
      if (minYear > maxYear) {
        toast('warn', '対象年の範囲が逆になっています');
        return null;
      }
      return { op: 'shiftYears', delta, minYear, maxYear, wholeNumberOnly, targets, range };
    }
    const pairs = parsePairs();
    if (!pairs.length) {
      toast('warn', '対応表を読み取れませんでした', '「2024→2025」のように 1 行ずつ書いてください。');
      return null;
    }
    return { op: 'mapYears', pairs, wholeNumberOnly, targets, range };
  }

  async function doPreview() {
    const body = buildBody();
    if (!body) return;
    const outcome = await previewOperation(body);
    setState({ preview: { label: '年度更新の試算', outcome } });
    toast(
      outcome.changedCells ? 'info' : 'warn',
      outcome.changedCells ? `試算: ${outcome.summary}` : '変更対象は見つかりませんでした',
      outcome.changedCells ? '右側のパネルで内訳を確認できます。' : undefined,
    );
  }

  async function doApply() {
    const body = buildBody();
    if (!body) return;
    const outcome = await runOperation(body);
    setState({ preview: null });
    toast(outcome.changedCells ? 'success' : 'info', outcome.summary);
  }

  async function doReplaceText(preview: boolean) {
    if (!replaceFind) {
      toast('warn', '検索する文字列を入力してください');
      return;
    }
    const range = toRangeSpec(rangeMode, a1, selectionA1) ?? { kind: 'used' as const };
    const body: StepBody = {
      op: 'replaceText',
      find: replaceFind,
      replace: replaceTo,
      matchCase: false,
      wholeCell,
      targets: {
        values: targets.values,
        formulas: targets.formulas,
        sheetNames: targets.sheetNames,
        fileNames: targets.fileNames,
      },
      range,
    };
    if (preview) {
      const outcome = await previewOperation(body);
      setState({ preview: { label: '文字列置換の試算', outcome } });
      toast(outcome.changedCells ? 'info' : 'warn', `試算: ${outcome.summary}`);
    } else {
      const outcome = await runOperation(body);
      setState({ preview: null });
      toast(outcome.changedCells ? 'success' : 'info', outcome.summary);
    }
  }

  const sample =
    mode === 'shift'
      ? `${minYear}→${minYear + delta}, ${minYear + 1}→${minYear + 1 + delta}, … , ${maxYear}→${maxYear + delta}`
      : parsePairs()
          .map((p) => `${p.from}→${p.to}`)
          .join(', ') || '(未設定)';

  return (
    <>
      <RGroup title="置換の方法">
        <RCol>
          <Field label="">
            <select value={mode} onChange={(e) => setMode(e.target.value as 'shift' | 'map')}>
              <option value="shift">年をまとめてずらす (推奨)</option>
              <option value="map">対応表で指定する</option>
            </select>
          </Field>
          {mode === 'shift' ? (
            <>
              <Field label="ずらす年数">
                <input
                  type="number"
                  style={{ width: 54 }}
                  value={delta}
                  onChange={(e) => setDelta(parseInt(e.target.value, 10) || 0)}
                />
                <span style={{ color: 'var(--text-dim)' }}>年</span>
              </Field>
              <Field label="対象">
                <input
                  type="number"
                  style={{ width: 62 }}
                  value={minYear}
                  onChange={(e) => setMinYear(parseInt(e.target.value, 10) || 0)}
                />
                <span>〜</span>
                <input
                  type="number"
                  style={{ width: 62 }}
                  value={maxYear}
                  onChange={(e) => setMaxYear(parseInt(e.target.value, 10) || 0)}
                />
              </Field>
            </>
          ) : (
            <textarea
              value={pairsText}
              onChange={(e) => setPairsText(e.target.value)}
              rows={4}
              style={{
                width: 178,
                fontSize: 11,
                fontFamily: 'Consolas, monospace',
                border: '1px solid var(--border)',
                borderRadius: 3,
                padding: 4,
                resize: 'vertical',
              }}
              placeholder={'2024→2025\n2025→2026'}
            />
          )}
        </RCol>
      </RGroup>

      <RGroup title="置換する場所">
        <RCol>
          <Check label="セルの値" checked={targets.values} onChange={(v) => setTargets({ ...targets, values: v })} />
          <Check
            label="数式の中身"
            checked={targets.formulas}
            onChange={(v) => setTargets({ ...targets, formulas: v })}
            title="シート名を変更する場合は、参照が壊れないようこれも ON にしてください"
          />
          <Check
            label="シート名"
            checked={targets.sheetNames}
            onChange={(v) => setTargets({ ...targets, sheetNames: v })}
          />
          <Check
            label="ファイル名 (保存時)"
            checked={targets.fileNames}
            onChange={(v) => setTargets({ ...targets, fileNames: v })}
          />
          <Check
            label="和暦 (令和6年 など)"
            checked={targets.japaneseEra}
            onChange={(v) => setTargets({ ...targets, japaneseEra: v })}
            disabled={mode !== 'shift'}
            title="「ずらす」モードでのみ使えます"
          />
          <Check
            label="数字の途中は対象外"
            checked={wholeNumberOnly}
            onChange={setWholeNumberOnly}
            title="ON のとき 20240401 のような並びの中の 2024 は置換しません"
          />
        </RCol>
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

      <RGroup title="実行">
        <BigButton icon="🔍" label={<>変更内容を<br />試算</>} disabled={!ready} onClick={() => void doPreview()} />
        <BigButton
          icon="📅"
          label={<>年度更新を<br />実行</>}
          primary
          disabled={!ready}
          onClick={() => void doApply()}
        />
        <div style={{ width: 224 }}>
          <NoteBox>
            <b>置換例:</b> <span className="mono">{sample}</span>
            <br />
            すべて<b>同時に</b>置換されるため、2024→2025 と 2025→2026 を同時に指定しても
            元の 2024 が 2026 まで進むことはありません。
          </NoteBox>
        </div>
      </RGroup>

      <RGroup title="文字列の置換">
        <ScopeBadge note="年度更新と同じ「対象」「適用先」が使われます" />
        <RCol>
          <Field label="検索">
            <input
              type="text"
              style={{ width: 128 }}
              value={replaceFind}
              placeholder="例: 令和6年度"
              onChange={(e) => setReplaceFind(e.target.value)}
            />
          </Field>
          <Field label="置換">
            <input
              type="text"
              style={{ width: 128 }}
              value={replaceTo}
              placeholder="例: 令和7年度"
              onChange={(e) => setReplaceTo(e.target.value)}
            />
          </Field>
          <Check label="セル全体が一致する場合のみ" checked={wholeCell} onChange={setWholeCell} />
          <div style={{ display: 'flex', gap: 4 }}>
            <Btn onClick={() => void doReplaceText(true)} disabled={!ready}>
              試算
            </Btn>
            <Btn kind="accent" onClick={() => void doReplaceText(false)} disabled={!ready}>
              置換する
            </Btn>
          </div>
        </RCol>
      </RGroup>
    </>
  );
}
