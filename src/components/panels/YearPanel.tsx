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
  toRangeSpec,
  type RangeMode,
} from '../ui';
import { DEFAULT_YEAR_TARGETS, type StepBody, type YearTargets } from '../../recipe/types';
import { rectToA1 } from '../../excel/cellRef';
import { runOperation, toast, useStore } from '../../state/store';

/**
 * 年度の一括更新。
 *
 * 「年が書いてあるところを、まとめて 1 年ずらす」という 1 つの目的に絞る。
 * 2023→2024, 2024→2025, 2025→2026 … を 1 回の走査で同時に置換するため、
 * 値が二重に進むことはない。
 */
export function YearPanel() {
  const s = useStore();
  const thisYear = new Date().getFullYear();
  const [delta, setDelta] = useState(1);
  const [minYear, setMinYear] = useState(thisYear - 10);
  const [maxYear, setMaxYear] = useState(thisYear + 10);
  const [wholeNumberOnly, setWholeNumberOnly] = useState(true);
  const [includeNumericCells, setIncludeNumericCells] = useState(false);
  const [targets, setTargets] = useState<YearTargets>(DEFAULT_YEAR_TARGETS);
  const [rangeMode, setRangeMode] = useState<RangeMode>('used');
  const [a1, setA1] = useState('');
  const [replaceFind, setReplaceFind] = useState('');
  const [replaceTo, setReplaceTo] = useState('');
  const [wholeCell, setWholeCell] = useState(false);

  const selectionA1 = s.selection ? rectToA1(s.selection) : null;
  const ready = s.books.length > 0;
  const sign = delta >= 0 ? '+' : '−';
  const abs = Math.abs(delta);

  async function doApply() {
    if (minYear > maxYear) {
      toast('warn', '対象にする年の範囲が逆になっています');
      return;
    }
    const range = toRangeSpec(rangeMode, a1, selectionA1) ?? { kind: 'used' as const };
    const body: StepBody = {
      op: 'shiftYears',
      delta,
      minYear,
      maxYear,
      wholeNumberOnly,
      includeNumericCells,
      targets,
      range,
    };
    const outcome = await runOperation(body);
    toast(
      outcome.changedCells ? 'success' : 'warn',
      outcome.summary,
      outcome.changedCells
        ? '画面で結果を確かめてから、「ファイル」タブで保存してください。'
        : `${minYear}〜${maxYear} 年の数字が見つかりませんでした。「対象にする年」の範囲をご確認ください。`,
    );
  }

  async function doReplaceText() {
    if (!replaceFind) {
      toast('warn', '置き換えたい文字を入力してください');
      return;
    }
    const range = toRangeSpec(rangeMode, a1, selectionA1) ?? { kind: 'used' as const };
    const outcome = await runOperation({
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
    });
    toast(outcome.changedCells ? 'success' : 'info', outcome.summary);
  }

  return (
    <>
      <RGroup title="年をずらす">
        <BigButton
          icon="📅"
          label={
            <>
              年を {sign}
              {abs} 年
              <br />
              ずらす
            </>
          }
          primary
          disabled={!ready}
          title="見つかった年をまとめて書き換えます。▾ でどのブック・シートに効かせるかを選べます"
          scopeMenu
          onClick={() => void doApply()}
        />
        <RCol>
          <Field label="ずらす年数">
            <input
              type="number"
              data-testid="year-delta"
              style={{ width: 54 }}
              value={delta}
              onChange={(e) => setDelta(parseInt(e.target.value, 10) || 0)}
            />
            <span style={{ color: 'var(--text-dim)' }}>年</span>
          </Field>
          <div style={{ width: 236 }}>
            <NoteBox>
              「2024年度」→「{2024 + delta}年度」、「2025年度」→「{2025 + delta}年度」…と、
              <b>見つかった年をすべて同時に</b>書き換えます。
              順番に置換しないので、2024 が {2024 + delta * 2} まで進むことはありません。
            </NoteBox>
          </div>
        </RCol>
      </RGroup>

      <RGroup title="どこに書いてある年を変えるか">
        <RCol>
          <Check
            label="セルの文字・数字"
            checked={targets.values}
            onChange={(v) => setTargets({ ...targets, values: v })}
            title="例: 「2024年度実績」→「2025年度実績」"
          />
          <Check
            label="数式の中身"
            checked={targets.formulas}
            onChange={(v) => setTargets({ ...targets, formulas: v })}
            title="例: ='2024年度'!B5 → ='2025年度'!B5。シート名を変えるときは必ず ON にしてください"
          />
          <Check
            label="シート名"
            testId="year-sheetnames"
            checked={targets.sheetNames}
            onChange={(v) => setTargets({ ...targets, sheetNames: v })}
            title="例: 「2024年度実績」タブ →「2025年度実績」タブ"
          />
          <Check
            label="ファイル名 (保存するとき)"
            checked={targets.fileNames}
            onChange={(v) => setTargets({ ...targets, fileNames: v })}
            title="例: 2025年度予算_東京支店.xlsx → 2026年度予算_東京支店.xlsx"
          />
          <Check
            label="和暦 (令和6年 → 令和7年)"
            checked={targets.japaneseEra}
            onChange={(v) => setTargets({ ...targets, japaneseEra: v })}
            title="「令和6年度」のような書き方も一緒にずらします"
          />
        </RCol>
        <div style={{ width: 216 }}>
          <NoteBox>
            チェックした場所の中から、<b>年に見える数字だけ</b>をこのツールが探して書き換えます。
            どの年が書いてあるかを前もって調べておく必要はありません。
          </NoteBox>
        </div>
      </RGroup>

      <RGroup title="間違って変えないための設定">
        <RCol>
          <Check
            label="数字だけのセルも年とみなす"
            testId="year-numeric"
            checked={includeNumericCells}
            onChange={setIncludeNumericCells}
            title="OFF のままにしておくと、数量 2,031 個・金額 2,025 円のような「たまたま年に見える数字」は変わりません"
          />
          <Check
            label="長い数字の途中は変えない"
            checked={wholeNumberOnly}
            onChange={setWholeNumberOnly}
            title="ON のとき 20240401 の中の 2024 は置換しません"
          />
          <Field label="対象にする年">
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
        </RCol>
        <div style={{ width: 214 }}>
          <NoteBox kind={includeNumericCells ? 'warn' : 'ok'}>
            {includeNumericCells ? (
              <>
                <b>数量 2,031 個</b>のように、年と同じ 4 桁の数字だけが入ったセルも
                <b>年として書き換わります</b>。心当たりがあるときはチェックを外してください。
              </>
            ) : (
              <>
                <b>✓ 数量 2,031 個・金額 2,025 円のような数字は変わりません。</b>
                「2025年度」のように文字と一緒に書かれた年だけを書き換えます。
              </>
            )}
          </NoteBox>
        </div>
      </RGroup>

      <RGroup title="シートの中のどこを見るか">
        <RangeSelector
          mode={rangeMode}
          a1={a1}
          selectionA1={selectionA1}
          onModeChange={setRangeMode}
          onA1Change={setA1}
          note="ここで決めた範囲の中だけを探します。ふつうは「データが入っている範囲全体」のままで結構です。"
        />
      </RGroup>

      <RGroup title="年ではない文字を置き換える">
        <ScopeBadge note="「年をずらす」と同じ「適用先」がそのまま使われます" />
        <RCol>
          <Field label="この文字を">
            <input
              type="text"
              style={{ width: 128 }}
              data-testid="replace-find"
              value={replaceFind}
              placeholder="例: 旧部署名"
              onChange={(e) => setReplaceFind(e.target.value)}
            />
          </Field>
          <Field label="この文字に">
            <input
              type="text"
              style={{ width: 128 }}
              data-testid="replace-to"
              value={replaceTo}
              placeholder="例: 新部署名"
              onChange={(e) => setReplaceTo(e.target.value)}
            />
          </Field>
          <Check label="セル全体が一致する場合のみ" checked={wholeCell} onChange={setWholeCell} />
          <Btn kind="accent" onClick={() => void doReplaceText()} disabled={!ready}>
            置き換える
          </Btn>
        </RCol>
        <div style={{ width: 210 }}>
          <NoteBox>
            <b>上の「年をずらす」との違い</b>
            <br />
            年をずらす … 年の数字を<b>ツールが探して</b>ずらす（何年と書いてあるか知らなくてよい）
            <br />
            こちら … <b>自分で指定した文字</b>をそのまま別の文字に置き換える（年とは関係ない書き換え用）
          </NoteBox>
        </div>
      </RGroup>
    </>
  );
}
