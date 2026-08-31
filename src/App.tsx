import { useEffect, useMemo, useState } from 'react';
import type ExcelJS from 'exceljs';
import { Ribbon } from './components/Ribbon';
import { FileTree } from './components/FileTree';
import { Grid } from './components/Grid';
import { SheetTabs } from './components/SheetTabs';
import { FormulaBar } from './components/FormulaBar';
import { StatusBar } from './components/StatusBar';
import { RightPanel } from './components/RightPanel';
import { BusyOverlay, Toasts } from './components/Toasts';
import { Btn, Modal, NoteBox } from './components/ui';
import { buildSheetView } from './excel/view';
import { getBlockedAttempts, onBlockedAttempt } from './security/networkGuard';
import { isHosted } from './security/selfCopy';
import { currentBook, currentSheet, setState, touch, useStore } from './state/store';

export default function App() {
  const s = useStore();
  const [showAbout, setShowAbout] = useState(false);
  const [blockedCount, setBlockedCount] = useState(getBlockedAttempts().length);
  useEffect(() => onBlockedAttempt((l) => setBlockedCount(l.length)), []);

  const book = currentBook();
  const ws = currentSheet();

  // ブックの内容が変わったときだけ描画用スナップショットを作り直す
  const view = useMemo(() => {
    if (!ws) return null;
    return buildSheetView(ws, ws.id);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [ws, s.currentBookId, s.currentSheetName, s.docVersion]);

  // 保存していない変更があるときに、うっかりタブを閉じるのを防ぐ
  useEffect(() => {
    const handler = (e: BeforeUnloadEvent) => {
      if (s.books.some((b) => b.dirty)) {
        e.preventDefault();
        e.returnValue = '';
      }
    };
    window.addEventListener('beforeunload', handler);
    return () => window.removeEventListener('beforeunload', handler);
  }, [s.books]);

  function commitEdit(row: number, col: number, text: string) {
    if (!ws || !book) return;
    const cell = ws.getRow(row).getCell(col);
    cell.value = parseInput(text);
    book.dirty = true;
    touch();
  }

  const docName = book
    ? `${s.renames[book.id] ?? book.fileName}${book.dirty ? ' *' : ''}${
        view ? ` — ${view.name}` : ''
      }`
    : '';

  return (
    <div className="app">
      <div className="titlebar">
        <span className="app-name">Excel 一括ロック / 年度更新ツール</span>
        <span className="doc-name">{docName}</span>
        <span className="spacer" />
        <span
          className={`offline-badge${blockedCount ? ' alarm' : ''}`}
          title={
            blockedCount
              ? `${blockedCount} 件の外部通信を遮断しました (F12 の「ネットワーク」タブでも確認できます)`
              : 'このツールは外部と通信しません。読み込んだファイルはブラウザー内だけで処理されます。'
          }
        >
          {blockedCount > 0
            ? `⚠️ 通信を ${blockedCount} 件遮断`
            : isHosted()
              ? '🔒 データは外部に出ません'
              : '🔒 完全オフライン動作'}
        </span>
        <button className="offline-badge" onClick={() => setShowAbout(true)} type="button">
          ？ 使い方
        </button>
      </div>

      <Ribbon />

      <div className="main">
        <FileTree />
        <div className="workspace">
          <FormulaBar view={view} onCommit={commitEdit} readOnly={!ws} />
          {view ? (
            <Grid
              view={view}
              selection={s.selection}
              anchor={s.anchor}
              showLockOverlay={s.showLockOverlay}
              readOnly={false}
              onSelect={(rect, anchor) => setState({ selection: rect, anchor })}
              onCommitEdit={commitEdit}
            />
          ) : (
            <div className="gridwrap">
              <div className="grid-placeholder">
                <div style={{ fontSize: 46 }}>📗</div>
                <div>
                  <b>Excel ファイルを読み込んでください</b>
                  <br />
                  「ファイル」タブ →「フォルダーを開く」で、
                  <br />
                  フォルダーの中の Excel をサブフォルダーごと一括で読み込めます。
                </div>
                <Btn kind="accent" onClick={() => setShowAbout(true)}>
                  はじめての方はこちら
                </Btn>
              </div>
            </div>
          )}
          <SheetTabs />
        </div>
        <RightPanel view={view} />
      </div>

      <StatusBar view={view} />
      <Toasts />
      <BusyOverlay />

      {showAbout && <AboutModal onClose={() => setShowAbout(false)} />}
    </div>
  );
}

/** 入力欄の文字列を ExcelJS の値に変換する */
function parseInput(text: string): ExcelJS.CellValue {
  const t = text.trim();
  if (t === '') return null;
  if (t.startsWith('=')) return { formula: t.slice(1) } as ExcelJS.CellFormulaValue;
  if (/^-?\d+(\.\d+)?$/.test(t)) return Number(t);
  if (/^-?[\d,]+(\.\d+)?$/.test(t) && t.includes(',')) {
    const n = Number(t.replace(/,/g, ''));
    if (!Number.isNaN(n)) return n;
  }
  return text;
}

function AboutModal({ onClose }: { onClose(): void }) {
  return (
    <Modal
      title="このツールについて"
      wide
      onClose={onClose}
      footer={
        <Btn kind="accent" onClick={onClose}>
          閉じる
        </Btn>
      }
    >
      <NoteBox kind="ok">
        <b>このツールは外部と一切通信しません。</b>
        読み込んだ Excel はブラウザーの中だけで処理され、社外はもちろん社内サーバーにも送られません。
        LAN を切断した状態でもすべての機能が動作します。
      </NoteBox>

      <h3 style={{ margin: '18px 0 6px' }}>できること</h3>
      <table className="plain">
        <tbody>
          <tr>
            <th style={{ width: 150 }}>フォルダー一括読み込み</th>
            <td>選んだフォルダーの下にある Excel を、サブフォルダーも含めてまとめて開きます。</td>
          </tr>
          <tr>
            <th>セルのロック</th>
            <td>
              一部のセルだけをロック / 逆に<b>指定した範囲以外を全部ロック</b>できます。
              複数ブック・全シートへ一度に適用できます。
            </td>
          </tr>
          <tr>
            <th>色分け</th>
            <td>ロックしていないセル (入力欄) だけをまとめて塗りつぶせます。</td>
          </tr>
          <tr>
            <th>年度の一括更新</th>
            <td>
              2024→2025、2025→2026 … をすべてのシートで<b>同時に</b>置換します。
              数式・シート名・ファイル名も対象にできます。
            </td>
          </tr>
          <tr>
            <th>手順書</th>
            <td>
              行った操作が自動で手順として記録されます。書き出しておけば、
              翌年は読み込んで<b>「すべて実行」</b>するだけで同じ作業が終わります。
            </td>
          </tr>
        </tbody>
      </table>

      <h3 style={{ margin: '18px 0 6px' }}>基本の流れ</h3>
      <ol style={{ paddingLeft: 20, margin: 0 }}>
        <li>「ファイル」タブ →「フォルダーを開く」</li>
        <li>「ロック」タブ → 入力させたい範囲を選んで「選択範囲以外をロック」</li>
        <li>続けて「シート保護を有効化」（これをしないとロックは効きません）</li>
        <li>「書式・色」タブ →「ロック解除セルを色分け」で入力欄を目立たせる</li>
        <li>「年度更新」タブ → 年度を +1 年ずらす</li>
        <li>「手順書」タブ → JSON と HTML を書き出して保管</li>
        <li>「ファイル」タブ → 保存</li>
      </ol>

      <h3 style={{ margin: '18px 0 6px' }}>注意していただきたいこと</h3>
      <NoteBox kind="warn">
        <ul style={{ margin: 0, paddingLeft: 18 }}>
          <li>
            <b>作業前に必ずバックアップを取ってください。</b>
          </li>
          <li>
            保存すると、ファイルはこのツールが書き直したものになります。
            グラフ・図形・ピボットテーブル・条件付き書式などは、まれに失われることがあります。
            重要なブックはコピーで試してから本番に適用してください。
          </li>
          <li>
            <b>.xlsm / .xltm（マクロ付き）はマクロが失われます。</b>
            これらのファイルは保存の対象から外すことをおすすめします。
          </li>
          <li>
            シート保護のパスワードは<b>手順書には保存されません</b>（安全のため）。
            毎回入力してください。
          </li>
        </ul>
      </NoteBox>

      <h3 style={{ margin: '18px 0 6px' }}>Excel のロックの仕組み</h3>
      <p style={{ margin: 0 }}>
        Excel ではすべてのセルに「ロック」という属性があり、<b>初期状態ではすべて ON</b> です。
        ただしこの属性は「シートの保護」を有効にして初めて意味を持ちます。
        そのため「一部だけ入力させたい」場合は、
        <b>入力させたい範囲のロックを外し → シートを保護する</b>、という 2 段階になります。
        このツールの「選択範囲以外をロック」＋「シート保護を有効化」が、その 2 段階に対応しています。
      </p>
    </Modal>
  );
}
