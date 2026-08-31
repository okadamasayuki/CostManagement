import { useEffect, useState } from 'react';
import type { SheetView } from '../excel/types';
import { rectCellCount, rectToA1 } from '../excel/cellRef';
import { getSendAttempts, onBlockedAttempt } from '../security/networkGuard';
import { describeScope, describeScopeShort } from '../recipe/describe';
import { useStore } from '../state/store';

export function StatusBar(props: { view: SheetView | null }) {
  const s = useStore();
  const { view } = props;
  const [blocked, setBlocked] = useState(getSendAttempts().length);
  useEffect(
    () => onBlockedAttempt((list) => setBlocked(list.filter((b) => b.kind === 'send').length)),
    [],
  );

  const dirty = s.books.filter((b) => b.dirty).length;

  return (
    <div className="statusbar">
      <span className="sb-item">
        {s.books.length ? `${s.books.length} ブック` : 'ファイル未読み込み'}
      </span>
      {view && (
        <>
          <span className="sb-item">シート: {view.name}</span>
          <span className="sb-item" title="このシートのロック済み / ロック解除セル数">
            🔒 {view.lockedCount.toLocaleString()} / 🔓 {view.unlockedCount.toLocaleString()}
          </span>
          <span className="sb-item">
            {view.isProtected ? `🛡️ 保護あり${view.hasPassword ? '(PW)' : ''}` : '⚠️ 保護なし'}
          </span>
        </>
      )}
      {s.selection && (
        <span className="sb-item">
          選択: {rectToA1(s.selection)} ({rectCellCount(s.selection).toLocaleString()} セル)
        </span>
      )}
      <span className="spacer" />
      {s.books.length > 0 && (
        <span className="sb-item" title={`操作の適用先: ${describeScope(s.scope)}`}>
          🎯 適用先: {describeScopeShort(s.scope)}
        </span>
      )}
      {dirty > 0 && <span className="sb-item">未保存 {dirty} ブック</span>}
      <span className="sb-item">記録 {s.recording ? 'ON' : 'OFF'} / {s.recipe.steps.length} 手順</span>
      <span
        className="sb-item"
        title={
          blocked === 0
            ? 'このツールは外部と通信しません'
            : `外部へデータを送ろうとした試みを ${blocked} 件遮断しました`
        }
      >
        {blocked === 0 ? '🔒 外部通信なし' : `⚠️ 送信を ${blocked} 件遮断`}
      </span>
    </div>
  );
}
