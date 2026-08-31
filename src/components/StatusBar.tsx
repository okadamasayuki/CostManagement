import { useEffect, useState } from 'react';
import type { SheetView } from '../excel/types';
import { rectCellCount, rectToA1 } from '../excel/cellRef';
import { getBlockedAttempts, onBlockedAttempt } from '../security/networkGuard';
import { setState, useStore } from '../state/store';

export function StatusBar(props: { view: SheetView | null }) {
  const s = useStore();
  const { view } = props;
  const [blocked, setBlocked] = useState(getBlockedAttempts().length);
  useEffect(() => onBlockedAttempt((list) => setBlocked(list.length)), []);

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
      {dirty > 0 && <span className="sb-item">未保存 {dirty} ブック</span>}
      <span className="sb-item">記録 {s.recording ? 'ON' : 'OFF'} / {s.recipe.steps.length} 手順</span>
      <span
        className="sb-item clickable"
        onClick={() => setState({ activeTab: 'security' })}
        title="外部通信の遮断状況を表示"
      >
        {blocked === 0 ? '🔒 外部通信なし' : `⚠️ 通信を ${blocked} 件遮断`}
      </span>
    </div>
  );
}
