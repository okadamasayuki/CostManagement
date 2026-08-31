import { dismissToast, useStore } from '../state/store';

export function Toasts() {
  const s = useStore();
  if (!s.toasts.length) return null;
  // 画面を覆ってしまわないよう、新しいものから 4 件だけ出す
  const visible = s.toasts.slice(-4);
  return (
    <div className="toasts">
      {visible.map((t) => (
        <div key={t.id} className={`toast ${t.kind}`} onClick={() => dismissToast(t.id)}>
          <div className="t-msg">{t.message}</div>
          {t.detail && <div className="t-detail">{t.detail}</div>}
        </div>
      ))}
    </div>
  );
}

export function BusyOverlay() {
  const s = useStore();
  if (!s.busy.active) return null;
  const pct = s.busy.total > 0 ? Math.round((s.busy.done / s.busy.total) * 100) : 0;
  return (
    <div className="busy-overlay">
      <div className="busy-card">
        <div className="label">処理中…</div>
        <div className={`progress${s.busy.total > 0 ? '' : ' indeterminate'}`}>
          <div style={{ width: s.busy.total > 0 ? `${pct}%` : undefined }} />
        </div>
        <div className="sub">{s.busy.label}</div>
      </div>
    </div>
  );
}
