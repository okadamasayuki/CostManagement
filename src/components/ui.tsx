import type { ReactNode } from 'react';
import type { OpScope } from '../excel/types';
import { listFolders } from '../excel/folders';
import { describeScope, describeScopeShort } from '../recipe/describe';
import { useStore } from '../state/store';
import type { RangeSpec } from '../recipe/types';
import { argbToCss, cssToArgb } from '../excel/format';

/** リボン内の機能グループ */
export function RGroup({ title, children }: { title: string; children: ReactNode }) {
  return (
    <div className="rgroup">
      <div className="rgroup-body">{children}</div>
      <div className="rgroup-title">{title}</div>
    </div>
  );
}

export function RCol({ children }: { children: ReactNode }) {
  return <div className="rcol">{children}</div>;
}

export function BigButton(props: {
  icon: string;
  label: ReactNode;
  onClick(): void;
  disabled?: boolean;
  primary?: boolean;
  title?: string;
}) {
  return (
    <button
      className={`rbtn-lg${props.primary ? ' primary' : ''}`}
      onClick={props.onClick}
      disabled={props.disabled}
      title={props.title}
      type="button"
    >
      <span className="ico" aria-hidden>
        {props.icon}
      </span>
      <span>{props.label}</span>
    </button>
  );
}

export function Btn(props: {
  children: ReactNode;
  onClick(): void;
  disabled?: boolean;
  kind?: 'normal' | 'danger' | 'accent';
  title?: string;
}) {
  const cls = props.kind === 'danger' ? 'rbtn danger' : props.kind === 'accent' ? 'rbtn accent' : 'rbtn';
  return (
    <button className={cls} onClick={props.onClick} disabled={props.disabled} title={props.title} type="button">
      {props.children}
    </button>
  );
}

export function Field({ label, children }: { label?: string; children: ReactNode }) {
  return (
    <div className="field">
      {label && <label>{label}</label>}
      {children}
    </div>
  );
}

export function Check(props: {
  label: string;
  checked: boolean;
  onChange(v: boolean): void;
  title?: string;
  disabled?: boolean;
}) {
  return (
    <label className="check" title={props.title}>
      <input
        type="checkbox"
        checked={props.checked}
        disabled={props.disabled}
        onChange={(e) => props.onChange(e.target.checked)}
      />
      {props.label}
    </label>
  );
}

/** Excel の標準的な塗りつぶし色 + 「塗りなし」 */
export const PALETTE: Array<{ argb: string; name: string }> = [
  { argb: 'FFFFFF00', name: '黄' },
  { argb: 'FFFFF2CC', name: '薄い黄' },
  { argb: 'FFFFC000', name: 'オレンジ' },
  { argb: 'FFFCE4D6', name: '薄いオレンジ' },
  { argb: 'FFFF0000', name: '赤' },
  { argb: 'FFFFC7CE', name: '薄い赤' },
  { argb: 'FF92D050', name: '黄緑' },
  { argb: 'FFC6EFCE', name: '薄い緑' },
  { argb: 'FF00B0F0', name: '水色' },
  { argb: 'FFDDEBF7', name: '薄い青' },
  { argb: 'FFB4C6E7', name: '青灰' },
  { argb: 'FFD9D9D9', name: '灰色' },
  { argb: 'FFA6A6A6', name: '濃い灰' },
  { argb: 'FFFFFFFF', name: '白' },
];

export function ColorSwatches(props: {
  value: string | null;
  onChange(argb: string | null): void;
  allowNone?: boolean;
}) {
  return (
    <div className="swatches">
      {PALETTE.map((p) => (
        <button
          key={p.argb}
          type="button"
          title={p.name}
          className={`swatch${props.value === p.argb ? ' selected' : ''}`}
          style={{ background: argbToCss(p.argb) }}
          onClick={() => props.onChange(p.argb)}
        />
      ))}
      <input
        type="color"
        className="swatch"
        title="その他の色を選ぶ"
        value={argbToCss(props.value ?? 'FFFFFFFF') ?? '#ffffff'}
        onChange={(e) => props.onChange(cssToArgb(e.target.value))}
        style={{ padding: 0, cursor: 'pointer' }}
      />
      {props.allowNone !== false && (
        <button
          type="button"
          className={`clear-fill${props.value === null ? ' selected' : ''}`}
          title="選んだセルの塗りつぶしを消して、色なしに戻します"
          onClick={() => props.onChange(null)}
        >
          🚫 色を消す
        </button>
      )}
    </div>
  );
}

/** 操作の適用範囲 (ブック / シート) の指定 */
export function ScopeSelector(props: {
  scope: OpScope;
  onChange(s: OpScope): void;
  /** 横並びにする (リボン下の適用先バー用) */
  inline?: boolean;
}) {
  const { scope, onChange } = props;
  const store = useStore();
  // 読み込んだファイルの相対パスから、実際に存在するフォルダーを出す
  const folders = listFolders(store.books);
  const targetCount =
    scope.books === 'folder'
      ? (folders.find((f) => f.path === (scope.bookFolder ?? ''))?.count ?? 0)
      : 0;

  const Wrap = props.inline
    ? ({ children }: { children: ReactNode }) => <div className="scope-inline">{children}</div>
    : RCol;

  return (
    <Wrap>
      <Field label="ブック">
        <select
          data-testid="scope-books"
          value={scope.books}
          onChange={(e) => onChange({ ...scope, books: e.target.value as OpScope['books'] })}
        >
          <option value="current">選択中のブックのみ</option>
          <option value="all">読み込んだ全ブック</option>
          <option value="selected">一覧で選んだブック ({store.selectedBookIds.length})</option>
          <option value="folder">フォルダーを指定…</option>
          <option value="glob">ファイル名で指定…</option>
        </select>
      </Field>
      {scope.books === 'selected' && (
        <div style={{ fontSize: 10.5, color: 'var(--text-dim)', paddingLeft: 2, width: 178 }}>
          左の一覧を <b>Shift</b> / <b>Ctrl</b> クリックで選べます (現在{' '}
          <b>{store.selectedBookIds.length}</b> ブック)
        </div>
      )}
      {scope.books === 'folder' && (
        <>
          <Field label="">
            <select
              data-testid="scope-folder"
              style={{ width: 178 }}
              value={scope.bookFolder ?? ''}
              onChange={(e) => onChange({ ...scope, bookFolder: e.target.value })}
            >
              {folders.map((f) => (
                <option key={f.path} value={f.path}>
                  {'　'.repeat(f.depth)}
                  {f.path ? '📁 ' : ''}
                  {f.label} ({f.count})
                </option>
              ))}
            </select>
          </Field>
          <div style={{ fontSize: 10.5, color: 'var(--text-dim)', paddingLeft: 2 }}>
            配下のサブフォルダーも含めて <b>{targetCount}</b> ブックが対象
          </div>
        </>
      )}
      {scope.books === 'glob' && (
        <Field label="">
          <input
            type="text"
            placeholder="例: *原価*.xlsx, 予算??.xlsx"
            style={{ width: 178 }}
            value={scope.bookGlob ?? ''}
            onChange={(e) => onChange({ ...scope, bookGlob: e.target.value })}
          />
        </Field>
      )}
      <Field label="シート">
        <select
          data-testid="scope-sheets"
          value={scope.sheets}
          onChange={(e) => onChange({ ...scope, sheets: e.target.value as OpScope['sheets'] })}
        >
          <option value="current">表示中のシートのみ</option>
          <option value="all">全シート</option>
          <option value="glob">シート名で指定…</option>
        </select>
      </Field>
      {scope.sheets === 'glob' && (
        <Field label="">
          <input
            type="text"
            placeholder="例: *年度*, 明細*"
            style={{ width: 178 }}
            value={scope.sheetGlob ?? ''}
            onChange={(e) => onChange({ ...scope, sheetGlob: e.target.value })}
          />
        </Field>
      )}
    </Wrap>
  );
}

/**
 * 今どこに適用されるのかを示す表示。
 * 「適用先」の設定はリボン内の離れた場所にあるため、
 * 実行ボタンのそばにも出して取り違えを防ぐ。
 */
export function ScopeBadge(props: { note?: string }) {
  const s = useStore();
  return (
    <div className="scope-badge" title={describeScope(s.scope)}>
      <span>🎯 適用先</span>
      <b>{describeScopeShort(s.scope)}</b>
      {props.note && <span className="note">{props.note}</span>}
    </div>
  );
}

export type RangeMode = 'selection' | 'used' | 'a1';

/** 操作対象のセル範囲の指定 */
export function RangeSelector(props: {
  mode: RangeMode;
  a1: string;
  selectionA1: string | null;
  onModeChange(m: RangeMode): void;
  onA1Change(v: string): void;
}) {
  return (
    <RCol>
      <Field label="範囲">
        <select
          data-testid="range-mode"
          value={props.mode}
          onChange={(e) => props.onModeChange(e.target.value as RangeMode)}
        >
          <option value="selection">画面で選択した範囲</option>
          <option value="used">データが入っている範囲全体</option>
          <option value="a1">アドレスで指定…</option>
        </select>
      </Field>
      {props.mode === 'a1' ? (
        <Field label="">
          <input
            type="text"
            placeholder="例: B4:F30, A:C, 1:5"
            style={{ width: 178 }}
            value={props.a1}
            onChange={(e) => props.onA1Change(e.target.value)}
          />
        </Field>
      ) : (
        <div style={{ fontSize: 10.5, color: 'var(--text-dim)', paddingLeft: 2 }}>
          {props.mode === 'selection'
            ? props.selectionA1
              ? `現在の選択: ${props.selectionA1}`
              : 'セルを選択してください'
            : 'シートごとに自動判定します'}
        </div>
      )}
    </RCol>
  );
}

/** RangeMode を実際の RangeSpec に変換する */
export function toRangeSpec(mode: RangeMode, a1: string, selectionA1: string | null): RangeSpec | null {
  if (mode === 'used') return { kind: 'used' };
  if (mode === 'a1') return a1.trim() ? { kind: 'a1', a1: a1.trim() } : null;
  return selectionA1 ? { kind: 'a1', a1: selectionA1 } : null;
}

export function Modal(props: {
  title: string;
  children: ReactNode;
  footer: ReactNode;
  onClose(): void;
  wide?: boolean;
}) {
  return (
    <div className="modal-backdrop" onMouseDown={(e) => e.target === e.currentTarget && props.onClose()}>
      <div className="modal" style={props.wide ? { maxWidth: 860 } : undefined}>
        <div className="modal-head">{props.title}</div>
        <div className="modal-body">{props.children}</div>
        <div className="modal-foot">{props.footer}</div>
      </div>
    </div>
  );
}

export function NoteBox({
  kind = 'info',
  children,
  testId,
}: {
  kind?: 'info' | 'warn' | 'ok' | 'err';
  children: ReactNode;
  testId?: string;
}) {
  return (
    <div className={`note-box${kind === 'info' ? '' : ` ${kind}`}`} data-testid={testId}>
      {children}
    </div>
  );
}
