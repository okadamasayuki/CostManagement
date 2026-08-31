import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import type { CellView, SheetView } from '../excel/types';
import type { RangeRect } from '../excel/cellRef';
import { colToLetter, normalizeRect } from '../excel/cellRef';
import { axisIndexAt, axisOffset, axisSize, axisTotal } from '../excel/axis';
import { argbToCss, readableTextColor } from '../excel/format';

/**
 * 表計算グリッド。
 *
 * ・4 ペイン構成 (角/列見出し/行見出し/本体) でスクロール同期する。
 * ・可視範囲のセルだけを描画する (仮想化)。数千行あっても軽い。
 * ・ロック状態を斜線 / 黄色の網掛けで重ねて表示できる。
 */

const ROW_HEADER_W = 52;
const COL_HEADER_H = 22;
const OVERSCAN = 4;

export interface GridProps {
  view: SheetView;
  selection: RangeRect | null;
  anchor: { row: number; col: number } | null;
  showLockOverlay: boolean;
  readOnly: boolean;
  onSelect(rect: RangeRect, anchor: { row: number; col: number }): void;
  onCommitEdit(row: number, col: number, text: string): void;
}


export function Grid(props: GridProps) {
  const { view, selection, anchor, showLockOverlay, readOnly } = props;
  const bodyRef = useRef<HTMLDivElement>(null);
  const colHeadInner = useRef<HTMLDivElement>(null);
  const rowHeadInner = useRef<HTMLDivElement>(null);
  const dragging = useRef<{ row: number; col: number } | null>(null);

  const [scroll, setScroll] = useState({ top: 0, left: 0 });
  const [size, setSize] = useState({ w: 800, h: 500 });
  const [editing, setEditing] = useState<{ row: number; col: number; value: string } | null>(null);

  const totalW = useMemo(() => axisTotal(view.cols), [view.cols]);
  const totalH = useMemo(() => axisTotal(view.rows), [view.rows]);
  const colX = (c: number) => axisOffset(view.cols, c);
  const rowY = (r: number) => axisOffset(view.rows, r);
  const colW = (c: number) => axisSize(view.cols, c);
  const rowH = (r: number) => axisSize(view.rows, r);

  // 表示中のシートが変わったらスクロール位置を戻す
  useEffect(() => {
    if (bodyRef.current) {
      bodyRef.current.scrollTop = 0;
      bodyRef.current.scrollLeft = 0;
    }
    setScroll({ top: 0, left: 0 });
    setEditing(null);
  }, [view.name, view.index]);

  useLayoutEffect(() => {
    const el = bodyRef.current;
    if (!el) return;
    const ro = new ResizeObserver(() => {
      setSize({ w: el.clientWidth, h: el.clientHeight });
    });
    ro.observe(el);
    setSize({ w: el.clientWidth, h: el.clientHeight });
    return () => ro.disconnect();
  }, []);

  const rafId = useRef(0);
  const onScroll = useCallback(() => {
    const el = bodyRef.current;
    if (!el) return;
    // 見出しの追従は React を経由せず直接動かして滑らかにする
    if (colHeadInner.current) {
      colHeadInner.current.style.transform = `translateX(${-el.scrollLeft}px)`;
    }
    if (rowHeadInner.current) {
      rowHeadInner.current.style.transform = `translateY(${-el.scrollTop}px)`;
    }
    if (rafId.current) return;
    rafId.current = requestAnimationFrame(() => {
      rafId.current = 0;
      setScroll({ top: el.scrollTop, left: el.scrollLeft });
    });
  }, []);

  const firstRow = Math.max(1, axisIndexAt(view.rows, scroll.top) - OVERSCAN);
  const lastRow = Math.min(view.rowCount, axisIndexAt(view.rows, scroll.top + size.h) + OVERSCAN);
  const firstCol = Math.max(1, axisIndexAt(view.cols, scroll.left) - OVERSCAN);
  const lastCol = Math.min(view.colCount, axisIndexAt(view.cols, scroll.left + size.w) + OVERSCAN);

  const select = useCallback(
    (rect: RangeRect, a: { row: number; col: number }) => props.onSelect(rect, a),
    [props],
  );

  const scrollIntoView = useCallback(
    (row: number, col: number) => {
      const el = bodyRef.current;
      if (!el) return;
      const x = colX(col);
      const w = colW(col);
      const y = rowY(row);
      const h = rowH(row);
      if (x < el.scrollLeft) el.scrollLeft = x;
      else if (x + w > el.scrollLeft + el.clientWidth) el.scrollLeft = x + w - el.clientWidth;
      if (y < el.scrollTop) el.scrollTop = y;
      else if (y + h > el.scrollTop + el.clientHeight) el.scrollTop = y + h - el.clientHeight;
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [view.rows, view.cols],
  );

  // 名前ボックスなど、グリッドの外から選択位置が変わったときに追従する。
  // シートが広いと、選んだセルが画面外のままになってしまうため。
  useEffect(() => {
    if (anchor) scrollIntoView(anchor.row, anchor.col);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [anchor?.row, anchor?.col]);

  // --- マウス操作 ---------------------------------------------------------

  const cellFromEvent = useCallback(
    (e: React.MouseEvent): { row: number; col: number } | null => {
      const el = bodyRef.current;
      if (!el) return null;
      const rect = el.getBoundingClientRect();
      const x = e.clientX - rect.left + el.scrollLeft;
      const y = e.clientY - rect.top + el.scrollTop;
      if (x < 0 || y < 0 || x > totalW || y > totalH) return null;
      return { row: axisIndexAt(view.rows, y), col: axisIndexAt(view.cols, x) };
    },
    [totalH, totalW, view.rows, view.cols],
  );

  const onMouseDown = (e: React.MouseEvent) => {
    if (e.button !== 0) return;
    const c = cellFromEvent(e);
    if (!c) return;
    bodyRef.current?.focus();
    setEditing(null);
    if (e.shiftKey && anchor) {
      select(normalizeRect(anchor, c), anchor);
    } else {
      dragging.current = c;
      select({ top: c.row, left: c.col, bottom: c.row, right: c.col }, c);
    }
  };

  const onMouseMove = (e: React.MouseEvent) => {
    if (!dragging.current || e.buttons !== 1) return;
    const c = cellFromEvent(e);
    if (!c) return;
    select(normalizeRect(dragging.current, c), dragging.current);
  };

  const onMouseUp = () => {
    dragging.current = null;
  };

  const onDoubleClick = (e: React.MouseEvent) => {
    if (readOnly) return;
    const c = cellFromEvent(e);
    if (!c) return;
    startEdit(c.row, c.col, view.cells.get(`${c.row}:${c.col}`)?.raw ?? '');
  };

  const startEdit = (row: number, col: number, initial: string) => {
    setEditing({ row, col, value: initial });
  };

  const commitEdit = (move: 'down' | 'right' | null) => {
    if (!editing) return;
    props.onCommitEdit(editing.row, editing.col, editing.value);
    const { row, col } = editing;
    setEditing(null);
    bodyRef.current?.focus();
    if (move === 'down') moveTo(Math.min(row + 1, view.rowCount), col);
    if (move === 'right') moveTo(row, Math.min(col + 1, view.colCount));
  };

  const moveTo = useCallback(
    (row: number, col: number) => {
      const r = Math.min(Math.max(1, row), view.rowCount);
      const c = Math.min(Math.max(1, col), view.colCount);
      select({ top: r, left: c, bottom: r, right: c }, { row: r, col: c });
      scrollIntoView(r, c);
    },
    [select, scrollIntoView, view.colCount, view.rowCount],
  );

  // --- キーボード操作 -----------------------------------------------------

  const onKeyDown = (e: React.KeyboardEvent) => {
    if (editing) return;
    const a = anchor ?? { row: 1, col: 1 };
    const sel = selection ?? { top: a.row, left: a.col, bottom: a.row, right: a.col };
    const extend = e.shiftKey;

    const applyMove = (dr: number, dc: number) => {
      e.preventDefault();
      if (extend) {
        const cur = { row: sel.bottom === a.row ? sel.top : sel.bottom, col: sel.right === a.col ? sel.left : sel.right };
        const target = {
          row: Math.min(Math.max(1, cur.row + dr), view.rowCount),
          col: Math.min(Math.max(1, cur.col + dc), view.colCount),
        };
        select(normalizeRect(a, target), a);
        scrollIntoView(target.row, target.col);
      } else {
        moveTo(a.row + dr, a.col + dc);
      }
    };

    switch (e.key) {
      case 'ArrowDown': return applyMove(1, 0);
      case 'ArrowUp': return applyMove(-1, 0);
      case 'ArrowLeft': return applyMove(0, -1);
      case 'ArrowRight': return applyMove(0, 1);
      case 'PageDown': return applyMove(20, 0);
      case 'PageUp': return applyMove(-20, 0);
      case 'Tab':
        e.preventDefault();
        return moveTo(a.row, a.col + (e.shiftKey ? -1 : 1));
      case 'Enter':
        e.preventDefault();
        if (readOnly) return moveTo(a.row + 1, a.col);
        return startEdit(a.row, a.col, view.cells.get(`${a.row}:${a.col}`)?.raw ?? '');
      case 'F2':
        e.preventDefault();
        if (readOnly) return;
        return startEdit(a.row, a.col, view.cells.get(`${a.row}:${a.col}`)?.raw ?? '');
      case 'Home':
        e.preventDefault();
        return moveTo(e.ctrlKey ? 1 : a.row, 1);
      case 'Escape':
        return;
      case 'Delete':
      case 'Backspace':
        if (readOnly) return;
        e.preventDefault();
        props.onCommitEdit(a.row, a.col, '');
        return;
      default:
        break;
    }
    // 印字可能な文字が押されたら編集を開始する (Excel と同じ挙動)
    if (!readOnly && e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
      e.preventDefault();
      startEdit(a.row, a.col, e.key);
    }
  };

  // --- 描画 ---------------------------------------------------------------

  const cells: React.ReactNode[] = [];
  for (let r = firstRow; r <= lastRow; r++) {
    const y = rowY(r);
    const h = rowH(r);
    for (let c = firstCol; c <= lastCol; c++) {
      const key = `${r}:${c}`;
      const cv = view.cells.get(key);
      if (cv?.mergedHidden) continue;
      const x = colX(c);
      let w = colW(c);
      let hh = h;
      if (cv?.mergeSpan) {
        w = colX(Math.min(c + cv.mergeSpan.cols, view.colCount + 1)) - x;
        hh = rowY(Math.min(r + cv.mergeSpan.rows, view.rowCount + 1)) - y;
      }
      cells.push(
        <CellBox
          key={key}
          cv={cv}
          x={x}
          y={y}
          w={w}
          h={hh}
          spillW={spillWidth(view, cv, r, c, w)}
          // シートの内容がある範囲の外は、見栄えのために描いているだけなので
          // ロック状態の網掛けは出さない
          showLockOverlay={
            showLockOverlay && r <= view.contentBottom && c <= view.contentRight
          }
        />,
      );
    }
  }

  const colHeads: React.ReactNode[] = [];
  for (let c = firstCol; c <= lastCol; c++) {
    const sel = selection ? c >= selection.left && c <= selection.right : false;
    colHeads.push(
      <div
        key={c}
        className={`colhead${sel ? ' sel' : ''}`}
        style={{ left: colX(c), width: colW(c), height: COL_HEADER_H, top: 0 }}
        onMouseDown={() => {
          const rect = { top: 1, bottom: view.rowCount, left: c, right: c };
          select(rect, { row: 1, col: c });
        }}
      >
        {colToLetter(c)}
      </div>,
    );
  }

  const rowHeads: React.ReactNode[] = [];
  for (let r = firstRow; r <= lastRow; r++) {
    const sel = selection ? r >= selection.top && r <= selection.bottom : false;
    rowHeads.push(
      <div
        key={r}
        className={`rowhead${sel ? ' sel' : ''}`}
        style={{ top: rowY(r), height: rowH(r), width: ROW_HEADER_W, left: 0 }}
        onMouseDown={() => {
          const rect = { top: r, bottom: r, left: 1, right: view.colCount };
          select(rect, { row: r, col: 1 });
        }}
      >
        {r}
      </div>,
    );
  }

  const selBox = selection
    ? {
        left: colX(selection.left),
        top: rowY(selection.top),
        width: colX(Math.min(selection.right + 1, view.colCount + 1)) - colX(selection.left),
        height: rowY(Math.min(selection.bottom + 1, view.rowCount + 1)) - rowY(selection.top),
      }
    : null;

  const editBox = editing
    ? {
        left: colX(editing.col),
        top: rowY(editing.row),
        width: colW(editing.col),
        height: rowH(editing.row),
      }
    : null;

  return (
    <div className="gridwrap">
      <div
        className="grid-corner"
        onMouseDown={() =>
          select({ top: 1, left: 1, bottom: view.rowCount, right: view.colCount }, { row: 1, col: 1 })
        }
        title="シート全体を選択"
      />
      <div className="grid-colheads">
        <div ref={colHeadInner} style={{ position: 'relative', width: totalW, height: COL_HEADER_H }}>
          {colHeads}
        </div>
      </div>
      <div className="grid-rowheads">
        <div ref={rowHeadInner} style={{ position: 'relative', height: totalH, width: ROW_HEADER_W }}>
          {rowHeads}
        </div>
      </div>
      <div
        className="grid-body"
        ref={bodyRef}
        tabIndex={0}
        onScroll={onScroll}
        onMouseDown={onMouseDown}
        onMouseMove={onMouseMove}
        onMouseUp={onMouseUp}
        onDoubleClick={onDoubleClick}
        onKeyDown={onKeyDown}
      >
        <div className="grid-canvas" style={{ width: totalW, height: totalH }}>
          {cells}
          {selBox && <div className="selection-rect" style={selBox} />}
          {editing && editBox && (
            <input
              className="cell-editor"
              style={editBox}
              autoFocus
              value={editing.value}
              onChange={(e) => setEditing({ ...editing, value: e.target.value })}
              onBlur={() => commitEdit(null)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  e.preventDefault();
                  commitEdit('down');
                } else if (e.key === 'Tab') {
                  e.preventDefault();
                  commitEdit('right');
                } else if (e.key === 'Escape') {
                  e.preventDefault();
                  setEditing(null);
                  bodyRef.current?.focus();
                }
                e.stopPropagation();
              }}
            />
          )}
        </div>
      </div>
    </div>
  );
}

/**
 * Excel と同じく、右隣が空のセルには文字をはみ出して表示する。
 * これがないと見出し行などが途中で切れて読めなくなる。
 * 戻り値ははみ出し後の表示幅 (はみ出さない場合は元の幅)。
 */
function spillWidth(
  view: SheetView,
  cv: CellView | undefined,
  row: number,
  col: number,
  width: number,
): number {
  if (!cv || !cv.text || cv.mergeSpan || cv.mergedHidden) return width;
  // 右寄せ・中央寄せ・数値は Excel でもはみ出さない (### になる)
  if (cv.align === 'right' || cv.align === 'center') return width;
  let w = width;
  for (let c = col + 1; c <= view.colCount && c <= col + 20; c++) {
    const n = view.cells.get(`${row}:${c}`);
    if (n && (n.text || n.mergedHidden)) break;
    w += axisSize(view.cols, c);
  }
  return w;
}

interface CellBoxProps {
  cv: CellView | undefined;
  x: number;
  y: number;
  w: number;
  h: number;
  spillW: number;
  showLockOverlay: boolean;
}

function CellBox({ cv, x, y, w, h, spillW, showLockOverlay }: CellBoxProps) {
  const locked = cv?.locked ?? true;
  const bg = argbToCss(cv?.fillArgb);
  const classes = ['cell'];
  if (cv?.align === 'right') classes.push('num');
  if (cv?.align === 'center') classes.push('ctr');
  if (cv?.kind === 'formula') classes.push('formula-cell');
  if (cv?.kind === 'error') classes.push('err');
  if (showLockOverlay) classes.push(locked ? 'ov-locked' : 'ov-unlocked');

  const color = cv?.fontArgb
    ? argbToCss(cv.fontArgb)
    : bg
      ? readableTextColor(cv?.fillArgb)
      : undefined;

  const spills = spillW > w;
  if (spills) classes.push('spill');

  return (
    <div
      className={classes.join(' ')}
      style={{
        left: x,
        top: y,
        width: w,
        height: h,
        background: bg,
        color,
        fontWeight: cv?.bold ? 700 : undefined,
        fontStyle: cv?.italic ? 'italic' : undefined,
      }}
      title={cv?.text && cv.text.length > 12 ? cv.text : undefined}
    >
      <span className="cell-text" style={spills ? { maxWidth: spillW - 7 } : undefined}>
        {cv?.text}
      </span>
    </div>
  );
}
