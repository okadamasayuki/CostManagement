import { useEffect, useState } from 'react';
import type { SheetView } from '../excel/types';
import { addrToA1, parseA1Range, rectToA1 } from '../excel/cellRef';
import { setState, useStore } from '../state/store';

export function FormulaBar(props: {
  view: SheetView | null;
  onCommit(row: number, col: number, text: string): void;
  readOnly: boolean;
}) {
  const s = useStore();
  const { view } = props;
  const anchor = s.anchor;
  const cv = view && anchor ? view.cells.get(`${anchor.row}:${anchor.col}`) : undefined;

  const [draft, setDraft] = useState('');
  const [nameBox, setNameBox] = useState('');
  const [editing, setEditing] = useState(false);

  useEffect(() => {
    setDraft(cv?.raw ?? '');
    setEditing(false);
  }, [cv?.raw, anchor?.row, anchor?.col, view?.name]);

  useEffect(() => {
    setNameBox(s.selection ? rectToA1(s.selection) : anchor ? addrToA1(anchor) : '');
  }, [s.selection, anchor]);

  const locked = cv?.locked ?? true;

  return (
    <div className="formulabar">
      <div className="namebox">
        <input
          value={nameBox}
          onChange={(e) => setNameBox(e.target.value)}
          placeholder="A1"
          title="セル範囲を入力して Enter で移動・選択"
          onKeyDown={(e) => {
            if (e.key !== 'Enter' || !view) return;
            const rect = parseA1Range(nameBox, view.rowCount, view.colCount);
            if (!rect) return;
            setState({
              selection: rect,
              anchor: { row: rect.top, col: rect.left },
            });
          }}
        />
      </div>
      <div className="fx" title="数式バー">
        fx
      </div>
      <input
        className="formula-input"
        value={draft}
        placeholder={anchor ? '' : 'セルを選択してください'}
        disabled={!anchor || props.readOnly}
        onChange={(e) => {
          setDraft(e.target.value);
          setEditing(true);
        }}
        onKeyDown={(e) => {
          if (e.key === 'Enter' && anchor) {
            props.onCommit(anchor.row, anchor.col, draft);
            setEditing(false);
          } else if (e.key === 'Escape') {
            setDraft(cv?.raw ?? '');
            setEditing(false);
          }
        }}
        onBlur={() => {
          if (editing && anchor) props.onCommit(anchor.row, anchor.col, draft);
          setEditing(false);
        }}
      />
      {anchor && (
        <div className={`lock-chip ${locked ? 'locked' : 'unlocked'}`} title="このセルのロック状態">
          {locked ? '🔒 ロック済み' : '🔓 入力可能'}
        </div>
      )}
    </div>
  );
}
