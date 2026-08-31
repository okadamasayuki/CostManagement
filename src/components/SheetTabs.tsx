import type ExcelJS from 'exceljs';
import { getSheetProtection } from '../excel/exceljsCompat';
import { currentBook, selectSheet, useStore } from '../state/store';

export function SheetTabs() {
  const s = useStore();
  const book = currentBook();
  if (!book || book.loadError) return <div className="sheettabs" />;

  const sheets: ExcelJS.Worksheet[] = [];
  book.wb.eachSheet((ws) => sheets.push(ws));

  return (
    <div className="sheettabs">
      {sheets.map((ws) => {
        const protectedSheet = Boolean(getSheetProtection(ws));
        return (
          <div
            key={ws.id}
            className={`sheettab${ws.name === s.currentSheetName ? ' active' : ''}`}
            onClick={() => selectSheet(ws.name)}
            title={
              protectedSheet
                ? `${ws.name} — シート保護が有効です`
                : `${ws.name} — シート保護は無効 (セルのロックは効きません)`
            }
          >
            {protectedSheet && <span className="prot">🛡️</span>}
            <span>{ws.name}</span>
            {ws.state !== 'visible' && <span className="prot" title="非表示シート">👁️‍🗨️</span>}
          </div>
        );
      })}
    </div>
  );
}
