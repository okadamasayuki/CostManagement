// 年度更新後のファイルが狙いどおりか検査する
const ExcelJS = require('exceljs');
const { readdirSync } = require('node:fs');
const { join } = require('node:path');
const ROOT = process.env.REPORT_DIR ?? '/root/報告共有フォルダー/2025年度報告';
(async () => {
  let n = 0, ok = 0; const ng = [];
  for (const 支店 of readdirSync(ROOT)) {
    for (const f of readdirSync(join(ROOT, 支店))) {
      n++;
      const p = [];
      if (!f.startsWith('2026年度_')) p.push(`ファイル名が ${f}`);
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile(join(ROOT, 支店, f));
      const names = wb.worksheets.map((w) => w.name);
      if (names[0] !== '2026年度' || names[1] !== '2025年度実績') p.push(`シート名 ${names.join(',')}`);
      const ws = wb.getWorksheet('2026年度');
      if (!ws) { ng.push(`${f}: 2026年度シートが無い`); continue; }
      if (String(ws.getCell('A1').value) !== '2026年度 数量報告書') p.push(`A1=${ws.getCell('A1').value}`);
      const head = [2,3,4].map((c) => ws.getRow(5).getCell(c).value).join('|');
      if (head !== '2024年度実績|2025年度実績|2026年度計画') p.push(`見出し=${head}`);
      // 数量 (年に見える 4 桁) が守られているか
      if (ws.getCell('B13').value !== 2031) p.push(`B13=${ws.getCell('B13').value} (2031 のはず)`);
      if (ws.getCell('C13').value !== 2018) p.push(`C13=${ws.getCell('C13').value} (2018 のはず)`);
      // 黄色の入力欄とロック・保護が保たれているか
      if (!ws.sheetProtection) p.push('シート保護が消えた');
      if (ws.getCell('D6').protection?.locked !== false) p.push('D6 が入力できない');
      if (ws.getCell('D6').fill?.fgColor?.argb !== 'FFFFFF00') p.push('D6 の黄色が消えた');
      if (ws.getCell('B6').protection?.locked === false) p.push('B6 のロックが外れた');
      if (ws.getCell('E6').value?.formula !== 'IF(D6="","",D6-C6)') p.push(`E6 の数式=${JSON.stringify(ws.getCell('E6').value)}`);
      if (!String(ws.getCell('A3').value).includes('2026年4月1日')) p.push(`A3=${ws.getCell('A3').value}`);
      if (p.length) ng.push(`${支店}/${f}: ${p.join(' / ')}`); else ok++;
    }
  }
  console.log(`検査: ${n} ファイル / 合格 ${ok} / 不合格 ${ng.length}`);
  ng.slice(0, 5).forEach((m) => console.log('  ✗', m));
})();
