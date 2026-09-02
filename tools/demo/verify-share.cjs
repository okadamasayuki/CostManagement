// 収録で実際に共有フォルダーが書き換わったかを確かめる
const ExcelJS = require('exceljs');
const { readdirSync, statSync } = require('node:fs');
const { join } = require('node:path');
const ROOT = process.env.SHARE_DIR ?? '/root/共有フォルダー/原価管理/2025年度予算';
(async () => {
  let n = 0, ok = 0, ng = [];
  for (const 支店 of readdirSync(ROOT)) {
    for (const f of readdirSync(join(ROOT, 支店))) {
      n++;
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile(join(ROOT, 支店, f));
      const ws = wb.getWorksheet('2025年度予算');
      const p = [];
      if (!ws.sheetProtection) p.push('保護なし');
      for (const ref of ['C6','C17','E6','B20']) if (ws.getCell(ref).protection?.locked !== false) p.push(`${ref}がロック`);
      for (const ref of ['A1','A6','B6','D6','A18']) if (ws.getCell(ref).protection?.locked === false) p.push(`${ref}が解除`);
      if (ws.getCell('C6').fill?.fgColor?.argb !== 'FFFFFF00') p.push('黄色が消えた');
      if (ws.getCell('D6').value?.formula !== 'IF(C6="","",C6-B6)') p.push('数式が壊れた');
      if (p.length) ng.push(`${支店}/${f}: ${p.join(',')}`); else ok++;
    }
  }
  console.log(`検査: ${n} ファイル / 合格 ${ok} / 不合格 ${ng.length}`);
  ng.slice(0, 5).forEach((m) => console.log('  ✗', m));
})();
