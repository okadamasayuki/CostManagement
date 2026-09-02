/**
 * 2 本目の動画用データ。
 * 「各支店から数量を報告してもらう様式」を、前年の作業が終わった状態で作る。
 *   ・2023年度 / 2024年度 の実績 … 数値が入っていてロック済み
 *   ・2025年度 の計画            … 空欄・黄色・入力できる
 *   ・シート保護あり
 */
const ExcelJS = require('exceljs');
const { mkdirSync, rmSync } = require('node:fs');
const { join } = require('node:path');

const BASE = '/root/報告共有フォルダー';
const ROOT = join(BASE, '2025年度報告');
rmSync(BASE, { recursive: true, force: true });

const YELLOW = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
const HEAD   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
const GRAY   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
const T = { style: 'thin', color: { argb: 'FFB0B0B0' } };
const B = { top: T, left: T, bottom: T, right: T };
const OPEN = { locked: false };

const 支店 = [
  ['東京支店', ['製造部','営業部','管理部','技術部','購買部','物流部','品質保証部','開発部']],
  ['大阪支店', ['製造部','営業部','管理部','技術部','購買部','物流部','品質保証部','開発部']],
  ['名古屋支店', ['製造部','営業部','管理部','技術部','購買部','物流部','品質保証部']],
  ['福岡支店', ['製造部','営業部','管理部','技術部','購買部','物流部','品質保証部']],
];

// 数量。あえて 4 桁の少量品目も混ぜてある (年と紛らわしい数量の例)
const 品目 = [
  ['鋼材 SS400',       184_000], ['アルミ板 A5052',   96_500],
  ['樹脂ペレット',      312_000], ['ベアリング 6204',   48_200],
  ['モーター 750W',      6_400], ['配線ハーネス',      27_800],
  ['基板 ASSY',         15_600], ['特注シャフト',       2_030],
  ['防振ゴム',          73_400], ['塗料 (下塗り)',     11_250],
  ['梱包材',           128_000], ['ラベル',           205_000],
];

function budgetSheet(wb, 支店名, 部門, seed) {
  const ws = wb.addWorksheet('2025年度');
  [22, 15, 15, 15, 13, 11, 26].forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  ws.getCell('A1').value = '2025年度 数量報告書';
  ws.getCell('A1').font = { bold: true, size: 15 };
  ws.getCell('A2').value = `${支店名}　${部門}`;
  ws.getCell('A2').font = { size: 12 };
  ws.getCell('A3').value = '作成日: 2025年4月1日　／　前回報告: 2024年度　／　単位: 個';
  ws.getCell('A3').font = { size: 10, color: { argb: 'FF666666' } };

  ['品目', '2023年度実績', '2024年度実績', '2025年度計画', '増減', '増減率', '備考'].forEach((h, i) => {
    const c = ws.getRow(5).getCell(i + 1);
    c.value = h; c.font = { bold: true }; c.fill = HEAD; c.border = B;
    c.alignment = { horizontal: 'center' };
  });

  品目.forEach(([name, base], i) => {
    const r = 6 + i;
    const small = base < 5000;   // 少量品目は年と紛らわしい 4 桁になる
    const j1 = small ? 1 : 1 + ((seed * 3 + i * 7) % 9 - 4) / 100;
    const j2 = small ? 1 : 1 + ((seed * 5 + i * 11) % 9 - 4) / 100;
    ws.getCell(`A${r}`).value = name;                     ws.getCell(`A${r}`).border = B;
    // 過去の実績 = 変更されたくない (ロックのまま)
    ws.getCell(`B${r}`).value = small ? 2031 : Math.round(base * j1);    ws.getCell(`B${r}`).numFmt = '#,##0';
    ws.getCell(`B${r}`).border = B; ws.getCell(`B${r}`).fill = GRAY;
    ws.getCell(`C${r}`).value = small ? 2018 : Math.round(base * j2);    ws.getCell(`C${r}`).numFmt = '#,##0';
    ws.getCell(`C${r}`).border = B; ws.getCell(`C${r}`).fill = GRAY;
    // ▼ 支店に入力してもらう欄 (空欄・黄色・入力できる)
    ws.getCell(`D${r}`).fill = YELLOW; ws.getCell(`D${r}`).numFmt = '#,##0';
    ws.getCell(`D${r}`).border = B;    ws.getCell(`D${r}`).protection = OPEN;
    ws.getCell(`E${r}`).value = { formula: `IF(D${r}="","",D${r}-C${r})` };
    ws.getCell(`E${r}`).numFmt = '#,##0;[Red]-#,##0'; ws.getCell(`E${r}`).border = B;
    ws.getCell(`F${r}`).value = { formula: `IF(D${r}="","",D${r}/C${r}-1)` };
    ws.getCell(`F${r}`).numFmt = '0.0%'; ws.getCell(`F${r}`).border = B;
    ws.getCell(`G${r}`).fill = YELLOW; ws.getCell(`G${r}`).border = B;
    ws.getCell(`G${r}`).protection = OPEN;
  });

  const t = 6 + 品目.length;
  ws.getCell(`A${t}`).value = '合計';
  ws.getCell(`A${t}`).font = { bold: true }; ws.getCell(`A${t}`).fill = HEAD; ws.getCell(`A${t}`).border = B;
  for (const col of ['B','C','D','E']) {
    ws.getCell(`${col}${t}`).value = { formula: `SUM(${col}6:${col}${t - 1})` };
    ws.getCell(`${col}${t}`).numFmt = '#,##0';
    ws.getCell(`${col}${t}`).font = { bold: true };
    ws.getCell(`${col}${t}`).fill = HEAD; ws.getCell(`${col}${t}`).border = B;
  }
  ['F','G'].forEach((c) => { ws.getCell(`${c}${t}`).fill = HEAD; ws.getCell(`${c}${t}`).border = B; });

  ws.getCell(`A${t + 2}`).value = '記入者'; ws.getCell(`A${t + 2}`).font = { bold: true };
  ws.getCell(`B${t + 2}`).fill = YELLOW; ws.getCell(`B${t + 2}`).border = B; ws.getCell(`B${t + 2}`).protection = OPEN;
  ws.getCell(`A${t + 3}`).value = '記入日'; ws.getCell(`A${t + 3}`).font = { bold: true };
  ws.getCell(`B${t + 3}`).fill = YELLOW; ws.getCell(`B${t + 3}`).border = B; ws.getCell(`B${t + 3}`).protection = OPEN;

  ws.getCell(`A${t + 5}`).value =
    '※ 黄色の「2025年度計画」欄のみ入力してください。2023年度・2024年度の実績は変更できません。';
  ws.getCell(`A${t + 5}`).font = { size: 10, color: { argb: 'FFC00000' } };
  return ws;
}

function actualSheet(wb, seed) {
  const ws = wb.addWorksheet('2024年度実績');
  [22, 16, 16].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  ws.getCell('A1').value = '2024年度 数量実績（確定）';
  ws.getCell('A1').font = { bold: true, size: 13 };
  ws.getCell('A2').value = '確定日: 2025年5月20日';
  ws.getCell('A2').font = { size: 10, color: { argb: 'FF666666' } };
  ['品目', '2024年度実績', '2023年度実績'].forEach((h, i) => {
    const c = ws.getRow(4).getCell(i + 1);
    c.value = h; c.font = { bold: true }; c.fill = HEAD; c.border = B;
  });
  品目.forEach(([name, base], i) => {
    const r = 5 + i;
    ws.getCell(`A${r}`).value = name; ws.getCell(`A${r}`).border = B;
    ws.getCell(`B${r}`).value = base < 5000 ? 2018 : Math.round(base * (1 + ((seed * 5 + i * 11) % 9 - 4) / 100));
    ws.getCell(`B${r}`).numFmt = '#,##0'; ws.getCell(`B${r}`).border = B;
    ws.getCell(`C${r}`).value = base < 5000 ? 2031 : Math.round(base * (1 + ((seed * 3 + i * 7) % 9 - 4) / 100));
    ws.getCell(`C${r}`).numFmt = '#,##0'; ws.getCell(`C${r}`).border = B;
  });
  return ws;
}

async function main() {
  let n = 0;
  for (const [支店名, 部門s] of 支店) {
    const dir = join(ROOT, 支店名);
    mkdirSync(dir, { recursive: true });
    for (const 部門 of 部門s) {
      const wb = new ExcelJS.Workbook();
      wb.creator = '原価管理課';
      const a = budgetSheet(wb, 支店名, 部門, ++n);
      const b = actualSheet(wb, n);
      // 前年の作業でシート保護までかけてある状態にする
      await a.protect('', { selectLockedCells: true, selectUnlockedCells: true });
      await b.protect('', { selectLockedCells: true, selectUnlockedCells: true });
      await wb.xlsx.writeFile(join(dir, `2025年度_数量報告書_${支店名}_${部門}.xlsx`));
    }
  }
  console.log(`${n} ファイルを作成しました: ${ROOT}`);
}
void main();
