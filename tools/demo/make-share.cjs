/**
 * 動画用の「共有フォルダー」を実データで作る。
 * 支店ごとのフォルダーに、黄色い入力欄を持つ予算入力表を配置する。
 */
const ExcelJS = require('exceljs');
const { mkdirSync, rmSync } = require('node:fs');
const { join } = require('node:path');

const ROOT = process.env.SHARE_DIR ?? '/root/共有フォルダー/原価管理/2025年度予算';
rmSync('/home/user/CostManagement/.test-build/video/共有フォルダー', { recursive: true, force: true });

const YELLOW = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
const HEAD   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
const THIN   = { style: 'thin', color: { argb: 'FFB0B0B0' } };
const BORDER = { top: THIN, left: THIN, bottom: THIN, right: THIN };

const 支店 = [
  ['東京支店', ['製造部', '営業部', '管理部', '技術部', '購買部', '物流部', '品質保証部', '開発部']],
  ['大阪支店', ['製造部', '営業部', '管理部', '技術部', '購買部', '物流部', '品質保証部', '開発部']],
  ['名古屋支店', ['製造部', '営業部', '管理部', '技術部', '購買部', '物流部', '品質保証部']],
  ['福岡支店', ['製造部', '営業部', '管理部', '技術部', '購買部', '物流部', '品質保証部']],
];

const 費目 = [
  ['材料費', 12_400_000], ['労務費', 28_600_000], ['外注加工費', 9_800_000],
  ['electric', 0], ['減価償却費', 6_200_000], ['修繕費', 3_100_000],
  ['旅費交通費', 1_450_000], ['通信費', 620_000], ['消耗品費', 2_380_000],
  ['支払手数料', 940_000], ['保険料', 1_180_000], ['雑費', 530_000],
];
費目[3] = ['水道光熱費', 4_750_000];

function buildBudgetSheet(wb, 支店名, 部門, seed) {
  const ws = wb.addWorksheet('2025年度予算');
  ws.views = [{ showGridLines: true }];
  ws.getColumn(1).width = 20;
  ws.getColumn(2).width = 16;
  ws.getColumn(3).width = 16;
  ws.getColumn(4).width = 14;
  ws.getColumn(5).width = 30;

  ws.getCell('A1').value = '2025年度 予算入力表';
  ws.getCell('A1').font = { bold: true, size: 15 };
  ws.getCell('A2').value = `${支店名}　${部門}`;
  ws.getCell('A2').font = { size: 12 };
  ws.getCell('A3').value = '提出期限: 2026年3月31日　／　単位: 円';
  ws.getCell('A3').font = { size: 10, color: { argb: 'FF666666' } };

  const heads = ['費目', '2024年度実績', '2025年度予算', '増減', '備考'];
  heads.forEach((h, i) => {
    const c = ws.getRow(5).getCell(i + 1);
    c.value = h;
    c.font = { bold: true };
    c.fill = HEAD;
    c.border = BORDER;
    c.alignment = { horizontal: 'center' };
  });

  費目.forEach(([name, base], i) => {
    const r = 6 + i;
    const jitter = 1 + ((seed * 7 + i * 13) % 11 - 5) / 100;
    ws.getCell(`A${r}`).value = name;
    ws.getCell(`A${r}`).border = BORDER;
    ws.getCell(`B${r}`).value = Math.round(base * jitter / 1000) * 1000;
    ws.getCell(`B${r}`).numFmt = '#,##0';
    ws.getCell(`B${r}`).border = BORDER;
    // ▼ ここが支店に入力してもらう欄 (黄色)
    ws.getCell(`C${r}`).fill = YELLOW;
    ws.getCell(`C${r}`).numFmt = '#,##0';
    ws.getCell(`C${r}`).border = BORDER;
    ws.getCell(`D${r}`).value = { formula: `IF(C${r}="","",C${r}-B${r})` };
    ws.getCell(`D${r}`).numFmt = '#,##0;[Red]-#,##0';
    ws.getCell(`D${r}`).border = BORDER;
    ws.getCell(`E${r}`).fill = YELLOW;      // 備考も入力欄
    ws.getCell(`E${r}`).border = BORDER;
  });

  const t = 6 + 費目.length;
  ws.getCell(`A${t}`).value = '合計';
  ws.getCell(`A${t}`).font = { bold: true };
  ws.getCell(`A${t}`).fill = HEAD;
  ws.getCell(`A${t}`).border = BORDER;
  for (const col of ['B', 'C', 'D']) {
    ws.getCell(`${col}${t}`).value = { formula: `SUM(${col}6:${col}${t - 1})` };
    ws.getCell(`${col}${t}`).numFmt = '#,##0';
    ws.getCell(`${col}${t}`).font = { bold: true };
    ws.getCell(`${col}${t}`).fill = HEAD;
    ws.getCell(`${col}${t}`).border = BORDER;
  }
  ws.getCell(`E${t}`).fill = HEAD;
  ws.getCell(`E${t}`).border = BORDER;

  ws.getCell(`A${t + 2}`).value = '記入者';
  ws.getCell(`A${t + 2}`).font = { bold: true };
  ws.getCell(`B${t + 2}`).fill = YELLOW;
  ws.getCell(`B${t + 2}`).border = BORDER;
  ws.getCell(`A${t + 3}`).value = '記入日';
  ws.getCell(`A${t + 3}`).font = { bold: true };
  ws.getCell(`B${t + 3}`).fill = YELLOW;
  ws.getCell(`B${t + 3}`).border = BORDER;
}

function buildGuideSheet(wb) {
  const ws = wb.addWorksheet('記入要領');
  ws.getColumn(1).width = 90;
  ws.getCell('A1').value = '記入要領';
  ws.getCell('A1').font = { bold: true, size: 14 };
  [
    '1. 黄色のセルにのみ数値を入力してください。',
    '2. 「2025年度予算」欄は円単位、税抜きで記入してください。',
    '3. 前年度実績から 10% 以上増減する費目は、備考欄に理由を記入してください。',
    '4. 合計欄は自動計算されます。入力の必要はありません。',
    '5. 記入後、ファイル名は変更せずに共有フォルダーへ戻してください。',
  ].forEach((t, i) => { ws.getCell(`A${3 + i}`).value = t; });
}

async function main() {
  let n = 0;
  for (const [支店名, 部門s] of 支店) {
    const dir = join(ROOT, 支店名);
    mkdirSync(dir, { recursive: true });
    for (const 部門 of 部門s) {
      const wb = new ExcelJS.Workbook();
      wb.creator = '原価管理課';
      buildBudgetSheet(wb, 支店名, 部門, ++n);
      buildGuideSheet(wb);
      await wb.xlsx.writeFile(join(dir, `2025年度予算_${支店名}_${部門}.xlsx`));
    }
  }
  console.log(`${n} ファイルを作成しました`);
}
void main();
