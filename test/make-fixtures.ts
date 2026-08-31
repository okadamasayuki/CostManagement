/**
 * 動作確認用のサンプル Excel を生成する。
 * 原価管理の年次更新という想定で、年度・数式・複数シートを含む。
 *   npm run fixtures
 */
import ExcelJS from 'exceljs';
import { mkdirSync } from 'node:fs';
import { join } from 'node:path';

const OUT = process.argv[2] ?? 'sample';
const YEAR = 2024;

function buildCostSheet(wb: ExcelJS.Workbook, sheetName: string, dept: string): void {
  const ws = wb.addWorksheet(sheetName);
  ws.getColumn(1).width = 18;
  ws.getColumn(2).width = 14;
  ws.getColumn(3).width = 14;
  ws.getColumn(4).width = 14;
  ws.getColumn(5).width = 26;

  ws.getCell('A1').value = `${YEAR}年度 原価管理表（${dept}）`;
  ws.getCell('A1').font = { bold: true, size: 14 };
  ws.getCell('A2').value = `作成: ${YEAR}年4月1日 / 前年度: ${YEAR - 1}年度`;

  const headers = ['費目', `${YEAR - 1}年度実績`, `${YEAR}年度予算`, '差額', '備考'];
  headers.forEach((h, i) => {
    const cell = ws.getRow(4).getCell(i + 1);
    cell.value = h;
    cell.font = { bold: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
  });

  const items = ['材料費', '労務費', '外注費', '経費', '減価償却費'];
  items.forEach((name, i) => {
    const r = 5 + i;
    ws.getCell(`A${r}`).value = name;
    ws.getCell(`B${r}`).value = 1_000_000 * (i + 1);
    ws.getCell(`C${r}`).value = 1_100_000 * (i + 1);
    // 予算欄は「入力してもらう欄」として黄色にしておく。
    // 「色からロックを設定」の動作確認に使う。
    ws.getCell(`C${r}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFFF00' },
    };
    ws.getCell(`D${r}`).value = { formula: `C${r}-B${r}` };
    ws.getCell(`E${r}`).value = i === 0 ? `${YEAR}年度は単価改定を反映` : '';
    for (const col of ['B', 'C', 'D']) ws.getCell(`${col}${r}`).numFmt = '#,##0';
  });

  const total = 5 + items.length;
  ws.getCell(`A${total}`).value = '合計';
  ws.getCell(`A${total}`).font = { bold: true };
  for (const col of ['B', 'C', 'D']) {
    ws.getCell(`${col}${total}`).value = { formula: `SUM(${col}5:${col}${total - 1})` };
    ws.getCell(`${col}${total}`).numFmt = '#,##0';
    ws.getCell(`${col}${total}`).font = { bold: true };
  }
}

async function makeWorkbook(path: string, dept: string): Promise<void> {
  const wb = new ExcelJS.Workbook();
  buildCostSheet(wb, `${YEAR}年度`, dept);
  buildCostSheet(wb, `${YEAR - 1}年度`, dept);

  const summary = wb.addWorksheet('集計');
  summary.getColumn(1).width = 24;
  summary.getCell('A1').value = `${YEAR}年度 集計`;
  summary.getCell('A1').font = { bold: true, size: 13 };
  summary.getCell('A3').value = `${YEAR}年度 合計`;
  summary.getCell('B3').value = { formula: `'${YEAR}年度'!C10` };
  summary.getCell('A4').value = `${YEAR - 1}年度 合計`;
  summary.getCell('B4').value = { formula: `'${YEAR - 1}年度'!C10` };

  await wb.xlsx.writeFile(path);
  console.log(`  作成: ${path}`);
}

/**
 * ファイル名を ASCII にするかどうか。
 * 自動テスト (Playwright) はコンテナのロケールの都合で日本語パスの
 * ファイルを input へ渡せないため、テスト用は ASCII 名で生成する。
 * シート名・セルの中身は日本語のままなので、動作確認としては十分。
 */
const ASCII = process.argv.includes('--ascii');

async function main(): Promise<void> {
  const dirs = ASCII ? ['tokyo', 'osaka'] : ['東京', '大阪'];
  for (const d of dirs) mkdirSync(join(OUT, d), { recursive: true });
  console.log(`サンプルを ${OUT}/ に作成します`);
  const name = (suffix: string) =>
    ASCII ? `cost_${YEAR}${suffix ? `_${suffix}` : ''}.xlsx` : `原価管理${YEAR}${suffix ? `_${suffix}` : ''}.xlsx`;
  await makeWorkbook(join(OUT, name('')), '全社');
  await makeWorkbook(join(OUT, dirs[0], name(ASCII ? 'tokyo' : '東京')), '東京支店');
  await makeWorkbook(join(OUT, dirs[1], name(ASCII ? 'osaka' : '大阪')), '大阪支店');
  console.log('完了');
}

void main();
