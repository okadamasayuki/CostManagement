/**
 * お試し用のサンプル Excel を、その場で作る。
 *
 * Excel が入っていない端末 (Mac など) でも操作感を確かめられるように、
 * 操作説明の動画で使っているのと同じ 30 ファイルをブラウザーの中で組み立てる。
 * 外部からは何も取ってこない (動画の収録に使ったものと同じ作りをここに写してある)。
 */
import ExcelJS from 'exceljs';

export type SampleKind = 'budget' | 'report';

export interface SampleFile {
  /** フォルダー階層を含む相対パス */
  relPath: string;
  fileName: string;
  data: ArrayBuffer;
}

export interface SampleInfo {
  kind: SampleKind;
  title: string;
  rootFolder: string;
  files: number;
  /** 何を試せるサンプルなのか */
  summary: string;
  points: string[];
}

const 支店: Array<[string, string[]]> = [
  ['東京支店', ['製造部', '営業部', '管理部', '技術部', '購買部', '物流部', '品質保証部', '開発部']],
  ['大阪支店', ['製造部', '営業部', '管理部', '技術部', '購買部', '物流部', '品質保証部', '開発部']],
  ['名古屋支店', ['製造部', '営業部', '管理部', '技術部', '購買部', '物流部', '品質保証部']],
  ['福岡支店', ['製造部', '営業部', '管理部', '技術部', '購買部', '物流部', '品質保証部']],
];

const YELLOW: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
const HEAD: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
const GRAY: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
const THIN: Partial<ExcelJS.Border> = { style: 'thin', color: { argb: 'FFB0B0B0' } };
const BORDER: Partial<ExcelJS.Borders> = { top: THIN, left: THIN, bottom: THIN, right: THIN };
const OPEN: ExcelJS.Protection = { locked: false } as ExcelJS.Protection;

export const SAMPLES: SampleInfo[] = [
  {
    kind: 'budget',
    title: '① 予算入力表（動画①と同じもの）',
    rootFolder: '原価管理/2025年度予算',
    files: 30,
    summary: '黄色いセルが支店に入力してもらう欄。ロックはまだ何もかかっていません。',
    points: [
      '「色からロックを設定」で 黄色だけ入力できる状態にする',
      '「シート保護を有効化」で実際に効かせる',
      '保存して、読み込み直して結果を見る',
    ],
  },
  {
    kind: 'report',
    title: '② 数量報告書（動画②と同じもの）',
    rootFolder: '報告共有フォルダー/2025年度報告',
    files: 30,
    summary:
      '前年の作業が終わった状態。2023・2024 年度の実績はロック済みで、2025 年度の記入欄だけ黄色です。',
    points: [
      '「年を +1 年ずらす」で 30 ファイルの年度をまとめて進める',
      'ファイル名・シート名・見出しにも年が入っている',
      '数量 2,031 個 / 2,018 個 が年と間違われないことを確かめられる',
    ],
  },
];

/** 動画①と同じ「予算入力表」 */
function budgetWorkbook(支店名: string, 部門: string, seed: number): ExcelJS.Workbook {
  const 費目: Array<[string, number]> = [
    ['材料費', 12_400_000], ['労務費', 28_600_000], ['外注加工費', 9_800_000],
    ['水道光熱費', 4_750_000], ['減価償却費', 6_200_000], ['修繕費', 3_100_000],
    ['旅費交通費', 1_450_000], ['通信費', 620_000], ['消耗品費', 2_380_000],
    ['支払手数料', 940_000], ['保険料', 1_180_000], ['雑費', 530_000],
  ];

  const wb = new ExcelJS.Workbook();
  wb.creator = '原価管理課';
  const ws = wb.addWorksheet('2025年度予算');
  [20, 16, 16, 14, 30].forEach((w, i) => {
    ws.getColumn(i + 1).width = w;
  });

  ws.getCell('A1').value = '2025年度 予算入力表';
  ws.getCell('A1').font = { bold: true, size: 15 };
  ws.getCell('A2').value = `${支店名}　${部門}`;
  ws.getCell('A2').font = { size: 12 };
  ws.getCell('A3').value = '提出期限: 2026年3月31日　／　単位: 円';
  ws.getCell('A3').font = { size: 10, color: { argb: 'FF666666' } };

  ['費目', '2024年度実績', '2025年度予算', '増減', '備考'].forEach((h, i) => {
    const c = ws.getRow(5).getCell(i + 1);
    c.value = h;
    c.font = { bold: true };
    c.fill = HEAD;
    c.border = BORDER;
    c.alignment = { horizontal: 'center' };
  });

  費目.forEach(([name, base], i) => {
    const r = 6 + i;
    const jitter = 1 + (((seed * 7 + i * 13) % 11) - 5) / 100;
    ws.getCell(`A${r}`).value = name;
    ws.getCell(`A${r}`).border = BORDER;
    ws.getCell(`B${r}`).value = Math.round((base * jitter) / 1000) * 1000;
    ws.getCell(`B${r}`).numFmt = '#,##0';
    ws.getCell(`B${r}`).border = BORDER;
    // ▼ ここが支店に入力してもらう欄 (黄色)
    ws.getCell(`C${r}`).fill = YELLOW;
    ws.getCell(`C${r}`).numFmt = '#,##0';
    ws.getCell(`C${r}`).border = BORDER;
    ws.getCell(`D${r}`).value = { formula: `IF(C${r}="","",C${r}-B${r})` };
    ws.getCell(`D${r}`).numFmt = '#,##0;[Red]-#,##0';
    ws.getCell(`D${r}`).border = BORDER;
    ws.getCell(`E${r}`).fill = YELLOW;
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

  const guide = wb.addWorksheet('記入要領');
  guide.getColumn(1).width = 90;
  guide.getCell('A1').value = '記入要領';
  guide.getCell('A1').font = { bold: true, size: 14 };
  [
    '1. 黄色のセルにのみ数値を入力してください。',
    '2. 「2025年度予算」欄は円単位、税抜きで記入してください。',
    '3. 前年度実績から 10% 以上増減する費目は、備考欄に理由を記入してください。',
    '4. 合計欄は自動計算されます。入力の必要はありません。',
    '5. 記入後、ファイル名は変更せずに共有フォルダーへ戻してください。',
  ].forEach((t2, i) => {
    guide.getCell(`A${3 + i}`).value = t2;
  });
  return wb;
}

/** 動画②と同じ「数量報告書」(前年の作業が終わった状態) */
function reportWorkbook(支店名: string, 部門: string, seed: number): ExcelJS.Workbook {
  // あえて 4 桁の少量品目を混ぜてある (年と紛らわしい数量の例)
  const 品目: Array<[string, number]> = [
    ['鋼材 SS400', 184_000], ['アルミ板 A5052', 96_500],
    ['樹脂ペレット', 312_000], ['ベアリング 6204', 48_200],
    ['モーター 750W', 6_400], ['配線ハーネス', 27_800],
    ['基板 ASSY', 15_600], ['特注シャフト', 2_030],
    ['防振ゴム', 73_400], ['塗料 (下塗り)', 11_250],
    ['梱包材', 128_000], ['ラベル', 205_000],
  ];

  const wb = new ExcelJS.Workbook();
  wb.creator = '原価管理課';
  const ws = wb.addWorksheet('2025年度');
  [22, 15, 15, 15, 13, 11, 26].forEach((w, i) => {
    ws.getColumn(i + 1).width = w;
  });

  ws.getCell('A1').value = '2025年度 数量報告書';
  ws.getCell('A1').font = { bold: true, size: 15 };
  ws.getCell('A2').value = `${支店名}　${部門}`;
  ws.getCell('A2').font = { size: 12 };
  ws.getCell('A3').value = '作成日: 2025年4月1日　／　前回報告: 2024年度　／　単位: 個';
  ws.getCell('A3').font = { size: 10, color: { argb: 'FF666666' } };

  ['品目', '2023年度実績', '2024年度実績', '2025年度計画', '増減', '増減率', '備考'].forEach((h, i) => {
    const c = ws.getRow(5).getCell(i + 1);
    c.value = h;
    c.font = { bold: true };
    c.fill = HEAD;
    c.border = BORDER;
    c.alignment = { horizontal: 'center' };
  });

  品目.forEach(([name, base], i) => {
    const r = 6 + i;
    const small = base < 5000;
    const j1 = small ? 1 : 1 + (((seed * 3 + i * 7) % 9) - 4) / 100;
    const j2 = small ? 1 : 1 + (((seed * 5 + i * 11) % 9) - 4) / 100;
    ws.getCell(`A${r}`).value = name;
    ws.getCell(`A${r}`).border = BORDER;
    // 過去の実績 = 変更されたくない (ロックのまま)
    ws.getCell(`B${r}`).value = small ? 2031 : Math.round(base * j1);
    ws.getCell(`B${r}`).numFmt = '#,##0';
    ws.getCell(`B${r}`).border = BORDER;
    ws.getCell(`B${r}`).fill = GRAY;
    ws.getCell(`C${r}`).value = small ? 2018 : Math.round(base * j2);
    ws.getCell(`C${r}`).numFmt = '#,##0';
    ws.getCell(`C${r}`).border = BORDER;
    ws.getCell(`C${r}`).fill = GRAY;
    // ▼ 支店に入力してもらう欄 (空欄・黄色・入力できる)
    ws.getCell(`D${r}`).fill = YELLOW;
    ws.getCell(`D${r}`).numFmt = '#,##0';
    ws.getCell(`D${r}`).border = BORDER;
    ws.getCell(`D${r}`).protection = OPEN;
    ws.getCell(`E${r}`).value = { formula: `IF(D${r}="","",D${r}-C${r})` };
    ws.getCell(`E${r}`).numFmt = '#,##0;[Red]-#,##0';
    ws.getCell(`E${r}`).border = BORDER;
    ws.getCell(`F${r}`).value = { formula: `IF(D${r}="","",D${r}/C${r}-1)` };
    ws.getCell(`F${r}`).numFmt = '0.0%';
    ws.getCell(`F${r}`).border = BORDER;
    ws.getCell(`G${r}`).fill = YELLOW;
    ws.getCell(`G${r}`).border = BORDER;
    ws.getCell(`G${r}`).protection = OPEN;
  });

  const t = 6 + 品目.length;
  ws.getCell(`A${t}`).value = '合計';
  ws.getCell(`A${t}`).font = { bold: true };
  ws.getCell(`A${t}`).fill = HEAD;
  ws.getCell(`A${t}`).border = BORDER;
  for (const col of ['B', 'C', 'D', 'E']) {
    ws.getCell(`${col}${t}`).value = { formula: `SUM(${col}6:${col}${t - 1})` };
    ws.getCell(`${col}${t}`).numFmt = '#,##0';
    ws.getCell(`${col}${t}`).font = { bold: true };
    ws.getCell(`${col}${t}`).fill = HEAD;
    ws.getCell(`${col}${t}`).border = BORDER;
  }
  for (const c of ['F', 'G']) {
    ws.getCell(`${c}${t}`).fill = HEAD;
    ws.getCell(`${c}${t}`).border = BORDER;
  }

  ws.getCell(`A${t + 2}`).value = '記入者';
  ws.getCell(`A${t + 2}`).font = { bold: true };
  ws.getCell(`B${t + 2}`).fill = YELLOW;
  ws.getCell(`B${t + 2}`).border = BORDER;
  ws.getCell(`B${t + 2}`).protection = OPEN;
  ws.getCell(`A${t + 3}`).value = '記入日';
  ws.getCell(`A${t + 3}`).font = { bold: true };
  ws.getCell(`B${t + 3}`).fill = YELLOW;
  ws.getCell(`B${t + 3}`).border = BORDER;
  ws.getCell(`B${t + 3}`).protection = OPEN;

  ws.getCell(`A${t + 5}`).value =
    '※ 黄色の「2025年度計画」欄のみ入力してください。2023年度・2024年度の実績は変更できません。';
  ws.getCell(`A${t + 5}`).font = { size: 10, color: { argb: 'FFC00000' } };

  // 前年の作業が終わっている状態なので、シート保護もかかっている
  ws.protect('', { selectLockedCells: true, selectUnlockedCells: true });

  const actual = wb.addWorksheet('2024年度実績');
  [22, 16, 16].forEach((w, i) => {
    actual.getColumn(i + 1).width = w;
  });
  actual.getCell('A1').value = '2024年度 数量実績（確定）';
  actual.getCell('A1').font = { bold: true, size: 13 };
  actual.getCell('A2').value = '確定日: 2025年5月20日';
  actual.getCell('A2').font = { size: 10, color: { argb: 'FF666666' } };
  ['品目', '2024年度実績', '2023年度実績'].forEach((h, i) => {
    const c = actual.getRow(4).getCell(i + 1);
    c.value = h;
    c.font = { bold: true };
    c.fill = HEAD;
    c.border = BORDER;
  });
  品目.forEach(([name, base], i) => {
    const r = 5 + i;
    actual.getCell(`A${r}`).value = name;
    actual.getCell(`A${r}`).border = BORDER;
    actual.getCell(`B${r}`).value =
      base < 5000 ? 2018 : Math.round(base * (1 + (((seed * 5 + i * 11) % 9) - 4) / 100));
    actual.getCell(`B${r}`).numFmt = '#,##0';
    actual.getCell(`B${r}`).border = BORDER;
    actual.getCell(`C${r}`).value =
      base < 5000 ? 2031 : Math.round(base * (1 + (((seed * 3 + i * 7) % 9) - 4) / 100));
    actual.getCell(`C${r}`).numFmt = '#,##0';
    actual.getCell(`C${r}`).border = BORDER;
  });
  actual.protect('', { selectLockedCells: true, selectUnlockedCells: true });
  return wb;
}

/**
 * サンプルを作る。1 ファイルごとに制御を返し、進み具合を知らせる。
 */
export async function buildSample(
  kind: SampleKind,
  onProgress?: (done: number, total: number) => void,
): Promise<SampleFile[]> {
  const info = SAMPLES.find((s) => s.kind === kind);
  if (!info) throw new Error(`未知のサンプル: ${kind}`);
  const out: SampleFile[] = [];
  const total = info.files;
  let seed = 0;
  for (const [支店名, 部門s] of 支店) {
    for (const 部門 of 部門s) {
      seed += 1;
      const wb =
        kind === 'budget'
          ? budgetWorkbook(支店名, 部門, seed)
          : reportWorkbook(支店名, 部門, seed);
      const fileName =
        kind === 'budget'
          ? `2025年度予算_${支店名}_${部門}.xlsx`
          : `2025年度_数量報告書_${支店名}_${部門}.xlsx`;
      const data = (await wb.xlsx.writeBuffer()) as ArrayBuffer;
      out.push({ relPath: `${info.rootFolder}/${支店名}/${fileName}`, fileName, data });
      onProgress?.(out.length, total);
      // 画面が固まらないよう 1 件ごとに制御を返す
      await new Promise((r) => setTimeout(r, 0));
    }
  }
  return out;
}
