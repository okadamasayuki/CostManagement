/**
 * 変換エンジンの検証。
 *
 * 実際に xlsx を生成 → 操作を適用 → 書き出し → 読み直して検証する
 * ラウンドトリップ試験。ブラウザーではなく Node 上で ExcelJS を直接
 * 使うため、UI とは独立してロジックの正しさを確認できる。
 *
 *   npm run test
 */
import ExcelJS from 'exceljs';
import assert from 'node:assert/strict';
import { applyStep, type OpContext } from '../src/excel/ops';
import type { LoadedWorkbook } from '../src/excel/types';
import type { RecipeStep, StepBody } from '../src/recipe/types';
import { DEFAULT_PROTECT_OPTIONS, DEFAULT_YEAR_TARGETS } from '../src/recipe/types';
import { replaceYearsInString, shiftMapper, pairMapper, replaceEraYears } from '../src/excel/yearShift';
import { parseRecipe, runRecipe } from '../src/recipe/runner';
import { recipeToJson } from '../src/recipe/document';
import { isCellLocked } from '../src/excel/view';

let passed = 0;
let failed = 0;
const failures: string[] = [];

async function test(name: string, fn: () => Promise<void> | void): Promise<void> {
  try {
    await fn();
    passed++;
    console.log(`  \x1b[32m✓\x1b[0m ${name}`);
  } catch (e) {
    failed++;
    const msg = e instanceof Error ? e.message : String(e);
    failures.push(`${name}\n    ${msg.split('\n').join('\n    ')}`);
    console.log(`  \x1b[31m✗\x1b[0m ${name}`);
  }
}

function section(name: string): void {
  console.log(`\n\x1b[1m${name}\x1b[0m`);
}

// --------------------------------------------------------------------------
// テスト用のブックを作る
// --------------------------------------------------------------------------

function makeBook(id: string, fileName: string, build: (wb: ExcelJS.Workbook) => void): LoadedWorkbook {
  const wb = new ExcelJS.Workbook();
  build(wb);
  return { id, relPath: fileName, fileName, wb, dirty: false, sizeBytes: 0 };
}

function ctx(books: LoadedWorkbook[], sheetName?: string): OpContext {
  return {
    books,
    currentBookId: books[0]?.id ?? null,
    currentSheetName: sheetName ?? null,
    selection: null,
  };
}

function step(body: StepBody, scope: RecipeStep['scope'] = { books: 'all', sheets: 'all' }): RecipeStep {
  return { id: 't', label: 't', enabled: true, scope, body };
}

/** 書き出して読み直す (実際に Excel が読むファイルと同じ経路を通す) */
async function roundTrip(book: LoadedWorkbook): Promise<ExcelJS.Workbook> {
  const buf = await book.wb.xlsx.writeBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf as ArrayBuffer);
  return wb;
}

// --------------------------------------------------------------------------

async function main(): Promise<void> {
  section('年の同時置換 (連鎖しないこと)');

  await test('2023→2024, 2024→2025 を同時に適用しても連鎖しない', () => {
    const m = shiftMapper(1, 2020, 2030);
    assert.equal(replaceYearsInString('2023年度', m, true), '2024年度');
    assert.equal(replaceYearsInString('2024年度', m, true), '2025年度');
    // 1 つの文字列に複数の年があっても、それぞれ 1 回だけ置換される
    assert.equal(replaceYearsInString('2023年度〜2024年度', m, true), '2024年度〜2025年度');
  });

  await test('対応表でも連鎖しない (2024→2025 と 2025→2026 を同時指定)', () => {
    const m = pairMapper([
      { from: 2024, to: 2025 },
      { from: 2025, to: 2026 },
    ]);
    assert.equal(replaceYearsInString('2024', m, true), '2025');
    assert.equal(replaceYearsInString('2025', m, true), '2026');
    assert.equal(replaceYearsInString('2024と2025', m, true), '2025と2026');
  });

  await test('対象年の範囲外は置換しない', () => {
    const m = shiftMapper(1, 2020, 2030);
    assert.equal(replaceYearsInString('1998年', m, true), null);
    assert.equal(replaceYearsInString('9999', m, true), null);
  });

  await test('数字の途中 (20240401) は既定で置換しない', () => {
    const m = shiftMapper(1, 2020, 2030);
    assert.equal(replaceYearsInString('20240401', m, true), null);
    // wholeNumberOnly=false なら置換する
    assert.equal(replaceYearsInString('20240401', m, false), '20250401');
  });

  await test('和暦をずらせる (令和6年 → 令和7年)', () => {
    assert.equal(replaceEraYears('令和6年度', 1), '令和7年度');
    assert.equal(replaceEraYears('令和元年', 1), '令和2年');
    assert.equal(replaceEraYears('平成31年', 1), '平成32年');
    assert.equal(replaceEraYears('ただの文字', 1), null);
  });

  // ------------------------------------------------------------------------
  section('セルのロック');

  await test('選択範囲だけをロック解除できる', async () => {
    const book = makeBook('b1', 'a.xlsx', (wb) => {
      const ws = wb.addWorksheet('明細');
      for (let r = 1; r <= 5; r++) for (let c = 1; c <= 5; c++) ws.getRow(r).getCell(c).value = r * c;
    });
    await applyStep(step({ op: 'setLock', range: { kind: 'a1', a1: 'B2:C3' }, locked: false }), ctx([book]));

    const wb = await roundTrip(book);
    const ws = wb.getWorksheet('明細')!;
    assert.equal(isCellLocked(ws.getCell('B2')), false, 'B2 は解除されるはず');
    assert.equal(isCellLocked(ws.getCell('C3')), false, 'C3 は解除されるはず');
    assert.equal(isCellLocked(ws.getCell('A1')), true, 'A1 はロックのままのはず');
    assert.equal(isCellLocked(ws.getCell('D4')), true, 'D4 はロックのままのはず');
  });

  await test('「選択範囲以外をロック」で指定範囲だけ入力可能になる', async () => {
    const book = makeBook('b1', 'a.xlsx', (wb) => {
      const ws = wb.addWorksheet('入力');
      for (let r = 1; r <= 4; r++) for (let c = 1; c <= 4; c++) ws.getRow(r).getCell(c).value = 'x';
      // 事前にあちこち解除しておく (これらは再びロックされるはず)
      ws.getCell('A1').protection = { locked: false };
      ws.getCell('D4').protection = { locked: false };
    });
    await applyStep(
      step({ op: 'lockAllExcept', range: { kind: 'a1', a1: 'B2:B3' }, alsoUnlockTarget: true }),
      ctx([book]),
    );

    const wb = await roundTrip(book);
    const ws = wb.getWorksheet('入力')!;
    assert.equal(isCellLocked(ws.getCell('B2')), false);
    assert.equal(isCellLocked(ws.getCell('B3')), false);
    assert.equal(isCellLocked(ws.getCell('A1')), true, '事前に解除されていた A1 が再ロックされるはず');
    assert.equal(isCellLocked(ws.getCell('D4')), true, '事前に解除されていた D4 が再ロックされるはず');
  });

  await test('スタイルを共有しているセルを巻き添えにしない', async () => {
    // 同じ書式のセルは ExcelJS 内部で style オブジェクトを共有するため、
    // その場書き換えをすると無関係なセルのロックまで変わってしまう。
    const src = new ExcelJS.Workbook();
    const sws = src.addWorksheet('S');
    for (let r = 1; r <= 3; r++) {
      const cell = sws.getRow(r).getCell(1);
      cell.value = r;
      cell.font = { bold: true }; // 3 セルとも同じ書式 = 同じ styleId
    }
    const buf = await src.xlsx.writeBuffer();

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf as ArrayBuffer);
    const book: LoadedWorkbook = {
      id: 'b1', relPath: 'a.xlsx', fileName: 'a.xlsx', wb, dirty: false, sizeBytes: 0,
    };

    await applyStep(step({ op: 'setLock', range: { kind: 'a1', a1: 'A2' }, locked: false }), ctx([book]));

    const out = await roundTrip(book);
    const ws = out.getWorksheet('S')!;
    assert.equal(isCellLocked(ws.getCell('A2')), false, 'A2 だけ解除されるはず');
    assert.equal(isCellLocked(ws.getCell('A1')), true, 'A1 は巻き添えにならないはず');
    assert.equal(isCellLocked(ws.getCell('A3')), true, 'A3 は巻き添えにならないはず');
    assert.equal(ws.getCell('A1').font?.bold, true, '元の書式は保たれるはず');
  });

  await test('シート保護をパスワード付きで有効化できる', async () => {
    const book = makeBook('b1', 'a.xlsx', (wb) => {
      wb.addWorksheet('保護').getCell('A1').value = 1;
    });
    await applyStep(
      step({ op: 'protectSheet', password: 'himitsu', options: DEFAULT_PROTECT_OPTIONS }),
      ctx([book]),
    );
    const wb = await roundTrip(book);
    const sp = (wb.getWorksheet('保護') as unknown as { sheetProtection: Record<string, unknown> })
      .sheetProtection;
    assert.ok(sp, 'シート保護が保存されているはず');
    assert.equal(sp.sheet, true);
    assert.ok(sp.hashValue, 'パスワードのハッシュが保存されているはず');
    assert.equal(String(sp.hashValue).includes('himitsu'), false, '平文が残っていないこと');
  });

  // ------------------------------------------------------------------------
  section('塗りつぶし');

  await test('ロックしていないセルだけをまとめて塗れる', async () => {
    const book = makeBook('b1', 'a.xlsx', (wb) => {
      const ws = wb.addWorksheet('色');
      for (let r = 1; r <= 3; r++) for (let c = 1; c <= 3; c++) ws.getRow(r).getCell(c).value = 'v';
      ws.getCell('B2').protection = { locked: false };
    });
    await applyStep(
      step({ op: 'fillByLockState', target: 'unlocked', colorArgb: 'FFFFFF00', onlyUsedRange: false }),
      ctx([book]),
    );
    const wb = await roundTrip(book);
    const ws = wb.getWorksheet('色')!;
    const fill = ws.getCell('B2').fill as { fgColor?: { argb?: string } };
    assert.equal(fill?.fgColor?.argb, 'FFFFFF00', 'B2 が黄色になるはず');
    assert.equal(ws.getCell('A1').fill, undefined, 'ロック済みの A1 は塗られないはず');
  });

  // ------------------------------------------------------------------------
  section('年度更新 (ブック全体)');

  await test('値・数式・シート名・ファイル名をまとめて更新できる', async () => {
    const book = makeBook('b1', '原価管理2024.xlsx', (wb) => {
      const ws = wb.addWorksheet('2024年度');
      ws.getCell('A1').value = '2024年度 原価集計';
      ws.getCell('A2').value = 2024;
      ws.getCell('A3').value = { formula: "'2024年度'!B1+2023" };
      ws.getCell('A4').value = '20240401'; // 日付の連番: 変わらないはず
      ws.getCell('A5').value = 1998; // 対象範囲外: 変わらないはず
      wb.addWorksheet('集計');
    });

    const out = await applyStep(
      step({
        op: 'shiftYears',
        delta: 1,
        minYear: 2020,
        maxYear: 2030,
        wholeNumberOnly: true,
        targets: { ...DEFAULT_YEAR_TARGETS, fileNames: true },
        range: { kind: 'used' },
      }),
      ctx([book]),
    );

    const wb = await roundTrip(book);
    const ws = wb.getWorksheet('2025年度');
    assert.ok(ws, 'シート名が 2025年度 に変わるはず');
    assert.equal(ws!.getCell('A1').value, '2025年度 原価集計');
    assert.equal(ws!.getCell('A2').value, 2025, '数値セルも更新されるはず');
    assert.equal(
      (ws!.getCell('A3').value as { formula: string }).formula,
      "'2025年度'!B1+2024",
      '数式の中のシート名と年が更新されるはず',
    );
    assert.equal(ws!.getCell('A4').value, '20240401', '数字の途中は変えないはず');
    assert.equal(ws!.getCell('A5').value, 1998, '対象範囲外は変えないはず');
    assert.equal(out.fileRenames[0]?.to, '原価管理2025.xlsx', 'ファイル名も更新されるはず');
  });

  await test('複数ブック・全シートに一度に適用できる', async () => {
    const mk = (id: string, name: string) =>
      makeBook(id, name, (wb) => {
        for (const s of ['一覧', '明細']) {
          wb.addWorksheet(s).getCell('A1').value = '2024年度';
        }
      });
    const books = [mk('b1', '予算.xlsx'), mk('b2', '実績.xlsx')];
    const out = await applyStep(
      step(
        {
          op: 'shiftYears',
          delta: 1,
          minYear: 2020,
          maxYear: 2030,
          wholeNumberOnly: true,
          targets: DEFAULT_YEAR_TARGETS,
          range: { kind: 'used' },
        },
        { books: 'all', sheets: 'all' },
      ),
      ctx(books),
    );
    assert.equal(out.changedBooks, 2);
    assert.equal(out.changedSheets, 4, '2 ブック × 2 シート');
    for (const b of books) {
      const wb = await roundTrip(b);
      for (const s of ['一覧', '明細']) {
        assert.equal(wb.getWorksheet(s)!.getCell('A1').value, '2025年度');
      }
      assert.equal(b.dirty, true, '変更済みとして記録されるはず');
    }
  });

  await test('ファイル名パターンで対象ブックを絞れる', async () => {
    const mk = (id: string, name: string) =>
      makeBook(id, name, (wb) => {
        wb.addWorksheet('S').getCell('A1').value = '2024';
      });
    const books = [mk('b1', '原価管理_東京.xlsx'), mk('b2', '予算.xlsx')];
    await applyStep(
      step(
        {
          op: 'shiftYears',
          delta: 1,
          minYear: 2020,
          maxYear: 2030,
          wholeNumberOnly: true,
          targets: DEFAULT_YEAR_TARGETS,
          range: { kind: 'used' },
        },
        { books: 'glob', bookGlob: '*原価*', sheets: 'all' },
      ),
      ctx(books),
    );
    assert.equal(books[0].wb.getWorksheet('S')!.getCell('A1').value, '2025');
    assert.equal(books[1].wb.getWorksheet('S')!.getCell('A1').value, '2024', '対象外は変わらないはず');
  });

  await test('シート名の衝突時は改名を見送る', async () => {
    const book = makeBook('b1', 'a.xlsx', (wb) => {
      wb.addWorksheet('2024').getCell('A1').value = 'x';
      wb.addWorksheet('2025').getCell('A1').value = 'y';
    });
    await applyStep(
      step({
        op: 'shiftYears',
        delta: 1,
        minYear: 2020,
        maxYear: 2030,
        wholeNumberOnly: true,
        targets: { ...DEFAULT_YEAR_TARGETS, values: false, formulas: false },
        range: { kind: 'used' },
      }),
      ctx([book]),
    );
    const names: string[] = [];
    book.wb.eachSheet((ws) => names.push(ws.name));
    // 2024→2025 は衝突するので見送られ、2025→2026 は実行される
    assert.ok(names.includes('2024'), '衝突する改名は見送られるはず');
    assert.ok(names.includes('2026'));
    assert.equal(new Set(names).size, names.length, 'シート名が重複していないこと');
  });

  await test('試算 (dryRun) ではファイルが変わらない', async () => {
    const book = makeBook('b1', 'a.xlsx', (wb) => {
      wb.addWorksheet('S').getCell('A1').value = '2024年度';
    });
    const out = await applyStep(
      step({
        op: 'shiftYears',
        delta: 1,
        minYear: 2020,
        maxYear: 2030,
        wholeNumberOnly: true,
        targets: DEFAULT_YEAR_TARGETS,
        range: { kind: 'used' },
      }),
      ctx([book]),
      { dryRun: true },
    );
    assert.ok(out.changedCells > 0, '件数は数えられるはず');
    assert.equal(book.wb.getWorksheet('S')!.getCell('A1').value, '2024年度', '値は変わらないはず');
    assert.equal(book.dirty, false, '未変更のままのはず');
  });

  // ------------------------------------------------------------------------
  section('手順書');

  await test('書き出した JSON を読み直せる', () => {
    const recipe = {
      format: 'excel-lock-manager/recipe' as const,
      version: 1,
      title: '年次更新',
      description: 'テスト',
      createdAt: new Date().toISOString(),
      sourceHint: ['a.xlsx'],
      steps: [
        step({ op: 'setLock', range: { kind: 'a1' as const, a1: 'A1:B2' }, locked: false }),
        step({ op: 'protectSheet', password: 'x', options: DEFAULT_PROTECT_OPTIONS }),
      ],
    };
    const parsed = parseRecipe(recipeToJson(recipe));
    assert.equal(parsed.steps.length, 2);
    assert.equal(parsed.title, '年次更新');
    assert.equal(parsed.steps[0].body.op, 'setLock');
  });

  await test('本ツール以外の JSON は拒否する', () => {
    assert.throws(() => parseRecipe('{"foo":1}'), /手順書ではない/);
    assert.throws(() => parseRecipe('こわれた JSON'), /JSON として読み取れません/);
  });

  await test('手順の一部が壊れていても読み込める', () => {
    const json = JSON.stringify({
      format: 'excel-lock-manager/recipe',
      steps: [
        { body: { op: '存在しない操作' } },
        { body: { op: 'setLock', range: { kind: 'a1', a1: 'A1' }, locked: true } },
      ],
    });
    const parsed = parseRecipe(json);
    assert.equal(parsed.steps.length, 1, '不明な操作は読み飛ばすはず');
  });

  await test('手順書を読み込んで別のブック群に一括適用できる (翌年の運用)', async () => {
    // 1 年目に作った手順書を、翌年の新しいファイル群へそのまま適用する想定
    const recipeJson = JSON.stringify({
      format: 'excel-lock-manager/recipe',
      version: 1,
      title: '年次更新手順',
      description: '',
      createdAt: new Date().toISOString(),
      sourceHint: [],
      steps: [
        {
          id: 's1',
          label: '入力欄以外をロック',
          enabled: true,
          scope: { books: 'all', sheets: 'glob', sheetGlob: '*年度*' },
          body: { op: 'lockAllExcept', range: { kind: 'a1', a1: 'C5:C9' }, alsoUnlockTarget: true },
        },
        {
          id: 's2',
          label: '入力欄を色分け',
          enabled: true,
          scope: { books: 'all', sheets: 'glob', sheetGlob: '*年度*' },
          body: {
            op: 'fillByLockState',
            target: 'unlocked',
            colorArgb: 'FFFFF2CC',
            onlyUsedRange: false,
          },
        },
        {
          id: 's3',
          label: 'シート保護',
          enabled: true,
          scope: { books: 'all', sheets: 'all' },
          body: { op: 'protectSheet', options: DEFAULT_PROTECT_OPTIONS },
        },
        {
          id: 's4',
          label: '年度を +1 年ずらす',
          enabled: true,
          scope: { books: 'all', sheets: 'all' },
          body: {
            op: 'shiftYears',
            delta: 1,
            minYear: 2000,
            maxYear: 2099,
            wholeNumberOnly: true,
            targets: { ...DEFAULT_YEAR_TARGETS, fileNames: true },
            range: { kind: 'used' },
          },
        },
      ],
    });

    // 手順書を書いたときとは別の (翌年の) ファイル群
    const mk = (id: string, name: string, year: number) =>
      makeBook(id, name, (wb) => {
        const ws = wb.addWorksheet(`${year}年度`);
        ws.getCell('A1').value = `${year}年度 原価管理表`;
        for (let r = 5; r <= 9; r++) {
          ws.getCell(`A${r}`).value = `費目${r - 4}`;
          ws.getCell(`C${r}`).value = 1000 * r;
        }
        wb.addWorksheet('集計').getCell('A1').value = `${year}年度 集計`;
      });
    const books = [mk('b1', '原価管理2025.xlsx', 2025), mk('b2', '原価管理2025_東京.xlsx', 2025)];

    const recipe = parseRecipe(recipeJson);
    const report = await runRecipe(recipe, ctx(books));

    assert.equal(report.steps.filter((s) => s.error).length, 0, 'エラーなく完了するはず');
    assert.ok(report.totalChangedCells > 0);
    assert.equal(report.fileRenames.length, 2, '2 冊ともファイル名が変わるはず');
    assert.equal(report.fileRenames[0].to, '原価管理2026.xlsx');

    for (const b of books) {
      const wb = await roundTrip(b);
      const ws = wb.getWorksheet('2026年度');
      assert.ok(ws, `${b.fileName}: シート名が 2026年度 になるはず`);
      assert.equal(ws!.getCell('A1').value, '2026年度 原価管理表');
      assert.equal(isCellLocked(ws!.getCell('C5')), false, '入力欄が解除されるはず');
      assert.equal(isCellLocked(ws!.getCell('A5')), true, '費目名はロックされるはず');
      const fill = ws!.getCell('C5').fill as { fgColor?: { argb?: string } };
      assert.equal(fill?.fgColor?.argb, 'FFFFF2CC');
      assert.ok((ws as unknown as { sheetProtection: unknown }).sheetProtection, 'シート保護が有効なはず');
      // 集計シートは sheetGlob に一致しないのでロック操作の対象外、
      // ただし年度更新とシート保護 (sheets:'all') は適用される
      const sum = wb.getWorksheet('集計')!;
      assert.equal(sum.getCell('A1').value, '2026年度 集計');
      assert.ok((sum as unknown as { sheetProtection: unknown }).sheetProtection);
    }
  });

  await test('無効にした手順は実行されない', async () => {
    const book = makeBook('b1', 'a.xlsx', (wb) => {
      wb.addWorksheet('S').getCell('A1').value = '2024年度';
    });
    const recipe = parseRecipe(
      JSON.stringify({
        format: 'excel-lock-manager/recipe',
        steps: [
          {
            id: 's1',
            enabled: false,
            scope: { books: 'all', sheets: 'all' },
            body: {
              op: 'shiftYears',
              delta: 1,
              minYear: 2000,
              maxYear: 2099,
              wholeNumberOnly: true,
              targets: DEFAULT_YEAR_TARGETS,
              range: { kind: 'used' },
            },
          },
        ],
      }),
    );
    const report = await runRecipe(recipe, ctx([book]));
    assert.equal(report.steps.length, 0);
    assert.equal(book.wb.getWorksheet('S')!.getCell('A1').value, '2024年度');
  });

  // ------------------------------------------------------------------------
  section('全体の流れ (ロック → 保護 → 色分け → 年度更新)');

  await test('典型的な運用を通しで実行できる', async () => {
    const book = makeBook('b1', '原価管理2024.xlsx', (wb) => {
      const ws = wb.addWorksheet('2024年度');
      ws.getCell('A1').value = '2024年度 原価管理表';
      ws.getCell('A3').value = '費目';
      ws.getCell('B3').value = '金額';
      for (let r = 4; r <= 8; r++) {
        ws.getCell(`A${r}`).value = `費目${r - 3}`;
        ws.getCell(`B${r}`).value = 1000 * (r - 3);
      }
      ws.getCell('B9').value = { formula: 'SUM(B4:B8)' };
    });
    const c = ctx([book], '2024年度');

    // ① 入力欄 (B4:B8) 以外をすべてロック
    await applyStep(step({ op: 'lockAllExcept', range: { kind: 'a1', a1: 'B4:B8' }, alsoUnlockTarget: true }), c);
    // ② シートを保護
    await applyStep(step({ op: 'protectSheet', options: DEFAULT_PROTECT_OPTIONS }), c);
    // ③ 入力欄を色分け
    await applyStep(
      step({ op: 'fillByLockState', target: 'unlocked', colorArgb: 'FFFFF2CC', onlyUsedRange: false }),
      c,
    );
    // ④ 年度を +1
    await applyStep(
      step({
        op: 'shiftYears',
        delta: 1,
        minYear: 2020,
        maxYear: 2030,
        wholeNumberOnly: true,
        targets: { ...DEFAULT_YEAR_TARGETS, fileNames: true },
        range: { kind: 'used' },
      }),
      c,
    );

    const wb = await roundTrip(book);
    const ws = wb.getWorksheet('2025年度');
    assert.ok(ws, 'シート名が更新されているはず');
    assert.equal(ws!.getCell('A1').value, '2025年度 原価管理表');
    assert.equal(isCellLocked(ws!.getCell('B4')), false, '入力欄は解除されているはず');
    assert.equal(isCellLocked(ws!.getCell('A1')), true, '見出しはロックされているはず');
    assert.equal(isCellLocked(ws!.getCell('B9')), true, '合計欄はロックされているはず');
    const fill = ws!.getCell('B5').fill as { fgColor?: { argb?: string } };
    assert.equal(fill?.fgColor?.argb, 'FFFFF2CC', '入力欄に色が付いているはず');
    assert.ok(
      (ws as unknown as { sheetProtection: unknown }).sheetProtection,
      'シート保護が有効なはず',
    );
    assert.equal(
      (ws!.getCell('B9').value as { formula: string }).formula,
      'SUM(B4:B8)',
      '年を含まない数式は変わらないはず',
    );
  });

  // ------------------------------------------------------------------------
  console.log(`\n${'─'.repeat(60)}`);
  if (failed) {
    console.log(`\x1b[31m${failed} 件失敗 / ${passed} 件成功\x1b[0m\n`);
    for (const f of failures) console.log(`  \x1b[31m✗\x1b[0m ${f}\n`);
    process.exit(1);
  }
  console.log(`\x1b[32m${passed} 件すべて成功\x1b[0m`);
}

void main();
