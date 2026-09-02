/**
 * ビルド済みの単一 HTML を実際のブラウザーで動かす通し試験。
 *
 *  ・グリッドが描画されること
 *  ・ロック / 色分け / 年度更新が UI から実行できること
 *  ・外部通信が 1 件も発生しないこと (最重要)
 *
 *   npm run build && npm run test:ui
 */
import { chromium } from 'playwright';
import { fileURLToPath } from 'node:url';
import { dirname, join, resolve } from 'node:path';
import { existsSync, mkdirSync, readFileSync } from 'node:fs';
import { createServer } from 'node:http';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const DIST = join(root, 'dist', 'index.html');
const SAMPLE = join(root, '.test-build', 'sample');
const SHOTS = join(root, '.test-build', 'shots');
const WIDE_FILE = join(root, '.test-build', 'wide.xlsx');
/** 2 年分の比較用。年フォルダーの下に同じ様式が 1 つずつ入っている。 */
const YEARS = join(root, '.test-build', 'years');

let passed = 0;
const failures = [];

async function check(name, fn) {
  try {
    await fn();
    passed++;
    console.log(`  \x1b[32m✓\x1b[0m ${name}`);
  } catch (e) {
    failures.push(`${name}\n    ${e.message}`);
    console.log(`  \x1b[31m✗\x1b[0m ${name}\n      \x1b[90m${String(e.message).split('\n')[0].slice(0, 200)}\x1b[0m`);
  }
}

function assert(cond, msg) {
  if (!cond) throw new Error(msg);
}

/**
 * 条件が満たされるまで待つ。
 * 固定時間の待機だと CI の速度差で不安定になるため、こちらを使う。
 */
/** ツリーの ✕ を押して、読み込んだブックを全部閉じる */
async function closeAllBooks(page) {
  for (let i = 0; i < 20; i++) {
    const n = await page.locator('.tree-file').count();
    if (n === 0) return;
    await page.locator('.tree-file .icon-btn').first().click();
    await new Promise((r) => setTimeout(r, 60));
  }
  throw new Error('ブックを閉じきれない');
}

/** 文字を手がかりにセルの class を読む (ロック表示の確認用) */
async function cellClass(page, text) {
  return page
    .locator('.cell', { hasText: text })
    .first()
    .evaluate((el) => el.className);
}

/** 文字を手がかりにセルの背景色を読む */
async function bg(page, text) {
  return page
    .locator('.cell', { hasText: text })
    .first()
    .evaluate((el) => getComputedStyle(el).backgroundColor);
}

async function waitUntil(fn, msg, timeout = 20000) {
  const deadline = Date.now() + timeout;
  let last;
  for (;;) {
    try {
      if (await fn()) return;
    } catch (e) {
      last = e.message;
    }
    if (Date.now() > deadline) {
      throw new Error(`${msg}${last ? ` (最後のエラー: ${last})` : ''}`);
    }
    await new Promise((r) => setTimeout(r, 100));
  }
}

/**
 * ブラウザーを起動する前に UTF-8 のロケールを保証する。
 *
 * ロケールが POSIX のままだと Chromium が download 属性の日本語ファイル名を
 * 落とし、拡張子ごと "download" になってしまう。実機 (Windows / macOS) では
 * 起きないが、そのままだと「保存したファイル名」を検証できない。
 * ダブルクリックで開ける .html として保存されることは、共有フォルダー運用の
 * 前提なので、ここを確かめられるようにしておく。
 */
if (!process.env.LC_ALL && !process.env.LANG) process.env.LC_ALL = 'C.UTF-8';

const requests = [];
const consoleErrors = [];

// この開発コンテナには Chromium が同梱されている。
// CI など無い環境では Playwright が用意したものを使う。
const CONTAINER_CHROMIUM = '/opt/pw-browsers/chromium-1194/chrome-linux/chrome';
const executablePath =
  process.env.CHROMIUM_PATH ?? (existsSync(CONTAINER_CHROMIUM) ? CONTAINER_CHROMIUM : undefined);

const browser = await chromium.launch({
  executablePath,
  args: [
    '--no-sandbox',
    '--disable-dev-shm-usage',
    // Chromium 自身のバックグラウンド通信を止める。
    // このテスト環境は外部へ出られないため、有効なままだと起動が待たされる。
    '--disable-background-networking',
    '--disable-component-update',
    '--disable-sync',
    '--disable-default-apps',
    '--no-first-run',
    '--no-default-browser-check',
    '--metrics-recording-only',
    '--disable-domain-reliability',
    '--disable-client-side-phishing-detection',
  ],
});
const page = await browser.newPage({ viewport: { width: 1600, height: 950 } });

// file:// 以外へのリクエストをすべて記録する
page.on('request', (r) => {
  if (!r.url().startsWith('file://')) requests.push(`${r.method()} ${r.url()}`);
});
page.on('console', (m) => {
  if (m.type() === 'error') consoleErrors.push(m.text());
});
page.on('pageerror', (e) => consoleErrors.push(`pageerror: ${e.message}`));

// 必要なサンプルが無いまま進むと Playwright が待ち続けて時間を浪費するため、
// 最初に確かめて、無ければすぐ落とす
for (const required of [
  DIST,
  WIDE_FILE,
  join(SAMPLE, 'cost_2024.xlsx'),
  join(YEARS, '2025', 'plan_2025.xlsx'),
]) {
  if (!existsSync(required)) {
    console.error(`必要なファイルがありません: ${required}\n  npm run test:ui から実行してください。`);
    process.exit(1);
  }
}

mkdirSync(SHOTS, { recursive: true });
await page.goto(`file://${DIST}`);
await page.waitForSelector('.app');

console.log('\n\x1b[1m画面の初期表示\x1b[0m');

await check('タイトルバーにオフライン表示が出る', async () => {
  const text = await page.textContent('[data-testid="guard-badge"]');
  assert(text.includes('完全オフライン'), `実際: ${text}`);
});

await check('ファイル未読み込み時の案内が出る', async () => {
  assert(await page.isVisible('.grid-placeholder'), 'プレースホルダが見えない');
});

console.log('\n\x1b[1mファイルの読み込み\x1b[0m');

await check('フォルダーを指定するとサブフォルダーの Excel も全部読み込まれる', async () => {
  // 注: Playwright はこのコンテナのロケールでは日本語パスのファイルを
  //     input へ渡せないため、サンプルは ASCII 名で生成している
  //     (シート名・セルの内容は日本語のまま)。
  await page.setInputFiles('input[webkitdirectory]', SAMPLE);
  await page.waitForSelector('.tree-file', { timeout: 20000 });
  const n = await page.locator('.tree-file').count();
  assert(n === 3, `ファイル数が ${n} 件`);
});

await check('フォルダー階層がツリーに再現される', async () => {
  const folders = await page.locator('.tree-folder').allTextContents();
  assert(folders.some((f) => f.includes('tokyo')), `フォルダー: ${folders.join(', ')}`);
  assert(folders.some((f) => f.includes('osaka')), `フォルダー: ${folders.join(', ')}`);
});

await check('グリッドにセルの内容が描画される', async () => {
  await page.waitForSelector('.cell');
  const text = await page.textContent('.grid-canvas');
  assert(text.includes('原価管理表'), 'A1 の値が見えない');
  assert(text.includes('材料費'), '費目が見えない');
  assert(text.includes('1,000,000'), '桁区切りの表示形式が効いていない');
});

await check('シートタブが 3 枚出る', async () => {
  const n = await page.locator('.sheettab').count();
  assert(n === 3, `タブ数が ${n}`);
});

await page.screenshot({ path: join(SHOTS, '01-loaded.png') });

await check('「ファイル」タブから不要な項目が消えている', async () => {
  await page.click('.ribbon-tab:has-text("ファイル")');
  const panel = await page.textContent('.ribbon-panel');
  assert(!panel.includes('すべて閉じる'), '「すべて閉じる」が残っている');
  assert(!panel.includes('変更を破棄'), '「変更を破棄して読み直す」が残っている');
  assert(!panel.includes('サブフォルダーも含めて'), '開くの説明文が残っている');
  assert(!panel.includes('直接上書き保存できます'), '開くの説明文が残っている');
  // 必要な機能は残っていること
  assert(panel.includes('フォルダーを'), 'フォルダーを開くが消えている');
  assert(panel.includes('ZIP'), '保存が消えている');
});

console.log('\n\x1b[1m読み込み時のロック状態\x1b[0m');

await check('「全シートのロックを外す」で読み込むと最初から全解除になる', async () => {
  // いったん閉じてから、設定を変えて読み込み直す
  await page.click('.ribbon-tab:has-text("ファイル")');
  await closeAllBooks(page);

  await page.selectOption('[data-testid="initial-lock"]', 'unlock');
  await page.setInputFiles('input[webkitdirectory]', SAMPLE);
  await waitUntil(async () => (await page.locator('.cell').count()) > 0, '読み込めない');
  await waitUntil(async () => {
    const status = await page.textContent('.statusbar');
    // 🔒 ロック済み が 0 件になっているはず
    return /🔒 0 \//.test(status);
  }, `ロックが解除されていない: ${await page.textContent('.statusbar')}`);

  // 画面上もロックの網掛けが消えていること
  assert(
    (await page.locator('.cell.ov-locked').count()) === 0,
    'ロック済みの表示が残っている',
  );

  // 設定を戻して読み込み直す
  await closeAllBooks(page);
  await page.selectOption('[data-testid="initial-lock"]', 'keep');
  await page.setInputFiles('input[webkitdirectory]', SAMPLE);
  await waitUntil(async () => (await page.locator('.cell').count()) > 0, '読み込み直せない');
});

console.log('\n\x1b[1m大きなシートの読み込み\x1b[0m');

await check('GR 列より右のデータも読み込める', async () => {
  await page.click('.ribbon-tab:has-text("ファイル")');
  await closeAllBooks(page);
  await page.setInputFiles('input[type="file"]:not([webkitdirectory])', [WIDE_FILE]);
  await waitUntil(async () => (await page.locator('.cell').count()) > 0, '読み込めない');

  // ZZ1 (702 列目) へ移動して中身を確認する
  await page.fill('.namebox input', 'ZZ1');
  await page.press('.namebox input', 'Enter');
  await waitUntil(
    async () => (await page.inputValue('.formula-input')) === '最果ての値',
    `ZZ1 が読めない: ${await page.inputValue('.formula-input')}`,
  );
});

await check('AMJ 列 (Excel の 1024 列目) まで移動できる', async () => {
  await page.fill('.namebox input', 'AMJ2000');
  await page.press('.namebox input', 'Enter');
  await waitUntil(
    async () => (await page.inputValue('.namebox input')) === 'AMJ2000',
    'Excel の範囲内に移動できない',
  );
});

await check('大きなシートでも操作がファイル全体に効く', async () => {
  await page.click('.ribbon-tab:has-text("年度更新")');
  await page.selectOption('[data-testid="scope-books"]', 'all');
  await page.selectOption('[data-testid="scope-sheets"]', 'all');
  await page.click('.rbtn-lg:has-text("年度更新を")');
  await waitUntil(
    async () => (await page.locator('.toast').last().textContent()).includes('箇所'),
    '年度更新が終わらない',
  );
  await page.fill('.namebox input', 'ZZ2');
  await page.press('.namebox input', 'Enter');
  await waitUntil(
    async () => (await page.inputValue('.formula-input')) === '2025年度',
    `右端の年度が更新されない: ${await page.inputValue('.formula-input')}`,
  );
  // 後片付け
  await page.click('.ribbon-tab:has-text("ファイル")');
  await closeAllBooks(page);
  await page.setInputFiles('input[webkitdirectory]', SAMPLE);
  await waitUntil(async () => (await page.locator('.cell').count()) > 0, '読み込み直せない');
  await page.click('.ribbon-tab:has-text("年度更新")');
  await page.selectOption('[data-testid="scope-books"]', 'current');
  await page.selectOption('[data-testid="scope-sheets"]', 'current');
});

console.log('\n\x1b[1mフォルダーを指定して一括処理\x1b[0m');

await check('ツリーのフォルダーから対象を指定できる', async () => {
  const tokyo = page.locator('.tree-folder', { hasText: 'tokyo' }).first();
  await tokyo.hover();
  await tokyo.locator('.icon-btn').click();
  await waitUntil(
    async () => (await page.locator('.scope-banner').count()) > 0,
    '対象フォルダーの表示が出ない',
  );
  const banner = await page.textContent('.scope-banner');
  assert(/tokyo/.test(banner), `対象の表示が想定外: ${banner}`);
});

await check('リボンの適用先にもフォルダー指定が反映される', async () => {
  await page.click('.ribbon-tab:has-text("年度更新")');
  assert(
    (await page.inputValue('[data-testid="scope-books"]')) === 'folder',
    '適用先がフォルダー指定になっていない',
  );
  const folder = await page.inputValue('[data-testid="scope-folder"]');
  assert(folder.endsWith('tokyo'), `選ばれているフォルダー: ${folder}`);
});

await check('そのフォルダー配下のブックだけが対象になる', async () => {
  // 実際には変更せず、試算で対象を確かめる (以降のテストに影響させないため)
  await page.selectOption('[data-testid="scope-sheets"]', 'all');
  await page.click('.rbtn-lg:has-text("変更内容を")');
  await page.waitForSelector('.rp-title:has-text("試算結果")', { timeout: 15000 });
  await waitUntil(
    async () => (await page.locator('[data-testid="preview-details"] li').count()) > 0,
    '試算の内訳が出ない',
  );
  const where = await page.locator('[data-testid="preview-details"] .where').allTextContents();
  assert(where.length > 0, '内訳が空');
  assert(
    where.every((w) => w.includes('tokyo')),
    `tokyo 以外が対象に入っている: ${where.join(' | ')}`,
  );
  // サブフォルダーを含めて 1 ブック × 3 シートが対象
  assert(where.length === 3, `対象シート数が ${where.length}: ${where.join(' | ')}`);
});

await check('最上位を選ぶと全ブックが対象になる', async () => {
  await page.selectOption('[data-testid="scope-folder"]', '');
  await page.click('.rbtn-lg:has-text("変更内容を")');
  await waitUntil(async () => {
    const w = await page.locator('[data-testid="preview-details"] .where').allTextContents();
    return w.some((x) => x.includes('osaka')) && w.some((x) => x.includes('tokyo'));
  }, '最上位を選んでも全ブックが対象にならない');
});

await check('文字列の置換もフォルダー指定の中だけに効く', async () => {
  // 「文字列の置換」はリボンの右端にあり独立して見えるが、
  // 年度更新と同じ適用先が使われる。それが画面から分かることも確かめる。
  const tokyoPath = await page
    .locator('[data-testid="scope-folder"] option')
    .evaluateAll((els) => els.map((e) => e.value).find((v) => v.endsWith('tokyo')));
  await page.selectOption('[data-testid="scope-folder"]', tokyoPath);
  const badges = await page.locator('.rgroup:has-text("文字列の置換") .scope-badge').allTextContents();
  assert(badges.length === 1, '文字列の置換に適用先の表示がない');
  assert(/tokyo/.test(badges[0]), `適用先の表示が想定外: ${badges[0]}`);
  assert(/同じ/.test(badges[0]), '適用先が共通である旨の説明がない');

  await page.fill('.rgroup:has-text("文字列の置換") input[placeholder*="令和6"]', '原価管理表');
  await page.fill('.rgroup:has-text("文字列の置換") input[placeholder*="令和7"]', '原価集計表');
  await page.click('.rgroup:has-text("文字列の置換") .rbtn:has-text("試算")');
  await waitUntil(async () => {
    const w = await page.locator('[data-testid="preview-details"] .where').allTextContents();
    return w.length > 0;
  }, '文字列置換の試算結果が出ない');
  const where = await page.locator('[data-testid="preview-details"] .where').allTextContents();
  assert(
    where.every((w) => w.includes('tokyo')),
    `フォルダー外まで対象になっている: ${where.join(' | ')}`,
  );
});

await check('フォルダー指定を解除して元に戻せる', async () => {
  await page.click('.scope-banner .icon-btn');
  await waitUntil(
    async () => (await page.locator('.scope-banner').count()) === 0,
    'フォルダー指定が解除されない',
  );
  // 以降のテストのため、適用先を既定に戻す
  await page.selectOption('[data-testid="scope-books"]', 'current');
  await page.selectOption('[data-testid="scope-sheets"]', 'current');
  await page.fill('.rgroup:has-text("文字列の置換") input[placeholder*="令和6"]', '');
  await page.fill('.rgroup:has-text("文字列の置換") input[placeholder*="令和7"]', '');
  await page.click('.rp-title:has-text("試算結果") .icon-btn');
});

await check('適用先はステータスバーにも常に出る', async () => {
  const status = await page.textContent('.statusbar');
  assert(/適用先/.test(status), `ステータスバー: ${status}`);
  assert(/このブック/.test(status), `解除後の適用先が想定外: ${status}`);
});

console.log('\n\x1b[1m複数ファイルの選択\x1b[0m');

await check('Shift クリックで範囲選択できる', async () => {
  const files = page.locator('.tree-file');
  await files.first().click();
  await files.nth(2).click({ modifiers: ['Shift'] });
  await waitUntil(
    async () => (await page.locator('.scope-banner').count()) > 0,
    '選択の表示が出ない',
  );
  const banner = await page.textContent('.scope-banner');
  assert(/選択した 3 ブック/.test(banner), `表示: ${banner}`);
  assert((await page.locator('.tree-file.picked').count()) === 3, '目印が 3 件でない');
});

await check('選んだブックだけが処理の対象になる', async () => {
  await page.click('.ribbon-tab:has-text("年度更新")');
  assert(
    (await page.inputValue('[data-testid="scope-books"]')) === 'selected',
    '適用先が「選んだブック」になっていない',
  );
  await page.selectOption('[data-testid="scope-sheets"]', 'all');
  await page.click('.rbtn-lg:has-text("変更内容を")');
  await waitUntil(
    async () => (await page.locator('[data-testid="preview-details"] li').count()) > 0,
    '試算が出ない',
  );
  const where = await page.locator('[data-testid="preview-details"] .where').allTextContents();
  const books = new Set(where.map((w) => w.split(' / ')[0]));
  assert(books.size === 3, `対象ブック数が ${books.size}: ${[...books].join(', ')}`);
});

await check('Ctrl クリックで 1 件ずつ外せる', async () => {
  await page.locator('.tree-file').nth(1).click({ modifiers: ['Control'] });
  await waitUntil(
    async () => (await page.locator('.tree-file.picked').count()) === 2,
    '選択を外せない',
  );
  // 後片付け
  await page.locator('.tree-file').first().click();
  await page.click('.ribbon-tab:has-text("年度更新")');
  await page.selectOption('[data-testid="scope-books"]', 'current');
  await page.selectOption('[data-testid="scope-sheets"]', 'current');
  await page.click('.rp-title:has-text("試算結果") .icon-btn');
});

console.log('\n\x1b[1m色からロックを設定\x1b[0m');

await check('使われている色の一覧が出る', async () => {
  await page.click('.ribbon-tab:has-text("書式")');
  await page.click('.rbtn-lg:has-text("色から")');
  await page.waitForSelector('.modal:has-text("色からロックを設定")');
  await waitUntil(
    async () => (await page.locator('.modal [data-color]').count()) > 0,
    '色の一覧が出ない',
  );
  const keys = await page.locator('.modal [data-color]').evaluateAll((els) =>
    els.map((e) => e.getAttribute('data-color')),
  );
  // サンプルは見出しが薄い青、予算欄が黄色
  assert(keys.includes('argb:FFFFFF00'), `黄色が見つからない: ${keys.join(', ')}`);
  assert(keys.includes('argb:FFD9E1F2'), `見出し色が見つからない: ${keys.join(', ')}`);
});

await check('選んだ色のセルだけロックを解除できる', async () => {
  // 黄色 = 予算欄 C5:C9 の 5 セル
  await page.click('.modal [data-color="argb:FFFFFF00"]');
  await page.locator('.modal .check:has-text("🔓 ロック解除する") input[type="radio"]').check();
  await page.click('.modal-foot .rbtn:has-text("試算")');
  await waitUntil(
    async () => (await page.locator('.note-box:has-text("試算")').count()) > 0,
    '試算結果が出ない',
  );
  const preview = await page.locator('.note-box:has-text("試算")').textContent();
  assert(/5 箇所/.test(preview), `試算の件数が想定外: ${preview}`);

  await page.click('.modal-foot .rbtn.accent');
  await waitUntil(async () => (await page.locator('.modal').count()) === 0, 'ダイアログが閉じない');
  await waitUntil(
    async () => (await page.locator('.cell.ov-unlocked').count()) === 5,
    'ロック解除されたセルが 5 つにならない',
  );
});

await check('手順書に色の指定が記録される', async () => {
  await page.click('.ribbon-tab:has-text("手順書")');
  await waitUntil(
    async () => (await page.locator('.step-item:has-text("塗られているセル")').count()) > 0,
    '色の手順が記録されていない',
  );
  const body = await page.locator('.step-item:has-text("塗られているセル")').last().textContent();
  assert(/ロック解除する/.test(body), `手順の説明が想定外: ${body}`);
});

console.log('\n\x1b[1mセルの選択とロック\x1b[0m');

await check('セルをクリックして選択できる', async () => {
  const cell = page.locator('.cell', { hasText: '材料費' }).first();
  await cell.click();
  await page.waitForSelector('.selection-rect');
  const nameBox = await page.inputValue('.namebox input');
  assert(nameBox === 'A5', `選択セルが ${nameBox}`);
});

await check('数式バーに値が出る', async () => {
  const v = await page.inputValue('.formula-input');
  assert(v === '材料費', `数式バーが「${v}」`);
});

await check('ドラッグで範囲選択できる', async () => {
  const from = await page.locator('.cell', { hasText: '1,100,000' }).first().boundingBox();
  const to = await page.locator('.cell', { hasText: '5,500,000' }).first().boundingBox();
  await page.mouse.move(from.x + 5, from.y + 5);
  await page.mouse.down();
  await page.mouse.move(to.x + 5, to.y + 5, { steps: 6 });
  await page.mouse.up();
  const nameBox = await page.inputValue('.namebox input');
  assert(nameBox.includes(':'), `範囲になっていない: ${nameBox}`);
});

await check('「選択範囲以外をロック」を実行できる', async () => {
  await page.click('.ribbon-tab:has-text("ロック")');
  await page.click('.rbtn-lg:has-text("選択範囲以外を")');
  // 直近のトースト (読み込み完了の通知がまだ残っていることがある)
  await waitUntil(
    async () => /セル|変更/.test(await page.locator('.toast').last().textContent()),
    'ロックの実行結果が表示されない',
  );
});

await check('ロック解除セルが黄色で可視化される', async () => {
  const n = await page.locator('.cell.ov-unlocked').count();
  assert(n > 0, 'ロック解除セルの網掛けが見えない');
});

await check('ロック操作だけでシート保護まで自動でかかる', async () => {
  // 「ロック操作のあと自動で保護する」が既定で ON のため、
  // シート保護のボタンを押さなくても保護済みになるはず
  await waitUntil(
    async () => (await page.textContent('.statusbar')).includes('保護あり'),
    'ロック操作だけではシート保護がかからない',
  );
});

await check('自動保護は手順書にも 2 手順として残る', async () => {
  await page.click('.ribbon-tab:has-text("手順書")');
  await waitUntil(
    async () => (await page.locator('.step-item:has-text("シートの保護を有効")').count()) > 0,
    'シート保護の手順が記録されていない',
  );
  await page.click('.ribbon-tab:has-text("ロック")');
});

await page.screenshot({ path: join(SHOTS, '02-locked.png') });

console.log('\n\x1b[1m色分け\x1b[0m');

await check('「範囲を塗る」の項目は無くなっている', async () => {
  await page.click('.ribbon-tab:has-text("書式")');
  const panel = await page.textContent('.ribbon-panel');
  assert(!panel.includes('範囲を塗る'), '「範囲を塗る」が残っている');
  assert(!panel.includes('選択範囲を'), '「選択範囲を塗りつぶす」が残っている');
  assert(panel.includes('塗りつぶしの色'), '色のパレットが消えている');
});

await check('セルを選んで色を押すと、その場で塗られる', async () => {
  // A5 (材料費) を選ぶ
  await page.fill('.namebox input', 'A5');
  await page.press('.namebox input', 'Enter');
  await waitUntil(
    async () => (await page.inputValue('.formula-input')) === '材料費',
    'A5 が選べない',
  );
  await page.click('.ribbon-tab:has-text("書式")');
  // パレットの「赤」を押す
  await page.click('.swatches .swatch[title="赤"]');
  await waitUntil(async () => {
    const cell = page.locator('.cell', { hasText: '材料費' }).first();
    const bg = await cell.evaluate((el) => getComputedStyle(el).backgroundColor);
    return bg === 'rgb(255, 0, 0)';
  }, '色を押してもセルが塗られない');
});

await check('「色を消す」を押すと塗りつぶしが解除される', async () => {
  // 斜線の四角では意味が伝わらないため、文字のボタンにしてある
  const label = await page.textContent('.swatches .clear-fill');
  assert(/色を消す/.test(label), `ボタンの表記: ${label}`);
  await page.click('.swatches .clear-fill');
  await waitUntil(async () => {
    const cell = page.locator('.cell', { hasText: '材料費' }).first();
    const bg = await cell.evaluate((el) => getComputedStyle(el).backgroundColor);
    return bg !== 'rgb(255, 0, 0)';
  }, '塗りなしを押しても解除されない');
});

await check('セルを選んでいないときは色を選ぶだけで塗らない', async () => {
  // シートを切り替えて選択を解除する
  await page.click('.sheettab:has-text("集計")');
  await waitUntil(async () => (await page.inputValue('.namebox input')) === '', '選択が解除されない');
  const before = await page.locator('.cell[style*="background"]').count();
  await page.click('.swatches .swatch[title="黄"]');
  await waitUntil(
    async () => (await page.locator('.toast').last().textContent()).includes('色を選びました'),
    '色を選んだ案内が出ない',
  );
  assert(
    (await page.locator('.cell[style*="background"]').count()) === before,
    '選択が無いのに塗られている',
  );
  await page.click('.sheettab:has-text("2024年度")');
  await waitUntil(async () => (await page.locator('.cell').count()) > 0, 'シートが戻らない');
});

await check('ロック解除セルを一括で塗れる', async () => {
  await page.click('.ribbon-tab:has-text("書式")');
  await page.click('.rbtn-lg:has-text("ロック解除セル")');
  await waitUntil(
    async () => (await page.locator('.cell[style*="background"]').count()) > 0,
    '塗られたセルが見つからない',
  );
});

await page.screenshot({ path: join(SHOTS, '03-colored.png') });

console.log('\n\x1b[1m条件を指定して塗る\x1b[0m');

await check('数値が入っているセルだけを塗れる', async () => {
  await page.click('.ribbon-tab:has-text("書式")');
  await page.click('.rbtn-lg:has-text("条件を指定して")');
  await page.waitForSelector('.modal:has-text("条件を指定して")');
  await page.selectOption('[data-testid="cond-kind"]', 'number');
  await page.click('.modal-foot .rbtn:has-text("試算")');
  await waitUntil(
    async () => (await page.locator('.note-box:has-text("試算")').count()) > 0,
    '試算が出ない',
  );
  await page.click('.modal-foot .rbtn.accent');
  await waitUntil(async () => (await page.locator('.modal').count()) === 0, '閉じない');

  // 金額のセルが塗られ、費目名は塗られていないこと
  // 既定の色 FFFFF2CC = rgb(255, 242, 204)。塗っていないセルは白。
  await waitUntil(async () => (await bg(page, '1,000,000')) === 'rgb(255, 242, 204)', '数値が塗られない');
  assert((await bg(page, '材料費')) === 'rgb(255, 255, 255)', '文字まで塗られている');
});

await check('金額の大きさで絞り込める', async () => {
  await page.click('.rbtn-lg:has-text("条件を指定して")');
  await page.waitForSelector('.modal');
  await page.selectOption('[data-testid="cond-kind"]', 'number');
  await page.locator('.modal .check:has-text("値の大きさ") input').check();
  await page.fill('[data-testid="cond-num-a"]', '4000000');
  await page.selectOption('[data-testid="cond-num-op"]', 'gt');
  await page.click('.modal .swatches .swatch[title="赤"]');
  await page.click('.modal-foot .rbtn.accent');
  await waitUntil(async () => (await page.locator('.modal').count()) === 0, '閉じない');

  await waitUntil(async () => (await bg(page, '5,000,000')) === 'rgb(255, 0, 0)', '400万超が赤くない');
  assert((await bg(page, '1,000,000')) !== 'rgb(255, 0, 0)', '400万以下まで赤い');
});

await check('条件が手順書に記録される', async () => {
  await page.click('.ribbon-tab:has-text("手順書")');
  await waitUntil(
    async () => (await page.locator('.step-item:has-text("数値が入っているセル")').count()) > 0,
    '条件の手順が記録されていない',
  );
  const body = await page.locator('.step-item:has-text("数値が入っているセル")').last().textContent();
  assert(/4,000,000/.test(body), `条件が残っていない: ${body}`);
});

console.log('\n\x1b[1m年度更新\x1b[0m');

await check('全ブック・全シートを対象にできる', async () => {
  await page.click('.ribbon-tab:has-text("年度更新")');
  await page.selectOption('[data-testid="scope-books"]', 'all');
  await page.selectOption('[data-testid="scope-sheets"]', 'all');
  assert((await page.inputValue('[data-testid="scope-books"]')) === 'all', 'ブックの指定が反映されない');
  assert((await page.inputValue('[data-testid="scope-sheets"]')) === 'all', 'シートの指定が反映されない');
});

await check('試算では値が変わらない', async () => {
  await page.click('.rbtn-lg:has-text("変更内容を")');
  await page.waitForSelector('.rp-title:has-text("試算結果")', { timeout: 10000 });
  const text = await page.textContent('.grid-canvas');
  assert(text.includes('2024年度'), '試算なのに値が変わっている');
});

await check('年度更新を実行すると 2024 が 2025 になる', async () => {
  await page.click('.rbtn-lg:has-text("年度更新を")');
  await waitUntil(
    async () => (await page.textContent('.grid-canvas')).includes('2025年度'),
    '2025 に更新されない',
  );
  const text = await page.textContent('.grid-canvas');
  assert(!text.includes('2024年度 原価管理表'), '2024 が残っている');
});

await check('シート名も 2025年度 に変わる', async () => {
  const tabs = await page.locator('.sheettab').allTextContents();
  assert(tabs.some((t) => t.includes('2025年度')), `タブ: ${tabs.join(', ')}`);
  assert(tabs.some((t) => t.includes('2024年度')), '前年度シートが 2024年度 になっているはず');
});

await check('他のブックにも適用されている', async () => {
  await page.locator('.tree-file').nth(1).click();
  await waitUntil(
    async () => (await page.textContent('.grid-canvas')).includes('2025年度'),
    '2 冊目が更新されていない',
  );
});

await page.screenshot({ path: join(SHOTS, '04-year-updated.png') });

console.log('\n\x1b[1m手順書\x1b[0m');

await check('操作が手順として記録されている', async () => {
  await page.click('.ribbon-tab:has-text("手順書")');
  await page.waitForSelector('.step-item');
  const n = await page.locator('.step-item').count();
  assert(n >= 4, `記録された手順が ${n} 件`);
});

// 注: このコンテナはロケールの都合で download 属性の日本語ファイル名を
//     落としてしまう (実機の Windows / Chrome では問題ない)。
//     そのためファイル名ではなく中身を検証する。
await check('手順書 (JSON) を書き出せて、再実行できる形式になっている', async () => {
  const dl = page.waitForEvent('download', { timeout: 15000 });
  await page.click('.rbtn-lg:has-text("JSON")');
  const download = await dl;
  const path = join(root, '.test-build', 'recipe.json');
  await download.saveAs(path);
  const recipe = JSON.parse(readFileSync(path, 'utf8'));
  assert(recipe.format === 'excel-lock-manager/recipe', `format: ${recipe.format}`);
  assert(recipe.steps.length >= 4, `手順数: ${recipe.steps.length}`);
  const ops = recipe.steps.map((s) => s.body.op);
  assert(ops.includes('lockAllExcept'), `ops: ${ops.join(',')}`);
  assert(ops.includes('protectSheet'), `ops: ${ops.join(',')}`);
  assert(ops.includes('shiftYears'), `ops: ${ops.join(',')}`);
  const protect = recipe.steps.find((s) => s.body.op === 'protectSheet');
  assert(!protect.body.password, 'パスワードが手順書に保存されてしまっている');
});

await check('手順書 (HTML) を書き出せて、読める内容になっている', async () => {
  const dl = page.waitForEvent('download', { timeout: 15000 });
  await page.click('.rbtn-lg:has-text("HTML")');
  const download = await dl;
  const path = join(root, '.test-build', 'tejunsho.html');
  await download.saveAs(path);
  const html = readFileSync(path, 'utf8');
  assert(html.includes('<!doctype html>'), 'HTML になっていない');
  assert(html.includes('作業手順'), '見出しがない');
  assert(html.includes('ロック'), '操作内容が書かれていない');
  assert(html.includes("connect-src 'none'"), '手順書側にも通信遮断の指定が必要');
});

await page.screenshot({ path: join(SHOTS, '05-recipe.png') });

console.log('\n\x1b[1m保存\x1b[0m');

await check('ZIP で保存でき、中身が更新後の Excel になっている', async () => {
  await page.click('.ribbon-tab:has-text("ファイル")');
  await page.click('.rbtn-lg:has-text("ZIP")');
  await page.waitForSelector('.modal');
  const dl = page.waitForEvent('download', { timeout: 60000 });
  await page.click('.modal-foot .rbtn.accent');
  const download = await dl;
  const path = join(root, '.test-build', 'out.zip');
  await download.saveAs(path);

  // ZIP を展開して、実際に Excel として読めるか / 年度が更新されているかを見る
  const zip = await JSZip.loadAsync(readFileSync(path));
  const names = Object.keys(zip.files).filter((n) => n.endsWith('.xlsx'));
  assert(names.length === 3, `ZIP 内の xlsx が ${names.length} 件`);
  assert(
    names.some((n) => n.includes('tokyo/')) && names.some((n) => n.includes('osaka/')),
    `フォルダー構成が保たれていない: ${names.join(', ')}`,
  );

  // 3 冊すべてで年度が更新されていること
  for (const name of names) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(await zip.file(name).async('nodebuffer'));
    const ws = wb.getWorksheet('2025年度');
    assert(ws, `${name}: シート名が更新されていない`);
    assert(
      String(ws.getCell('A1').value).includes('2025年度'),
      `${name}: A1 が「${ws.getCell('A1').value}」`,
    );
    // 数式は年を含まないので変わらないこと
    assert(
      ws.getCell('B10').value?.formula === 'SUM(B5:B9)',
      `${name}: 合計の数式が壊れている (${JSON.stringify(ws.getCell('B10').value)})`,
    );
  }

  // 画面で操作した 1 冊にロック / 保護 / 色が保存されていること
  // (フォルダー読み込み直後にどのブックが選ばれるかは読み込み順によるため、
  //  「どれか 1 冊に入っていること」を確認する)
  let operated = 0;
  for (const name of names) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(await zip.file(name).async('nodebuffer'));
    const ws = wb.getWorksheet('2025年度');
    if (!ws.sheetProtection) continue;
    const unlocked = [];
    ws.eachRow((row) =>
      row.eachCell((cell) => {
        if (cell.protection?.locked === false) unlocked.push(cell);
      }),
    );
    assert(unlocked.length > 0, `${name}: 保護済みなのにロック解除セルが 1 つもない`);
    assert(
      unlocked.every((c) => c.fill?.fgColor?.argb),
      `${name}: ロック解除セルに色が付いていない`,
    );
    assert(
      ws.getCell('A5').protection?.locked !== false,
      `${name}: 見出しがロックされていない`,
    );
    operated++;
  }
  assert(operated === 1, `シート保護が保存されたブックが ${operated} 冊`);
});

console.log('\n\x1b[1m2 年分を見比べて記入欄を判定\x1b[0m');

await check('2 年分のフォルダーを読み込める', async () => {
  await page.click('.ribbon-tab:has-text("ファイル")');
  await closeAllBooks(page);
  // 判定結果 (様式をロック) が分かるよう、全解除の状態から始める
  await page.selectOption('[data-testid="initial-lock"]', 'unlock');
  await page.setInputFiles('input[webkitdirectory]', YEARS);
  await waitUntil(async () => (await page.locator('.tree-file').count()) === 2, '2 年分が読み込めない');

  // 2 冊を見比べる操作なので、「選択中のブックのみ」のままでは組が作れない。
  // わざと 1 冊だけの指定に戻し、開いたときに広がることを次で確かめる。
  await page.click('.ribbon-tab:has-text("書式")');
  await page.selectOption('.ribbon-panel [data-testid="scope-books"]', 'current');
  await page.selectOption('.ribbon-panel [data-testid="scope-sheets"]', 'current');
});

await check('試算で「毎年書き換わる欄」が数えられる', async () => {
  await page.click('.ribbon-tab:has-text("書式")');
  await page.click('.rbtn-lg:has-text("2 年分を見比べて")');
  await page.waitForSelector('.modal:has-text("2 年分を見比べて")');
  // 開いた時点で、読み込んだ全ブック・全シートが対象になっていること
  assert(
    (await page.inputValue('.modal [data-testid="scope-books"]')) === 'all',
    '適用先が全ブックに広がらない',
  );
  assert(
    (await page.inputValue('.modal [data-testid="scope-sheets"]')) === 'all',
    '適用先が全シートに広がらない',
  );
  await page.click('.modal-foot .rbtn:has-text("判定してみる")');
  await page.waitForSelector('[data-testid="detect-preview"]', { timeout: 20000 });
  const text = await page.textContent('[data-testid="detect-preview"]');
  assert(/1 組/.test(text), `組数が合わない: ${text}`);
  // 予算額の 5 セルだけが毎年書き換わっている
  assert(/記入欄 5 セル/.test(text), `判定結果: ${text}`);
});

await check('判定した根拠 (変化した値) が例として出る', async () => {
  const samples = await page.textContent('[data-testid="detect-samples"]');
  assert(/B4/.test(samples), `例に B4 が出ない: ${samples}`);
  assert(/3,?400,?000/.test(samples), `前年の値が出ない: ${samples}`);
  assert(/3,?500,?000/.test(samples), `今年の値が出ない: ${samples}`);
  // 年の数字だけ違う表題や、数式の合計欄は記入欄にしない
  assert(!/予算入力表/.test(samples), `表題を記入欄と判定している: ${samples}`);
  assert(!/A9|B9/.test(samples), `合計欄を記入欄と判定している: ${samples}`);
});

await check('記入欄が塗られ、様式はロックされる', async () => {
  await page.click('.modal-foot .rbtn.accent');
  await waitUntil(async () => (await page.locator('.modal').count()) === 0, '閉じない');

  // 判定は新しい方 (2025) のファイルに反映される
  await page.locator('.tree-file:has-text("2025")').click();
  await waitUntil(
    async () => (await page.textContent('.grid-canvas')).includes('2025年度'),
    '2025 のブックが開けない',
  );

  assert((await bg(page, '3,500,000')) === 'rgb(255, 242, 204)', '記入欄が塗られていない');
  assert((await bg(page, '材料費')) === 'rgb(255, 255, 255)', '様式まで塗られている');
  assert(/ov-locked/.test(await cellClass(page, '材料費')), '様式がロックされていない');
  assert(!/ov-locked/.test(await cellClass(page, '3,500,000')), '記入欄がロックされている');
});

await check('古い方のファイルは変更されない', async () => {
  await page.locator('.tree-file:has-text("2024")').click();
  await waitUntil(
    async () => (await page.textContent('.grid-canvas')).includes('2024年度'),
    '2024 のブックが開けない',
  );
  assert((await bg(page, '3,400,000')) === 'rgb(255, 255, 255)', '前年のファイルまで塗られている');
});

await check('判定が手順書に記録される', async () => {
  await page.click('.ribbon-tab:has-text("手順書")');
  // 手順名は入力欄なので、本文 (説明) の方で探す
  await waitUntil(
    async () => (await page.locator('.step-item:has-text("年違いの同じ様式")').count()) > 0,
    '判定の手順が記録されていない',
  );
  const body = await page.locator('.step-item:has-text("年違いの同じ様式")').last().textContent();
  assert(/毎年記入されている欄/.test(body), `判定の内容が残っていない: ${body}`);
  assert(/様式.*ロック/.test(body), `様式のロックが残っていない: ${body}`);
});

await page.screenshot({ path: join(SHOTS, '05b-detect.png') });

console.log('\n\x1b[1mオフライン利用 (ツール本体の保存)\x1b[0m');

// 共有フォルダーを模したパス (日本語・空白・入れ子)。
// 社内の共有フォルダーはこういう名前になることが多いため、そのまま試す。
const SHARE = join(root, '.test-build', '社内共有', '原価管理 共通', '01 ツール');
let sharedToolPath = null;

await check('ツール本体をダブルクリックできる .html として保存できる', async () => {
  await page.click('.ribbon-tab:has-text("ファイル")');
  const before = requests.length;
  const dl = page.waitForEvent('download', { timeout: 15000 });
  await page.click('.rbtn-lg:has-text("ツール本体を")');
  const download = await dl;

  const name = download.suggestedFilename();
  assert(name.endsWith('.html'), `保存名に .html が付かない: ${name}`);

  mkdirSync(SHARE, { recursive: true });
  sharedToolPath = join(SHARE, name);
  await download.saveAs(sharedToolPath);

  // 保存自体がメモリー上の DOM から作られ、通信を伴わないこと
  assert(requests.length === before, 'ツール保存で通信が発生している');

  const html = readFileSync(sharedToolPath, 'utf8');
  assert(html.startsWith('<!doctype html>'), 'DOCTYPE がない');
  assert(html.includes("connect-src 'none'"), 'CSP が失われている');
  // 外部から読み込む資源が 1 つも無いこと (共有フォルダーでは取りに行けない)
  assert(!/<script[^>]+src=/i.test(html), '外部スクリプトの読み込みが残っている');
  assert(!/<link[^>]+href="http/i.test(html), '外部スタイルの読み込みが残っている');
});

await check('共有フォルダーから、通信できない状態でも単体で動く', async () => {
  // ネットワークを完全に遮断した状態で開く。
  // 社内で LAN から切り離しても動くことの確認。
  const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } });
  await ctx.setOffline(true);
  const fresh = await ctx.newPage();
  const freshRequests = [];
  fresh.on('request', (r) => {
    if (!r.url().startsWith('file://')) freshRequests.push(r.url());
  });
  const freshErrors = [];
  fresh.on('pageerror', (e) => freshErrors.push(e.message));

  await fresh.goto(`file://${sharedToolPath}`);
  await fresh.waitForSelector('.app', { timeout: 15000 });

  // 保存直後の状態 (ファイル未読み込み) で開けること
  assert(await fresh.isVisible('.grid-placeholder'), '保存したファイルが初期状態で開かない');
  // ローカル起動として認識され、赤い警告が出ないこと
  const badge = await fresh.textContent('[data-testid="guard-badge"]');
  assert(badge.includes('完全オフライン'), `タイトルバー: ${badge}`);

  // 保存したファイルでも Excel が読めること
  await fresh.setInputFiles('input[webkitdirectory]', SAMPLE);
  await fresh.waitForSelector('.cell', { timeout: 20000 });
  const text = await fresh.textContent('.grid-canvas');
  assert(text.includes('原価管理表'), '保存したファイルで Excel を読み込めない');
  assert(freshErrors.length === 0, `保存したファイルで JS エラー: ${freshErrors[0]}`);
  assert(freshRequests.length === 0, `保存したファイルが通信している: ${freshRequests[0]}`);
  await ctx.close();
});

console.log('\n\x1b[1m外部通信の遮断\x1b[0m');

await check('外部へのリクエストが 1 件も発生していない', () => {
  assert(requests.length === 0, `${requests.length} 件発生: ${requests.slice(0, 5).join(', ')}`);
});

await check('fetch が使用不能になっている', async () => {
  const r = await page.evaluate(async () => {
    try {
      await fetch('https://example.com');
      return 'ALLOWED';
    } catch (e) {
      return e.name;
    }
  });
  assert(r === 'NetworkBlockedError', `fetch の結果: ${r}`);
});

await check('XMLHttpRequest が使用不能になっている', async () => {
  const r = await page.evaluate(() => {
    try {
      const x = new XMLHttpRequest();
      x.open('GET', 'https://example.com');
      return 'ALLOWED';
    } catch (e) {
      return e.name;
    }
  });
  assert(r === 'NetworkBlockedError', `XHR の結果: ${r}`);
});

await check('WebSocket が使用不能になっている', async () => {
  const r = await page.evaluate(() => {
    try {
      new WebSocket('wss://example.com');
      return 'ALLOWED';
    } catch (e) {
      return e.name;
    }
  });
  assert(r === 'NetworkBlockedError', `WebSocket の結果: ${r}`);
});

await check('遮断がステータスバーとタイトルバーに表示される', async () => {
  const text = await page.textContent('.statusbar');
  assert(text.includes('遮断'), `ステータスバー: ${text}`);
  const badge = await page.textContent('[data-testid="guard-badge"]');
  assert(badge.includes('遮断'), `タイトルバー: ${badge}`);
});

await check('セキュリティタブは無い', async () => {
  const tabs = await page.locator('.ribbon-tab').allTextContents();
  assert(!tabs.some((t) => t.includes('セキュリティ')), `タブ: ${tabs.join(', ')}`);
  assert(tabs.length === 5, `タブ数: ${tabs.length}`);
});

await check('どのタブに切り替えてもリボンの高さが変わらない', async () => {
  const tabs = await page.locator('.ribbon-tab').allTextContents();
  const heights = [];
  for (const t of tabs) {
    await page.click(`.ribbon-tab:has-text("${t}")`);
    heights.push(await page.locator('.ribbon-panel').evaluate((el) => el.clientHeight));
  }
  assert(
    new Set(heights).size === 1,
    `タブごとに高さが違う: ${tabs.map((t, i) => `${t}=${heights[i]}`).join(', ')}`,
  );
});

await page.screenshot({ path: join(SHOTS, '06-tabs.png') });

// --------------------------------------------------------------------------
// GitHub Pages のようにサーバーから配信した場合の検証。
// ページ本体の読み込み以外に通信が発生しないことを確かめる。
// --------------------------------------------------------------------------
console.log('\n\x1b[1mサーバー配信時 (GitHub Pages を想定)\x1b[0m');

const server = createServer((req, res) => {
  if (req.url === '/favicon.ico') {
    res.writeHead(404).end();
    return;
  }
  res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
  res.end(readFileSync(DIST));
});
await new Promise((r) => server.listen(0, '127.0.0.1', r));
const origin = `http://127.0.0.1:${server.address().port}`;

const hostedPage = await browser.newPage({ viewport: { width: 1600, height: 950 } });
const hostedRequests = [];
hostedPage.on('request', (r) => hostedRequests.push(r.url()));
const hostedErrors = [];
hostedPage.on('pageerror', (e) => hostedErrors.push(e.message));
await hostedPage.goto(`${origin}/`);
await hostedPage.waitForSelector('.app');

await check('サーバー配信であることが画面に表示される', async () => {
  const badge = await hostedPage.textContent('[data-testid="guard-badge"]');
  assert(badge.includes('外部に出ません'), `バッジ: ${badge}`);
  // 起動元の説明とダウンロードの案内は「ファイル」タブに出る
  await hostedPage.click('.ribbon-tab:has-text("ファイル")');
  const body = await hostedPage.textContent('.ribbon-panel');
  assert(body.includes('サーバーから開いています'), `起動元の案内がない: ${body.slice(0, 200)}`);
  assert(body.includes('ツール本体を保存'), 'オフライン保存の案内がない');
});

await check('ページ本体以外に通信が発生しない', async () => {
  const baseline = hostedRequests.length;
  await hostedPage.click('.ribbon-tab:has-text("ファイル")');
  await hostedPage.setInputFiles('input[webkitdirectory]', SAMPLE);
  await hostedPage.waitForSelector('.cell', { timeout: 20000 });
  await hostedPage.click('.ribbon-tab:has-text("年度更新")');
  await hostedPage.click('.rbtn-lg:has-text("年度更新を")');
  await waitUntil(
    async () => (await hostedPage.textContent('.grid-canvas')).includes('2025年度'),
    'サーバー配信時に年度更新が反映されない',
  );
  const added = hostedRequests.slice(baseline);
  assert(added.length === 0, `操作中に ${added.length} 件の通信: ${added.join(', ')}`);
  // 最初の読み込みもドキュメント 1 件だけ (全て単一 HTML に同梱されているため)
  assert(baseline === 1, `初期読み込みが ${baseline} 件 (期待: 1 件): ${hostedRequests.join(', ')}`);
});

await check('ブラウザーが favicon を探しに行かない', async () => {
  // アイコンの指定が無いと /favicon.ico が要求され、CSP に引っかかって
  // 「遮断しました」と赤く出てしまう。データ送信ではないので紛らわしい。
  const hasIcon = await hostedPage.locator('link[rel="icon"]').count();
  assert(hasIcon > 0, 'favicon の指定が無い');
  const iconHref = await hostedPage.locator('link[rel="icon"]').getAttribute('href');
  assert(iconHref.startsWith('data:'), `アイコンが外部参照になっている: ${iconHref.slice(0, 40)}`);
  assert(
    !hostedRequests.some((u) => u.includes('favicon')),
    `favicon が要求されている: ${hostedRequests.join(', ')}`,
  );
});

await check('資源の読み込みが止まっただけなら警告色にしない', async () => {
  // 拡張機能などが画像を差し込むと CSP に引っかかるが、
  // データが出ようとしたわけではないので赤くしない
  await hostedPage.evaluate(() => {
    const img = document.createElement('img');
    img.src = '/blocked-by-csp.png';
    document.body.appendChild(img);
  });
  await new Promise((r) => setTimeout(r, 500));
  const badge = await hostedPage.textContent('[data-testid="guard-badge"]');
  assert(!badge.includes('遮断'), `資源の遮断で警告色になっている: ${badge}`);
  const tip = await hostedPage.locator('[data-testid="guard-badge"]').getAttribute('title');
  assert(/外部リソースの読み込みを/.test(tip), `説明が出ていない: ${tip}`);
});

await check('本当の送信は遮断して警告する', async () => {
  await hostedPage.evaluate(() => {
    try {
      // eslint-disable-next-line no-undef
      fetch('https://example.com/leak');
    } catch {
      /* 遮断される */
    }
  });
  await waitUntil(
    async () => (await hostedPage.textContent('[data-testid="guard-badge"]')).includes('送信を'),
    '送信の遮断が警告として出ない',
  );
  const tip = await hostedPage.locator('[data-testid="guard-badge"]').getAttribute('title');
  assert(/送ろうとした試み/.test(tip), `説明が出ていない: ${tip}`);
  assert(/example\.com/.test(tip), `宛先が出ていない: ${tip}`);
});

await check('サーバー配信でも Excel を処理できる', async () => {
  const text = await hostedPage.textContent('.grid-canvas');
  assert(text.includes('2025年度'), '年度更新が効いていない');
  // 上のテストで意図的に起こした遮断は除く
  const real = hostedErrors.filter((e) => !e.includes('本ツールでは使用できません'));
  assert(real.length === 0, `JS エラー: ${real[0]}`);
});

await hostedPage.screenshot({ path: join(SHOTS, '07-hosted.png') });
await hostedPage.close();
server.close();

await check('JS エラーが出ていない', () => {
  const real = consoleErrors.filter((e) => !e.includes('外部通信ガード') && !e.includes('NetworkBlocked'));
  assert(real.length === 0, real.slice(0, 3).join(' | '));
});

await browser.close();

console.log(`\n${'─'.repeat(60)}`);
if (failures.length) {
  console.log(`\x1b[31m${failures.length} 件失敗 / ${passed} 件成功\x1b[0m\n`);
  for (const f of failures) console.log(`  \x1b[31m✗\x1b[0m ${f}\n`);
  process.exit(1);
}
console.log(`\x1b[32m${passed} 件すべて成功\x1b[0m  (スクリーンショット: .test-build/shots/)`);
