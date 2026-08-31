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

mkdirSync(SHOTS, { recursive: true });
await page.goto(`file://${DIST}`);
await page.waitForSelector('.app');

console.log('\n\x1b[1m画面の初期表示\x1b[0m');

await check('タイトルバーにオフライン表示が出る', async () => {
  const text = await page.textContent('.offline-badge');
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

await check('シート保護を有効化できる', async () => {
  await page.click('.rbtn-lg:has-text("シート保護を")');
  await waitUntil(
    async () => (await page.textContent('.statusbar')).includes('保護あり'),
    'シート保護がステータスバーに反映されない',
  );
});

await page.screenshot({ path: join(SHOTS, '02-locked.png') });

console.log('\n\x1b[1m色分け\x1b[0m');

await check('ロック解除セルを一括で塗れる', async () => {
  await page.click('.ribbon-tab:has-text("書式")');
  await page.click('.rbtn-lg:has-text("ロック解除セル")');
  await waitUntil(
    async () => (await page.locator('.cell[style*="background"]').count()) > 0,
    '塗られたセルが見つからない',
  );
});

await page.screenshot({ path: join(SHOTS, '03-colored.png') });

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

console.log('\n\x1b[1mオフライン利用 (ツール本体の保存)\x1b[0m');

await check('ツール本体を保存でき、それ単体で動く', async () => {
  await page.click('.ribbon-tab:has-text("ファイル")');
  const before = requests.length;
  const dl = page.waitForEvent('download', { timeout: 15000 });
  await page.click('.rbtn-lg:has-text("ツール本体を")');
  const download = await dl;
  const savedPath = join(root, '.test-build', 'self-copy.html');
  await download.saveAs(savedPath);

  // 保存自体がメモリー上の DOM から作られ、通信を伴わないこと
  assert(requests.length === before, 'ツール保存で通信が発生している');

  const html = readFileSync(savedPath, 'utf8');
  assert(html.startsWith('<!doctype html>'), 'DOCTYPE がない');
  assert(html.includes("connect-src 'none'"), 'CSP が失われている');

  // 保存したファイルを別のページとして開き、実際に動くか確かめる
  const fresh = await browser.newPage({ viewport: { width: 1400, height: 900 } });
  const freshRequests = [];
  fresh.on('request', (r) => {
    if (!r.url().startsWith('file://')) freshRequests.push(r.url());
  });
  const freshErrors = [];
  fresh.on('pageerror', (e) => freshErrors.push(e.message));
  await fresh.goto(`file://${savedPath}`);
  await fresh.waitForSelector('.app', { timeout: 15000 });

  // 保存直後の状態 (ファイル未読み込み) で開けること
  assert(await fresh.isVisible('.grid-placeholder'), '保存したファイルが初期状態で開かない');

  // 保存したファイルでも Excel が読めること
  await fresh.setInputFiles('input[webkitdirectory]', SAMPLE);
  await fresh.waitForSelector('.cell', { timeout: 20000 });
  const text = await fresh.textContent('.grid-canvas');
  assert(text.includes('原価管理表'), '保存したファイルで Excel を読み込めない');
  assert(freshErrors.length === 0, `保存したファイルで JS エラー: ${freshErrors[0]}`);
  assert(freshRequests.length === 0, `保存したファイルが通信している: ${freshRequests[0]}`);
  await fresh.close();
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

await check('遮断がステータスバーに表示される', async () => {
  const text = await page.textContent('.statusbar');
  assert(text.includes('遮断'), `ステータスバー: ${text}`);
});

await page.click('.ribbon-tab:has-text("セキュリティ")');
await page.waitForSelector('.rgroup-title:has-text("遮断した通信")');
await page.screenshot({ path: join(SHOTS, '06-security.png') });

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
  const badge = await hostedPage.textContent('.offline-badge');
  assert(badge.includes('外部に出ません'), `バッジ: ${badge}`);
  await hostedPage.click('.ribbon-tab:has-text("セキュリティ")');
  const body = await hostedPage.textContent('.ribbon-panel');
  assert(body.includes('起動元: サーバー'), '起動元の表示がない');
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

await check('サーバー配信でも Excel を処理できる', async () => {
  const text = await hostedPage.textContent('.grid-canvas');
  assert(text.includes('2025年度'), '年度更新が効いていない');
  assert(hostedErrors.length === 0, `JS エラー: ${hostedErrors[0]}`);
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
