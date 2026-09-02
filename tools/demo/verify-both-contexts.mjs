/**
 * 動画とまったく同じ手順を、
 *   A: GitHub Pages と同じ「サーバー配信」
 *   B: ダウンロードした HTML を「ダブルクリックで開いた状態 (file://)」
 * の両方で実行し、出来上がった Excel を突き合わせる。
 */
import { chromium } from 'playwright';
import { createServer } from 'node:http';
import { readFileSync, writeFileSync, readdirSync, rmSync, cpSync, mkdirSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { execSync } from 'node:child_process';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';

const ROOT = process.env.REPO_ROOT ?? process.cwd();
const TOOL = join(ROOT, '.test-build', 'video', 'Excel一括ロック_年度更新ツール.html');
const WORK = join(ROOT, '.test-build', 'abwf');
rmSync(WORK, { recursive: true, force: true });
mkdirSync(WORK, { recursive: true });

const html = readFileSync(join(ROOT, 'dist', 'index.html'));
const srv = createServer((_q, res) => { res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' }); res.end(html); });
await new Promise((r) => srv.listen(0, '127.0.0.1', r));
const ORIGIN = `http://127.0.0.1:${srv.address().port}/`;

const browser = await chromium.launch({
  executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome',
  args: ['--no-sandbox','--disable-dev-shm-usage','--no-first-run','--disable-background-networking'],
});

const wait = async (p, f, m, t = 240000) => { const e = Date.now() + t; for (;;) { try { if (await f()) return; } catch {} if (Date.now() > e) throw new Error(m); await new Promise((r) => setTimeout(r, 150)); } };

/** 実フォルダーを本物のまま渡す (OS のダイアログの代わり。以降はアプリの通常経路) */
async function makePage(url, getRoot) {
  const page = await browser.newPage({ viewport: { width: 1600, height: 950 } });
  await page.exposeBinding('__fsList', async (_s, rel) => readdirSync(rel ? join(getRoot(), rel) : getRoot(), { withFileTypes: true }).map((d) => ({ name: d.name, kind: d.isDirectory() ? 'directory' : 'file' })));
  await page.exposeBinding('__fsRead', async (_s, rel) => readFileSync(join(getRoot(), rel)).toString('base64'));
  await page.exposeBinding('__fsWrite', async (_s, rel, b64) => { writeFileSync(join(getRoot(), rel), Buffer.from(b64, 'base64')); return true; });
  await page.exposeBinding('__fsRemove', async (_s, rel) => { rmSync(join(getRoot(), rel), { force: true }); return true; });
  await page.exposeBinding('__fsRootName', async () => getRoot().split('/').pop());
  await page.addInitScript(() => {
    const X = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    const d64 = (b) => Uint8Array.from(atob(b), (c) => c.charCodeAt(0));
    const e64 = (u) => { let s = ''; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); };
    const mf = (n, r) => ({ kind: 'file', name: n,
      async getFile() { return new File([d64(await window.__fsRead(r))], n, { type: X }); },
      async createWritable() { const p = []; return { async write(d) { p.push(d); }, async close() { await window.__fsWrite(r, e64(new Uint8Array(await new Blob(p).arrayBuffer()))); } }; } });
    const md = (n, r) => ({ kind: 'directory', name: n,
      async *values() { for (const it of await window.__fsList(r)) { const c = r ? `${r}/${it.name}` : it.name; yield it.kind === 'directory' ? md(it.name, c) : mf(it.name, c); } },
      async getFileHandle(n2) { return mf(n2, r ? `${r}/${n2}` : n2); },
      async removeEntry(n2) { await window.__fsRemove(r ? `${r}/${n2}` : n2); } });
    window.showDirectoryPicker = async () => md(await window.__fsRootName(), '');
  });
  await page.goto(url);
  await page.waitForSelector('.app');
  return page;
}

/** 動画①: 黄色以外をロック → シート保護 → 上書き保存 */
async function workflow1(page) {
  await page.click('.ribbon-tab:has-text("ファイル")');
  await page.click('.rbtn-lg:has-text("フォルダーを開く")');
  await wait(page, async () => (await page.locator('.tree-file').count()) === 30, '読み込めない');
  await page.click('.ribbon-tab:has-text("書式")');
  await page.selectOption('.ribbon-panel [data-testid="scope-books"]', 'all');
  await page.selectOption('.ribbon-panel [data-testid="scope-sheets"]', 'all');
  await page.click('.rbtn-lg:has-text("色から")');
  await page.waitForSelector('.modal');
  await wait(page, async () => (await page.locator('.modal [data-color]').count()) > 0, '色一覧が出ない');
  const keys = await page.locator('.modal [data-color]').evaluateAll((els) => els.map((e) => e.getAttribute('data-color')));
  await page.click(`.modal [data-color="${keys.find((k) => /FFFF00/i.test(k))}"]`);
  await page.locator('.modal .check:has-text("以外の") input[type="radio"]').check();
  await page.locator('.modal .check:has-text("🔒 ロックする") input[type="radio"]').check();
  await page.click('.modal-foot .rbtn.accent');
  await wait(page, async () => (await page.locator('.modal').count()) === 0, '閉じない');
  await page.click('.ribbon-tab:has-text("ロック")');
  await page.selectOption('.ribbon-panel [data-testid="scope-books"]', 'all');
  await page.selectOption('.ribbon-panel [data-testid="scope-sheets"]', 'all');
  await page.click('.rbtn-lg:has-text("シート保護を有効化")');
  await wait(page, async () => (await page.textContent('.rp-body')).includes('保護'), '保護できない');
  await page.click('.ribbon-tab:has-text("ファイル")');
  await page.click('.rbtn-lg:has-text("元の場所へ")');
  await page.waitForSelector('.modal');
  await page.click('.modal-foot .rbtn.accent');
  await wait(page, async () => /件を上書き|新しい名前で保存/.test(await page.textContent('body')), '上書きできない');
  await new Promise((r) => setTimeout(r, 2500));
}

/** 動画②: 年度を +1 (対象 2023〜2025 / ファイル名も) → 上書き保存 (元は削除) */
async function workflow2(page) {
  await page.click('.ribbon-tab:has-text("ファイル")');
  await page.click('.rbtn-lg:has-text("フォルダーを開く")');
  await wait(page, async () => (await page.locator('.tree-file').count()) === 30, '読み込めない');
  await page.click('.ribbon-tab:has-text("年度更新")');
  await page.selectOption('.ribbon-panel [data-testid="scope-books"]', 'all');
  await page.selectOption('.ribbon-panel [data-testid="scope-sheets"]', 'all');
  await page.locator('.check:has-text("ファイル名") input').check();
  const nums = page.locator('.rgroup:has(.rgroup-title:text-is("置換の方法")) input[type="number"]');
  await nums.nth(1).fill('2023');
  await nums.nth(2).fill('2025');
  await page.click('.rbtn-lg:has-text("年度更新を")');
  await wait(page, async () => (await page.textContent('.grid-canvas')).includes('2026年度'), '年度更新されない');
  await new Promise((r) => setTimeout(r, 1500));
  await page.click('.ribbon-tab:has-text("ファイル")');
  await page.click('.rbtn-lg:has-text("元の場所へ")');
  await page.waitForSelector('.modal');
  const del = page.locator('.modal .check:has-text("元の名前のファイルを削除する") input');
  await del.scrollIntoViewIfNeeded();
  await del.check();
  await page.click('.modal-foot .rbtn.accent');
  await wait(page, async () => /件を上書き|新しい名前で保存/.test(await page.textContent('body')), '上書きできない');
  await new Promise((r) => setTimeout(r, 2500));
}

/** フォルダーの中身を、比較できる形にまとめる */
async function snapshot(dir) {
  const out = {};
  const walk = (d, prefix) => {
    for (const e of readdirSync(d, { withFileTypes: true })) {
      const p = join(d, e.name);
      const rel = prefix ? `${prefix}/${e.name}` : e.name;
      if (e.isDirectory()) walk(p, rel);
      else out[rel] = p;
    }
  };
  walk(dir, '');
  const result = {};
  for (const [rel, p] of Object.entries(out)) {
    const zip = await JSZip.loadAsync(readFileSync(p));
    const parts = {};
    for (const name of Object.keys(zip.files).sort()) {
      if (zip.files[name].dir || /docProps\/core\.xml$/.test(name)) continue;
      parts[name] = await zip.file(name).async('string');
    }
    result[rel] = parts;
  }
  return result;
}

// 前回の収録で処理済みになっているので、毎回まっさらから作り直す
console.log('データを作り直しています…');
execSync(`node ${join(ROOT, '.test-build/video/make-share.cjs')}`, { stdio: 'ignore' });
execSync(`rm -rf /root/共有フォルダー && cp -r "${join(ROOT, '.test-build/video/共有フォルダー')}" /root/`, { stdio: 'ignore' });
execSync(`node ${join(ROOT, '.test-build/video/make-report.cjs')}`, { stdio: 'ignore' });

const run = (cmd) => { try { return execSync(cmd, { encoding: 'utf8' }).trim().split('\n')[0]; } catch (e) { return `失敗: ${e.message.split('\n')[0]}`; } };

const cases = [
  { name: '動画① 黄色以外をロック', src: '/root/共有フォルダー', run: workflow1, inner: '原価管理/2025年度予算',
    verify: (d) => run(`SHARE_DIR=${JSON.stringify(d)} node ${join(ROOT, '.test-build/video/verify-share.cjs')}`) },
  { name: '動画② 年度更新',        src: '/root/報告共有フォルダー', run: workflow2, inner: '2025年度報告',
    verify: (d) => run(`SHARE_DIR=${JSON.stringify(d)} node ${join(ROOT, '.test-build/video/verify-report.cjs')}`) },
];

for (const c of cases) {
  console.log(`\n===== ${c.name} =====`);
  const snaps = {};
  for (const [label, url] of [['A サーバー配信', ORIGIN], ['B ローカル (file://)', `file://${TOOL}`]]) {
    const work = join(WORK, `${c.name.slice(0, 4)}-${label[0]}`);
    rmSync(work, { recursive: true, force: true });
    cpSync(c.src, work, { recursive: true });
    const target = join(work, c.inner);
    const page = await makePage(url, () => target);
    const t0 = Date.now();
    await c.run(page);
    const files = readdirSync(join(target, readdirSync(target)[0]));
    console.log(`  ${label}: 完了 ${((Date.now() - t0) / 1000).toFixed(1)} 秒 / 先頭フォルダーのファイル数 ${files.length}`);
    await page.close();
    snaps[label] = await snapshot(target);
    // 出来上がりが「正しい」かどうかも見る (両方が同じように間違っていないこと)
    const check = await c.verify(target);
    console.log(`    → 検査: ${check}`);
  }
  const A = snaps['A サーバー配信'], B = snaps['B ローカル (file://)'];
  const namesA = Object.keys(A).sort(), namesB = Object.keys(B).sort();
  console.log(`  ファイル数: A=${namesA.length} B=${namesB.length}`);
  console.log(`  ファイル名の一致: ${JSON.stringify(namesA) === JSON.stringify(namesB) ? 'はい' : 'いいえ'}`);
  let same = 0, diff = [];
  for (const n of namesA) {
    const pa = A[n] ?? {}, pb = B[n] ?? {};
    for (const part of new Set([...Object.keys(pa), ...Object.keys(pb)])) {
      if (pa[part] === pb[part]) same++; else diff.push(`${n} :: ${part}`);
    }
  }
  console.log(`  中身の比較: 一致 ${same} パーツ / 相違 ${diff.length} パーツ`);
  diff.slice(0, 3).forEach((d) => console.log(`    ✗ ${d}`));
}
await browser.close();
srv.close();
