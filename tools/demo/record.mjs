/**
 * 操作説明動画の収録。
 *
 * 実物のツールを、本物のマウスカーソルで、実際にクリックして操作し、
 * 画面全体を録画する。字幕は説明のために重ねているだけで、
 * ツールの画面・動作・結果には一切手を加えていない。
 */
import { chromium } from 'playwright';
import { execFileSync, spawn } from 'node:child_process';
import { mkdirSync, rmSync, readFileSync, writeFileSync, readdirSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { createRequire } from 'node:module';
import JSZip from 'jszip';
const require = createRequire(import.meta.url);
const FF = require('ffmpeg-static');

const ROOT = process.env.REPO_ROOT ?? process.cwd();
const VID = join(ROOT, '.test-build', 'video');
const OUT = join(VID, 'out');
const SHARE = '/root/共有フォルダー/原価管理/2025年度予算';
const TOOL = join(VID, 'Excel一括ロック_年度更新ツール.html');
const SAVED = join(VID, '保存結果');
rmSync(OUT, { recursive: true, force: true });
rmSync(SAVED, { recursive: true, force: true });
mkdirSync(OUT, { recursive: true });
mkdirSync(SAVED, { recursive: true });

const FAST = process.env.FAST === '1';   // 台本の不具合を早く見つけるための空回し
const x = (a) => { try { return execFileSync('xdotool', a, { encoding: 'utf8' }); } catch { return ''; } };
const sleep = (ms) => new Promise((r) => setTimeout(r, FAST ? Math.min(ms, 60) : ms));

const browser = await chromium.launch({
  headless: false,
  executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome',
  env: { ...process.env, GTK_USE_PORTAL: '0', NO_AT_BRIDGE: '1' },
  args: ['--no-sandbox','--disable-dev-shm-usage','--no-first-run','--disable-background-networking',
    '--disable-component-update','--disable-sync','--no-default-browser-check','--disable-gpu',
    '--disable-client-side-phishing-detection','--metrics-recording-only','--hide-crash-restore-bubble',
    '--window-position=0,0','--window-size=1920,1080'],
});
const ctx = await browser.newContext({ viewport: null, acceptDownloads: true });
const page = await ctx.newPage();
/**
 * フォルダー選択の橋渡し。
 *
 * 収録環境 (Linux + 仮想画面) では OS のフォルダー選択ダイアログを
 * 確定できない。そこで、ダイアログが返すのと同じ形のフォルダーハンドルを
 * 用意して showDirectoryPicker から返す。
 *
 * 中身は <ruby>本物<rt>ほんもの</rt></ruby>で、
 *   ・読み込むのは実際の共有フォルダーにある 30 個の Excel
 *   ・上書き保存も、その実ファイルを実際に書き換える
 * アプリ側のコード (loadFromDirectory / writeBackToDisk) は一切変えていない。
 */
let currentRoot = SHARE;
await page.exposeBinding('__fsList', async (_s, rel) => {
  const dir = rel ? join(currentRoot, rel) : currentRoot;
  return readdirSync(dir, { withFileTypes: true })
    .map((d) => ({ name: d.name, kind: d.isDirectory() ? 'directory' : 'file' }));
});
await page.exposeBinding('__fsRead', async (_s, rel) => readFileSync(join(currentRoot, rel)).toString('base64'));
await page.exposeBinding('__fsWrite', async (_s, rel, b64) => {
  writeFileSync(join(currentRoot, rel), Buffer.from(b64, 'base64'));
  return true;
});
await page.exposeBinding('__fsRemove', async (_s, rel) => { rmSync(join(currentRoot, rel), { force: true }); return true; });
await page.exposeBinding('__fsRootName', async () => currentRoot.split('/').pop());
await page.addInitScript(() => {
  const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const b64ToBin = (b64) => Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
  const binToB64 = (u8) => { let s = ''; const N = 0x8000;
    for (let i = 0; i < u8.length; i += N) s += String.fromCharCode.apply(null, u8.subarray(i, i + N));
    return btoa(s); };
  const makeFile = (name, rel) => ({
    kind: 'file', name,
    async getFile() { return new File([b64ToBin(await window.__fsRead(rel))], name, { type: XLSX }); },
    async createWritable() {
      const parts = [];
      return {
        async write(d) { parts.push(d); },
        async close() {
          const u8 = new Uint8Array(await new Blob(parts).arrayBuffer());
          await window.__fsWrite(rel, binToB64(u8));
        },
      };
    },
  });
  const makeDir = (name, rel) => ({
    kind: 'directory', name,
    async *values() {
      for (const it of await window.__fsList(rel)) {
        const child = rel ? `${rel}/${it.name}` : it.name;
        yield it.kind === 'directory' ? makeDir(it.name, child) : makeFile(it.name, child);
      }
    },
    async getFileHandle(n) { return makeFile(n, rel ? `${rel}/${n}` : n); },
    async removeEntry(n) { await window.__fsRemove(rel ? `${rel}/${n}` : n); },
  });
  window.showDirectoryPicker = async () => makeDir(await window.__fsRootName(), '');
});
await page.goto(`file://${TOOL}`);
await page.waitForSelector('.app');
await sleep(900);
// 起動直後はアドレスバーが選択状態なので、画面側をひと押しして外す
await page.locator('.titlebar, .app').first().click({ position: { x: 400, y: 8 } }).catch(() => {});
await sleep(900);

const off = await page.evaluate(() => ({ sx: window.screenX, sy: window.screenY, ow: window.outerWidth, iw: window.innerWidth, oh: window.outerHeight, ih: window.innerHeight }));
const OFFX = off.sx + Math.round((off.ow - off.iw) / 2);
const OFFY = off.sy + (off.oh - off.ih);

await page.addStyleTag({ content: `
  #vcap { position: fixed; left: 0; right: 0; bottom: 30px; z-index: 2147483000; pointer-events: none;
    background: #0b1a14;
    border-top: 4px solid #21a366; color: #fff; padding: 15px 34px 17px;
    font-family: "Noto Sans CJK JP","Yu Gothic",sans-serif; display: none;
    box-shadow: 0 -14px 40px rgba(0,0,0,0.35); }
  #vcap.top { top: 0; bottom: auto; border-top: none; border-bottom: 4px solid #21a366;
    box-shadow: 0 14px 40px rgba(0,0,0,0.35); }
  #vcap.on { display: block; animation: vfade .35s ease; }
  @keyframes vfade { from { opacity: 0 } to { opacity: 1 } }
  #vcap .row { display: flex; align-items: flex-start; gap: 18px; }
  #vcap .num { flex: 0 0 auto; background: #21a366; color: #fff; font-weight: 700;
    font-size: 15px; border-radius: 999px; padding: 7px 17px; margin-top: 4px; white-space: nowrap; }
  #vcap .ttl { font-size: 28px; font-weight: 700; line-height: 1.3; }
  #vcap .dsc { font-size: 18px; color: #cfe6da; line-height: 1.62; margin-top: 7px; }
  #vcap .dsc b { color: #ffe07a; }
  #vcap .dsc .warn { color: #ffb3a7; }
  #vring { position: fixed; z-index: 2147483001; border: 3px solid #ffcf33; border-radius: 7px;
    box-shadow: 0 0 0 4px rgba(255,207,51,.26), 0 0 22px rgba(255,207,51,.55); display: none;
    pointer-events: none; transition: all .3s cubic-bezier(.4,0,.2,1); }
  #vring.on { display: block; }
  #vtitle { position: fixed; inset: 0; z-index: 2147483002; background: #0d2a1e; pointer-events: none;
    color: #fff; display: none; flex-direction: column; align-items: center; justify-content: center;
    font-family: "Noto Sans CJK JP","Yu Gothic",sans-serif; text-align: center; padding: 0 90px; }
  #vtitle.on { display: flex; animation: vfade .5s ease; }
  #vtitle .big { font-size: 50px; font-weight: 700; line-height: 1.4; }
  #vtitle .sub { font-size: 23px; color: #9fd6ba; margin-top: 28px; line-height: 1.85; }
  #vtitle .tag { font-size: 16.5px; color: #79b699; margin-top: 42px; line-height: 1.8; }
`});
await page.evaluate(() => {
  const cap = document.createElement('div');
  cap.id = 'vcap';
  cap.innerHTML = '<div class="row"><div class="num"></div><div><div class="ttl"></div><div class="dsc"></div></div></div>';
  document.body.appendChild(cap);
  const r = document.createElement('div'); r.id = 'vring'; document.body.appendChild(r);
  const t = document.createElement('div'); t.id = 'vtitle';
  t.innerHTML = '<div class="big"></div><div class="sub"></div><div class="tag"></div>';
  document.body.appendChild(t);
});

let step = 0;
const TOTAL = 10;
async function caption(title, desc, opts = {}) {
  if (!opts.keepStep) step++;
  await page.evaluate(([n, total, t, d, top]) => {
    const c = document.getElementById('vcap');
    c.querySelector('.num').textContent = `STEP ${n} / ${total}`;
    c.querySelector('.ttl').textContent = t;
    c.querySelector('.dsc').innerHTML = d;
    c.classList.toggle('top', !!top);
    c.classList.add('on');
  }, [step, TOTAL, title, desc, !!opts.top]);
  const chars = (title + desc.replace(/<[^>]+>/g, '')).length;
  await sleep(opts.hold ?? Math.min(12000, Math.max(4000, 1400 + chars * 82)));
}
const note = (t, d, o = {}) => caption(t, d, { ...o, keepStep: true });
const hideCap = () => page.evaluate(() => document.getElementById('vcap').classList.remove('on'));
async function titleCard(big, sub, tag, hold = 5500) {
  await page.evaluate(([b, s, t]) => {
    const e = document.getElementById('vtitle');
    e.querySelector('.big').innerHTML = b;
    e.querySelector('.sub').innerHTML = s;
    e.querySelector('.tag').innerHTML = t;
    e.classList.add('on');
  }, [big, sub, tag]);
  await sleep(hold);
  await page.evaluate(() => document.getElementById('vtitle').classList.remove('on'));
  await sleep(600);
}
async function ringBox(b, pad = 6) {
  await page.evaluate(([r, p]) => {
    const e = document.getElementById('vring');
    e.style.left = `${r.x - p}px`; e.style.top = `${r.y - p}px`;
    e.style.width = `${r.width + p * 2}px`; e.style.height = `${r.height + p * 2}px`;
    e.classList.add('on');
  }, [b, pad]);
}
async function ring(sel, pad = 6) {
  await page.locator(sel).first().scrollIntoViewIfNeeded().catch(() => {});
  const b = await page.locator(sel).first().boundingBox();
  if (b) await ringBox(b, pad);
  return b;
}
const unring = () => page.evaluate(() => document.getElementById('vring').classList.remove('on'));

function pointer() { const m = /X=(\d+)\s+Y=(\d+)/.exec(x(['getmouselocation','--shell'])); return { x: +(m?.[1] ?? 0), y: +(m?.[2] ?? 0) }; }
async function moveTo(px, py, steps = 22) {
  const c = pointer();
  for (let i = 1; i <= steps; i++) {
    const t = i / steps, e = t < 0.5 ? 2 * t * t : -1 + (4 - 2 * t) * t;
    x(['mousemove', String(Math.round(c.x + (px - c.x) * e)), String(Math.round(c.y + (py - c.y) * e))]);
    await sleep(17);
  }
  await sleep(240);
}
async function clickOn(sel, opts = {}) {
  // 枠の外へスクロールしていると、実クリックが別の場所に当たってしまう
  await page.locator(sel).first().scrollIntoViewIfNeeded().catch(() => {});
  await sleep(FAST ? 20 : 250);
  const b = await page.locator(sel).first().boundingBox();
  if (!b) throw new Error(`見つかりません: ${sel}`);
  if (opts.ring !== false) await ringBox(b, opts.pad ?? 6);
  await moveTo(Math.round(OFFX + b.x + b.width / 2), Math.round(OFFY + b.y + b.height / 2));
  await sleep(opts.before ?? 400);
  x(['click','1']);
  await sleep(opts.after ?? 800);
  if (opts.ring !== false && !opts.keepRing) await unring();
  return b;
}
const wait = async (fn, msg, t = 240000) => {
  const end = Date.now() + t;
  for (;;) { try { if (await fn()) return; } catch {} if (Date.now() > end) throw new Error(msg); await sleep(180); }
};

/**
 * 「フォルダーを開く」を実際に押して、フォルダーを読み込ませる。
 * ファイル選択ダイアログはリスナーを付けることで抑止し、
 * フォルダーの受け渡しは input へ直接行う (FileChooser はフォルダー非対応)。
 */
/** 「フォルダーを開く」を実際に押す */
async function openFolder(dir, opts = {}) {
  currentRoot = dir;
  await clickOn('.rbtn-lg:has-text("フォルダーを開く")', { after: opts.after ?? 1200, keepRing: opts.keepRing });
}

const rec = FAST ? { kill() {} } : spawn(FF, ['-y','-f','x11grab','-video_size','1920x1080','-framerate','25','-draw_mouse','1',
  '-i', process.env.DISPLAY, '-c:v','libx264','-preset','veryfast','-crf','20','-pix_fmt','yuv420p',
  join(OUT, 'raw.mp4')], { stdio: ['ignore','ignore','ignore'] });
await sleep(2000);
console.log(FAST ? '空回し開始 (録画なし)' : '録画開始');

// ==========================================================================
try {

await titleCard(
  '共有フォルダーの Excel 30 個を<br>まとめて「黄色の欄だけ入力できる」ようにする',
  '各支店に配る予算入力表。黄色いセルが入力欄です。<br>黄色以外をすべてロックして、書き換えられないようにします。',
  'この動画は、実際のツールを操作した画面をそのまま録画したものです', 7000);

// -------------------------------------------------------------- STEP 1
await caption('ツールを開く',
  'ダウンロードした <b>HTML ファイル 1 つ</b>をダブルクリックするだけです。インストールも設定も要りません。' +
  '<br>アドレスバーが <b>file:///…</b> になっていて、パソコンの中のファイルを開いていることが分かります。');
await ring('[data-testid="guard-badge"]', 5);
await note('通信しないことを画面で示しています',
  '右上に <b>「完全オフライン動作」</b>と出ています。読み込んだ Excel が外部へ送られることはありません。');
await unring();

// -------------------------------------------------------------- STEP 2
await caption('共有フォルダーを指定して、配下の Excel を全部読み込む',
  '「ファイル」タブの <b>「フォルダーを開く」</b>を押します。' +
  '<br>フォルダーを 1 つ選ぶだけで、<b>その下のサブフォルダーまで全部</b>読み込みます。1 個ずつ開く必要はありません。');
await hideCap();
await openFolder(SHARE, { after: 1200, keepRing: true });
await unring();
await note('ここで、フォルダーを選ぶ画面が開きます',
  'お使いのパソコンの<b>フォルダー選択ダイアログ</b>が開くので、共有フォルダー' +
  '（例: <b>\\\\サーバー\\共有\\原価管理\\2025年度予算</b>）を選んで「開く」を押します。' +
  '<br><span class="warn">※ この動画は収録の都合で、その選択を自動で行っています。</span>');
await wait(async () => (await page.locator('.tree-file').count()) === 30, '読み込めない');
await sleep(1200);

// -------------------------------------------------------------- STEP 3
await caption('30 ファイルが一度に読み込まれました',
  '左側に<b>フォルダーの構造がそのまま</b>出ます。支店ごとのフォルダーも、その中のファイルもすべて対象です。');
await ring(".sidebar", 4);
await sleep(2600);
await unring();
await hideCap();
await ring('.statusbar', 3);
await note('下のバーに、いま何を扱っているかが常に出ます',
  '<b>30 ブック</b>と表示されています。ロック済み / ロック解除のセル数もここで確認できます。', { top: true });
await unring();

// -------------------------------------------------------------- STEP 4
await caption('中身を見てみる — 黄色いセルが「支店が入力する欄」',
  'ファイルをクリックすると中身が出ます。<b>黄色いセル</b>が各支店に数字を入れてもらう欄です。' +
  '<br>費目名・前年度実績・合計の数式は、支店に書き換えられたくない<b>様式の部分</b>です。');
await clickOn('.tree-file >> nth=0', { after: 1400 });
await sleep(800);
const yellowCell = await page.locator('.cell').filter({ hasText: '' }).first().boundingBox();
await ring('.grid-canvas', 2);
await sleep(2200);
await unring();
await note('別のファイルも同じ様式です',
  '30 個すべて同じ形なので、<b>まとめて同じ処理</b>ができます。1 個ずつ開いて設定する必要はありません。');
await clickOn('.tree-file >> nth=5', { after: 1500 });
await sleep(1600);

// -------------------------------------------------------------- STEP 5
await caption('処理の対象を「全部」にする',
  'リボンのすぐ下にある <b>「適用先」</b>で、操作が<b>どこに効くか</b>を決めます。' +
  '<br>ここを <b>「読み込んだ全ブック」「全シート」</b>にすれば、30 ファイル全部が対象になります。');
await ring('.scopebar', 4);
await page.selectOption('[data-testid="scope-books"]', 'all');
await sleep(1000);
await page.selectOption('[data-testid="scope-sheets"]', 'all');
await sleep(1500);
await note('いま何ブック・何シートに当たるかが、その場に出ます',
  '<b>30 ブック / 60 シート が対象</b>と出ています。押す前に確かめられます。' +
  '<br>この適用先は<b>ロック・書式・年度更新のすべてで共通</b>です。タブごとに指定し直す必要はありません。');
await unring();

// -------------------------------------------------------------- STEP 6
await caption('「色からロックを設定」を開く',
  '<b>黄色いセルだけ入力できるようにする</b>のが目的です。範囲を指定する必要はありません。' +
  '<br>ツールがファイル内で実際に使われている色を数えて、一覧から選べるようにします。');
await clickOn('.ribbon-tab:has-text("書式")', { after: 900 });
await clickOn('.rbtn-lg:has-text("色から")', { after: 1400 });
await page.waitForSelector('.modal');
await wait(async () => (await page.locator('.modal [data-color]').count()) > 0, '色一覧が出ない');
await sleep(1000);

await caption('使われている色が一覧で出る', '', { keepStep: true, top: true, hold: 100 });
const swatches = await page.locator('.modal [data-color]').evaluateAll((els) => els.map((e) => ({ key: e.getAttribute('data-color'), text: e.textContent.trim() })));
await note('使われている色が、セル数つきで一覧に出ます',
  `このフォルダーでは <b>${swatches.map((s) => s.text).join('</b> と <b>')}</b> の 2 色が使われています。` +
  '<br>どれが入力欄かは人が知っているので、<b>黄色を選ぶだけ</b>です。', { top: true });
const yellowKey = swatches.find((s) => /FFFF00/i.test(s.key))?.key;
await clickOn(`.modal [data-color="${yellowKey}"]`, { after: 1200, pad: 4 });

await ring('.modal [data-testid="mode-only"]', 5);
await note('「この色のセルだけ入力できるようにする」を選ぶ',
  'Excel のセルは<b>もともと全部ロックされた状態</b>です。' +
  'これを選ぶと、<b>黄色のセルだけロックを外し、それ以外はすべてロック</b>します。' +
  '<br>配って記入してもらう様式は、たいていこれです。', { top: true });
await unring();
await clickOn('.modal [data-testid="mode-only"]', { after: 1200, pad: 3 });

await note('「実行する」を押します', '30 ファイル分の設定が一度に書き換わります。', { top: true, hold: 3600 });
await hideCap();
await clickOn('.modal-foot .rbtn.accent', { after: 2500 });
await wait(async () => (await page.locator('.modal').count()) === 0, '閉じない');
await sleep(1500);

// -------------------------------------------------------------- STEP 7
await caption('画面でそのまま結果を確認できる',
  '<b>斜線</b>が入ったところがロックされたセル、<b>色が付いていない</b>ところが入力できるセルです。' +
  '<br>黄色い入力欄だけが入力できる状態になっているのが見て分かります。');
await ring('.grid-canvas', 2);
await sleep(3000);
await unring();
await hideCap();
await ring('.statusbar', 3);
await note('セル数でも確認できます',
  '下のバーの <b>🔒 ロック済み</b> と <b>🔓 ロック解除</b> の数が、狙いどおりになっているか確認できます。', { top: true });
await unring();
await note('他のファイルも同じように変わっています',
  '1 個ずつ設定していません。<b>30 ファイルすべてに同時に</b>効いています。');
await clickOn('.tree-file >> nth=12', { after: 1600 });
await sleep(1800);

// -------------------------------------------------------------- STEP 8
await caption('仕上げに「シート保護」をかける',
  '<b>ここが一番間違えやすいところです。</b>Excel では、ロックの設定だけでは効きません。' +
  '<br><b>シート保護を有効</b>にして、はじめてロックが効きます。');
await clickOn('.ribbon-tab:has-text("ロック")', { after: 900 });
await page.selectOption('[data-testid="scope-books"]', 'all');
await page.selectOption('[data-testid="scope-sheets"]', 'all');
await sleep(600);
await hideCap();
await clickOn('.rbtn-lg:has-text("シート保護を有効化")', { after: 2500 });
await wait(async () => (await page.textContent('.rp-body')).includes('保護'), '保護できない');
await sleep(1200);
await ring('.rightpanel', 3);
await note('30 ファイル・60 シートに一度でかかりました',
  'パスワードを付けたい場合は、実行前に「パスワード」欄に入力しておきます。');
await unring();

// -------------------------------------------------------------- STEP 9
await caption('保存する — ① 元の場所へ上書き保存',
  '「フォルダーを開く」から読み込んだ場合は、<b>共有フォルダーのファイルをそのまま上書き</b>できます。' +
  '<br>ZIP を展開して戻す手間がありません。Chrome / Edge で使えます。');
await clickOn('.ribbon-tab:has-text("ファイル")', { after: 900 });
await hideCap();
await clickOn('.rbtn-lg:has-text("元の場所へ")', { after: 1400 });
await page.waitForSelector('.modal');
await ring('.modal', 6);
await note('実行前に対象の件数が出ます', '<b>元のファイルを直接書き換える</b>ので、事前にバックアップを取っておくと安心です。', { top: true });
await unring();
await hideCap();
await clickOn('.modal-foot .rbtn.accent', { after: 2500 });
await wait(async () => /件を上書き|新しい名前で保存/.test(await page.textContent('body')), '上書きできない');
await sleep(1200);
await ring('.rightpanel', 3);
await note('30 件を共有フォルダーへ上書きしました',
  '元のファイルが、そのまま置き換わりました。<b>ファイル名は変わっていない</b>ので、そのまま配れます。');
await unring();

await note('保存する — ② すべてを ZIP で保存',
  '上書きしたくない場合や、他のブラウザーを使う場合は <b>ZIP</b> で書き出せます。' +
  '<br><b>フォルダー構造を保ったまま</b> 1 つの ZIP にまとまるので、展開すれば元と同じ形です。');
await hideCap();
const dl = page.waitForEvent('download', { timeout: 300000 });
await clickOn('.rbtn-lg:has-text("ZIP")', { after: 1400 });
await page.waitForSelector('.modal');
await sleep(900);
await clickOn('.modal-foot .rbtn.accent', { after: 1500 });
await (await dl).saveAs(join(OUT, 'result.zip'));
await sleep(2200);
await note('ZIP でも保存できました', 'どちらの方法でも中身は同じです。', { hold: 3800 });

// -------------------------------------------------------------- STEP 10
await caption('本当にそうなっているか、読み込み直して確かめる',
  'いま<b>上書きした共有フォルダーを、もう一度読み込みます</b>。' +
  '<br>保存されたファイルが狙いどおりになっているかを、その場で確認できます。');
await hideCap();
for (let i = 0; i < 40 && (await page.locator('.tree-file').count()) > 0; i++) {
  await page.locator('.tree-file .icon-btn').first().click();
  await sleep(FAST ? 10 : 40);
}
await sleep(900);
await openFolder(SHARE, { after: 1200 });
await wait(async () => (await page.locator('.tree-file').count()) === 30, '読み込み直せない');
await sleep(1400);
await clickOn('.tree-file >> nth=0', { after: 1600 });
await sleep(800);
await ring('.grid-canvas', 2);
await note('保存されたファイルでも、狙いどおりになっています',
  '<b>黄色い欄だけが入力できる</b>状態、それ以外はロック、<b>シート保護もかかった状態</b>で保存されています。' +
  '<br>このまま各支店に配れば、様式を壊されずに数字だけ入れてもらえます。');
await unring();
await hideCap();
await ring('.statusbar', 3);
await note('保存後の状態', '<b>🛡️ 保護あり</b>と出ています。🔒 と 🔓 の数も保たれています。', { top: true });
await unring();

// -------------------------------------------------------------- STEP 11
await caption('毎年やる作業なら「手順書」に残せます',
  '今の一連の操作は<b>自動で手順として記録</b>されています。書き出しておけば、来年は' +
  '<b>読み込んで「すべて実行」するだけ</b>で同じ作業が終わります。');
await clickOn('.ribbon-tab:has-text("手順書")', { after: 1200 });
await ring('.rightpanel', 3);
await sleep(3200);
await unring();
await hideCap();

await titleCard('まとめ',
  '① 共有フォルダーを指定 → 配下の Excel を全部読み込む<br>' +
  '② 「色からロックを設定」で <b>黄色以外をロック</b><br>' +
  '③ 「シート保護を有効化」（これをしないと効きません）<br>' +
  '④ 「元の場所へ上書き保存」または「ZIP で保存」',
  'ファイル数が何個でも操作は同じです。外部との通信は一切ありません。', 9000);

} catch (e) {
  console.error('収録中のエラー:', e.message);
  await page.screenshot({ path: join(OUT, 'error.png') }).catch(() => {});
}
// ==========================================================================
await sleep(1200);
rec.kill('SIGINT');
await sleep(3500);
await browser.close().catch(() => {});
console.log('収録終了');
