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
const OUT = join(VID, 'out2');
const SHARE = '/root/報告共有フォルダー/2025年度報告';
const TOOL = join(VID, 'Excel一括ロック_年度更新ツール.html');
const SAVED = join(VID, '保存結果2');
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
const TOTAL = 11;
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
  '30 ファイルの年度を<br>一度に 1 年ぶん進める',
  '各支店へ配る数量報告書。2023・2024 年度の実績はロック済み、<br>2025 年度の記入欄だけが黄色で入力できる状態です。<br>これを丸ごと <b>2026 年度版</b>に更新します。',
  'この動画は、実際のツールを操作した画面をそのまま録画したものです', 8000);

// -------------------------------------------------------------- STEP 1
await caption('ツールを開く',
  'ダウンロードした <b>HTML ファイル 1 つ</b>をダブルクリックするだけです。' +
  '<br>アドレスバーが <b>file:///…</b> になっていて、パソコンの中のファイルを開いていることが分かります。');
await ring('[data-testid="guard-badge"]', 5);
await note('読み込んだ Excel が外へ出ることはありません',
  '右上に <b>「完全オフライン動作」</b>と出ています。ネットワークに接続していなくても動きます。');
await unring();

// -------------------------------------------------------------- STEP 2
await caption('共有フォルダーを指定して、配下の Excel を全部読み込む',
  '「ファイル」タブの <b>「フォルダーを開く」</b>を押し、報告書が入っている共有フォルダーを選びます。' +
  '<br>フォルダーを 1 つ選ぶだけで、<b>支店ごとのサブフォルダーまで全部</b>読み込みます。');
await hideCap();
await openFolder(SHARE, { after: 1200, keepRing: true });
await unring();
await note('ここで、フォルダーを選ぶ画面が開きます',
  'お使いのパソコンの<b>フォルダー選択ダイアログ</b>が開くので、共有フォルダーを選んで「開く」を押します。' +
  '<br><span class="warn">※ この動画は収録の都合で、その選択を自動で行っています。</span>');
await wait(async () => (await page.locator('.tree-file').count()) === 30, '読み込めない');
await sleep(1200);
await ring(".sidebar", 4);
await note('30 ファイルが一度に読み込まれました', '支店ごとのフォルダーがそのまま出ます。1 個ずつ開く必要はありません。');
await unring();

// -------------------------------------------------------------- STEP 3
await caption('いまのファイルの状態を確認する',
  '<b>2023年度・2024年度の実績</b>には数値が入っていて、<b>斜線</b>＝ロックされています（書き換えられません）。' +
  '<br><b>2025年度計画</b>の欄は<b>黄色</b>で、空欄のまま入力できる状態です。ここに支店が数字を入れます。');
await clickOn('.tree-file >> nth=0', { after: 1500 });
await sleep(800);
await ring('.grid-canvas', 2);
await sleep(3000);
await unring();
await hideCap();
await ring('.statusbar', 3);
await note('シート保護もかかっています',
  '<b>🛡️ 保護あり</b>と出ています。前年の作業でロックと保護が済んでいる状態です。', { top: true });
await unring();
await note('年は「あちこち」に入っています',
  '<b>ファイル名</b>（2025年度_数量報告書_…）、<b>シート名</b>（2025年度 / 2024年度実績）、' +
  '表題、作成日、見出し（2023年度実績 / 2024年度実績 / 2025年度計画）、注記。' +
  '<br>これを手で直すと、30 ファイル × 何箇所もの修正になります。');
await ring('.sheettabs, .sheettab', 4);
await sleep(2500);
await unring();

// -------------------------------------------------------------- STEP 4
await caption('やりたいこと — すべての年を 1 年ぶん進める',
  '<b>2023 → 2024</b>、<b>2024 → 2025</b>、<b>2025 → 2026</b>。' +
  '<br>このツールは<b>すべてを同時に置き換える</b>ので、2023 が 2024 になってさらに 2025 に…という' +
  '<b>連鎖は起きません</b>。');

// -------------------------------------------------------------- STEP 5
await caption('まず「適用先」を全部にする',
  'リボンのすぐ下の <b>「適用先」</b>を「読み込んだ全ブック」「全シート」にします。' +
  '<br>ここは<b>ロック・書式・年度更新のすべてで共通</b>です。タブごとに指定し直す必要はありません。');
await ring('.scopebar', 4);
await page.selectOption('[data-testid="scope-books"]', 'all');
await sleep(1000);
await page.selectOption('[data-testid="scope-sheets"]', 'all');
await sleep(1400);
await note('いま何ブック・何シートに当たるかが、その場に出ます',
  '<b>30 ブック / 60 シート が対象</b>と出ています。押す前に確かめられます。');
await unring();
await clickOn('.ribbon-tab:has-text("年度更新")', { after: 1000 });
await ring('.ribbon-panel .rgroup:has(.rgroup-title:text-is("年をずらす"))', 5);
await note('やることは「年を +1 年ずらす」の 1 つだけです',
  '<b>何年と書いてあるかを調べておく必要はありません。</b>' +
  'ツールが年に見える数字を探して、<b>見つかったものを全部まとめて</b> 1 年進めます。' +
  '<br>2023→2024、2024→2025、2025→2026 を<b>同時に</b>置き換えるので、' +
  '2023 が 2025 まで進んでしまう<b>連鎖は起きません</b>。');
await unring();

await ring('.ribbon-panel .rgroup:has(.rgroup-title:text-is("どこに書いてある年を変えるか"))', 5);
await note('どこに書いてある年を変えるかを選びます',
  '<b>セルの文字・数字</b>・<b>数式の中身</b>・<b>シート名</b>は最初から入っています。' +
  '<br>今回は<b>ファイル名</b>にも年が入っているので、これも入れます。');
await unring();
await clickOn('.ribbon-panel .check:has-text("ファイル名")', { after: 1200, pad: 3 });

// -------------------------------------------------------------- STEP 6
await caption('落とし穴 — 数量にも「年に見える 4 桁」がある',
  'この表の <b>「特注シャフト」</b>は数量が <b>2,031 個 / 2,018 個</b>。' +
  '<br>4 桁の数字なので、何も考えずに置き換えると<b>数量まで 1 増えてしまいます</b>。');
await hideCap();
await clickOn('.tree-file >> nth=0', { after: 1200 });
const shaft = await page.locator('.cell', { hasText: '特注シャフト' }).first().boundingBox();
if (shaft) await ringBox({ x: shaft.x, y: shaft.y, width: shaft.width * 3.4, height: shaft.height }, 3);
await note('数量の 2,031 / 2,018 は「年」ではありません',
  '実務では<b>ここが一番の落とし穴</b>です。');
await unring();

// -------------------------------------------------------------- STEP 7
await caption('このツールは、最初からその分を守るようにしてあります',
  '<b>「数字だけのセルも年とみなす」は、はじめから外れています。</b>' +
  '<br>数量 2,031 個のように<b>数字だけ</b>が入ったセルは触らず、' +
  '「2025年度」のように<b>文字と一緒に書かれた年</b>だけを書き換えます。');
await clickOn('.ribbon-tab:has-text("年度更新")', { after: 900 });
await ring('.ribbon-panel .rgroup:has(.rgroup-title:text-is("間違って変えないための設定"))', 5);
await sleep(3200);
await note('外したままで結構です',
  'ここに<b>チェックを入れたときだけ</b>、数量 2,031 も年として扱われます。' +
  '<br>ふつうの年度更新では、外したままにしてください。');
await unring();

// -------------------------------------------------------------- STEP 8
await caption('実行する — 30 ファイルを一度に',
  '<b>1 個ずつ開いて直す必要はありません。</b>ワンクリックで 30 ファイル・60 シート分が更新されます。');
await hideCap();
await clickOn('.rbtn-lg:has-text("年を")', { after: 2500 });
await wait(async () => (await page.textContent('.grid-canvas')).includes('2026年度'), '年度更新されない');
await sleep(1500);
await ring('.rightpanel', 3);
await note('更新できました', '変更した件数が一覧で出ます。この内容はそのまま「手順書」にも記録されます。');
await unring();

// -------------------------------------------------------------- STEP 9
await caption('結果を何ファイルか見て確かめる',
  '<b>シート名</b>・<b>見出し</b>・<b>表題</b>が 1 年ぶん進んでいるか、いくつか開いて確認します。');
await hideCap();
await clickOn('.tree-file >> nth=0', { after: 1500 });
await sleep(600);
await ring('.grid-canvas', 2);
await note('① 見出しが 2024 / 2025 / 2026 年度になりました',
  '<b>2,031 個 / 2,018 個 の数量はそのまま</b>です。狙ったところだけが変わっています。');
await unring();
for (const [i, label] of [[7, '②'], [14, '③'], [21, '④'], [28, '⑤']]) {
  await hideCap();
  await clickOn(`.tree-file >> nth=${i}`, { after: 1300 });
  await sleep(500);
  const tabs = (await page.locator('.sheettab').allTextContents()).join(' / ');
  await note(`${label} 別の支店のファイルも同じように更新されています`,
    `シートも <b>${tabs}</b> に変わっています。`, { hold: FAST ? 60 : 4200 });
}
await hideCap();
await ring('.sidebar', 4);
await note('ファイル名も 2026年度 に変わります',
  '<b>ファイル名</b>の年は<b>保存するときに</b>反映されます。一覧では変更予定として扱われます。');
await unring();

// -------------------------------------------------------------- STEP 10
await caption('元の場所へ上書き保存する',
  '「フォルダーを開く」から読み込んでいるので、<b>共有フォルダーのファイルをそのまま上書き</b>できます。' +
  '<br>ファイル名も <b>2026年度_…</b> に変わって保存されます。');
await clickOn('.ribbon-tab:has-text("ファイル")', { after: 1000 });
await hideCap();
await clickOn('.rbtn-lg:has-text("元の場所へ")', { after: 1500 });
await page.waitForSelector('.modal');
await ring('.modal', 6);
await note('実行前に、変わるファイル名が一覧で確認できます',
  '<b>2025年度_… → 2026年度_…</b> と変わることが分かります。' +
  '<br><b>元のファイルを直接書き換える</b>ので、事前にバックアップを取っておくと安心です。', { top: true });
await unring();
await ring('.modal .check:has-text("元の名前のファイルを削除する")', 4);
await note('前年のファイルを「残す」か「消す」かを選べます',
  '既定は<b>残す</b>です。この場合、フォルダーには <b>2025年度_… と 2026年度_… の両方</b>が並びます。' +
  '<br>今回は<b>入れ替える</b>運用なので、チェックを入れて元のファイルを削除します。', { top: true });
await unring();
await hideCap();
await clickOn('.modal .check:has-text("元の名前のファイルを削除する")', { after: 1200, pad: 3 });
await sleep(700);
await clickOn('.modal-foot .rbtn.accent', { after: 2500 });
await wait(async () => /新しい名前で保存|件を上書き/.test(await page.textContent('body')), '上書きできない');
await sleep(1500);
await ring('.rightpanel', 3);
await note('30 件を新しい名前で保存しました',
  '<b>上書きした件数と、名前が変わって作り直した件数を分けて</b>知らせます。' +
  '<br>今回は元の 2025年度_… を削除したので、フォルダーには 2026年度_… だけが残ります。');
await unring();

// -------------------------------------------------------------- STEP 11
await caption('保存されたファイルを読み込み直して確かめる',
  '<b>いま上書きした共有フォルダーを、もう一度読み込みます。</b>');
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
await sleep(700);
await ring('.sidebar', 4);
await note('ファイル名が 2026年度 になっています', '共有フォルダーの実ファイルが置き換わっているのが分かります。');
await unring();
await ring('.grid-canvas', 2);
await note('中身も 2026 年度版になっています',
  '<b>黄色の入力欄とロック・シート保護はそのまま</b>です。このまま各支店に配れば、' +
  '<b>2026年度計画</b>の欄だけ入力してもらえます。');
await unring();
await hideCap();

await titleCard('まとめ',
  '① 共有フォルダーを指定 → 配下の Excel を全部読み込む<br>' +
  '② リボン下の<b>適用先を「全ブック・全シート」</b>に<br>' +
  '③ 「年度更新」タブで<b>ファイル名</b>にもチェックを入れる<br>' +
  '④ <b>「年を +1 年ずらす」を 1 回押す</b>（30 ファイルが一度に更新）<br>' +
  '⑤ 何ファイルか開いて確かめ、「元の場所へ上書き保存」',
  'ロックや黄色の色分けはそのまま保たれます。外部との通信は一切ありません。', 10000);

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
