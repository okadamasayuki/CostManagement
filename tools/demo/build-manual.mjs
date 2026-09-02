/**
 * 共有フォルダーに置ける「動画つき説明書」を 1 ファイルで作る。
 *
 * 動画を base64 で埋め込み、再生ボタンを押したときに Blob に戻して渡す。
 * (data: URI のまま再生すると、シークのたびに読み直しが起きて重いため)
 * 外部から読み込むものは 1 つも無く、ダブルクリックだけで開ける。
 *
 * 埋め込むのは WebM (VP9)。ロイヤリティフリーなので、どのブラウザーにも
 * 必ず入っている (H.264 は入っていない版のブラウザーが存在する)。
 * 持ち出して PowerPoint などで使いたい場合のために、MP4 も別に置いてある。
 */
import { readFileSync, writeFileSync, statSync, mkdirSync } from 'node:fs';
import { join, dirname } from 'node:path';

const ROOT = process.env.REPO_ROOT ?? process.cwd();
const OUT = process.argv[2] ?? join(ROOT, '_site', 'manual-offline.html');
const VIDEOS = join(ROOT, 'docs', 'videos');

const items = [
  {
    id: 'v1',
    file: 'lock-by-color.webm',
    title: '① 黄色い入力欄だけ入力できるようにする',
    steps: [
      'ダウンロードした HTML をダブルクリックして開く',
      '共有フォルダーを指定して、配下の Excel を 30 個まとめて読み込む',
      '「色からロックを設定」で <b>黄色以外をすべてロック</b>（試算 → 実行）',
      '画面で結果を確認（斜線＝ロック／黄色＝入力できる）',
      '「シート保護を有効化」（<b>これをしないとロックは効きません</b>）',
      '「元の場所へ上書き保存」と「ZIP で保存」',
      '保存したファイルを読み込み直して確認',
    ],
    note: null,
  },
  {
    id: 'v2',
    file: 'year-update.webm',
    title: '② 年度をまとめて 1 年進める',
    steps: [
      '2023・2024 年度の実績はロック済み、2025 年度の記入欄だけ黄色、という状態から始める',
      '年は<b>ファイル名・シート名・表題・見出し・注記</b>のあちこちに入っている',
      '「年度更新」タブで適用先を全ブック・全シートに',
      '<b>対象にする年の範囲を絞る</b> — 数量にある「年に見える 4 桁」を守るため',
      '「試算」で確認してから実行（30 ファイルが一度に更新）',
      '結果を何ファイルか開いて確認',
      '「元の場所へ上書き保存」（前年のファイルを残すか消すかも選べます）',
    ],
    note:
      '<b>年度更新は「年の数字」を書き換える機能です。</b>実績の数値そのものが列ごと動くわけではありません。' +
      '数量などに年と同じ 4 桁の数字があると年とみなされることがあるので、' +
      '<b>対象の年を実際に使う範囲に絞り、必ず「試算」で確認</b>してください。動画ではその手順も収録しています。',
  },
];

const mb = (n) => `${(n / 1048576).toFixed(1)}MB`;
let total = 0;
const blocks = items.map((it) => {
  const path = join(VIDEOS, it.file);
  const size = statSync(path).size;
  total += size;
  const b64 = readFileSync(path).toString('base64');
  return { ...it, b64, size };
});

const html = `<!doctype html>
<html lang="ja">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <!-- 外部とは一切通信しない。動画はこのファイルの中に入っている。 -->
    <meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; img-src data:; media-src blob:; connect-src 'none'; form-action 'none'; frame-src 'none'; object-src 'none'; base-uri 'none';" />
    <meta name="referrer" content="no-referrer" />
    <link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 16 16'%3E%3Crect width='16' height='16' rx='3' fill='%23217346'/%3E%3Cpath d='M4 4h3l1.5 2.5L10 4h3l-3 4 3 4h-3L8.5 9.5 7 12H4l3-4z' fill='white'/%3E%3C/svg%3E" />
    <title>操作説明（動画つき） — Excel 一括ロック / 年度更新ツール</title>
    <style>
      :root { --green:#217346; --green-dark:#0f5132; --ink:#1f2328; --dim:#57606a; --line:#d8dee4; --bg:#f6f8fa; }
      * { box-sizing:border-box; }
      body { margin:0; font-family:"Yu Gothic","Hiragino Kaku Gothic ProN","Noto Sans JP","Meiryo",sans-serif;
             color:var(--ink); background:var(--bg); line-height:1.75; }
      header { background:var(--green-dark); color:#fff; padding:22px 24px; }
      header .inner { max-width:1040px; margin:0 auto; }
      header h1 { margin:0; font-size:20px; }
      header p { margin:6px 0 0; font-size:13.5px; color:#b7e4c7; }
      main { max-width:1040px; margin:0 auto; padding:28px 24px 64px; }
      .lead { background:#fff; border:1px solid var(--line); border-left:5px solid var(--green);
              border-radius:6px; padding:16px 20px; margin:0 0 30px; }
      .lead p { margin:0 0 8px; } .lead p:last-child { margin-bottom:0; }
      section { background:#fff; border:1px solid var(--line); border-radius:8px; padding:22px 24px 26px; margin-bottom:30px; }
      h2 { margin:0 0 4px; font-size:19px; color:var(--green-dark); }
      .meta { color:var(--dim); font-size:13px; margin:0 0 16px; }
      .player { position:relative; background:#000; border-radius:6px; border:1px solid var(--line); overflow:hidden; }
      .player video { width:100%; display:block; }
      .poster { display:flex; align-items:center; justify-content:center; flex-direction:column; gap:14px;
                aspect-ratio:16/9; color:#cfe6da; cursor:pointer; background:#0d2a1e; }
      .poster .play { width:76px; height:76px; border-radius:50%; background:var(--green); color:#fff;
                      display:flex; align-items:center; justify-content:center; font-size:30px; }
      .poster span { font-size:14px; }
      h3 { font-size:14.5px; margin:20px 0 8px; color:var(--dim); letter-spacing:.04em; }
      ol.steps { margin:0; padding-left:22px; font-size:14.5px; } ol.steps li { margin-bottom:5px; }
      b { color:#8a4b00; }
      .note { background:#fffbe6; border:1px solid #ffe08a; border-radius:6px; padding:13px 16px; font-size:13.5px; margin-top:18px; }
      footer { color:var(--dim); font-size:13px; text-align:center; padding:0 24px 44px; }
      @media (max-width:640px){ main{padding:18px 14px 48px;} section{padding:16px 15px 20px;} }
    </style>
  </head>
  <body>
    <header>
      <div class="inner">
        <h1>操作説明（動画つき）</h1>
        <p>このファイル 1 つに動画が入っています。ネットワークにつながっていなくても再生できます。</p>
      </div>
    </header>
    <main>
      <div class="lead">
        <p><b>実際のツールを操作した画面をそのまま録画したもの</b>です。作り物の画面ではありません。説明の字幕だけを重ねています。</p>
        <p>どちらも、共有フォルダーにある <b>30 個の Excel</b> を対象にしています。ファイル数が何個でも操作は同じです。</p>
      </div>
${blocks
  .map(
    (b) => `      <section>
        <h2>${b.title}</h2>
        <p class="meta">1600×900 ／ 約 ${mb(b.size)}</p>
        <div class="player" id="p-${b.id}">
          <div class="poster" data-for="${b.id}">
            <div class="play">▶</div>
            <span>クリックで再生します</span>
          </div>
        </div>
        <h3>この動画で分かること</h3>
        <ol class="steps">
${b.steps.map((s) => `          <li>${s}</li>`).join('\n')}
        </ol>
${b.note ? `        <div class="note">${b.note}</div>\n` : ''}      </section>
`,
  )
  .join('')}      <div class="note">
        <b>収録環境について。</b>この動画は Linux 上で収録しているため、フォルダーを選ぶ画面の見た目は Windows と異なります。
        収録の都合でフォルダーの選択を自動化していますが、読み込むファイルも保存先も
        <b>実際の共有フォルダーの実ファイル</b>で、ツールの動作・結果には手を加えていません。
      </div>
    </main>
    <footer>Excel 一括ロック / 年度更新ツール — 社内向け・完全オフライン動作</footer>

${blocks.map((b) => `    <script type="text/plain" id="d-${b.id}">${b.b64}</script>`).join('\n')}
    <script>
      // 押されたときだけ Blob に戻す。開いた瞬間に全部展開すると重いため。
      document.querySelectorAll('.poster').forEach(function (poster) {
        poster.addEventListener('click', function () {
          var id = poster.getAttribute('data-for');
          var b64 = document.getElementById('d-' + id).textContent.trim();
          var bin = atob(b64);
          var buf = new Uint8Array(bin.length);
          for (var i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
          var url = URL.createObjectURL(new Blob([buf], { type: 'video/webm' }));
          var v = document.createElement('video');
          v.controls = true;
          v.playsInline = true;
          v.src = url;
          poster.parentNode.replaceChild(v, poster);
          v.play();
        });
      });
    </script>
  </body>
</html>
`;

mkdirSync(dirname(OUT), { recursive: true });
writeFileSync(OUT, html);
console.log(`作成: ${OUT}`);
console.log(`  動画 ${blocks.length} 本 / 元サイズ ${mb(total)} → ファイル全体 ${mb(statSync(OUT).size)}`);
