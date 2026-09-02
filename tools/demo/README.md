# 操作説明動画の収録

`docs/videos/*.mp4` を作り直すための一式です。2 本あります。

| 台本 | 動画 | 内容 |
| --- | --- | --- |
| `record.mjs` | `lock-by-color.mp4` | 黄色い入力欄だけ入力できるようにする |
| `record2.mjs` | `year-update.mp4` | 年度をまとめて 1 年進める |

実物のツールを**本物のマウスカーソルで実際にクリックして**動かし、画面全体を録画します。
説明の字幕は重ねていますが、ツールの画面・動作・結果には手を加えていません。

## ① の動画で収録している内容

共有フォルダーの Excel 30 個を、黄色い入力欄だけ入力できる状態にして配る、という運用です。

1. ダウンロードした HTML をダブルクリックで開く
2. 「フォルダーを開く」で共有フォルダーを指定 → 配下の 30 ファイルを一括読み込み
3. 中身の確認（黄色いセルが支店の入力欄）
4. 適用先を「全ブック・全シート」に
5. 「色からロックを設定」で **黄色以外をロック**（試算 → 実行）
6. 画面で結果を確認（斜線 = ロック / 黄色 = 入力できる）
7. 「シート保護を有効化」
8. 「元の場所へ上書き保存」と「ZIP で保存」
9. 保存したファイルを**読み込み直して**結果を確認

## ② の動画で収録している内容

2023・2024 年度の実績がロック済み、2025 年度の記入欄だけ黄色、という状態から始めます。

1. ダウンロードした HTML をダブルクリックで開く
2. 共有フォルダーを指定 → 30 ファイルを一括読み込み
3. 年が入っている場所の確認（ファイル名・シート名・表題・見出し・注記）
4. 「年度更新」タブで適用先を全ブック・全シートに、ファイル名も対象に
5. **対象にする年の範囲を絞る** — 数量にある「年に見える 4 桁」を守るため
6. 「試算」で確認（8 セル → 6 セルに減ることを見せる）→ 実行
7. 結果を 5 ファイル確認
8. 「元の場所へ上書き保存」（前年のファイルを削除する選択肢も見せる）
9. 読み込み直して確認

## 必要なもの

- Xvfb, openbox, xdotool（`apt-get install -y xvfb openbox xdotool`）
- `npm install --no-save ffmpeg-static`
- Playwright の Chromium

## 手順

```sh
npm run build                       # dist/index.html を作る
mkdir -p .test-build/video
cp dist/index.html .test-build/video/Excel一括ロック_年度更新ツール.html

# ① 黄色以外をロックする動画
node tools/demo/make-share.cjs                    # 共有フォルダー (30 ファイル) を作る
tools/demo/run-x.sh node tools/demo/record.mjs    # 収録 → .test-build/video/out/raw.mp4
node tools/demo/verify-share.cjs                  # 結果の検査

# ② 年度更新の動画
node tools/demo/make-report.cjs                   # 報告書 30 ファイルを作る
tools/demo/run-x.sh node tools/demo/record2.mjs   # 収録 → .test-build/video/out2/raw.mp4
node tools/demo/verify-report.cjs                 # 結果の検査
```

書き出しは `ffmpeg -i raw.mp4 -c:v libx264 -preset slow -crf 23 -pix_fmt yuv420p -movflags +faststart 出力.mp4`。

`FAST=1` を付けると、待ち時間を詰めて録画せずに空回しします（台本の不具合を早く見つける用）。

## 収録環境についての注意

- **フォルダー選択ダイアログ**: 収録環境 (Linux + 仮想画面) では OS のダイアログを確定できません
  （確定した瞬間に Chromium が落ちます）。そのため `showDirectoryPicker` が返すのと同じ形の
  フォルダーハンドルを用意して差し込んでいます。読み込むファイルも、上書き保存の書き込み先も
  **実際の共有フォルダーの実ファイル**で、アプリ側のコードは一切変えていません。
- **ロケール**: `LC_ALL=C.UTF-8` 相当が要ります（POSIX のままだと Chromium が
  ダウンロード名の日本語を落とします）。
