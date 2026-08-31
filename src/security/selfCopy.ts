/**
 * 自分自身 (この HTML ファイル) を保存する機能。
 *
 * 本ツールは全てを 1 枚の HTML に同梱しているため、起動直後の DOM を
 * そのまま文字列にすれば、元のファイルと同じものが再現できる。
 *
 * これにより GitHub Pages などから開いた場合でも、
 *   「1 度だけ開いてダウンロード → あとは共有フォルダーから完全オフラインで使う」
 * という運用ができる。ダウンロードにネットワークを使わない
 * (メモリー上の DOM から作る) ため、「通信 0 件」の性質も保たれる。
 */

let snapshot: string | null = null;

/**
 * React が描画する前に呼ぶこと。
 * 描画後だと画面の状態を含んだ HTML になってしまう。
 */
export function captureSelf(): void {
  try {
    const doctype = document.doctype ? `<!doctype ${document.doctype.name}>\n` : '';
    snapshot = doctype + document.documentElement.outerHTML;
  } catch {
    snapshot = null;
  }
}

export function getSelfCopy(): string | null {
  return snapshot;
}

/** サーバー (GitHub Pages など) から開かれているか。file:// なら false。 */
export function isHosted(): boolean {
  return location.protocol === 'http:' || location.protocol === 'https:';
}

export const OFFLINE_FILE_NAME = 'Excel一括ロック_年度更新ツール.html';
