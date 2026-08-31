import type { LoadedWorkbook } from './types';

/**
 * 読み込んだブックの相対パスから、フォルダーの一覧を組み立てる。
 *
 * 「このフォルダー配下をまとめて処理する」という指定のために使う。
 * 配下 (孫フォルダー以下も含む) のファイル数を数えるので、
 * 選ぶ前に何件が対象になるか分かる。
 */

export interface FolderEntry {
  /** 相対パス。最上位は空文字。 */
  path: string;
  /** 画面に出す名前 */
  label: string;
  /** 階層の深さ (字下げ用) */
  depth: number;
  /** 配下にあるブックの数 (孫以下も含む) */
  count: number;
}

/** relPath が folder の配下にあるか。folder が空文字なら全て対象。 */
export function isUnderFolder(relPath: string, folder: string): boolean {
  const f = folder.replace(/^\/+|\/+$/g, '');
  if (!f) return true;
  return relPath.startsWith(`${f}/`);
}

/** ブック一覧に含まれるフォルダーを、階層順に列挙する */
export function listFolders(books: LoadedWorkbook[]): FolderEntry[] {
  const counts = new Map<string, number>();

  for (const b of books) {
    const parts = b.relPath.split('/');
    parts.pop(); // ファイル名を除く
    // 途中のフォルダーすべてに 1 件ずつ加算する (孫以下も数えるため)
    let acc = '';
    for (const part of parts) {
      acc = acc ? `${acc}/${part}` : part;
      counts.set(acc, (counts.get(acc) ?? 0) + 1);
    }
  }

  const entries: FolderEntry[] = [...counts.entries()]
    .map(([path, count]) => ({
      path,
      label: path.split('/').pop() ?? path,
      depth: path.split('/').length - 1,
      count,
    }))
    .sort((a, b) => a.path.localeCompare(b.path, 'ja'));

  // 最上位 (フォルダー指定なし = 全部) を先頭に置く
  return [{ path: '', label: 'すべて (最上位)', depth: 0, count: books.length }, ...entries];
}
