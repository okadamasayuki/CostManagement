import JSZip from 'jszip';
import type { LoadedWorkbook } from './types';
import type { FileSystemFileHandleLike } from './fsTypes';

/**
 * 保存処理。すべてブラウザ内で完結し、外部へは一切送信しない。
 * ダウンロードは blob: URL 経由なのでネットワークを使わない。
 */

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

export interface SaveNameOverride {
  [bookId: string]: string;
}

async function toBuffer(book: LoadedWorkbook): Promise<ArrayBuffer> {
  return (await book.wb.xlsx.writeBuffer()) as ArrayBuffer;
}

function triggerDownload(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  // すぐに revoke するとダウンロードが始まらないブラウザがあるので少し待つ
  setTimeout(() => URL.revokeObjectURL(url), 30_000);
}

export function downloadBlob(blob: Blob, fileName: string): void {
  triggerDownload(blob, fileName);
}

export function downloadText(text: string, fileName: string, mime: string): void {
  triggerDownload(new Blob([text], { type: `${mime};charset=utf-8` }), fileName);
}

/** 1 ブックをダウンロード保存する */
export async function saveOne(book: LoadedWorkbook, nameOverride?: string): Promise<void> {
  const buf = await toBuffer(book);
  triggerDownload(new Blob([buf], { type: XLSX_MIME }), nameOverride ?? book.fileName);
}

export interface ZipProgress {
  done: number;
  total: number;
  current: string;
}

/** フォルダ構成を保ったまま ZIP にまとめてダウンロードする */
export async function saveAllAsZip(
  books: LoadedWorkbook[],
  zipName: string,
  nameOverrides: SaveNameOverride = {},
  onProgress?: (p: ZipProgress) => void,
): Promise<void> {
  const zip = new JSZip();
  let done = 0;
  for (const book of books) {
    if (book.loadError) continue;
    onProgress?.({ done, total: books.length, current: book.relPath });
    const buf = await toBuffer(book);
    zip.file(applyRename(book.relPath, nameOverrides[book.id]), buf);
    done++;
    await new Promise((r) => setTimeout(r, 0));
  }
  onProgress?.({ done, total: books.length, current: 'ZIP を生成中…' });
  const blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' });
  triggerDownload(blob, zipName);
}

/** 相対パスの末尾のファイル名だけ差し替える */
export function applyRename(relPath: string, newName: string | undefined): string {
  if (!newName) return relPath;
  const idx = relPath.lastIndexOf('/');
  return idx < 0 ? newName : `${relPath.slice(0, idx + 1)}${newName}`;
}

export interface WriteBackResult {
  written: number;
  failed: Array<{ relPath: string; reason: string }>;
}

/**
 * File System Access API で開いたフォルダへ直接上書き保存する。
 * ファイル名が変わる場合は、同じフォルダーに新しい名前で作成する
 * (元のファイルは残す。誤操作時に元へ戻せるようにするため)。
 */
export async function writeBackToDisk(
  books: LoadedWorkbook[],
  nameOverrides: SaveNameOverride = {},
  onProgress?: (p: ZipProgress) => void,
): Promise<WriteBackResult> {
  const result: WriteBackResult = { written: 0, failed: [] };
  let done = 0;
  for (const book of books) {
    done++;
    if (book.loadError) continue;
    if (!book.handle) {
      result.failed.push({
        relPath: book.relPath,
        reason: 'フォルダーから開いていないため書き戻せません',
      });
      continue;
    }
    onProgress?.({ done, total: books.length, current: book.relPath });
    try {
      const newName = nameOverrides[book.id];
      let target = book.handle as unknown as FileSystemFileHandleLike;

      if (newName && newName !== book.fileName) {
        if (!book.parentDir) {
          result.failed.push({
            relPath: book.relPath,
            reason: `ファイル名を「${newName}」へ変更できません (フォルダー情報がありません)`,
          });
          continue;
        }
        target = await book.parentDir.getFileHandle(newName, { create: true });
      }

      const buf = await toBuffer(book);
      const writable = await target.createWritable();
      await writable.write(buf);
      await writable.close();
      result.written++;
      book.dirty = false;
    } catch (e) {
      result.failed.push({
        relPath: book.relPath,
        reason: e instanceof Error ? e.message : String(e),
      });
    }
    await new Promise((r) => setTimeout(r, 0));
  }
  return result;
}
