import ExcelJS from 'exceljs';
import type { LoadedWorkbook, SkippedFile } from './types';
import type { FileSystemDirectoryHandleLike, FileSystemFileHandleLike } from './fsTypes';

/** 読み込む拡張子 */
const EXCEL_EXT = /\.(xlsx|xlsm|xltx|xltm)$/i;
/** マクロや特殊機能を含む可能性があり、保存時に欠落しうる形式 */
const MACRO_EXT = /\.(xlsm|xltm)$/i;
/** Excel が作る一時ファイル */
const TEMP_FILE = /(^|\/)~\$/;

export interface LoadResult {
  books: LoadedWorkbook[];
  skipped: SkippedFile[];
}

export interface LoadProgress {
  done: number;
  total: number;
  current: string;
}

let idSeq = 0;
function nextId(): string {
  idSeq += 1;
  return `wb${idSeq}`;
}

export function isMacroFormat(fileName: string): boolean {
  return MACRO_EXT.test(fileName);
}

async function parseOne(
  file: File,
  relPath: string,
  handle?: FileSystemFileHandleLike,
  parentDir?: FileSystemDirectoryHandleLike,
): Promise<LoadedWorkbook> {
  const base: LoadedWorkbook = {
    id: nextId(),
    relPath,
    fileName: file.name,
    wb: new ExcelJS.Workbook(),
    handle: handle as unknown as FileSystemFileHandle | undefined,
    parentDir,
    sourceFile: file,
    dirty: false,
    sizeBytes: file.size,
  };
  try {
    const buf = await file.arrayBuffer();
    await base.wb.xlsx.load(buf);
  } catch (e) {
    base.loadError = e instanceof Error ? e.message : String(e);
  }
  return base;
}

/**
 * <input type="file"> / <input webkitdirectory> から読み込む。
 * フォルダ選択時は webkitRelativePath に階層が入るので、それを相対パスとして使う。
 */
export async function loadFromFiles(
  fileList: FileList | File[],
  onProgress?: (p: LoadProgress) => void,
): Promise<LoadResult> {
  const files = Array.from(fileList);
  const books: LoadedWorkbook[] = [];
  const skipped: SkippedFile[] = [];

  const targets = files.filter((f) => {
    const path = (f as File & { webkitRelativePath?: string }).webkitRelativePath || f.name;
    if (TEMP_FILE.test(path)) {
      skipped.push({ relPath: path, reason: 'Excel の一時ファイル (~$)' });
      return false;
    }
    if (!EXCEL_EXT.test(f.name)) {
      skipped.push({ relPath: path, reason: 'Excel ファイルではない' });
      return false;
    }
    return true;
  });

  let done = 0;
  for (const f of targets) {
    const relPath = (f as File & { webkitRelativePath?: string }).webkitRelativePath || f.name;
    onProgress?.({ done, total: targets.length, current: relPath });
    books.push(await parseOne(f, relPath));
    done++;
    // UI を固まらせないため 1 件ごとに制御を返す
    await yieldToUi();
  }
  onProgress?.({ done, total: targets.length, current: '' });
  return { books, skipped };
}

/**
 * File System Access API のディレクトリハンドルから再帰的に読み込む。
 * ファイルハンドルを保持するので、同じ場所へ上書き保存できる。
 */
export async function loadFromDirectory(
  dir: FileSystemDirectoryHandleLike,
  onProgress?: (p: LoadProgress) => void,
): Promise<LoadResult> {
  const found: Array<{
    handle: FileSystemFileHandleLike;
    relPath: string;
    parent: FileSystemDirectoryHandleLike;
  }> = [];
  const skipped: SkippedFile[] = [];

  async function walk(d: FileSystemDirectoryHandleLike, prefix: string): Promise<void> {
    for await (const entry of d.values()) {
      const path = prefix ? `${prefix}/${entry.name}` : entry.name;
      if (entry.kind === 'directory') {
        await walk(entry as FileSystemDirectoryHandleLike, path);
      } else {
        if (TEMP_FILE.test(entry.name)) {
          skipped.push({ relPath: path, reason: 'Excel の一時ファイル (~$)' });
          continue;
        }
        if (!EXCEL_EXT.test(entry.name)) continue;
        found.push({ handle: entry as FileSystemFileHandleLike, relPath: path, parent: d });
      }
    }
  }
  await walk(dir, dir.name);

  const books: LoadedWorkbook[] = [];
  let done = 0;
  for (const item of found) {
    onProgress?.({ done, total: found.length, current: item.relPath });
    const file = await item.handle.getFile();
    books.push(await parseOne(file, item.relPath, item.handle, item.parent));
    done++;
    await yieldToUi();
  }
  onProgress?.({ done, total: found.length, current: '' });
  return { books, skipped };
}

function yieldToUi(): Promise<void> {
  return new Promise((r) => setTimeout(r, 0));
}
