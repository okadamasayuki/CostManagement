/**
 * File System Access API の型定義。
 * TypeScript の標準 lib には未収録なので最小限だけ宣言する。
 * この API を使うと「開いたフォルダにそのまま上書き保存」ができる
 * (Chrome / Edge のみ。Firefox / Safari では ZIP ダウンロードにフォールバック)
 */

export interface FileSystemHandleLike {
  kind: 'file' | 'directory';
  name: string;
}

export interface FileSystemFileHandleLike extends FileSystemHandleLike {
  kind: 'file';
  getFile(): Promise<File>;
  createWritable(): Promise<{
    write(data: BufferSource | Blob | string): Promise<void>;
    close(): Promise<void>;
  }>;
}

export interface FileSystemDirectoryHandleLike extends FileSystemHandleLike {
  kind: 'directory';
  values(): AsyncIterableIterator<FileSystemFileHandleLike | FileSystemDirectoryHandleLike>;
  getFileHandle(name: string, options?: { create?: boolean }): Promise<FileSystemFileHandleLike>;
  getDirectoryHandle(
    name: string,
    options?: { create?: boolean },
  ): Promise<FileSystemDirectoryHandleLike>;
  removeEntry?(name: string, options?: { recursive?: boolean }): Promise<void>;
  queryPermission?(descriptor: { mode: 'read' | 'readwrite' }): Promise<PermissionState>;
  requestPermission?(descriptor: { mode: 'read' | 'readwrite' }): Promise<PermissionState>;
}

declare global {
  interface Window {
    showDirectoryPicker?: (options?: {
      mode?: 'read' | 'readwrite';
    }) => Promise<FileSystemDirectoryHandleLike>;
  }
}

export function supportsDirectoryPicker(): boolean {
  return typeof window.showDirectoryPicker === 'function';
}
