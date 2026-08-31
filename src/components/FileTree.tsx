import { useMemo, useState } from 'react';
import type { LoadedWorkbook } from '../excel/types';
import { isMacroFormat } from '../excel/loader';
import { closeBook, selectBook, setState, toast, useStore } from '../state/store';
import { isUnderFolder } from '../excel/folders';

/** 相対パスからフォルダー階層のツリーを組み立てる */
interface TreeNode {
  name: string;
  path: string;
  folders: Map<string, TreeNode>;
  files: LoadedWorkbook[];
}

function buildTree(books: LoadedWorkbook[]): TreeNode {
  const root: TreeNode = { name: '', path: '', folders: new Map(), files: [] };
  for (const b of books) {
    const parts = b.relPath.split('/');
    const fileName = parts.pop() ?? b.fileName;
    void fileName;
    let node = root;
    for (const part of parts) {
      let child = node.folders.get(part);
      if (!child) {
        child = { name: part, path: `${node.path}/${part}`, folders: new Map(), files: [] };
        node.folders.set(part, child);
      }
      node = child;
    }
    node.files.push(b);
  }
  return root;
}

export function FileTree() {
  const s = useStore();
  const [collapsed, setCollapsed] = useState<Set<string>>(new Set());
  const tree = useMemo(() => buildTree(s.books), [s.books]);

  if (!s.books.length) {
    return (
      <div className="sidebar">
        <div className="side-header">ファイル</div>
        <div className="empty-hint">
          ファイルが読み込まれていません。
          <br />
          <br />
          「ファイル」タブの
          <br />
          <b>「フォルダーを開く」</b>
          <br />
          から、対象フォルダーを選んでください。
          <br />
          <br />
          サブフォルダーの中の Excel も
          <br />
          まとめて読み込まれます。
        </div>
      </div>
    );
  }

  const toggle = (path: string) => {
    const next = new Set(collapsed);
    if (next.has(path)) next.delete(path);
    else next.add(path);
    setCollapsed(next);
  };

  const renderNode = (node: TreeNode, depth: number): React.ReactNode[] => {
    const out: React.ReactNode[] = [];
    for (const folder of node.folders.values()) {
      const isCollapsed = collapsed.has(folder.path);
      // 先頭の '/' を落として、相対パスの表記に合わせる
      const folderPath = folder.path.replace(/^\//, '');
      const isTarget = s.scope.books === 'folder' && (s.scope.bookFolder ?? '') === folderPath;
      const fileCount = s.books.filter((b) => isUnderFolder(b.relPath, folderPath)).length;
      out.push(
        <div
          key={folder.path}
          className={`tree-folder${isTarget ? ' target' : ''}`}
          style={{ paddingLeft: 8 + depth * 12 }}
          onClick={() => toggle(folder.path)}
        >
          <span>{isCollapsed ? '▸' : '▾'}</span>
          <span>📁</span>
          <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', flex: 1 }}>{folder.name}</span>
          <button
            className="icon-btn"
            title={
              isTarget
                ? 'このフォルダーを対象にしています'
                : `このフォルダー配下の ${fileCount} ブックを操作の対象にする`
            }
            onClick={(e) => {
              e.stopPropagation();
              setState({
                scope: { ...s.scope, books: 'folder', bookFolder: folderPath },
              });
              toast(
                'info',
                `「${folder.name}」配下の ${fileCount} ブックを対象にしました`,
                'ロック・色・年度更新の操作がこの範囲に適用されます。',
              );
            }}
          >
            {isTarget ? '🎯' : '⊙'}
          </button>
        </div>,
      );
      if (!isCollapsed) out.push(...renderNode(folder, depth + 1));
    }
    for (const b of node.files) {
      const rename = s.renames[b.id];
      out.push(
        <div
          key={b.id}
          className={`tree-file${b.id === s.currentBookId ? ' active' : ''}`}
          style={{ paddingLeft: 10 + depth * 12 }}
          onClick={() => selectBook(b.id)}
          title={`${b.relPath}\n${(b.sizeBytes / 1024).toFixed(0)} KB${
            rename ? `\n保存時のファイル名: ${rename}` : ''
          }`}
        >
          <span>{b.loadError ? '⚠️' : '📗'}</span>
          <span className="name">{rename ?? b.fileName}</span>
          {b.loadError && <span className="badge err">失敗</span>}
          {isMacroFormat(b.fileName) && !b.loadError && <span className="badge macro">マクロ</span>}
          {rename && <span className="badge renamed">改名</span>}
          {b.dirty && <span className="badge dirty">未保存</span>}
          <button
            className="icon-btn"
            title="このブックを閉じる"
            onClick={(e) => {
              e.stopPropagation();
              closeBook(b.id);
            }}
          >
            ✕
          </button>
        </div>,
      );
    }
    return out;
  };

  const dirty = s.books.filter((b) => b.dirty).length;

  return (
    <div className="sidebar">
      <div className="side-header">
        <span>ファイル ({s.books.length})</span>
        <span className="spacer" />
        {dirty > 0 && <span className="badge dirty">未保存 {dirty}</span>}
      </div>
      {s.scope.books === 'folder' && (
        <div className="scope-banner" title="リボンの「適用先」で変更できます">
          🎯 対象:{' '}
          <b>{s.scope.bookFolder ? `${s.scope.bookFolder}/` : 'すべて'}</b>
          <button
            className="icon-btn"
            title="フォルダーの指定をやめて、選択中のブックのみに戻す"
            onClick={() => setState({ scope: { ...s.scope, books: 'current' } })}
          >
            ✕
          </button>
        </div>
      )}
      <div className="tree">{renderNode(tree, 0)}</div>
      {s.skipped.length > 0 && (
        <div
          style={{
            padding: '6px 10px',
            fontSize: 10.5,
            color: 'var(--text-dim)',
            borderTop: '1px solid var(--border)',
          }}
          title={s.skipped.slice(0, 20).map((x) => `${x.relPath} — ${x.reason}`).join('\n')}
        >
          対象外として除いたファイル: {s.skipped.length} 件
        </div>
      )}
    </div>
  );
}
