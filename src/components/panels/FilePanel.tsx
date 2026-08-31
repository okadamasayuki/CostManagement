import { useRef, useState } from 'react';
import { BigButton, Btn, Check, Field, Modal, NoteBox, RCol, RGroup } from '../ui';
import { loadFromDirectory, loadFromFiles, isMacroFormat } from '../../excel/loader';
import { supportsDirectoryPicker } from '../../excel/fsTypes';
import { applyRename, downloadBlob, saveAllAsZip, saveOne, writeBackToDisk } from '../../excel/saver';
import { OFFLINE_FILE_NAME, getSelfCopy, isHosted } from '../../security/selfCopy';
import {
  addBooks,
  clearBusy,
  closeAll,
  currentBook,
  getState,
  revertBook,
  runOperation,
  setBusy,
  setState,
  toast,
  useStore,
} from '../../state/store';

export function FilePanel() {
  const s = useStore();
  const fileInput = useRef<HTMLInputElement>(null);
  const dirInput = useRef<HTMLInputElement>(null);
  const [confirmSave, setConfirmSave] = useState<null | 'zip' | 'disk'>(null);

  const book = currentBook();
  const hasBooks = s.books.length > 0;
  const canWriteBack = s.books.some((b) => b.handle);

  async function handleFiles(files: FileList | null) {
    if (!files || !files.length) return;
    setBusy('ファイルを読み込んでいます…');
    try {
      const result = await loadFromFiles(files, (p) =>
        setBusy(`ファイルを読み込んでいます… (${p.done}/${p.total})`, p.done, p.total),
      );
      finishLoad(result.books.length, result.skipped.length, result.books.filter((b) => b.loadError).length);
      addBooks(result.books, result.skipped);
      await applyInitialLock();
    } finally {
      clearBusy();
    }
  }

  async function handleDirectoryPicker() {
    if (!window.showDirectoryPicker) return;
    let dir;
    try {
      dir = await window.showDirectoryPicker({ mode: 'readwrite' });
    } catch {
      return; // ユーザーがキャンセルした
    }
    setBusy('フォルダーを読み込んでいます…');
    try {
      const result = await loadFromDirectory(dir, (p) =>
        setBusy(`フォルダーを読み込んでいます… (${p.done}/${p.total})`, p.done, p.total),
      );
      finishLoad(result.books.length, result.skipped.length, result.books.filter((b) => b.loadError).length);
      addBooks(result.books, result.skipped);
      await applyInitialLock();
    } catch (e) {
      toast('error', 'フォルダーの読み込みに失敗しました', e instanceof Error ? e.message : String(e));
    } finally {
      clearBusy();
    }
  }

  /**
   * 読み込み直後にロック状態をそろえる。
   *
   * Excel のセルは既定で全てロック済みなので、そのままだと
   * 「どこもロックされていない状態から始める」ことができない。
   * ここで一度そろえておくと、あとは入力させたい所だけを
   * 触ればよくなる。
   */
  async function applyInitialLock() {
    const mode = getState().initialLockMode;
    if (mode === 'keep') return;
    const r = await runOperation(
      { op: 'setLock', range: { kind: 'sheet' }, locked: mode === 'lock' },
      { scope: { books: 'all', sheets: 'all' } },
    );
    toast(
      'info',
      mode === 'lock'
        ? '読み込んだ全シートをロックしました'
        : '読み込んだ全シートのロックを外しました',
      r.summary,
    );
  }

  function finishLoad(loaded: number, skipped: number, failed: number) {
    if (loaded === 0) {
      toast('warn', 'Excel ファイルが見つかりませんでした', '.xlsx / .xlsm を含むフォルダーを選んでください。');
      return;
    }
    const parts = [`${loaded} ファイルを読み込みました`];
    if (failed) parts.push(`${failed} 件は読み込めませんでした`);
    if (skipped) parts.push(`${skipped} 件は対象外`);
    toast(failed ? 'warn' : 'success', parts[0], parts.slice(1).join(' / ') || undefined);
  }

  async function doSaveOne() {
    if (!book) return;
    setBusy('保存しています…');
    try {
      await saveOne(book, s.renames[book.id]);
      toast('success', `「${s.renames[book.id] ?? book.fileName}」を保存しました`);
    } catch (e) {
      toast('error', '保存に失敗しました', e instanceof Error ? e.message : String(e));
    } finally {
      clearBusy();
    }
  }

  async function doSaveZip() {
    setConfirmSave(null);
    setBusy('ZIP を作成しています…');
    try {
      const stamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
      await saveAllAsZip(getState().books, `excel_更新後_${stamp}.zip`, getState().renames, (p) =>
        setBusy(`ZIP を作成しています… (${p.done}/${p.total})`, p.done, p.total),
      );
      toast('success', 'ZIP ファイルを保存しました', 'フォルダー構成はそのまま保たれています。');
    } catch (e) {
      toast('error', 'ZIP の作成に失敗しました', e instanceof Error ? e.message : String(e));
    } finally {
      clearBusy();
    }
  }

  async function doWriteBack() {
    setConfirmSave(null);
    setBusy('元のフォルダーへ書き戻しています…');
    try {
      const r = await writeBackToDisk(getState().books, getState().renames, (p) =>
        setBusy(`書き戻しています… (${p.done}/${p.total})`, p.done, p.total),
      );
      if (r.failed.length) {
        toast(
          'warn',
          `${r.written} 件を上書きしました (${r.failed.length} 件は失敗)`,
          r.failed.slice(0, 5).map((f) => `${f.relPath}: ${f.reason}`).join('\n'),
        );
      } else {
        toast('success', `${r.written} 件を元のフォルダーへ上書きしました`);
      }
    } catch (e) {
      toast('error', '書き戻しに失敗しました', e instanceof Error ? e.message : String(e));
    } finally {
      clearBusy();
    }
  }

  const macroBooks = s.books.filter((b) => isMacroFormat(b.fileName));
  const renamedCount = Object.keys(s.renames).length;

  /**
   * このツール自身 (単一 HTML) をダウンロードする。
   * メモリー上の DOM から作るので、通信は一切発生しない。
   */
  function saveToolItself() {
    const html = getSelfCopy();
    if (!html) {
      toast('error', 'ツールの保存に失敗しました', 'お手数ですがブラウザーの「名前を付けて保存」をご利用ください。');
      return;
    }
    downloadBlob(new Blob([html], { type: 'text/html;charset=utf-8' }), OFFLINE_FILE_NAME);
    toast(
      'success',
      'ツール本体を保存しました',
      '共有フォルダーに置けば、以降はネットワークに接続しなくても使えます。',
    );
  }

  return (
    <>
      <input
        ref={fileInput}
        type="file"
        className="hidden-input"
        multiple
        accept=".xlsx,.xlsm,.xltx,.xltm"
        onChange={(e) => {
          void handleFiles(e.target.files);
          e.target.value = '';
        }}
      />
      <input
        ref={dirInput}
        type="file"
        className="hidden-input"
        {...({ webkitdirectory: '', directory: '' } as Record<string, string>)}
        multiple
        onChange={(e) => {
          void handleFiles(e.target.files);
          e.target.value = '';
        }}
      />

      <RGroup title="開く">
        <BigButton icon="📄" label={<>ファイルを<br />開く</>} onClick={() => fileInput.current?.click()} />
        <BigButton
          icon="📁"
          label={<>フォルダーを<br />開く</>}
          primary
          title="選んだフォルダーの下にある Excel ファイルを、サブフォルダーも含めてすべて読み込みます"
          onClick={() =>
            supportsDirectoryPicker() ? void handleDirectoryPicker() : dirInput.current?.click()
          }
        />
        <RCol>
          <div style={{ fontSize: 10.5, color: 'var(--text-dim)', width: 178, lineHeight: 1.6 }}>
            サブフォルダーも含めて
            <br />
            <b>.xlsx / .xlsm</b> をすべて読み込みます。
            <br />
            {supportsDirectoryPicker() ? (
              <span style={{ color: 'var(--ok)' }}>✓ 元のフォルダーへ直接上書き保存できます</span>
            ) : (
              <span style={{ color: 'var(--warn)' }}>
                ※ ZIP でのダウンロード保存になります
              </span>
            )}
          </div>
        </RCol>
      </RGroup>

      <RGroup title="読み込み時のロック">
        <RCol>
          <Field label="">
            <select
              data-testid="initial-lock"
              style={{ width: 186 }}
              value={s.initialLockMode}
              onChange={(e) =>
                setState({ initialLockMode: e.target.value as typeof s.initialLockMode })
              }
            >
              <option value="keep">そのまま (ファイルの設定を使う)</option>
              <option value="unlock">全シートのロックを外す</option>
              <option value="lock">全シートをロックする</option>
            </select>
          </Field>
          <div style={{ width: 190 }}>
            <NoteBox>
              Excel のセルは<b>既定で全てロック済み</b>です。
              「ロックを外す」を選ぶと、読み込んだ時点で
              まっさらな状態から始められます。
            </NoteBox>
          </div>
        </RCol>
      </RGroup>

      <RGroup title="保存">
        <BigButton
          icon="💾"
          label={<>このブックを<br />保存</>}
          disabled={!book}
          onClick={() => void doSaveOne()}
        />
        <BigButton
          icon="🗜️"
          label={<>すべてを<br />ZIP で保存</>}
          disabled={!hasBooks}
          onClick={() => setConfirmSave('zip')}
        />
        <BigButton
          icon="📤"
          label={<>元の場所へ<br />上書き保存</>}
          disabled={!canWriteBack}
          title={
            canWriteBack
              ? '「フォルダーを開く」で読み込んだファイルを、同じ場所へ上書きします'
              : 'この機能を使うには Chrome / Edge で「フォルダーを開く」から読み込んでください'
          }
          onClick={() => setConfirmSave('disk')}
        />
      </RGroup>

      <RGroup title="表示">
        <RCol>
          <Check
            label="ロック状態を色分け表示"
            checked={s.showLockOverlay}
            onChange={(v) => setState({ showLockOverlay: v })}
          />
          <div className="legend">
            <div className="legend-row">
              <span className="legend-chip locked" />
              <span>ロック済み (保護時に編集不可)</span>
            </div>
            <div className="legend-row">
              <span className="legend-chip unlocked" />
              <span>ロック解除 (入力できる)</span>
            </div>
          </div>
        </RCol>
      </RGroup>

      <RGroup title="オフライン利用">
        <BigButton
          icon="⬇️"
          label={<>ツール本体を<br />保存</>}
          primary={isHosted()}
          title="このツール (単一 HTML ファイル) をダウンロードします。通信は発生しません。"
          onClick={saveToolItself}
        />
        <RCol>
          <div style={{ width: 210 }}>
            <NoteBox kind={isHosted() ? 'warn' : 'ok'}>
              {isHosted() ? (
                <>
                  現在は<b>サーバーから開いています</b>。
                  左のボタンで本体を保存し、共有フォルダーなどに置いてお使いください。
                  以降はネットワーク接続なしで動きます。
                </>
              ) : (
                <>
                  <b>✓ ローカルのファイルから起動しています。</b>
                  ネットワークには一切接続していません。
                </>
              )}
            </NoteBox>
          </div>
        </RCol>
      </RGroup>

      <RGroup title="閉じる">
        <RCol>
          <Btn onClick={() => book && void revertBook(book.id)} disabled={!book?.dirty}>
            ↩ 変更を破棄して読み直す
          </Btn>
          <Btn
            kind="danger"
            onClick={() => {
              if (s.books.some((b) => b.dirty) && !confirm('保存していない変更があります。すべて閉じますか?'))
                return;
              closeAll();
              toast('info', 'すべてのブックを閉じました');
            }}
            disabled={!hasBooks}
          >
            ✕ すべて閉じる
          </Btn>
          <div style={{ fontSize: 10.5, color: 'var(--text-dim)' }}>
            {hasBooks ? `${s.books.length} ブック読み込み中` : '未読み込み'}
            {renamedCount > 0 && ` / ${renamedCount} 件はファイル名を変更して保存`}
          </div>
        </RCol>
      </RGroup>

      {confirmSave && (
        <Modal
          title={confirmSave === 'zip' ? 'すべてを ZIP で保存' : '元のフォルダーへ上書き保存'}
          onClose={() => setConfirmSave(null)}
          footer={
            <>
              <Btn onClick={() => setConfirmSave(null)}>キャンセル</Btn>
              <Btn kind="accent" onClick={() => (confirmSave === 'zip' ? void doSaveZip() : void doWriteBack())}>
                {confirmSave === 'zip' ? 'ZIP を作成する' : '上書き保存する'}
              </Btn>
            </>
          }
        >
          <p>
            対象: <b>{s.books.filter((b) => !b.loadError).length} ファイル</b>
            {confirmSave === 'zip'
              ? '（フォルダー構成を保ったまま 1 つの ZIP にまとめます）'
              : '（読み込んだ元のファイルを直接上書きします）'}
          </p>

          {confirmSave === 'disk' && (
            <NoteBox kind="warn">
              <b>元のファイルが置き換わります。</b>
              実行前にフォルダーのバックアップを取ることを強くおすすめします。
            </NoteBox>
          )}

          {renamedCount > 0 && (
            <>
              <h4 style={{ margin: '14px 0 6px' }}>ファイル名が変わるもの ({renamedCount} 件)</h4>
              <table className="plain">
                <thead>
                  <tr>
                    <th>変更前</th>
                    <th>変更後</th>
                  </tr>
                </thead>
                <tbody>
                  {s.books
                    .filter((b) => s.renames[b.id])
                    .slice(0, 12)
                    .map((b) => (
                      <tr key={b.id}>
                        <td>{b.relPath}</td>
                        <td>{applyRename(b.relPath, s.renames[b.id])}</td>
                      </tr>
                    ))}
                </tbody>
              </table>
              {confirmSave === 'disk' && (
                <NoteBox>
                  名前が変わるファイルは<b>同じフォルダーに新しい名前で作成</b>されます。元のファイルは残ります。
                </NoteBox>
              )}
            </>
          )}

          {macroBooks.length > 0 && (
            <>
              <h4 style={{ margin: '14px 0 6px' }}>マクロを含む可能性のあるファイル ({macroBooks.length} 件)</h4>
              <NoteBox kind="err">
                <b>.xlsm / .xltm は保存するとマクロ (VBA) が失われます。</b>
                これらのファイルはロック設定だけを手作業で反映するか、マクロを別途保管してください。
                <ul style={{ margin: '6px 0 0', paddingLeft: 18 }}>
                  {macroBooks.slice(0, 6).map((b) => (
                    <li key={b.id}>{b.relPath}</li>
                  ))}
                  {macroBooks.length > 6 && <li>ほか {macroBooks.length - 6} 件</li>}
                </ul>
              </NoteBox>
            </>
          )}

          <NoteBox>
            保存はすべてブラウザー内で行われます。ファイルの内容が外部へ送信されることはありません。
          </NoteBox>
        </Modal>
      )}
    </>
  );
}
