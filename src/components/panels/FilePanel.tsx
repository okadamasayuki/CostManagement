import { useRef, useState } from 'react';
import { BigButton, Btn, Check, Field, Modal, NoteBox, RCol, RGroup } from '../ui';
import { loadFromDirectory, loadFromFiles, loadGenerated, isMacroFormat } from '../../excel/loader';
import { SAMPLES, buildSample, type SampleKind } from '../../excel/samples';
import { supportsDirectoryPicker } from '../../excel/fsTypes';
import { applyRename, downloadBlob, saveAllAsZip, saveOne, writeBackToDisk } from '../../excel/saver';
import { OFFLINE_FILE_NAME, getSelfCopy, isHosted } from '../../security/selfCopy';
import {
  addBooks,
  clearBusy,
  currentBook,
  getState,
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
  const [showSample, setShowSample] = useState(false);

  const book = currentBook();
  const hasBooks = s.books.length > 0;
  const canWriteBack = s.books.some((b) => b.handle);
  /** 名前が変わったとき、元のファイルを消すか (既定は残す) */
  const [deleteRenamedOriginal, setDeleteRenamedOriginal] = useState(false);

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

  /**
   * お試し用のサンプルを作って、そのまま読み込む。
   * Excel が入っていない端末でも操作感を確かめられるようにするためのもの。
   */
  async function loadSample(kind: SampleKind) {
    const info = SAMPLES.find((x) => x.kind === kind);
    setShowSample(false);
    setBusy('サンプルを作っています…', 0, info?.files ?? 30);
    try {
      const files = await buildSample(kind, (done, total) =>
        setBusy(`サンプルを作っています… (${done}/${total})`, done, total),
      );
      const result = await loadGenerated(files, (p) =>
        setBusy(`読み込んでいます… (${p.done}/${p.total})`, p.done, p.total),
      );
      addBooks(result.books, []);
      toast(
        'success',
        `お試し用のサンプルを ${result.books.length} ファイル作りました`,
        `${info?.summary ?? ''}\n※ お試し用なので「元の場所へ上書き保存」はできません。結果を残すときは「ZIP で保存」を使ってください。`,
      );
    } catch (e) {
      toast('error', 'サンプルを作れませんでした', e instanceof Error ? e.message : String(e));
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
      const r = await writeBackToDisk(
        getState().books,
        getState().renames,
        (p) => setBusy(`書き戻しています… (${p.done}/${p.total})`, p.done, p.total),
        { deleteRenamedOriginal },
      );
      /**
       * 名前が変わったものは「上書き」ではなく新しい名前での作成になる。
       * そこを混ぜて数えると、元のファイルが残っていることに気付けないので分けて伝える。
       */
      const parts: string[] = [];
      const overwritten = r.written - r.renamed;
      if (overwritten > 0) parts.push(`${overwritten} 件を上書き`);
      if (r.renamed > 0) {
        parts.push(
          r.removedOriginals > 0
            ? `${r.renamed} 件を新しい名前で保存 (元のファイルは削除)`
            : `${r.renamed} 件を新しい名前で保存 (元のファイルは残っています)`,
        );
      }
      const summary = parts.join(' / ') || '変更はありませんでした';
      if (r.failed.length) {
        toast(
          'warn',
          `${summary} (${r.failed.length} 件は失敗)`,
          r.failed.slice(0, 5).map((f) => `${f.relPath}: ${f.reason}`).join('\n'),
        );
      } else {
        toast(
          'success',
          summary,
          r.renamed > 0 && r.removedOriginals === 0
            ? '元の名前のファイルも同じフォルダーに残っています。'
            : undefined,
        );
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
        <BigButton
          icon="🧪"
          label={<>お試し用の<br />サンプルを作る</>}
          title="動画と同じ 30 ファイルをその場で作ります。Excel が入っていない端末でも操作を試せます"
          onClick={() => setShowSample(true)}
        />
        <div style={{ width: 176 }}>
          <NoteBox>
            <b>Excel が無くても試せます。</b>
            「お試し用のサンプルを作る」で、説明動画と<b>同じ 30 ファイル</b>を
            その場で作って読み込みます。通信は発生しません。
          </NoteBox>
        </div>
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
                  <b>この「ページ」をサーバーから読み込んでいます。</b>
                  読み込んだ Excel の中身がサーバーへ送られることはありません
                  （送信は遮断されています）。
                  <br />
                  社内では、左のボタンで本体を保存して共有フォルダーに置く運用をおすすめします。
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

      {showSample && (
        <Modal
          title="お試し用のサンプルを作る"
          wide
          onClose={() => setShowSample(false)}
          footer={<Btn onClick={() => setShowSample(false)}>閉じる</Btn>}
        >
          <NoteBox kind="ok">
            <b>Excel が入っていない端末でも、操作をひととおり試せます。</b>
            説明動画で使っているのと<b>同じ 30 ファイル</b>を、このツールの中で作って読み込みます。
            外部から取ってくるものは何もありません。
          </NoteBox>

          {SAMPLES.map((info) => (
            <div
              key={info.kind}
              data-testid={`sample-${info.kind}`}
              style={{
                display: 'flex',
                gap: 14,
                alignItems: 'flex-start',
                border: '1px solid var(--border)',
                borderRadius: 4,
                padding: '10px 12px',
                marginTop: 12,
              }}
            >
              <div style={{ flex: 1 }}>
                <div style={{ fontWeight: 700, marginBottom: 4 }}>{info.title}</div>
                <div style={{ color: 'var(--text-dim)', marginBottom: 6 }}>
                  {info.rootFolder}/ …（{info.files} ファイル）
                  <br />
                  {info.summary}
                </div>
                <ul style={{ margin: 0, paddingLeft: 18, lineHeight: 1.7 }}>
                  {info.points.map((p) => (
                    <li key={p}>{p}</li>
                  ))}
                </ul>
              </div>
              <Btn
                kind="accent"
                onClick={() => void loadSample(info.kind)}
                title="作って、そのまま読み込みます"
              >
                作って読み込む
              </Btn>
            </div>
          ))}

          <NoteBox kind="warn">
            お試し用のファイルは<b>パソコンの中には作られません</b>（このツールの中だけにあります）。
            そのため「元の場所へ上書き保存」は使えません。
            結果を Excel で開いて確かめたいときは、<b>「ZIP で保存」</b>で取り出してください。
            <br />
            すでにファイルを読み込んでいる場合は、そこに<b>追加</b>で読み込まれます。
          </NoteBox>
        </Modal>
      )}

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
                <>
                  <NoteBox>
                    名前が変わるファイルは<b>同じフォルダーに新しい名前で作成</b>されます。
                    {deleteRenamedOriginal ? (
                      <>
                        <br />
                        <b>元の名前のファイルは削除します。</b>
                      </>
                    ) : (
                      <>
                        <br />
                        元の名前のファイルは<b>そのまま残ります</b>
                        （このフォルダーには新旧の両方が並びます）。
                      </>
                    )}
                  </NoteBox>
                  <Check
                    label="元の名前のファイルを削除する"
                    checked={deleteRenamedOriginal}
                    onChange={setDeleteRenamedOriginal}
                    title="前年の分を残しておきたい場合はオフのままにしてください"
                  />
                </>
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
