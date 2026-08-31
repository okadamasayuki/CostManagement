import { useEffect, useState } from 'react';
import { NoteBox, RCol, RGroup } from '../ui';
import { getBlockedAttempts, onBlockedAttempt, type BlockedAttempt } from '../../security/networkGuard';
import { isHosted } from '../../security/selfCopy';

/**
 * 社内利用時の説明用パネル。
 * 「本当に外部へ出ていないのか」を情報システム部門に説明できるよう、
 * 何をどう遮断しているかと、遮断の実績を表示する。
 */
export function SecurityPanel() {
  const [blocked, setBlocked] = useState<BlockedAttempt[]>(getBlockedAttempts());
  useEffect(() => onBlockedAttempt(setBlocked), []);

  return (
    <>
      <RGroup title="通信の遮断">
        <RCol>
          <div style={{ width: 300 }}>
            <NoteBox kind={blocked.length ? 'warn' : 'ok'}>
              {blocked.length === 0 ? (
                <>
                  <b>✓ 外部通信は発生していません。</b>
                  <br />
                  読み込んだファイルはブラウザーのメモリー内だけで処理され、
                  保存もローカルへの書き出しのみです。
                </>
              ) : (
                <>
                  <b>{blocked.length} 件の通信を遮断しました。</b>
                  <br />
                  下の一覧を確認してください。
                </>
              )}
            </NoteBox>
          </div>
          <div style={{ width: 300, marginTop: 8 }}>
            <NoteBox kind={isHosted() ? 'warn' : 'ok'}>
              {isHosted() ? (
                <>
                  <b>起動元: サーバー ({location.host})</b>
                  <br />
                  ページ自体はこのサーバーから読み込まれています。
                  Excel の中身が送信されることはありませんが、
                  社内規程で外部サイトの利用が制限されている場合は、
                  「ファイル」タブの<b>「ツール本体を保存」</b>で
                  ダウンロードして共有フォルダーからお使いください。
                </>
              ) : (
                <>
                  <b>起動元: ローカルファイル</b>
                  <br />
                  サーバーを経由していません。ネットワークから完全に切り離された状態です。
                </>
              )}
            </NoteBox>
          </div>
        </RCol>
      </RGroup>

      <RGroup title="対策の内容">
        <div style={{ width: 380, fontSize: 11, lineHeight: 1.75 }}>
          <table className="plain">
            <tbody>
              <tr>
                <th style={{ width: 130 }}>CSP</th>
                <td>
                  <code>connect-src &apos;none&apos;</code> でブラウザー自体が通信を禁止
                </td>
              </tr>
              <tr>
                <th>通信 API</th>
                <td>fetch / XHR / WebSocket / EventSource / sendBeacon / WebRTC を起動時に無効化</td>
              </tr>
              <tr>
                <th>外部リソース</th>
                <td>CDN・Web フォント・解析タグを一切使用しない (単一 HTML に同梱)</td>
              </tr>
              <tr>
                <th>ファイルの保存</th>
                <td>ブラウザーのダウンロード機能とローカルフォルダーへの書き込みのみ</td>
              </tr>
              <tr>
                <th>データの保持</th>
                <td>タブを閉じるとメモリー上のデータは消える (サーバー保存なし)</td>
              </tr>
            </tbody>
          </table>
        </div>
      </RGroup>

      <RGroup title="遮断した通信">
        <div style={{ width: 300 }}>
          {blocked.length === 0 ? (
            <div style={{ fontSize: 11, color: 'var(--text-dim)', padding: '6px 2px' }}>
              記録はありません。
            </div>
          ) : (
            <ul className="detail-list">
              {blocked
                .slice()
                .reverse()
                .map((b, i) => (
                  <li key={i}>
                    <b>{b.api}</b>
                    <span className="where">
                      {b.at.toLocaleTimeString('ja-JP')} — {b.target}
                    </span>
                  </li>
                ))}
            </ul>
          )}
        </div>
      </RGroup>

      <RGroup title="確認のしかた">
        <div style={{ width: 300 }}>
          <NoteBox>
            ブラウザーで <b>F12</b> →「ネットワーク」タブを開いたまま操作しても、
            通信が 1 件も記録されないことを確認できます。
            <br />
            オフライン (LAN ケーブルを抜いた状態) でもすべての機能が動きます。
          </NoteBox>
        </div>
      </RGroup>
    </>
  );
}
