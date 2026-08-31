/**
 * 外部通信ガード (多層防御の 2 層目)
 *
 * index.html の CSP (connect-src 'none') がブラウザレベルで通信を
 * 遮断しているが、それに加えてアプリ起動時に通信 API 自体を
 * 無効化しておく。目的は 2 つ:
 *
 *  1. 万一 CSP が効かない環境 (古いブラウザ、file:// の特殊な設定など)
 *     でも情報が外へ出ないようにする。
 *  2. 「遮断されたこと」を画面上で可視化し、監査・情報システム部門への
 *     説明材料にする。ブロック件数はステータスバーに表示される。
 *
 * 本アプリは Excel の読み書きを全てブラウザのメモリ内で行うため、
 * 通信 API は一切必要ない。したがって全面的に潰して問題ない。
 */

export interface BlockedAttempt {
  api: string;
  target: string;
  at: Date;
  /**
   * 'send'     … 外へデータを送ろうとした (本当に警告すべきもの)
   * 'resource' … 外部リソースの読み込みが止まっただけ
   *              (ブラウザーが探すアイコン、拡張機能が差し込む画像など)
   */
  kind: 'send' | 'resource';
}

const blocked: BlockedAttempt[] = [];
const listeners = new Set<(list: BlockedAttempt[]) => void>();

function record(api: string, target: unknown, kind: BlockedAttempt['kind'] = 'send'): void {
  blocked.push({ api, target: String(target).slice(0, 300), at: new Date(), kind });
  listeners.forEach((fn) => fn([...blocked]));
  // 開発者が気付けるようコンソールにも残す
  console.warn(`[外部通信ガード] ${api} への呼び出しを遮断しました:`, target);
}

export function onBlockedAttempt(fn: (list: BlockedAttempt[]) => void): () => void {
  listeners.add(fn);
  return () => listeners.delete(fn);
}

export function getBlockedAttempts(): BlockedAttempt[] {
  return [...blocked];
}

/** 外へデータを送ろうとした試みだけを返す */
export function getSendAttempts(): BlockedAttempt[] {
  return blocked.filter((b) => b.kind === 'send');
}

/** 画面に出す説明文 */
export function describeAttempt(b: BlockedAttempt): string {
  return `${b.at.toLocaleTimeString('ja-JP')}  ${b.api} → ${b.target}`;
}

class NetworkBlockedError extends Error {
  constructor(api: string) {
    super(
      `${api} は本ツールでは使用できません。` +
        `このツールは完全オフライン動作で、いかなる外部通信も行いません。`,
    );
    this.name = 'NetworkBlockedError';
  }
}

/** アプリの最初期に 1 度だけ呼ぶ。以降あらゆる通信 API が使用不能になる。 */
export function installNetworkGuard(): void {
  const w = window as unknown as Record<string, unknown>;

  const define = (name: string, value: unknown) => {
    try {
      Object.defineProperty(w, name, {
        value,
        writable: false,
        configurable: false,
      });
    } catch {
      /* 再定義できない環境では CSP に任せる */
    }
  };

  define('fetch', (input: unknown) => {
    record('fetch', input);
    return Promise.reject(new NetworkBlockedError('fetch'));
  });

  define('sendBeacon', undefined);
  try {
    Object.defineProperty(navigator, 'sendBeacon', {
      value: (url: unknown) => {
        record('navigator.sendBeacon', url);
        return false;
      },
      configurable: false,
    });
  } catch {
    /* noop */
  }

  class BlockedXHR {
    open(_method: string, url: string) {
      record('XMLHttpRequest', url);
      throw new NetworkBlockedError('XMLHttpRequest');
    }
    send() {
      throw new NetworkBlockedError('XMLHttpRequest');
    }
    setRequestHeader() {
      /* noop */
    }
    addEventListener() {
      /* noop */
    }
  }
  define('XMLHttpRequest', BlockedXHR);

  class BlockedWebSocket {
    constructor(url: string) {
      record('WebSocket', url);
      throw new NetworkBlockedError('WebSocket');
    }
  }
  define('WebSocket', BlockedWebSocket);

  class BlockedEventSource {
    constructor(url: string) {
      record('EventSource', url);
      throw new NetworkBlockedError('EventSource');
    }
  }
  define('EventSource', BlockedEventSource);

  class BlockedRTC {
    constructor() {
      record('RTCPeerConnection', '(WebRTC)');
      throw new NetworkBlockedError('RTCPeerConnection');
    }
  }
  define('RTCPeerConnection', BlockedRTC);
  define('webkitRTCPeerConnection', BlockedRTC);

  // CSP 違反も記録する。
  // ただし connect-src / form-action 以外は「読み込みが止まった」だけで、
  // データが外へ出ようとしたわけではないので区別する。
  const SEND_DIRECTIVES = ['connect-src', 'form-action'];
  window.addEventListener('securitypolicyviolation', (e) => {
    const kind = SEND_DIRECTIVES.some((d) => e.violatedDirective.startsWith(d))
      ? 'send'
      : 'resource';
    record(`CSP:${e.violatedDirective}`, e.blockedURI, kind);
  });
}
