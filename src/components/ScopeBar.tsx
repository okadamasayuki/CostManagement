import { ScopeSelector } from './ui';
import { countTargets } from '../excel/ops';
import { describeScope } from '../recipe/describe';
import { opContext, setState, useStore } from '../state/store';

/**
 * 「この操作をどこに当てるか」を決める、ただ 1 つの場所。
 *
 * 以前はタブごとに「対象」「適用先」が置かれていて、
 *   ・タブを変えるたびに指定し直すのか
 *   ・年度更新の「対象」とロックの「対象」は別物なのか
 * が分からなかった。実際にはずっと同じ 1 つの指定なので、
 * リボンの下に常に出しておき、押す前に「何ブック / 何シートに当たるか」も見せる。
 */
export function ScopeBar() {
  const s = useStore();
  const ready = s.books.length > 0;
  const n = ready ? countTargets(opContext(), s.scope) : { books: 0, sheets: 0 };

  return (
    <div className="scopebar" data-testid="scope-bar">
      <span className="sb-label">🎯 適用先</span>
      <div className="sb-fields">
        <ScopeSelector scope={s.scope} onChange={(scope) => setState({ scope })} inline />
      </div>
      <span className="sb-count" title={describeScope(s.scope)}>
        {ready ? (
          <>
            <b>{n.books}</b> ブック / <b>{n.sheets}</b> シート
          </>
        ) : (
          'ファイル未読み込み'
        )}
      </span>
      <span className="sb-note">
        ボタン横の <b>▾</b> からも、その操作の対象をその場で選べます
      </span>
    </div>
  );
}
