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
      <span className="sb-count">
        {ready ? (
          <>
            <b>{n.books}</b> ブック / <b>{n.sheets}</b> シート が対象
          </>
        ) : (
          'ファイルを読み込むと対象が出ます'
        )}
      </span>
      <span className="sb-note">
        ロック・書式・年度更新など<b>すべての操作</b>で、ここの指定が使われます
      </span>
      <span className="sb-desc" title={describeScope(s.scope)}>
        {describeScope(s.scope)}
      </span>
    </div>
  );
}
