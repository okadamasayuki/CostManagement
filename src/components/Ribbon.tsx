import type { RibbonTab } from '../state/store';
import { setState, useStore } from '../state/store';
import { FilePanel } from './panels/FilePanel';
import { LockPanel } from './panels/LockPanel';
import { FormatPanel } from './panels/FormatPanel';
import { YearPanel } from './panels/YearPanel';
import { RecipePanel } from './panels/RecipePanel';

const TABS: Array<{ id: RibbonTab; label: string }> = [
  { id: 'file', label: 'ファイル' },
  { id: 'lock', label: 'ロック' },
  { id: 'format', label: '書式・色' },
  { id: 'year', label: '年度更新' },
  { id: 'recipe', label: '手順書' },
];

export function Ribbon() {
  const s = useStore();
  return (
    <div className="ribbon">
      <div className="ribbon-tabs">
        {TABS.map((t) => (
          <button
            key={t.id}
            className={`ribbon-tab${s.activeTab === t.id ? ' active' : ''}`}
            onClick={() => setState({ activeTab: t.id })}
            type="button"
          >
            {t.label}
          </button>
        ))}
      </div>
      <div className="ribbon-panel">
        {s.activeTab === 'file' && <FilePanel />}
        {s.activeTab === 'lock' && <LockPanel />}
        {s.activeTab === 'format' && <FormatPanel />}
        {s.activeTab === 'year' && <YearPanel />}
        {s.activeTab === 'recipe' && <RecipePanel />}
      </div>
    </div>
  );
}
