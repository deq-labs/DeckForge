import type { Recipe } from '../../electron/api.ts';
import { relativeTime, splitPath } from '../logic.ts';

interface Props {
  recipes: Recipe[];
  activeId: string | null;
  busy: boolean;
  onOpen: (r: Recipe) => void;
  onRun: (r: Recipe) => void;
  onRemove: (id: string) => void;
}

export function RecipeList({ recipes, activeId, busy, onOpen, onRun, onRemove }: Props): React.JSX.Element {
  return (
    <div className="recipes">
      <h3>Recipes</h3>
      {recipes.length === 0 && <p className="muted small">Save a job as a recipe to rerun it next month with one click.</p>}
      <ul>
        {recipes.map((r) => (
          <li key={r.id} className={r.id === activeId ? 'active' : ''}>
            <button className="recipe" onClick={() => onOpen(r)} title={`${r.spec.templatePath}\n${r.spec.dataPath}`}>
              <span className="name">{r.name}</span>
              <span className="muted small">{splitPath(r.spec.dataPath).file} · {relativeTime(r.lastRunAt)}</span>
            </button>
            <div className="recipe-actions">
              <button className="link" disabled={busy} onClick={() => onRun(r)}>Run</button>
              <button className="link danger" disabled={busy} onClick={() => onRemove(r.id)}>Remove</button>
            </div>
          </li>
        ))}
      </ul>
    </div>
  );
}
