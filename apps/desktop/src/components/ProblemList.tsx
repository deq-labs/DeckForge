import type { ProblemView } from '../../electron/api.ts';

export function ProblemList({ problems, title }: { problems: ProblemView[]; title?: string }): React.JSX.Element | null {
  if (problems.length === 0) return null;
  return (
    <div className="problems">
      {title && <h3>{title}</h3>}
      <ul>
        {problems.map((p, i) => (
          <li key={`${i}:${p.message}`} className={p.severity}>
            <span className="badge">{p.severity === 'error' ? 'Fix' : 'Note'}</span>
            <div>
              <div>{p.message}</div>
              {p.hint && <div className="hint">{p.hint}</div>}
            </div>
          </li>
        ))}
      </ul>
    </div>
  );
}
