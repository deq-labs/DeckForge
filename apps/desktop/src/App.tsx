import { useCallback, useEffect, useState } from 'react';
import type { CheckView, GenerateView, JobOptions, JobSpec, LicenseView, PreviewTool, Recipe, TemplateInfo } from '../electron/api.ts';
import { DEFAULT_OPTIONS, LOCALES, countBySeverity, defaultOutDir, newId, plural, slotStatus, splitPath, stem } from './logic.ts';
import { ProblemList } from './components/ProblemList.tsx';
import { RecipeList } from './components/RecipeList.tsx';
import { LicensePanel } from './components/LicensePanel.tsx';

type Busy = null | 'inspect' | 'data' | 'check' | 'generate' | 'preview';

const api = window.deckforge;

export function App(): React.JSX.Element {
  const [template, setTemplate] = useState<TemplateInfo | null>(null);
  const [dataPath, setDataPath] = useState('');
  const [outDir, setOutDir] = useState('');
  const [options, setOptions] = useState<JobOptions>(DEFAULT_OPTIONS);
  const [checked, setChecked] = useState<CheckView | null>(null);
  const [result, setResult] = useState<GenerateView | null>(null);
  const [busy, setBusy] = useState<Busy>(null);
  const [error, setError] = useState<string | null>(null);
  const [recipes, setRecipes] = useState<Recipe[]>([]);
  const [recipeId, setRecipeId] = useState<string | null>(null);
  const [recipeName, setRecipeName] = useState('');
  const [license, setLicense] = useState<LicenseView | null>(null);
  const [showLicense, setShowLicense] = useState(false);
  const [previewTool, setPreviewTool] = useState<PreviewTool | null>(null);

  useEffect(() => {
    void api.recipes.list().then(setRecipes);
    void api.license.status().then(setLicense);
    void api.previewTool().then(setPreviewTool);
  }, []);

  // The main process can open a job by URL hash (used by the automated UI smoke test).
  useEffect(() => {
    const m = /^#job=(.+)$/.exec(window.location.hash);
    if (!m) return;
    const spec = JSON.parse(decodeURIComponent(m[1]!)) as JobSpec;
    void openRecipe({ id: 'launch', name: stem(spec.templatePath), spec }, true);
  }, []);

  /** Runs an action with a busy state; errors become a banner instead of a crash. */
  const act = useCallback(async <T,>(kind: Busy, f: () => Promise<T>): Promise<T | undefined> => {
    setBusy(kind);
    setError(null);
    try {
      return await f();
    } catch (e) {
      setError((e as Error).message);
      return undefined;
    } finally {
      setBusy(null);
    }
  }, []);

  const resetResults = (): void => {
    setChecked(null);
    setResult(null);
  };

  const loadTemplate = async (path: string, keepOutDir = false): Promise<void> => {
    const info = await act('inspect', () => api.inspect(path));
    if (!info) return;
    setTemplate(info);
    if (!keepOutDir) setOutDir(defaultOutDir(path));
    resetResults();
  };

  const spec: JobSpec | null = template && dataPath && outDir ? { templatePath: template.path, dataPath, outDir, options } : null;

  const chooseTemplate = async (): Promise<void> => {
    const p = await api.pickTemplate();
    if (p) {
      setRecipeId(null);
      setRecipeName(stem(p));
      await loadTemplate(p);
    }
  };
  const chooseData = async (): Promise<void> => {
    const p = await api.pickData();
    if (p) {
      setDataPath(p);
      resetResults();
    }
  };
  const createData = async (): Promise<void> => {
    if (!template) return;
    const p = await act('data', () => api.createDataFile(template.path));
    if (p) {
      setDataPath(p);
      resetResults();
    }
  };
  const chooseOut = async (): Promise<void> => {
    const p = await api.pickFolder(outDir || undefined);
    if (p) setOutDir(p);
  };
  const runCheck = async (): Promise<void> => {
    if (!spec) return;
    const r = await act('check', () => api.check(spec));
    if (r) {
      setChecked(r);
      setResult(null);
    }
  };
  const runGenerate = async (s: JobSpec | null = spec, rid: string | null = recipeId): Promise<void> => {
    if (!s) return;
    const r = await act('generate', () => api.generate(s));
    if (!r) return;
    setChecked(r);
    setResult(r);
    if (r.ok && rid) {
      const existing = recipes.find((x) => x.id === rid);
      if (existing) setRecipes(await api.recipes.save({ ...existing, spec: s, lastRunAt: new Date().toISOString(), lastDeckCount: r.files.length }));
    }
  };
  const saveRecipe = async (): Promise<void> => {
    if (!spec) return;
    const id = recipeId ?? newId();
    const prev = recipes.find((r) => r.id === id);
    const saved = await act(null, () => api.recipes.save({ ...prev, id, name: recipeName.trim() || stem(spec.templatePath), spec }));
    if (saved) {
      setRecipes(saved);
      setRecipeId(id);
    }
  };
  const openRecipe = async (r: Recipe, andRun: boolean): Promise<void> => {
    setRecipeId(r.id);
    setRecipeName(r.name);
    setDataPath(r.spec.dataPath);
    setOutDir(r.spec.outDir);
    setOptions(r.spec.options);
    await loadTemplate(r.spec.templatePath, true);
    if (andRun) await runGenerate(r.spec, r.id);
  };
  const removeRecipe = async (id: string): Promise<void> => {
    setRecipes(await api.recipes.remove(id));
    if (recipeId === id) setRecipeId(null);
  };
  const newJob = (): void => {
    setTemplate(null);
    setDataPath('');
    setOutDir('');
    setOptions(DEFAULT_OPTIONS);
    setRecipeId(null);
    setRecipeName('');
    setError(null);
    resetResults();
  };

  const counts = checked ? countBySeverity(checked.problems) : null;
  const trial = license?.mode !== 'licensed';

  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">
          <span className="logo" aria-hidden>◆</span> DeckForge <span className="edition">Lite</span>
        </div>
        <button className={`license-badge ${trial ? 'trial' : 'licensed'}`} onClick={() => setShowLicense(true)}>
          {license ? (trial ? 'Trial' : `Licensed${license.licensee ? ` to ${license.licensee}` : ''}`) : '…'}
        </button>
      </header>

      <aside className="sidebar">
        <button className="primary block" onClick={newJob}>+ New job</button>
        <RecipeList recipes={recipes} activeId={recipeId} busy={busy !== null} onOpen={(r) => void openRecipe(r, false)} onRun={(r) => void openRecipe(r, true)} onRemove={(id) => void removeRecipe(id)} />
      </aside>

      <main className="main">
        {error && (
          <div className="banner error" role="alert">
            <span>{error}</span>
            <button className="link" onClick={() => setError(null)}>Dismiss</button>
          </div>
        )}
        {trial && license && (
          <div className="banner info">
            {license.summary} <button className="link" onClick={() => setShowLicense(true)}>Enter a license key</button>
          </div>
        )}

        <section className="card">
          <h2><span className="step">1</span> Template</h2>
          <div className="row">
            <button onClick={() => void chooseTemplate()} disabled={busy !== null}>{template ? 'Change…' : 'Choose template…'}</button>
            <span className="path" title={template?.path}>{template ? template.fileName : 'A PowerPoint file with {{tags}} and named tables/charts'}</span>
          </div>
          {template && (
            <div className="slots">
              <div className="muted">{plural(template.slideCount, 'slide')} · {plural(template.slots.length, 'item')} to fill</div>
              <table className="slot-table">
                <thead>
                  <tr><th /><th>Name</th><th>Kind</th><th>Where</th><th>Needs</th></tr>
                </thead>
                <tbody>
                  {template.slots.map((s) => {
                    const st = slotStatus(s, checked);
                    return (
                      <tr key={`${s.kind}:${s.name}`} className={`status-${st}`}>
                        <td className="status" title={st}>{st === 'ok' ? '✓' : st === 'problem' ? '✗' : st === 'unsupported' ? '–' : '○'}</td>
                        <td className="mono">{s.kind === 'text' ? `{{${s.name}}}` : s.name}</td>
                        <td>{s.kind}</td>
                        <td>{s.where}</td>
                        <td className="muted">{s.detail}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
              {template.rules.length > 0 && <div className="muted small">Slide rules: {template.rules.join(' · ')}</div>}
              {template.problems.length > 0 && <ProblemList problems={template.problems} title="About this template" />}
            </div>
          )}
        </section>

        <section className={`card ${template ? '' : 'disabled'}`}>
          <h2><span className="step">2</span> Data</h2>
          <div className="row">
            <button onClick={() => void chooseData()} disabled={!template || busy !== null}>{dataPath ? 'Change…' : 'Choose workbook…'}</button>
            <button onClick={() => void createData()} disabled={!template || busy !== null}>{busy === 'data' ? 'Creating…' : 'Create data file…'}</button>
            <span className="path" title={dataPath}>{dataPath ? splitPath(dataPath).file : 'An Excel workbook laid out for this template'}</span>
          </div>
        </section>

        <section className={`card ${template ? '' : 'disabled'}`}>
          <h2><span className="step">3</span> Output</h2>
          <div className="row">
            <button onClick={() => void chooseOut()} disabled={!template || busy !== null}>Folder…</button>
            <span className="path" title={outDir}>{outDir || '—'}</span>
          </div>
          <div className="options">
            <label>Numbers and dates
              <select value={options.locale} onChange={(e) => setOptions({ ...options, locale: e.target.value })}>
                {LOCALES.map((l) => <option key={l.code} value={l.code}>{l.label}</option>)}
              </select>
            </label>
            <label>Currency
              <input value={options.currency} maxLength={3} onChange={(e) => setOptions({ ...options, currency: e.target.value.toUpperCase() })} />
            </label>
            <label>Missing data
              <select value={options.missing} onChange={(e) => setOptions({ ...options, missing: e.target.value as JobOptions['missing'] })}>
                <option value="error">Stop and tell me</option>
                <option value="keep">Leave {'{{tags}}'} visible</option>
                <option value="blank">Leave blank</option>
              </select>
            </label>
          </div>
        </section>

        <section className="actions">
          <button onClick={() => void runCheck()} disabled={!spec || busy !== null}>{busy === 'check' ? 'Checking…' : 'Check'}</button>
          <button className="primary" onClick={() => void runGenerate()} disabled={!spec || busy !== null}>{busy === 'generate' ? 'Generating…' : 'Generate decks'}</button>
          <span className="spacer" />
          <input className="recipe-name" placeholder="Recipe name" value={recipeName} onChange={(e) => setRecipeName(e.target.value)} disabled={!spec} />
          <button onClick={() => void saveRecipe()} disabled={!spec || busy !== null}>{recipes.some((r) => r.id === recipeId) ? 'Update recipe' : 'Save as recipe'}</button>
        </section>

        {checked && (
          <section className="card">
            <h2>{result ? 'Result' : 'Check'}</h2>
            <p className={checked.ok ? 'ok' : 'bad'}>
              {checked.ok
                ? result
                  ? `${plural(result.files.length, 'deck')} written${result.watermarked ? ' (trial footer added)' : ''}.`
                  : `Ready: ${plural(checked.deckCount, 'deck')} would be made${checked.batches.length ? ` (${checked.batches.join(', ')})` : ''}.`
                : `${plural(counts!.errors, 'problem')} to fix before generating.`}
              {counts && counts.warnings > 0 && ` ${plural(counts.warnings, 'warning')}.`}
            </p>
            {result && result.files.length > 0 && (
              <ul className="files">
                {result.files.map((f) => (
                  <li key={f}>
                    <span className="mono">{splitPath(f).file}</span>
                    <button className="link" onClick={() => void act(null, () => api.openPath(f))}>Open</button>
                    {previewTool && (
                      <button className="link" disabled={busy !== null} onClick={() => void act('preview', () => api.preview(f))}>
                        {busy === 'preview' ? 'Rendering…' : 'Preview'}
                      </button>
                    )}
                    <button className="link" onClick={() => void act(null, () => api.showInFolder(f))}>Show in folder</button>
                  </li>
                ))}
              </ul>
            )}
            <ProblemList problems={checked.problems} />
          </section>
        )}
      </main>

      {showLicense && license && <LicensePanel license={license} onChange={setLicense} onClose={() => setShowLicense(false)} />}
    </div>
  );
}
