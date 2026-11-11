import type { DeckData, Manifest, RowRecord, ValueEntry } from './types.ts';

/**
 * Data that fills every slot of a template with what the template already shows
 * (table rows and chart values from the template itself; text tags get a placeholder).
 * Used to pre-fill a new data workbook, and as smoke-test input.
 */
export function sampleData(manifest: Manifest): DeckData {
  const values: Record<string, ValueEntry> = {};
  for (const t of manifest.text) {
    const needsNumbers = t.fields.some((f) => f === 'change' || f === 'change_pct' || f === 'trend');
    values[t.name] = needsNumbers || t.fields.includes('previous') ? { value: 110, previous: 100 } : { value: `[${t.name}]` };
  }
  const tables: NonNullable<DeckData['tables']> = {};
  for (const t of manifest.tables) {
    tables[t.name] = { columns: [...t.columns], rows: t.sampleRows.map((r) => [...r]) };
  }
  const charts: NonNullable<DeckData['charts']> = {};
  for (const c of manifest.charts) {
    if (!c.chartType) continue;
    charts[c.name] = {
      categories: [...c.sample.categories],
      series: c.sample.series.map((s) => ({ name: s.name, values: [...s.values] })),
    };
  }
  const rows: Record<string, RowRecord[]> = {};
  for (const r of manifest.repeats) {
    rows[r.name] = [1, 2].map((n) => {
      const rec: RowRecord = {};
      for (const t of r.tags) {
        const numeric = t.fields.some((f) => f !== 'value');
        rec[t.name] = numeric ? { value: 110 * n, previous: 100 } : `[${t.name} ${n}]`;
      }
      return rec;
    });
  }
  return { values, tables, charts, rows };
}
