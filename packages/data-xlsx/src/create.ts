/** Builds a starter workbook in the Lite layout for a template ("Create data file" / `deckforge init-data`). */
import ExcelJS from 'exceljs';
import type { Worksheet } from 'exceljs';
import type { Manifest } from '@deckforge/pptx-engine';
import { KPI_SHEET } from './read.ts';

export interface CreateOptions {
  /** Pre-fill tables and charts with the data the template currently shows. Default true. */
  sample?: boolean;
}

const README = [
  'How to fill this workbook',
  '',
  'kpis sheet: one row per {{tag}} in the template.',
  '  key       the tag name, e.g. oee for {{oee}}',
  '  value     what to show',
  '  previous  optional; feeds {{oee.previous}}, {{oee.change}}, {{oee.change_pct}} and {{oee.trend}}',
  '  format    optional: text, number, percent, currency, date; add decimals like percent:1.',
  '            Leave it empty to use the cell\'s own Excel number format.',
  '',
  'meta sheet (optional): key, value columns for free text tags such as title or period.',
  '',
  'One sheet per table, named like the table in PowerPoint\'s Selection Pane.',
  '  Row 1 holds the column headers, matching the table\'s header row. Rows below are the table body;',
  '  the table grows or shrinks to fit.',
  '',
  'One sheet per chart, named like the chart.',
  '  Column A holds the categories (x axis), columns B onward one series each, names in row 1.',
  '',
  'Sheets whose names start with _ (like this one) are ignored.',
];

function header(ws: Worksheet, values: string[]): void {
  const row = ws.addRow(values);
  row.font = { bold: true };
  ws.views = [{ state: 'frozen', ySplit: 1 }];
}

function autoWidth(ws: Worksheet): void {
  ws.columns.forEach((col) => {
    let max = 8;
    col.eachCell?.({ includeEmpty: false }, (cell) => {
      max = Math.max(max, String(cell.value ?? '').length + 2);
    });
    col.width = Math.min(max, 60);
  });
}

/** Template table text that looks like a plain number becomes a number cell, so Excel treats it as one. */
function sampleCell(s: string): string | number {
  return /^-?\d+(\.\d+)?$/.test(s.trim()) ? Number(s) : s;
}

export async function createWorkbook(manifest: Manifest, options: CreateOptions = {}): Promise<Uint8Array> {
  const sample = options.sample ?? true;
  const wb = new ExcelJS.Workbook();
  wb.creator = 'DeckForge';

  const kpis = wb.addWorksheet(KPI_SHEET);
  header(kpis, ['key', 'value', 'previous', 'format', 'notes']);
  for (const t of manifest.text) {
    const extra = t.fields.filter((f) => f !== 'value').map((f) => `.${f}`);
    const note = `slide ${t.slides.join(', ')}${extra.length ? `; uses ${extra.join(', ')} (fill in previous)` : ''}`;
    kpis.addRow([t.name, null, null, null, note]);
  }
  autoWidth(kpis);

  for (const t of manifest.tables) {
    const ws = wb.addWorksheet(t.name);
    header(ws, t.columns);
    if (sample) for (const r of t.sampleRows) ws.addRow(r.map(sampleCell));
    autoWidth(ws);
  }

  for (const c of manifest.charts) {
    if (!c.chartType) continue;
    const ws = wb.addWorksheet(c.name);
    header(ws, [c.sample.categoryLabel ?? 'Category', ...c.sample.series.map((s) => s.name)]);
    if (sample) {
      c.sample.categories.forEach((cat, i) => ws.addRow([cat, ...c.sample.series.map((s) => s.values[i] ?? null)]));
    }
    autoWidth(ws);
  }

  const help = wb.addWorksheet('_readme');
  for (const line of README) help.addRow([line]);
  help.getRow(1).font = { bold: true, size: 14 };
  help.getColumn(1).width = 110;

  const out = await wb.xlsx.writeBuffer();
  return new Uint8Array(out as ArrayBuffer);
}
