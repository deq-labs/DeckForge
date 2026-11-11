/** Builds the small workbook embedded in each chart, so "Edit Data" in PowerPoint shows the filled data. */
import ExcelJS from 'exceljs';
import type { ChartData } from './types.ts';

export async function buildChartWorkbook(sheetName: string, categories: string[], data: ChartData, formatCodes: string[]): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'DeckForge';
  const ws = wb.addWorksheet(sheetName);
  ws.addRow([data.categoryLabel ?? '', ...data.series.map((s) => s.name)]);
  categories.forEach((cat, r) => {
    ws.addRow([cat, ...data.series.map((s) => s.values[r] ?? null)]);
  });
  data.series.forEach((_, i) => {
    const fmt = formatCodes[i];
    if (fmt && fmt !== 'General') ws.getColumn(i + 2).numFmt = fmt;
  });
  const out = await wb.xlsx.writeBuffer();
  return new Uint8Array(out as ArrayBuffer);
}
