/** Reads the Lite workbook layout (design-log 0003) into the engine's DeckData. */
import ExcelJS from 'exceljs';
import type { Worksheet } from 'exceljs';
import { SLOT_NAME_RE, validateFormat } from '@deckforge/pptx-engine';
import type { ChartData, DeckData, Issue, Manifest, Scalar, TableCell, TableData, ValueEntry } from '@deckforge/pptx-engine';
import { cellText, cellValue, formatFromNumFmt, isBlank } from './cells.ts';

export const KPI_SHEET = 'kpis';
export const META_SHEET = 'meta';
export const RESERVED_SHEETS = [KPI_SHEET, META_SHEET];
/** Sheets whose name starts with this are notes for people and are never read. */
export const IGNORED_PREFIX = '_';

export interface WorkbookResult {
  data: DeckData;
  issues: Issue[];
  /** Keys whose row exists but whose value cell is empty; they are left out of `data`. */
  blankKeys: string[];
}

const NOT_A_WORKBOOK = 'This file is not an Excel workbook (.xlsx), or it is damaged. Save it from Excel as "Excel Workbook (.xlsx)".';

export async function loadWorkbook(bytes: Uint8Array): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.load(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer);
  } catch {
    throw new Error(NOT_A_WORKBOOK);
  }
  // ExcelJS opens any zip (a .pptx, say) without complaint; a real workbook always has a sheet.
  if (wb.worksheets.length === 0) throw new Error(NOT_A_WORKBOOK);
  return wb;
}

function headerMap(ws: Worksheet): Map<string, number> {
  const m = new Map<string, number>();
  for (let c = 1; c <= ws.columnCount; c++) {
    const h = cellText(ws.getCell(1, c)).replace(/\s+/g, ' ').trim().toLowerCase();
    if (h && !m.has(h)) m.set(h, c);
  }
  return m;
}

const err = (message: string, slot?: string): Issue => ({ severity: 'error', code: 'data-layout', message, ...(slot ? { slot } : {}) });
const warn = (message: string, slot?: string): Issue => ({ severity: 'warning', code: 'data-layout', message, ...(slot ? { slot } : {}) });

const BATCH_NOT_YET = (sheet: string): Issue =>
  err(`Sheet '${sheet}' has a "batch" column. Batch decks (one deck per value) are not available yet; remove the column for now.`);

export async function readWorkbook(bytes: Uint8Array, manifest: Manifest): Promise<WorkbookResult> {
  const wb = await loadWorkbook(bytes);
  const issues: Issue[] = [];
  const sheets = new Map<string, Worksheet>();
  for (const ws of wb.worksheets) sheets.set(ws.name.toLowerCase(), ws);

  for (const slot of [...manifest.tables, ...manifest.charts]) {
    if (RESERVED_SHEETS.includes(slot.name.toLowerCase())) {
      issues.push(err(`The template has a ${slot.kind} named '${slot.name}' (slide ${slot.slide}), but that sheet name is reserved. Rename it in PowerPoint's Selection Pane.`, slot.name));
    }
  }

  const values: Record<string, ValueEntry> = {};
  const origin = new Map<string, string>();
  const blankKeys: string[] = [];
  for (const sheetName of [KPI_SHEET, META_SHEET]) {
    const ws = sheets.get(sheetName);
    if (ws) readKeyValues(ws, sheetName === KPI_SHEET, values, origin, blankKeys, issues);
  }

  const tables: Record<string, TableData> = {};
  for (const slot of manifest.tables) {
    const ws = sheets.get(slot.name.toLowerCase());
    if (ws) tables[slot.name] = readTable(ws, issues);
  }
  const charts: Record<string, ChartData> = {};
  for (const slot of manifest.charts) {
    const ws = sheets.get(slot.name.toLowerCase());
    if (ws) charts[slot.name] = readChart(ws, issues);
  }

  const known = new Set([...RESERVED_SHEETS, ...manifest.tables.map((t) => t.name.toLowerCase()), ...manifest.charts.map((c) => c.name.toLowerCase())]);
  const names = [...manifest.tables, ...manifest.charts].map((s) => s.name);
  for (const ws of wb.worksheets) {
    if (ws.name.startsWith(IGNORED_PREFIX) || known.has(ws.name.toLowerCase())) continue;
    issues.push({
      severity: 'warning',
      code: 'unused-data',
      slot: ws.name,
      message: `Sheet '${ws.name}' does not match any table or chart in the template, so it is ignored. ` +
        (names.length ? `Tables and charts in this template: ${names.join(', ')}.` : 'This template has no named tables or charts.'),
    });
  }

  return { data: { values, tables, charts }, issues, blankKeys };
}

function readKeyValues(
  ws: Worksheet,
  isKpi: boolean,
  values: Record<string, ValueEntry>,
  origin: Map<string, string>,
  blankKeys: string[],
  issues: Issue[],
): void {
  const cols = headerMap(ws);
  const keyCol = cols.get('key');
  const valueCol = cols.get('value');
  if (cols.has('batch')) issues.push(BATCH_NOT_YET(ws.name));
  if (!keyCol || !valueCol) {
    const expected = isKpi ? 'key, value, previous, format' : 'key, value';
    issues.push(err(`Sheet '${ws.name}' needs a header row with the columns ${expected}. Row 1 has: ${[...cols.keys()].join(', ') || '(nothing)'}.`));
    return;
  }
  const prevCol = isKpi ? cols.get('previous') : undefined;
  const fmtCol = isKpi ? cols.get('format') : undefined;

  for (let r = 2; r <= ws.rowCount; r++) {
    const key = cellText(ws.getCell(r, keyCol)).trim();
    if (key === '') continue;
    const where = `sheet '${ws.name}' row ${r}`;
    if (!SLOT_NAME_RE.test(key)) {
      issues.push(warn(`"${key}" (${where}) is not a valid tag name, so no tag can use it. Use letters, digits, _ and - only, no spaces.`, key));
      continue;
    }
    const lower = key.toLowerCase();
    const prev = origin.get(lower);
    if (prev) {
      issues.push(err(`"${key}" appears twice (${prev} and ${where}). Keep only one row for it.`, key));
      continue;
    }
    origin.set(lower, where);

    const valueCell = ws.getCell(r, valueCol);
    const value = cellValue(valueCell);
    if (!value.ok) {
      issues.push(err(`Sheet '${ws.name}' cell ${valueCell.address} ${value.reason}.`, key));
      continue;
    }
    if (isBlank(value.value)) {
      // An empty value is treated as missing, so a half-filled workbook is caught, not silently blank.
      blankKeys.push(lower);
      continue;
    }
    const entry: ValueEntry = { value: value.value };

    if (prevCol) {
      const prevCell = ws.getCell(r, prevCol);
      const p = cellValue(prevCell);
      if (!p.ok) {
        issues.push(err(`Sheet '${ws.name}' cell ${prevCell.address} ${p.reason}.`, key));
        continue;
      }
      if (!isBlank(p.value)) entry.previous = p.value;
    }

    const explicit = fmtCol ? cellText(ws.getCell(r, fmtCol)).trim() : '';
    if (explicit) {
      const problem = validateFormat(explicit);
      if (problem) {
        issues.push(err(`Sheet '${ws.name}' cell ${ws.getCell(r, fmtCol!).address}: ${problem}`, key));
        continue;
      }
      entry.format = explicit;
    } else if (typeof entry.value === 'number') {
      const f = formatFromNumFmt(valueCell.numFmt);
      if (f) entry.format = f;
    }
    values[key] = entry;
  }
}

function readTable(ws: Worksheet, issues: Issue[]): TableData {
  const cols: { name: string; col: number }[] = [];
  for (let c = 1; c <= ws.columnCount; c++) {
    const name = cellText(ws.getCell(1, c)).trim();
    if (name) cols.push({ name, col: c });
  }
  if (cols.some((c) => c.name.toLowerCase() === 'batch')) issues.push(BATCH_NOT_YET(ws.name));

  const rows: TableCell[][] = [];
  for (let r = 2; r <= ws.rowCount; r++) {
    const row: TableCell[] = [];
    let blank = true;
    for (const { col } of cols) {
      const cell = ws.getCell(r, col);
      const v = cellValue(cell);
      if (!v.ok) {
        issues.push(err(`Sheet '${ws.name}' cell ${cell.address} ${v.reason}.`, ws.name));
        row.push(null);
        continue;
      }
      if (!isBlank(v.value)) blank = false;
      const fmt = typeof v.value === 'number' ? formatFromNumFmt(cell.numFmt) : undefined;
      row.push(fmt ? { value: v.value, format: fmt } : v.value);
    }
    if (!blank) rows.push(row);
  }
  return { columns: cols.map((c) => c.name), rows };
}

function readChart(ws: Worksheet, issues: Issue[]): ChartData {
  const categoryLabel = cellText(ws.getCell(1, 1)).trim();
  const series: { name: string; col: number }[] = [];
  for (let c = 2; c <= ws.columnCount; c++) {
    const name = cellText(ws.getCell(1, c)).trim();
    if (!name) break; // series are the contiguous headed columns after A
    series.push({ name, col: c });
  }
  if (series.some((s) => s.name.toLowerCase() === 'batch')) issues.push(BATCH_NOT_YET(ws.name));

  const categories: Scalar[] = [];
  const values: (number | null)[][] = series.map(() => []);
  for (let r = 2; r <= ws.rowCount; r++) {
    const catCell = ws.getCell(r, 1);
    const cat = cellValue(catCell);
    const nums: (number | null)[] = [];
    let blank = !cat.ok || isBlank(cat.value);
    for (const { col } of series) {
      const cell = ws.getCell(r, col);
      const v = cellValue(cell);
      if (!v.ok) {
        issues.push(err(`Sheet '${ws.name}' cell ${cell.address} ${v.reason}.`, ws.name));
        nums.push(null);
        continue;
      }
      if (isBlank(v.value)) {
        nums.push(null);
        continue;
      }
      blank = false;
      const n = typeof v.value === 'number' ? v.value : typeof v.value === 'string' ? Number(v.value.trim()) : Number.NaN;
      if (!Number.isFinite(n)) {
        issues.push(err(`Sheet '${ws.name}' cell ${cell.address} should be a number but contains "${String(v.value)}".`, ws.name));
        nums.push(null);
      } else {
        nums.push(n);
      }
    }
    if (blank) continue;
    if (!cat.ok) issues.push(err(`Sheet '${ws.name}' cell ${catCell.address} ${cat.reason}.`, ws.name));
    categories.push(cat.ok ? cat.value : null);
    nums.forEach((n, i) => values[i]!.push(n));
  }
  return {
    ...(categoryLabel ? { categoryLabel } : {}),
    categories,
    series: series.map((s, i) => ({ name: s.name, values: values[i]! })),
  };
}
