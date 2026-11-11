/**
 * Reads the Lite workbook layout (design-log 0003, 0004, 0005) into the engine's DeckData,
 * one DeckData per batch.
 */
import ExcelJS from 'exceljs';
import type { Cell, Worksheet } from 'exceljs';
import { SLOT_NAME_RE, validateFormat } from '@deckforge/pptx-engine';
import type { ChartData, DeckData, Issue, Manifest, RowRecord, Scalar, TableCell, TableData, ValueEntry } from '@deckforge/pptx-engine';
import { cellText, cellValue, formatFromNumFmt, isBlank } from './cells.ts';

export const KPI_SHEET = 'kpis';
export const META_SHEET = 'meta';
export const BATCH_COLUMN = 'batch';
export const RESERVED_SHEETS = [KPI_SHEET, META_SHEET];
/** Sheets whose name starts with this are notes for people and are never read. */
export const IGNORED_PREFIX = '_';

export interface BatchDeck {
  /** The batch value, or null when the workbook has no batch column. */
  batch: string | null;
  data: DeckData;
}

export interface WorkbookResult {
  /** One entry per batch value, in the order they first appear; a single entry without batches. */
  decks: BatchDeck[];
  issues: Issue[];
  /** Keys whose row exists but whose value cell is empty; they are left out of the data. */
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

/** "Plant Manager" -> "plant_manager": how a repeat sheet's headers become tag names. */
export function headerToTag(header: string): string {
  return header.trim().toLowerCase().replace(/[^a-z0-9_-]+/g, '_').replace(/^_+|_+$/g, '');
}

const err = (message: string, slot?: string): Issue => ({ severity: 'error', code: 'data-layout', message, ...(slot ? { slot } : {}) });
const warn = (message: string, slot?: string): Issue => ({ severity: 'warning', code: 'data-layout', message, ...(slot ? { slot } : {}) });

interface Tagged<T> {
  batch: string | null;
  item: T;
}

/** Header text by column, and the batch column if there is one. */
function headers(ws: Worksheet): { cols: { name: string; col: number }[]; batchCol: number | undefined } {
  const cols: { name: string; col: number }[] = [];
  let batchCol: number | undefined;
  for (let c = 1; c <= ws.columnCount; c++) {
    const name = cellText(ws.getCell(1, c)).replace(/\s+/g, ' ').trim();
    if (!name) continue;
    if (name.toLowerCase() === BATCH_COLUMN && batchCol === undefined) batchCol = c;
    else cols.push({ name, col: c });
  }
  return { cols, batchCol };
}

function batchOf(ws: Worksheet, r: number, batchCol: number | undefined): string | null {
  if (batchCol === undefined) return null;
  const t = cellText(ws.getCell(r, batchCol)).trim();
  return t === '' ? null : t;
}

/** A numeric cell keeps its Excel number format as a DeckForge format. */
function withFormat(cell: Cell, value: Scalar): TableCell {
  const fmt = typeof value === 'number' ? formatFromNumFmt(cell.numFmt) : undefined;
  return fmt ? { value, format: fmt } : value;
}

export async function readWorkbook(bytes: Uint8Array, manifest: Manifest): Promise<WorkbookResult> {
  const wb = await loadWorkbook(bytes);
  const issues: Issue[] = [];
  const sheets = new Map<string, Worksheet>();
  for (const ws of wb.worksheets) sheets.set(ws.name.toLowerCase(), ws);
  const batchOrder: string[] = [];
  const noteBatch = (b: string | null): void => {
    if (b !== null && !batchOrder.includes(b)) batchOrder.push(b);
  };

  const named = [...manifest.tables, ...manifest.charts, ...manifest.repeats];
  for (const slot of named) {
    if (RESERVED_SHEETS.includes(slot.name.toLowerCase())) {
      const where = slot.kind === 'repeat' ? `"#repeat ${slot.name}"` : `a ${slot.kind} named '${slot.name}'`;
      issues.push(err(`The template has ${where}, but "${slot.name}" is a reserved sheet name. Rename it in the template.`, slot.name));
    }
  }

  // kpis / meta
  const kv: Tagged<{ key: string; entry: ValueEntry; where: string }>[] = [];
  const blankKeys: string[] = [];
  const seenKey = new Map<string, string>();
  for (const sheetName of [KPI_SHEET, META_SHEET]) {
    const ws = sheets.get(sheetName);
    if (ws) readKeyValues(ws, sheetName === KPI_SHEET, kv, seenKey, blankKeys, issues, noteBatch);
  }

  const tables = new Map<string, { columns: string[]; rows: Tagged<TableCell[]>[] }>();
  for (const slot of manifest.tables) {
    const ws = sheets.get(slot.name.toLowerCase());
    if (ws) tables.set(slot.name, readTable(ws, issues, noteBatch));
  }
  const charts = new Map<string, { categoryLabel: string; names: string[]; points: Tagged<{ category: Scalar; values: (number | null)[] }>[] }>();
  for (const slot of manifest.charts) {
    const ws = sheets.get(slot.name.toLowerCase());
    if (ws) charts.set(slot.name, readChart(ws, issues, noteBatch));
  }
  const repeats = new Map<string, Tagged<RowRecord>[]>();
  for (const slot of manifest.repeats) {
    const ws = sheets.get(slot.name.toLowerCase());
    if (ws) repeats.set(slot.name, readRepeat(ws, issues, noteBatch));
  }

  const known = new Set([...RESERVED_SHEETS, ...named.map((s) => s.name.toLowerCase())]);
  const names = named.map((s) => s.name);
  for (const ws of wb.worksheets) {
    if (ws.name.startsWith(IGNORED_PREFIX) || known.has(ws.name.toLowerCase())) continue;
    issues.push({
      severity: 'warning',
      code: 'unused-data',
      slot: ws.name,
      message:
        `Sheet '${ws.name}' does not match any table, chart or repeated slide in the template, so it is ignored. ` +
        (names.length ? `Sheets this template uses: ${names.join(', ')}.` : 'This template only uses the kpis sheet.'),
    });
  }

  // Assemble one DeckData per batch: shared rows (no batch) plus that batch's rows.
  const scopes: (string | null)[] = batchOrder.length ? batchOrder : [null];
  const inScope = (b: string | null, scope: string | null): boolean => b === null || b === scope;
  const decks = scopes.map((scope): BatchDeck => {
    const values: Record<string, ValueEntry> = {};
    const lowerToKey = new Map<string, string>();
    const put = (key: string, entry: ValueEntry): void => {
      const existing = lowerToKey.get(key.toLowerCase());
      if (existing !== undefined) delete values[existing];
      lowerToKey.set(key.toLowerCase(), key);
      values[key] = entry;
    };
    // Shared rows first; this batch's rows then override them.
    for (const { batch, item } of kv) if (batch === null) put(item.key, item.entry);
    if (scope !== null) for (const { batch, item } of kv) if (batch === scope) put(item.key, item.entry);
    const t: Record<string, TableData> = {};
    for (const [name, tbl] of tables) t[name] = { columns: tbl.columns, rows: tbl.rows.filter((r) => inScope(r.batch, scope)).map((r) => r.item) };
    const c: Record<string, ChartData> = {};
    for (const [name, ch] of charts) {
      const pts = ch.points.filter((p) => inScope(p.batch, scope)).map((p) => p.item);
      c[name] = {
        ...(ch.categoryLabel ? { categoryLabel: ch.categoryLabel } : {}),
        categories: pts.map((p) => p.category),
        series: ch.names.map((n, i) => ({ name: n, values: pts.map((p) => p.values[i] ?? null) })),
      };
    }
    const rows: Record<string, RowRecord[]> = {};
    for (const [name, rs] of repeats) rows[name] = rs.filter((r) => inScope(r.batch, scope)).map((r) => r.item);
    return { batch: scope, data: { values, tables: t, charts: c, rows } };
  });

  return { decks, issues, blankKeys };
}

function readKeyValues(
  ws: Worksheet,
  isKpi: boolean,
  out: Tagged<{ key: string; entry: ValueEntry; where: string }>[],
  seenKey: Map<string, string>,
  blankKeys: string[],
  issues: Issue[],
  noteBatch: (b: string | null) => void,
): void {
  const { cols, batchCol } = headers(ws);
  const col = (name: string): number | undefined => cols.find((c) => c.name.toLowerCase() === name)?.col;
  const keyCol = col('key');
  const valueCol = col('value');
  if (!keyCol || !valueCol) {
    const expected = isKpi ? 'key, value, previous, format' : 'key, value';
    issues.push(err(`Sheet '${ws.name}' needs a header row with the columns ${expected}. Row 1 has: ${cols.map((c) => c.name).join(', ') || '(nothing)'}.`));
    return;
  }
  const prevCol = isKpi ? col('previous') : undefined;
  const fmtCol = isKpi ? col('format') : undefined;

  for (let r = 2; r <= ws.rowCount; r++) {
    const key = cellText(ws.getCell(r, keyCol)).trim();
    if (key === '') continue;
    const where = `sheet '${ws.name}' row ${r}`;
    if (!SLOT_NAME_RE.test(key)) {
      issues.push(warn(`"${key}" (${where}) is not a valid tag name, so no tag can use it. Use letters, digits, _ and - only, no spaces.`, key));
      continue;
    }
    const batch = batchOf(ws, r, batchCol);
    noteBatch(batch);
    // Duplicates are checked before the row's own validity, so both problems get reported.
    const id = `${batch ?? ''}|${key.toLowerCase()}`;
    const prev = seenKey.get(id);
    if (prev) {
      issues.push(err(`"${key}" appears twice${batch ? ` for batch "${batch}"` : ''} (${prev} and ${where}). Keep only one row for it.`, key));
      continue;
    }
    seenKey.set(id, where);

    const valueCell = ws.getCell(r, valueCol);
    const value = cellValue(valueCell);
    if (!value.ok) {
      issues.push(err(`Sheet '${ws.name}' cell ${valueCell.address} ${value.reason}.`, key));
      continue;
    }
    if (isBlank(value.value)) {
      // An empty value is treated as missing, so a half-filled workbook is caught, not silently blank.
      if (!blankKeys.includes(key.toLowerCase())) blankKeys.push(key.toLowerCase());
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
    out.push({ batch, item: { key, entry, where } });
  }
}

function readTable(ws: Worksheet, issues: Issue[], noteBatch: (b: string | null) => void): { columns: string[]; rows: Tagged<TableCell[]>[] } {
  const { cols, batchCol } = headers(ws);
  const rows: Tagged<TableCell[]>[] = [];
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
      row.push(withFormat(cell, v.value));
    }
    if (blank) continue;
    const batch = batchOf(ws, r, batchCol);
    noteBatch(batch);
    rows.push({ batch, item: row });
  }
  return { columns: cols.map((c) => c.name), rows };
}

function readChart(
  ws: Worksheet,
  issues: Issue[],
  noteBatch: (b: string | null) => void,
): { categoryLabel: string; names: string[]; points: Tagged<{ category: Scalar; values: (number | null)[] }>[] } {
  const { cols, batchCol } = headers(ws);
  const categoryLabel = cellText(ws.getCell(1, 1)).trim();
  // Series: the contiguous headed columns after A (a batch column anywhere is skipped).
  const series: { name: string; col: number }[] = [];
  for (let c = 2; c <= ws.columnCount; c++) {
    if (c === batchCol) continue;
    const h = cols.find((x) => x.col === c);
    if (!h) break;
    series.push(h);
  }

  const points: Tagged<{ category: Scalar; values: (number | null)[] }>[] = [];
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
    const batch = batchOf(ws, r, batchCol);
    noteBatch(batch);
    points.push({ batch, item: { category: cat.ok ? cat.value : null, values: nums } });
  }
  return { categoryLabel, names: series.map((s) => s.name), points };
}

function readRepeat(ws: Worksheet, issues: Issue[], noteBatch: (b: string | null) => void): Tagged<RowRecord>[] {
  const { cols, batchCol } = headers(ws);
  const tagCols = cols
    .map((c) => ({ ...c, tag: headerToTag(c.name) }))
    .filter((c) => {
      if (SLOT_NAME_RE.test(c.tag)) return true;
      issues.push(warn(`Sheet '${ws.name}' column "${c.name}" can't be used as a tag name, so it is ignored. Use letters, digits and spaces, starting with a letter.`, ws.name));
      return false;
    });
  const out: Tagged<RowRecord>[] = [];
  for (let r = 2; r <= ws.rowCount; r++) {
    const rec: RowRecord = {};
    const previous = new Map<string, Scalar>();
    let blank = true;
    for (const c of tagCols) {
      const cell = ws.getCell(r, c.col);
      const v = cellValue(cell);
      if (!v.ok) {
        issues.push(err(`Sheet '${ws.name}' cell ${cell.address} ${v.reason}.`, ws.name));
        continue;
      }
      if (isBlank(v.value)) continue;
      blank = false;
      if (c.tag.endsWith('_previous')) {
        previous.set(c.tag.slice(0, -'_previous'.length), v.value);
        continue;
      }
      const tc = withFormat(cell, v.value);
      rec[c.tag] = tc !== null && typeof tc === 'object' && !(tc instanceof Date) ? tc : { value: v.value };
    }
    if (blank) continue;
    for (const [tag, prev] of previous) {
      const e = rec[tag];
      if (e && typeof e === 'object' && !(e instanceof Date)) e.previous = prev;
      else rec[`${tag}_previous`] = prev;
    }
    const batch = batchOf(ws, r, batchCol);
    noteBatch(batch);
    out.push({ batch, item: rec });
  }
  return out;
}
