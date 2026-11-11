/**
 * Builds the template packs: copies each PowerPoint-built template into template-packs/<pack>/
 * and writes a realistic sample workbook next to it, starting from the layout that
 * `deckforge init-data` would create (so the packs always match the product's own contract).
 *
 * Run: pnpm packs        Then check: pnpm --filter @deckforge/packs test
 * Preview images (optional, needs PowerPoint): see template-packs/README.md.
 */
import { copyFileSync, mkdirSync, readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import ExcelJS from 'exceljs';
import type { Worksheet } from 'exceljs';
import { Template } from '@deckforge/pptx-engine';
import { createWorkbook } from '@deckforge/data-xlsx';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
export const PACKS_DIR = join(root, 'template-packs');

type Cell = string | number | Date | null;
interface Kpi {
  key: string;
  value: Cell;
  previous?: Cell;
  /** Excel number format for the value/previous cells (DeckForge carries it over). */
  numFmt?: string;
  /** Batch value, for packs that make one deck per client/site. */
  batch?: string;
}
interface PackSpec {
  name: string;
  template: string;
  kpis: Kpi[];
  /** Sheet name -> rows (row 0 = header). Number formats per column index. */
  sheets: Record<string, { rows: Cell[][]; numFmts?: Record<number, string> }>;
}

const d = (y: number, m: number, day: number): Date => new Date(Date.UTC(y, m - 1, day));

const PACKS: PackSpec[] = [
  {
    name: 'ops-review',
    template: 'ppt-ops-review.pptx',
    kpis: [
      { key: 'title', value: 'October 2026 Operations Review' },
      { key: 'site', value: 'Riverside Plant' },
      { key: 'period_start', value: d(2026, 10, 1), numFmt: 'mmm d, yyyy' },
      { key: 'period_end', value: d(2026, 10, 31), numFmt: 'mmm d, yyyy' },
      { key: 'oee', value: 0.784, previous: 0.751, numFmt: '0.0%' },
      { key: 'output_units', value: 48210, previous: 45120, numFmt: '#,##0' },
      { key: 'scrap_rate', value: 0.018, previous: 0.021, numFmt: '0.0%' },
      { key: 'downtime_hours', value: 21.5, previous: 26, numFmt: '0.0' },
      {
        key: 'highlights',
        value:
          'Line 3 changeover time down 20% after SMED kaizen\nNew resin supplier qualified; first lots arrive Nov 10\nAction: PM backlog on Line 2 press to zero by Nov 15 (owner: M. Ortiz)',
      },
    ],
    sheets: {
      downtime: {
        rows: [
          ['Line', 'Cause', 'Minutes', 'Share'],
          ['Line 3', 'Material shortage', 380, 0.29],
          ['Line 1', 'Changeover', 340, 0.26],
          ['Line 2', 'Press jam', 300, 0.23],
          ['Line 4', 'Sensor fault', 170, 0.13],
          ['Line 1', 'Planned maintenance', 100, 0.08],
        ],
        numFmts: { 2: '#,##0', 3: '0%' },
      },
      downtime_by_line: {
        rows: [
          ['Line', 'Planned', 'Unplanned'],
          ['Line 1', 100, 340],
          ['Line 2', 60, 300],
          ['Line 3', 45, 380],
          ['Line 4', 30, 170],
        ],
      },
      oee_trend: {
        rows: [
          ['Month', 'OEE', 'Target'],
          ['May', 0.712, 0.75],
          ['Jun', 0.738, 0.75],
          ['Jul', 0.694, 0.75],
          ['Aug', 0.756, 0.75],
          ['Sep', 0.751, 0.75],
          ['Oct', 0.784, 0.75],
        ],
        numFmts: { 1: '0%', 2: '0%' },
      },
    },
  },
  {
    name: 'kpi-scorecard',
    template: 'ppt-kpi-scorecard.pptx',
    kpis: [
      { key: 'title', value: 'Quality & Delivery Scorecard' },
      { key: 'period', value: 'Q3 2026 · July – September' },
      { key: 'overall_status', value: 'on track, with 4 of 5 metrics at or better than target' },
      { key: 'incidents', value: 1 },
      { key: 'incident_notes', value: 'Minor hand laceration at Line 2 packing (Aug 14). Guard redesign approved; install scheduled Nov 3.' },
    ],
    sheets: {
      scorecard: {
        rows: [
          ['Metric', 'Target', 'Actual', 'Status'],
          ['On-time delivery', 0.95, 0.968, 'On track'],
          ['First pass yield', 0.98, 0.972, 'Watch'],
          ['Scrap rate', 0.02, 0.017, 'On track'],
          ['Customer complaints', 3, 2, 'On track'],
          ['Safety incidents', 0, 1, 'Off track'],
        ],
        // Mixed units per row: percents are pre-formatted by row below.
      },
      scrap_by_cause: {
        rows: [
          ['Cause', 'Scrap (kg)'],
          ['Setup', 412],
          ['Material', 296],
          ['Machine', 251],
          ['Operator', 138],
          ['Other', 57],
        ],
        numFmts: { 1: '#,##0' },
      },
      volume_vs_yield: {
        rows: [
          ['Month', 'Volume', 'Yield'],
          ['Jul', 12040, 0.962],
          ['Aug', 13380, 0.958],
          ['Sep', 12810, 0.971],
        ],
        numFmts: { 1: '#,##0', 2: '0.0%' },
      },
    },
  },
  {
    name: 'client-monthly',
    template: 'ppt-client-monthly.pptx',
    kpis: [
      // Shared by every client deck (empty batch):
      { key: 'month', value: 'October 2026' },
      { key: 'agency', value: 'Northline Marketing' },
      // One set per client:
      ...(
        [
          ['Harbor Dental', 'Harbor Dental Group', 18450, 16200, 212, 187, 0.0115, 0.0115, 7200, 6900],
          ['Summit Outdoor', 'Summit Outdoor Co.', 40210, 43890, 388, 402, 0.0096, 0.0092, 12500, 12500],
          ['Vela Coffee', 'Vela Coffee Roasters', 9870, 8120, 145, 98, 0.0147, 0.0121, 3100, 2600],
        ] as const
      ).flatMap(([batch, client, s, sp, l, lp, c, cp, sp2, spp]) => [
        { key: 'client', value: client, batch },
        { key: 'sessions', value: s, previous: sp, numFmt: '#,##0', batch },
        { key: 'leads', value: l, previous: lp, numFmt: '#,##0', batch },
        { key: 'conversion_rate', value: c, previous: cp, numFmt: '0.00%', batch },
        { key: 'spend', value: sp2, previous: spp, numFmt: '"$"#,##0', batch },
        { key: 'notes', value: `Next month for ${client}: shift 10% of budget to the best-converting channel and refresh ad creative.`, batch },
      ]),
    ],
    sheets: {
      leads_by_channel: {
        rows: [
          ['Channel', 'This month', 'Last month', 'batch'],
          ['Organic search', 81, 70, 'Harbor Dental'],
          ['Paid search', 74, 66, 'Harbor Dental'],
          ['Social', 31, 30, 'Harbor Dental'],
          ['Email', 26, 21, 'Harbor Dental'],
          ['Organic search', 142, 151, 'Summit Outdoor'],
          ['Paid search', 128, 140, 'Summit Outdoor'],
          ['Social', 79, 74, 'Summit Outdoor'],
          ['Email', 39, 37, 'Summit Outdoor'],
          ['Organic search', 52, 38, 'Vela Coffee'],
          ['Social', 61, 40, 'Vela Coffee'],
          ['Email', 32, 20, 'Vela Coffee'],
        ],
      },
      campaigns: {
        rows: [
          ['Campaign', 'Spend', 'Leads', 'Cost per lead', 'batch'],
          ['Spring whitening promo', 4200, 61, 68.85, 'Harbor Dental'],
          ['Brand search', 2100, 54, 38.89, 'Harbor Dental'],
          ['Retargeting', 900, 17, 52.94, 'Harbor Dental'],
          ['Fall gear launch', 7800, 172, 45.35, 'Summit Outdoor'],
          ['Brand search', 3100, 140, 22.14, 'Summit Outdoor'],
          ['Holiday early access', 1600, 41, 39.02, 'Summit Outdoor'],
          ['Subscription push', 1900, 77, 24.68, 'Vela Coffee'],
          ['Instagram reels', 1200, 51, 23.53, 'Vela Coffee'],
        ],
        numFmts: { 1: '"$"#,##0', 3: '"$"#,##0.00' },
      },
    },
  },
];

function header(ws: Worksheet): Map<string, number> {
  const m = new Map<string, number>();
  ws.getRow(1).eachCell((c, col) => m.set(String(c.value).toLowerCase(), col));
  return m;
}

/** Writes rows from row 1 down, overwriting what's there, and blanks any rows left below. */
function writeRows(ws: Worksheet, rows: Cell[][], firstRow = 1): void {
  const lastOld = ws.rowCount;
  rows.forEach((r, i) => {
    const row = ws.getRow(firstRow + i);
    row.values = r as ExcelJS.CellValue[];
    row.commit();
  });
  for (let n = firstRow + rows.length; n <= lastOld; n++) ws.getRow(n).values = [];
}

function fillSheet(ws: Worksheet, rows: Cell[][], numFmts: Record<number, string> = {}): void {
  writeRows(ws, rows);
  ws.getRow(1).font = { bold: true };
  for (let n = 2; n <= rows.length; n++) {
    for (const [col, fmt] of Object.entries(numFmts)) ws.getRow(n).getCell(Number(col) + 1).numFmt = fmt;
  }
  ws.columns.forEach((c) => (c.width = 18));
}

export async function buildPack(p: PackSpec): Promise<void> {
  const dir = join(PACKS_DIR, p.name);
  mkdirSync(dir, { recursive: true });
  const templatePath = join(dir, 'template.pptx');
  copyFileSync(join(root, 'fixtures/templates', p.template), templatePath);
  const manifest = (await Template.load(new Uint8Array(readFileSync(templatePath)))).slots();

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load((await createWorkbook(manifest, { sample: false })).buffer as ArrayBuffer);

  // kpis: rewrite the rows (keeping the header), adding a batch column only if the pack uses one.
  const kpis = wb.getWorksheet('kpis')!;
  const usesBatch = p.kpis.some((k) => k.batch);
  const cols = header(kpis);
  const notes = new Map<string, unknown>();
  kpis.eachRow((row, n) => {
    if (n > 1) notes.set(String(row.getCell(1).value), row.getCell(cols.get('notes')!).value);
  });
  if (usesBatch) kpis.getRow(1).getCell(6).value = 'batch';
  writeRows(
    kpis,
    p.kpis.map((k) => [k.key, k.value, k.previous ?? null, null, (notes.get(k.key) as Cell) ?? null, ...(usesBatch ? [k.batch ?? null] : [])]),
    2,
  );
  p.kpis.forEach((k, i) => {
    if (!k.numFmt) return;
    const row = kpis.getRow(i + 2);
    row.getCell(2).numFmt = k.numFmt;
    row.getCell(3).numFmt = k.numFmt;
  });
  for (const [name, sheet] of Object.entries(p.sheets)) {
    const ws = wb.getWorksheet(name) ?? wb.addWorksheet(name);
    fillSheet(ws, sheet.rows, sheet.numFmts);
  }
  // Scorecard: targets/actuals mix percents and counts, so format by row.
  if (p.name === 'kpi-scorecard') {
    const ws = wb.getWorksheet('scorecard')!;
    ws.eachRow((row, n) => {
      if (n === 1) return;
      const pct = typeof row.getCell(2).value === 'number' && (row.getCell(2).value as number) < 1 && (row.getCell(2).value as number) > 0;
      if (pct) for (const c of [2, 3]) row.getCell(c).numFmt = '0.0%';
    });
  }
  await wb.xlsx.writeFile(join(dir, 'sample-data.xlsx'));
  console.log(`built template-packs/${p.name}`);
}

export const PACK_NAMES = PACKS.map((p) => p.name);

if (process.argv[1] && fileURLToPath(import.meta.url) === resolve(process.argv[1])) {
  for (const p of PACKS) await buildPack(p);
}
