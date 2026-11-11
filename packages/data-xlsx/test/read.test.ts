import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import ExcelJS from 'exceljs';
import { Template } from '@deckforge/pptx-engine';
import type { Manifest } from '@deckforge/pptx-engine';
import { createWorkbook, formatFromNumFmt, readWorkbook } from '../src/index.ts';

const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(resolve(import.meta.dirname, '../../../fixtures/templates', name)));

async function opsManifest(): Promise<Manifest> {
  return (await Template.load(fixture('ops-review.pptx'))).slots();
}

async function workbook(build: (wb: ExcelJS.Workbook) => void): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  build(wb);
  return new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);
}

describe('formatFromNumFmt', () => {
  it.each([
    ['General', undefined],
    ['@', undefined],
    ['0', 'number:0'],
    ['#,##0.00', 'number:2'],
    ['0%', 'percent:0'],
    ['0.0%', 'percent:1'],
    ['"$"#,##0', 'currency:0'],
    ['[$€-407] #,##0.00', 'currency:2'],
    ['$#,##0.00_);[Red]($#,##0.00)', 'currency:2'],
    ['mm/dd/yyyy', 'date'],
    ['d-mmm-yy', 'date'],
    [undefined, undefined],
  ])('%s → %s', (numFmt, expected) => {
    expect(formatFromNumFmt(numFmt)).toBe(expected);
  });
});

describe('readWorkbook', () => {
  it('reads kpis, meta, tables and charts, carrying Excel number formats', async () => {
    const bytes = await workbook((wb) => {
      const k = wb.addWorksheet('KPIs'); // sheet names match case-insensitively
      k.addRow(['Key', 'Value', 'Previous', 'Format']);
      k.addRow(['oee', 0.853, 0.81, null]).getCell(2).numFmt = '0.0%';
      k.addRow(['downtime_total', 1234.5, null, 'number:0']);
      k.addRow(['formula', { formula: 'A1', result: 42 }, null, null]);
      k.addRow([null, 'ignored: no key', null, null]);
      const m = wb.addWorksheet('meta');
      m.addRow(['key', 'value']);
      m.addRow(['title', 'October']);
      const t = wb.addWorksheet('downtime');
      t.addRow(['Line', 'Cause', 'Minutes', 'Share']);
      const r = t.addRow(['Line 1', { richText: [{ text: 'Change' }, { text: 'over' }] }, 95, 0.25]);
      r.getCell(4).numFmt = '0%';
      t.addRow([]);
      t.addRow(['Line 2', 'Jam', 40, 0.1]);
      const c = wb.addWorksheet('oee_trend');
      c.addRow(['Month', 'OEE', 'Target']);
      c.addRow(['Jan', 0.71, 0.75]);
      c.addRow(['Feb', null, '0.75']);
      wb.addWorksheet('_notes').addRow(['ignored']);
    });

    const { decks, issues } = await readWorkbook(bytes, await opsManifest());
    expect(decks).toHaveLength(1);
    const data = decks[0]!.data;
    expect(data.values).toEqual({
      oee: { value: 0.853, previous: 0.81, format: 'percent:1' },
      downtime_total: { value: 1234.5, format: 'number:0' },
      formula: { value: 42 },
      title: { value: 'October' },
    });
    expect(data.tables).toEqual({
      downtime: {
        columns: ['Line', 'Cause', 'Minutes', 'Share'],
        rows: [
          ['Line 1', 'Changeover', 95, { value: 0.25, format: 'percent:0' }],
          ['Line 2', 'Jam', 40, 0.1],
        ],
      },
    });
    expect(data.charts).toEqual({
      oee_trend: {
        categoryLabel: 'Month',
        categories: ['Jan', 'Feb'],
        series: [
          { name: 'OEE', values: [0.71, null] },
          { name: 'Target', values: [0.75, 0.75] },
        ],
      },
    });
    expect(issues).toEqual([]);
  });

  it('explains layout problems with sheet and cell references', async () => {
    const bytes = await workbook((wb) => {
      const k = wb.addWorksheet('kpis');
      k.addRow(['key', 'value', 'previous', 'format']);
      k.addRow(['oee', 1, null, '0.0%']);
      k.addRow(['bad key', 1]);
      k.addRow(['title', { formula: 'B9' }]);
      const m = wb.addWorksheet('meta');
      m.addRow(['key', 'value']);
      m.addRow(['OEE', 'dup']);
      const c = wb.addWorksheet('oee_trend');
      c.addRow(['Month', 'OEE']);
      c.addRow(['Jan', 'n/a']);
      wb.addWorksheet('scrap').addRow(['x']);
    });
    const { issues } = await readWorkbook(bytes, await opsManifest());
    expect(issues.map((i) => `${i.severity}: ${i.message}`)).toEqual([
      expect.stringMatching(/^error: Sheet 'kpis' cell D2: "0.0%" is not a format DeckForge knows/),
      `warning: "bad key" (sheet 'kpis' row 3) is not a valid tag name, so no tag can use it. Use letters, digits, _ and - only, no spaces.`,
      `error: Sheet 'kpis' cell B4 has a formula with no saved result. Open the workbook in Excel, save it, and try again.`,
      `error: "OEE" appears twice (sheet 'kpis' row 2 and sheet 'meta' row 2). Keep only one row for it.`,
      `error: Sheet 'oee_trend' cell B2 should be a number but contains "n/a".`,
      `warning: Sheet 'scrap' does not match any table, chart or repeated slide in the template, so it is ignored. Sheets this template uses: downtime, downtime_by_line, oee_trend.`,
    ]);
  });

  it('rejects files that are not workbooks', async () => {
    await expect(readWorkbook(new TextEncoder().encode('nope'), await opsManifest())).rejects.toThrow(/not an Excel workbook/);
  });
});

describe('createWorkbook → readWorkbook → fill', () => {
  it('produces a workbook that, once the kpis are filled in, fills the template', async () => {
    const manifest = await opsManifest();
    const created = await createWorkbook(manifest);

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(created.buffer as ArrayBuffer);
    expect(wb.worksheets.map((w) => w.name)).toEqual(['kpis', 'downtime', 'downtime_by_line', 'oee_trend', '_readme']);
    const kpis = wb.getWorksheet('kpis')!;
    const keys: string[] = [];
    kpis.eachRow((row, n) => {
      if (n > 1) keys.push(String(row.getCell(1).value));
    });
    expect(keys).toEqual(['title', 'site', 'period_end', 'downtime_total']);
    kpis.eachRow((row, n) => {
      if (n > 1) row.getCell(2).value = `v${n}`;
    });
    const filled = new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);

    const { decks, issues, blankKeys } = await readWorkbook(filled, manifest);
    const data = decks[0]!.data;
    expect(blankKeys).toEqual([]);
    expect(issues).toEqual([]);
    expect(data.tables!.downtime!.rows).toEqual([
      ['Line 1', 'Changeover', 120],
      ['Line 2', 'Jam', 45],
    ]);
    const deck = await (await Template.load(fixture('ops-review.pptx'))).fill(data);
    expect(deck.issues).toEqual([]);
  });
});
