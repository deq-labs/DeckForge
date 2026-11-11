import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import ExcelJS from 'exceljs';
import { Template } from '@deckforge/pptx-engine';
import { createWorkbook, headerToTag, readWorkbook } from '../src/index.ts';

const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(resolve(import.meta.dirname, '../../../fixtures/templates', name)));

async function workbook(build: (wb: ExcelJS.Workbook) => void): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  build(wb);
  return new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);
}

describe('batch column', () => {
  it('splits rows into one deck per value, sharing rows with an empty batch', async () => {
    const manifest = (await Template.load(fixture('ops-review.pptx'))).slots();
    const bytes = await workbook((wb) => {
      const k = wb.addWorksheet('kpis');
      k.addRow(['key', 'value', 'batch']);
      k.addRow(['title', 'Monthly review', null]);
      k.addRow(['site', 'Plant A', 'A']);
      k.addRow(['site', 'Plant B', 'B']);
      k.addRow(['title', 'Plant B special', 'B']); // overrides the shared title for B
      const t = wb.addWorksheet('downtime');
      t.addRow(['Batch', 'Line', 'Cause', 'Minutes']);
      t.addRow(['A', 'L1', 'Jam', 5]);
      t.addRow(['B', 'L9', 'Setup', 7]);
      t.addRow([null, 'All', 'Shared', 1]);
      const c = wb.addWorksheet('oee_trend');
      c.addRow(['Month', 'OEE', 'batch']);
      c.addRow(['Jan', 0.7, 'A']);
      c.addRow(['Jan', 0.6, 'B']);
      c.addRow(['Feb', 0.8, 'B']);
    });
    const { decks, issues } = await readWorkbook(bytes, manifest);
    expect(issues).toEqual([]);
    expect(decks.map((d) => d.batch)).toEqual(['A', 'B']);
    const [a, b] = decks;
    expect(a!.data.values).toEqual({ title: { value: 'Monthly review' }, site: { value: 'Plant A' } });
    expect(b!.data.values).toEqual({ title: { value: 'Plant B special' }, site: { value: 'Plant B' } });
    expect(a!.data.tables!.downtime).toEqual({ columns: ['Line', 'Cause', 'Minutes'], rows: [['L1', 'Jam', 5], ['All', 'Shared', 1]] });
    expect(b!.data.charts!.oee_trend).toEqual({ categoryLabel: 'Month', categories: ['Jan', 'Feb'], series: [{ name: 'OEE', values: [0.6, 0.8] }] });
  });

  it('flags a key repeated within the same batch', async () => {
    const manifest = (await Template.load(fixture('ops-review.pptx'))).slots();
    const bytes = await workbook((wb) => {
      const k = wb.addWorksheet('kpis');
      k.addRow(['key', 'value', 'batch']);
      k.addRow(['site', 'x', 'A']);
      k.addRow(['site', 'y', 'A']);
    });
    const { issues } = await readWorkbook(bytes, manifest);
    expect(issues.map((i) => i.message)).toEqual([`"site" appears twice for batch "A" (sheet 'kpis' row 2 and sheet 'kpis' row 3). Keep only one row for it.`]);
  });
});

describe('repeat sheets', () => {
  it('maps headers to tags, merges "<tag> previous" and keeps formats', async () => {
    expect(headerToTag(' Plant Manager ')).toBe('plant_manager');
    expect(headerToTag('OEE %')).toBe('oee');
    const manifest = (await Template.load(fixture('l2-features.pptx'))).slots();
    const bytes = await workbook((wb) => {
      const s = wb.addWorksheet('Sites');
      s.addRow(['Site', 'OEE', 'OEE previous', 'Manager', '2025']);
      const r = s.addRow(['Plant A', 0.8, 0.7, 'Ann', 1]);
      r.getCell(2).numFmt = '0%';
      s.addRow([]);
      s.addRow(['Plant B', 0.6, null, 'Bo', 2]);
    });
    const { decks, issues } = await readWorkbook(bytes, manifest);
    expect(issues.map((i) => i.message)).toEqual([
      `Sheet 'Sites' column "2025" can't be used as a tag name, so it is ignored. Use letters, digits and spaces, starting with a letter.`,
    ]);
    expect(decks[0]!.data.rows).toEqual({
      sites: [
        { site: { value: 'Plant A' }, oee: { value: 0.8, format: 'percent:0', previous: 0.7 }, manager: { value: 'Ann' } },
        { site: { value: 'Plant B' }, oee: { value: 0.6 }, manager: { value: 'Bo' } },
      ],
    });
  });

  it('createWorkbook adds a sheet per repeat with its tag columns', async () => {
    const manifest = (await Template.load(fixture('l2-features.pptx'))).slots();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load((await createWorkbook(manifest)).buffer as ArrayBuffer);
    expect(wb.worksheets.map((w) => w.name)).toEqual(['kpis', 'site_downtime', 'scrap_share', 'volume_yield', 'sites', '_readme']);
    expect((wb.getWorksheet('sites')!.getRow(1).values as unknown[]).slice(1)).toEqual(['site', 'oee', 'oee previous', 'manager']);
  });
});
