import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import ExcelJS from 'exceljs';
import { Template } from '@deckforge/pptx-engine';
import { check, run, safeFileName } from '../src/index.ts';

const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(resolve(import.meta.dirname, '../../../fixtures/templates', name)));

async function workbook(build: (wb: ExcelJS.Workbook) => void): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  build(wb);
  return new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);
}

/** A complete workbook for l2-features, with a batch column on the kpis and sites sheets. */
async function l2Workbook(opts: { batchB?: boolean; omitManagerForB?: boolean } = {}): Promise<Uint8Array> {
  return workbook((wb) => {
    const k = wb.addWorksheet('kpis');
    k.addRow(['key', 'value', 'batch']);
    for (const [key, v] of [['title', 'Review'], ['presenter', 'Dana'], ['company', 'Acme'], ['incidents', 0], ['skip_appendix', 'no']] as const) k.addRow([key, v, null]);
    const s = wb.addWorksheet('sites');
    s.addRow(['Site', 'OEE', 'OEE previous', 'Manager', 'batch']);
    s.addRow(['Plant A', 0.8, 0.7, 'Ann', 'East/North']);
    if (opts.batchB !== false) s.addRow(['Plant B', 0.6, 0.7, opts.omitManagerForB ? null : 'Bo', 'West']);
    for (const name of ['site_downtime', 'scrap_share']) {
      const c = wb.addWorksheet(name);
      c.addRow(['Cause', 'Hours']);
      c.addRow(['Setup', 2]);
    }
    const v = wb.addWorksheet('volume_yield');
    v.addRow(['Month', 'Volume', 'Yield']);
    v.addRow(['Jul', 1, 0.9]);
  });
}

describe('run', () => {
  it('produces one deck per batch with safe file names', async () => {
    const tpl = await Template.load(fixture('l2-features.pptx'));
    const result = await run(tpl, await l2Workbook(), { templateName: 'l2' });
    expect(result.ok).toBe(true);
    expect(result.batches).toEqual(['East/North', 'West']);
    expect(result.decks.map((d) => [d.batch, d.fileName])).toEqual([
      ['East/North', 'l2-East-North.pptx'],
      ['West', 'l2-West.pptx'],
    ]);
    const east = (await Template.load(result.decks[0]!.bytes)).slots();
    expect(east.slideCount).toBe(4); // title, Plant A, charts, appendix (incidents = 0 hides safety)
  });

  it('is all or nothing, and tags each problem with its batch and a hint', async () => {
    const tpl = await Template.load(fixture('l2-features.pptx'));
    const result = await run(tpl, await l2Workbook({ omitManagerForB: true }), { templateName: 'l2' });
    expect(result.ok).toBe(false);
    expect(result.decks).toEqual([]);
    const err = result.problems.find((p) => p.severity === 'error')!;
    expect(err).toMatchObject({
      batch: 'West',
      message: '[West] No data for {{manager}} (slide 2).',
      hint: 'Add a column "manager" to the "sites" sheet.',
    });
  });

  it('names a single deck {template}.filled.pptx', async () => {
    const tpl = await Template.load(fixture('l2-features.pptx'));
    const bytes = await workbook((wb) => {
      const k = wb.addWorksheet('kpis');
      k.addRow(['key', 'value']);
      for (const [key, v] of [['title', 'T'], ['presenter', 'P'], ['company', 'C'], ['incidents', 1], ['skip_appendix', 'yes']] as const) k.addRow([key, v]);
      wb.addWorksheet('sites').addRow(['site', 'oee', 'manager']);
      for (const name of ['site_downtime', 'scrap_share']) wb.addWorksheet(name).addRow(['x', 'y']);
      wb.addWorksheet('volume_yield').addRow(['x', 'a', 'b']);
    });
    const result = await run(tpl, bytes, { templateName: 'l2' });
    expect(result.problems.filter((p) => p.severity === 'error')).toEqual([]);
    expect(result.decks.map((d) => d.fileName)).toEqual(['l2.filled.pptx']);
  });
});

describe('check', () => {
  it('reports template problems with no workbook', async () => {
    const result = await check(await Template.load(fixture('l2-features.pptx')), undefined, { templateName: 'l2' });
    expect(result.ok).toBe(true);
    expect(result.deckCount).toBe(0);
    expect(result.problems.map((p) => p.code)).toEqual(['slide-rule', 'misplaced-tag']);
  });

  it('dry-runs every batch', async () => {
    const result = await check(await Template.load(fixture('l2-features.pptx')), await l2Workbook(), { templateName: 'l2' });
    expect(result).toMatchObject({ ok: true, deckCount: 2, batches: ['East/North', 'West'] });
    expect('decks' in result).toBe(false);
  });
});

describe('safeFileName', () => {
  it('replaces characters Windows forbids', () => {
    expect(safeFileName('a/b\\c:d*e?"f<g>h|i')).toBe('a-b-c-d-e--f-g-h-i');
    expect(safeFileName('trailing. ')).toBe('trailing');
    expect(safeFileName('  ')).toBe('deck');
  });
});
