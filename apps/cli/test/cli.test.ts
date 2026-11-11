import { copyFileSync, existsSync, mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import ExcelJS from 'exceljs';
import { Template } from '@deckforge/pptx-engine';
import { batch, check, fill, initData, slots } from '../src/commands.ts';
import { UserError } from '../src/files.ts';
import type { Io } from '../src/report.ts';

const FIXTURE = resolve(import.meta.dirname, '../../../fixtures/templates/ops-review.pptx');

let dir: string;
let io: Io & { lines: string[] };

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'deckforge-cli-'));
  copyFileSync(FIXTURE, join(dir, 'ops.pptx'));
  const lines: string[] = [];
  io = { lines, out: (l) => lines.push(l), err: (l) => lines.push(l) };
});
afterEach(() => rmSync(dir, { recursive: true, force: true }));

const p = (name: string): string => join(dir, name);
// eslint-disable-next-line no-control-regex
const text = (): string => io.lines.join('\n').replace(/\x1b\[[0-9;]*m/g, '');

async function setKpis(file: string, values: Record<string, unknown>): Promise<void> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(file);
  wb.getWorksheet('kpis')!.eachRow((row, n) => {
    const key = String(row.getCell(1).value);
    if (n > 1 && key in values) row.getCell(2).value = values[key] as ExcelJS.CellValue;
  });
  await wb.xlsx.writeFile(file);
}

describe('deckforge CLI', () => {
  it('slots lists what the template can be filled with', async () => {
    expect(await slots(p('ops.pptx'), {}, io)).toBe(0);
    expect(text()).toContain('{{downtime_total}}  (slide 2)');
    expect(text()).toContain('downtime  (slide 2): Line | Cause | Minutes');
    expect(text()).toContain('oee_trend  (slide 4, line): 2 series — OEE, Target');
  });

  it('init-data, fill in the kpis, fill: produces a deck', async () => {
    expect(await initData(p('ops.pptx'), {}, io)).toBe(0);
    expect(existsSync(p('ops.data.xlsx'))).toBe(true);

    // Unfilled kpis are reported, and no deck is written.
    expect(await fill(p('ops.pptx'), p('ops.data.xlsx'), {}, io)).toBe(1);
    expect(text()).toContain('Fill in the value for "title" in the kpis sheet');
    expect(existsSync(p('ops.filled.pptx'))).toBe(false);

    await setKpis(p('ops.data.xlsx'), { title: 'Ops Review', site: 'Plant A', period_end: 'Oct 2026', downtime_total: 165 });
    expect(await fill(p('ops.pptx'), p('ops.data.xlsx'), {}, io)).toBe(0);
    const deck = await Template.load(new Uint8Array(readFileSync(p('ops.filled.pptx'))));
    expect(deck.slots().text).toEqual([]); // every tag filled
    expect(deck.slots().tables[0]!.sampleRows).toEqual([
      ['Line 1', 'Changeover', '120'],
      ['Line 2', 'Jam', '45'],
    ]);
  });

  it('--missing keep writes the deck and leaves tags visible', async () => {
    await initData(p('ops.pptx'), { output: p('d.xlsx') }, io);
    expect(await fill(p('ops.pptx'), p('d.xlsx'), { output: p('out.pptx'), missing: 'keep' }, io)).toBe(0);
    const deck = await Template.load(new Uint8Array(readFileSync(p('out.pptx'))));
    expect(deck.slots().text.map((t) => t.name)).toEqual(['title', 'site', 'period_end', 'downtime_total']);
  });

  it('protects existing files and the template', async () => {
    await initData(p('ops.pptx'), {}, io);
    await expect(initData(p('ops.pptx'), {}, io)).rejects.toThrow(UserError);
    await expect(fill(p('ops.pptx'), p('ops.data.xlsx'), { output: p('OPS.pptx') }, io)).rejects.toThrow(/overwrite the template/);
  });

  it('explains missing and wrong input files', async () => {
    await expect(fill(p('nope.pptx'), p('x.xlsx'), {}, io)).rejects.toThrow(/Cannot find the template/);
    await expect(fill(p('ops.pptx'), p('ops.pptx'), { output: p('o.pptx') }, io)).rejects.toThrow(/not an Excel workbook/);
  });
});

describe('deckforge batch / check', () => {
  async function batchWorkbook(file: string): Promise<void> {
    await initData(p('ops.pptx'), { output: file }, io);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(file);
    const k = wb.getWorksheet('kpis')!;
    k.getCell(1, 6).value = 'batch';
    k.eachRow((row, n) => {
      if (n > 1) row.getCell(2).value = `v${n}`;
    });
    k.addRow(['site', 'Plant B', null, null, null, 'B']);
    k.getRow(3).getCell(6).value = 'A'; // the existing "site" row belongs to batch A
    await wb.xlsx.writeFile(file);
  }

  it('writes one deck per batch value', async () => {
    await batchWorkbook(p('b.xlsx'));
    expect(await batch(p('ops.pptx'), p('b.xlsx'), { outDir: p('out') }, io)).toBe(0);
    expect(existsSync(p('out/ops-A.pptx'))).toBe(true);
    expect(existsSync(p('out/ops-B.pptx'))).toBe(true);
    expect(text()).toContain('Wrote 2 decks');
  });

  it('fill refuses a batched workbook and points to batch', async () => {
    await batchWorkbook(p('b.xlsx'));
    await expect(fill(p('ops.pptx'), p('b.xlsx'), {}, io)).rejects.toThrow(/deckforge batch/);
  });

  it('check reports readiness without writing files', async () => {
    expect(await check(p('ops.pptx'), undefined, {}, io)).toBe(0);
    expect(text()).toContain('ops.pptx: ready (0 errors, 0 warnings).');
    await batchWorkbook(p('b.xlsx'));
    expect(await check(p('ops.pptx'), p('b.xlsx'), {}, io)).toBe(0);
    expect(text()).toContain('ready (0 errors, 0 warnings; 2 decks would be made)');
    expect(existsSync(p('ops-decks'))).toBe(false);
  });
});
